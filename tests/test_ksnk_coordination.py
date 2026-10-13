"""Concurrency tests use synthetic jobs and OS locks, never user Word documents."""
import asyncio
from contextlib import nullcontext
import json
import os
from pathlib import Path
import subprocess
import sys
import threading
import time
import uuid
import hashlib

import pytest
from word_document_server.core.coordination import DispatchError, ProcessLock
from word_document_server.core.sta_dispatch import STAWorker
from word_document_server.core.tool_dispatch import OfflineWorker, dispatch_tool
from word_document_server.core.path_policy import check_path, input_paths


def test_another_process_cannot_enter_and_can_enter_after_release():
    key = 'ksnk-test-' + uuid.uuid4().hex
    code = '''
import json, sys
from word_document_server.core.coordination import ProcessLock, DispatchError
try:
    with ProcessLock(sys.argv[1]): print('acquired')
except DispatchError as e: print(e.code)
'''
    def child():
        return subprocess.run([sys.executable, '-c', code, key], capture_output=True,
                              text=True, timeout=15, check=True).stdout.strip()
    with ProcessLock(key):
        assert child() == 'APP_BUSY'
    assert child() == 'acquired'


@pytest.mark.skipif(sys.platform != 'win32', reason='Windows abandoned mutex semantics')
def test_dead_owner_reports_uncertain_state_before_next_edit():
    import win32event
    import win32api
    key = 'ksnk-abandoned-' + uuid.uuid4().hex
    name = 'Local\\word-mcp-ksnk-v2-' + hashlib.sha256(key.encode()).hexdigest()
    # Keep the kernel object alive after the child dies to exercise abandonment.
    observer = win32event.CreateMutex(None, False, name)
    code = '''
import os, sys
from word_document_server.core.coordination import ProcessLock
with ProcessLock(sys.argv[1]): os._exit(23)
'''
    try:
        child = subprocess.run([sys.executable, '-c', code, key], capture_output=True, timeout=15)
        assert child.returncode == 23, child.stderr
        with pytest.raises(DispatchError) as error:
            with ProcessLock(key): pytest.fail('Must not start after an abandoned operation')
        assert error.value.code == 'APP_STATE_UNCERTAIN'
        # After the refused call, an explicit following request can proceed.
        with ProcessLock(key): pass
    finally: win32api.CloseHandle(observer)


def test_failed_tool_releases_live_lock_and_worker_recovers():
    key = 'ksnk-test-' + uuid.uuid4().hex
    worker = STAWorker(guard=lambda: ProcessLock(key))
    async def failure(): raise ValueError('synthetic failure')
    async def following(): return 'recovered'
    async def run():
        with pytest.raises(ValueError): await worker.call(failure)
        assert await worker.call(following) == 'recovered'
    try: asyncio.run(run())
    finally: worker.close()
    assert worker.status()['recent'][0]['outcome'] == 'tool_exception'
    assert worker.status()['running'] is None


def test_waiting_edit_expires_without_touching_document():
    started, release = threading.Event(), threading.Event()
    writes = []
    worker = STAWorker(queue_wait_seconds=0.04)
    async def first(): started.set(); assert release.wait(3)
    async def edit(): writes.append('must not run')
    async def run():
        task = asyncio.create_task(worker.call(first))
        while not started.is_set(): await asyncio.sleep(.005)
        queued = asyncio.create_task(worker.call(edit))
        await asyncio.sleep(.08)
        assert worker.status()['queue_depth'] == 1
        release.set(); await task
        with pytest.raises(DispatchError, match='expired') as error:
            await queued
        assert error.value.code == 'QUEUE_WAIT_EXPIRED'
    try: asyncio.run(run())
    finally: release.set(); worker.close()
    assert writes == []
    assert worker.status()['recent'][-1]['operation_started'] is False


def test_running_cancel_does_not_repeat_or_release_lock_early():
    started, release, finished = threading.Event(), threading.Event(), threading.Event()
    writes = []
    key = 'ksnk-test-' + uuid.uuid4().hex
    worker = STAWorker(guard=lambda: ProcessLock(key))
    async def edit(secret):
        started.set(); assert release.wait(3)
        writes.append(secret); finished.set()
    async def run():
        task = asyncio.create_task(worker.call(edit, 'synthetic confidential text'))
        while not started.is_set(): await asyncio.sleep(.005)
        task.cancel()
        with pytest.raises(asyncio.CancelledError): await task
        # Main thread is a different mutex owner, even in the same process.
        with pytest.raises(DispatchError):
            with ProcessLock(key): pass
        assert worker.status()['running']['tool'] == 'edit'
        release.set()
        while not finished.is_set(): await asyncio.sleep(.005)
    try: asyncio.run(run())
    finally: release.set(); worker.close()
    assert writes == ['synthetic confidential text']
    assert 'confidential' not in json.dumps(worker.status())
    assert worker.status()['client_detachments'] == 1


def test_full_live_queue_rejects_before_start():
    started, release = threading.Event(), threading.Event()
    worker = STAWorker(capacity=1)
    async def job(block=False):
        if block: started.set(); assert release.wait(3)
    async def run():
        running = asyncio.create_task(worker.call(job, True))
        while not started.is_set(): await asyncio.sleep(.005)
        queued = asyncio.create_task(worker.call(job)); await asyncio.sleep(.02)
        with pytest.raises(DispatchError) as error: await worker.call(job)
        assert error.value.code == 'QUEUE_FULL'
        release.set(); await running; await queued
    try: asyncio.run(run())
    finally: release.set(); worker.close()


def test_dead_sta_rejects_new_work_instead_of_hanging():
    def failed_pump(): raise RuntimeError('synthetic pump failure')
    worker = STAWorker(pump=failed_pump)
    async def read(): return 'ok'
    async def run():
        assert await worker.call(read) == 'ok'
        limit = time.monotonic() + 3
        while not worker.closed.is_set() and time.monotonic() < limit:
            await asyncio.sleep(.01)
        assert worker.closed.is_set()
        with pytest.raises(DispatchError) as error: await worker.call(read)
        assert error.value.code == 'WORKER_CLOSED'
    try: asyncio.run(run())
    finally: worker.close()
    assert worker.status()['fatal_error'] == 'RuntimeError'


def test_cancelled_running_file_job_keeps_slot_until_completion():
    started, release, finished = threading.Event(), threading.Event(), threading.Event()
    writes = []
    worker = OfflineWorker(workers=1, capacity=1)
    def edit():
        started.set(); assert release.wait(3); writes.append('once'); finished.set()
    async def run():
        task = asyncio.create_task(worker.call(edit, (), {}, []))
        while not started.is_set(): await asyncio.sleep(.005)
        task.cancel()
        with pytest.raises(asyncio.CancelledError): await task
        with pytest.raises(DispatchError) as error: await worker.call(edit, (), {}, [])
        assert error.value.code == 'QUEUE_FULL'
        release.set()
        while not finished.is_set(): await asyncio.sleep(.005)
    try: asyncio.run(run())
    finally: release.set(); worker.executor.shutdown(wait=True)
    assert writes == ['once']


def test_independent_files_run_concurrently_but_same_file_is_busy(tmp_path):
    worker = OfflineWorker(workers=2)
    started, release = threading.Event(), threading.Event()
    entered = []
    def first(): started.set(); assert release.wait(3); entered.append('first')
    async def second(): entered.append('second'); return 'second'
    async def run():
        path = str(tmp_path/'a.docx')
        running = asyncio.create_task(worker.call(first, (), {}, [path]))
        while not started.is_set(): await asyncio.sleep(.005)
        with pytest.raises(DispatchError) as error:
            await worker.call(second, (), {}, [path])
        assert error.value.code == 'FILE_BUSY'
        assert await worker.call(second, (), {}, [str(tmp_path/'b.docx')]) == 'second'
        assert not release.is_set()  # genuinely overlapped the blocked first job
        release.set(); await running
    try: asyncio.run(run())
    finally: release.set(); worker.close()
    assert entered == ['second', 'first']


def test_path_policy_blocks_traversal_sibling_and_junction(tmp_path, monkeypatch):
    root = tmp_path/'allowed'; root.mkdir()
    outside = tmp_path/'allowed-sibling'; outside.mkdir()
    monkeypatch.setenv('WORD_MCP_ALLOWED_ROOTS', str(root))
    assert check_path(root/'new.docx') == root/'new.docx'
    for path in [root/'..'/'outside.docx', outside/'file.docx']:
        with pytest.raises(DispatchError): check_path(path)
    link = root/'escape'
    if os.name == 'nt':
        # Junction creation does not need Windows developer mode/symlink rights.
        result = subprocess.run(['cmd', '/c', 'mklink', '/J', str(link), str(outside)],
                                capture_output=True, timeout=10)
        assert result.returncode == 0, result.stderr
    else: link.symlink_to(outside, target_is_directory=True)
    try:
        with pytest.raises(DispatchError): check_path(link/'file.docx')
    finally:
        # Remove the link itself, never recursively remove the target directory.
        if os.name == 'nt': os.rmdir(link)
        else: link.unlink()
    assert outside.is_dir()


def test_denied_path_never_enters_tool(tmp_path, monkeypatch):
    monkeypatch.setenv('WORD_MCP_ALLOWED_ROOTS', str(tmp_path/'allowed'))
    calls = []
    def create_document(filename: str): calls.append(filename)
    result = json.loads(asyncio.run(dispatch_tool(create_document)(str(tmp_path/'outside.docx'))))
    assert result['code'] == 'PATH_NOT_ALLOWED' and result['operation_started'] is False
    assert calls == []


def test_all_source_paths_and_derived_outputs_are_checked(tmp_path, monkeypatch):
    monkeypatch.setenv('WORD_MCP_ALLOWED_ROOTS', str(tmp_path/'allowed'))
    with pytest.raises(DispatchError):
        input_paths({'source_filenames': [str(tmp_path/'allowed'/'a.docx'), str(tmp_path/'outside.docx')]})
    paths = input_paths({'source_filename': str(tmp_path/'allowed'/'a')}, tool_name='copy_document')
    assert os.path.normcase(str(tmp_path/'allowed'/'a_copy.docx')) in paths


def test_config_invalid_fails_closed_and_unset_preserves_access(tmp_path, monkeypatch):
    monkeypatch.delenv('WORD_MCP_ALLOWED_ROOTS', raising=False)
    assert check_path(tmp_path/'file.docx') == tmp_path/'file.docx'
    monkeypatch.setenv('WORD_MCP_ALLOWED_ROOTS', 'relative-path')
    with pytest.raises(DispatchError) as error: check_path(tmp_path/'file.docx')
    assert error.value.code == 'PATH_POLICY_INVALID'
