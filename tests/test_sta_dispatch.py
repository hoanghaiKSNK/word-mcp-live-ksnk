import asyncio
import inspect
import threading
import time
from types import SimpleNamespace

from word_document_server.core import sta_dispatch


def test_serialized_worker_does_not_block_loop_and_cancels_queued_edit():
    events = []; started = threading.Event(); release = threading.Event()
    worker = sta_dispatch.STAWorker(lambda: events.append(('init', threading.get_ident())),
                                    lambda: events.append(('uninit', threading.get_ident())))

    async def operation(value):
        events.append((value, threading.get_ident()))
        if value == 'first':
            started.set(); assert release.wait(3)
        return value

    async def run():
        first = asyncio.create_task(worker.call(operation, 'first'))
        while not started.is_set(): await asyncio.sleep(0.005)
        cancelled = asyncio.create_task(worker.call(operation, 'cancelled'))
        await asyncio.sleep(0.02)
        cancelled.cancel()
        try: await cancelled
        except asyncio.CancelledError: pass
        second = asyncio.create_task(worker.call(operation, 'second'))
        # The event loop remains available while a COM-like operation blocks.
        await asyncio.sleep(0.03)
        release.set()
        assert await first == 'first'
        assert await second == 'second'
    try:
        asyncio.run(run())
    finally:
        release.set(); worker.close()
    assert [e[0] for e in events] == ['init', 'first', 'second', 'uninit']
    assert len({e[1] for e in events}) == 1


def test_nested_dispatch_and_schema_are_preserved(monkeypatch):
    worker = sta_dispatch.STAWorker(); monkeypatch.setattr(sta_dispatch, 'worker', worker)
    monkeypatch.setattr(sta_dispatch.sys, 'platform', 'win32')
    async def word_live_read(filename: str = None, page: int = 1) -> str:
        """Synthetic schema."""
        return str(page)
    module = SimpleNamespace(word_live_read=word_live_read)
    sta_dispatch.install_sta_dispatch(module)
    assert inspect.signature(module.word_live_read) == inspect.signature(word_live_read)
    assert module.word_live_read.__doc__ == word_live_read.__doc__
    async def outer(): return await module.word_live_read(page=2)
    try: assert asyncio.run(worker.call(outer)) == '2'
    finally: worker.close()
