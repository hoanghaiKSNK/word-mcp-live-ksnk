"""Run independent closed-file jobs in a bounded pool; live Word uses its STA."""
import asyncio
import atexit
from collections import deque
from concurrent.futures import ThreadPoolExecutor
from contextlib import ExitStack
import functools
import inspect
import json
import os
import threading
import time
import uuid

from word_document_server.core.coordination import DispatchError, ProcessLock
from word_document_server.core.path_policy import input_paths, roots


class OfflineWorker:
    def __init__(self, workers=4, capacity=32, queue_wait_seconds=30):
        if workers < 1 or capacity < workers or queue_wait_seconds <= 0:
            raise ValueError('Invalid offline pool limits')
        self.executor = ThreadPoolExecutor(max_workers=workers, thread_name_prefix='word-mcp-file')
        self.workers, self.capacity = workers, capacity
        self.queue_wait_seconds = queue_wait_seconds
        self.slots = threading.BoundedSemaphore(capacity)
        self.lock = threading.Lock()
        self.pending = self.rejected = 0
        self.active = {}
        self.recent = deque(maxlen=100)

    def status(self):
        with self.lock:
            active = []
            for record in self.active.values():
                item = dict(record)
                item['elapsed_seconds'] = round(time.monotonic() - item.pop('_start'), 4)
                active.append(item)
            return {'workers': self.workers, 'capacity': self.capacity,
                    'pending_including_running': self.pending,
                    'queue_depth': max(0, self.pending - len(active)),
                    'running': active, 'rejected': self.rejected,
                    'recent': [dict(item) for item in self.recent]}

    def _execute(self, function, args, kwargs, paths, record):
        start = time.monotonic()
        queued = record.pop('_queued')
        record['queue_wait_seconds'] = round(start - queued, 4)
        try:
            if start - queued >= self.queue_wait_seconds:
                raise DispatchError('QUEUE_WAIT_EXPIRED', 'File job expired before starting')
            with ExitStack() as stack:
                for path in paths:
                    stack.enter_context(ProcessLock('file:' + path, busy_code='FILE_BUSY'))
                record['operation_started'] = True
                with self.lock:
                    self.active[record['request_id']] = {'request_id': record['request_id'],
                                                        'tool': record['tool'], '_start': start}
                result = function(*args, **kwargs)
                if inspect.isawaitable(result):
                    result = asyncio.run(result)
                record['outcome'] = 'returned'
                return result
        except DispatchError as error:
            record['outcome'] = error.code
            error.request_id = record['request_id']
            raise
        except BaseException:
            record['outcome'] = 'tool_exception'
            raise
        finally:
            record['execution_seconds'] = round(time.monotonic() - start, 4)
            with self.lock:
                self.active.pop(record['request_id'], None)
                if not record['operation_started']:
                    self.rejected += 1
                self.recent.append(record)

    async def call(self, function, args, kwargs, paths):
        if not self.slots.acquire(blocking=False):
            with self.lock:
                self.rejected += 1
            raise DispatchError('QUEUE_FULL', 'Closed-file pool is full; no operation started')
        record = {'request_id': uuid.uuid4().hex, 'tool': function.__name__,
                  '_queued': time.monotonic(), 'operation_started': False}
        with self.lock:
            self.pending += 1
        try:
            future = self.executor.submit(self._execute, function, args, kwargs, paths, record)
        except BaseException:
            with self.lock:
                self.pending -= 1
            self.slots.release()
            raise
        def completed(done):
            with self.lock:
                self.pending -= 1
                if done.cancelled():
                    record.pop('_queued', None)
                    record['outcome'] = 'cancelled_before_start'
                    self.recent.append(record)
            self.slots.release()
        future.add_done_callback(completed)
        return await asyncio.wrap_future(future)

    def close(self):
        self.executor.shutdown(wait=False, cancel_futures=True)


offline = OfflineWorker(workers=int(os.getenv('WORD_MCP_OFFLINE_WORKERS', '4')))
atexit.register(offline.close)


def dispatch_tool(function):
    """Keep the original schema while moving blocking file work off the MCP loop."""
    signature = inspect.signature(function)
    live = function.__name__.startswith('word_live_') or function.__name__ == 'word_screen_capture'
    @functools.wraps(function)
    async def dispatched(*args, **kwargs):
        try:
            arguments = signature.bind(*args, **kwargs)
            arguments.apply_defaults()
            paths = input_paths(arguments.arguments, live=live, tool_name=function.__name__)
            if live:
                result = function(*args, **kwargs)
                return await result if inspect.isawaitable(result) else result
            return await offline.call(function, args, kwargs, paths)
        except DispatchError as error:
            return json.dumps(error.as_dict())
    return dispatched


def diagnostics():
    from importlib.metadata import version
    from word_document_server.core.sta_dispatch import worker
    return {'version': version('word-mcp-live'), 'live': worker.status(),
            'offline': offline.status(), 'path_policy_enabled': bool(roots()),
            'legacy_clients_coordinate': False,
            'note': 'Restart old MCP sessions to use this version. A started operation may finish after client cancellation; never blindly retry an edit.'}
