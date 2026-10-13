"""Serialize live Word operations on one COM STA without blocking MCP's loop.

Queued cancellation prevents the operation from starting. Once Word has started
an operation it cannot safely be interrupted: it finishes before the next one.
Mutations are never retried automatically. Only JSON/primitive results cross
the thread boundary, never COM proxies.
"""
import asyncio
import atexit
import concurrent.futures
import functools
import inspect
import queue
import sys
import threading


class STAWorker:
    def __init__(self, initialize=None, uninitialize=None, capacity=32):
        self.queue = queue.Queue(capacity)
        self.initialize = initialize
        self.uninitialize = uninitialize
        self.closed = threading.Event()
        self.lock = threading.Lock()
        self.thread = None

    def _run(self):
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)
        initialized = False
        init_error = None
        try:
            try:
                if self.initialize:
                    self.initialize()
                initialized = True
            except Exception as error:
                init_error = error
            while not self.closed.is_set():
                try:
                    future, function, args, kwargs = self.queue.get(timeout=0.1)
                except queue.Empty:
                    continue
                try:
                    if not future.set_running_or_notify_cancel():
                        continue
                    if init_error is not None:
                        future.set_exception(init_error)
                        continue
                    try:
                        result = function(*args, **kwargs)
                        if inspect.isawaitable(result):
                            result = loop.run_until_complete(result)
                        future.set_result(result)
                    except BaseException as error:
                        future.set_exception(error)
                finally:
                    self.queue.task_done()
        finally:
            while True:
                try:
                    future, *_ = self.queue.get_nowait()
                except queue.Empty:
                    break
                future.cancel()
                self.queue.task_done()
            loop.close()
            if initialized and self.uninitialize:
                self.uninitialize()

    async def call(self, function, *args, **kwargs):
        if threading.current_thread() is self.thread:
            return await function(*args, **kwargs)
        future = concurrent.futures.Future()
        with self.lock:
            if self.closed.is_set():
                raise RuntimeError('Word STA worker is closed')
            if self.thread is None:
                self.thread = threading.Thread(target=self._run, name='word-mcp-STA', daemon=True)
                self.thread.start()
            try:
                self.queue.put_nowait((future, function, args, kwargs))
            except queue.Full:
                raise RuntimeError('Word queue is full; no operation was started')
        return await asyncio.wrap_future(future)

    def close(self):
        with self.lock:
            self.closed.set()
        if self.thread and threading.current_thread() is not self.thread:
            self.thread.join(timeout=2)


def _initialize_com():
    import pythoncom
    pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)


def _uninitialize_com():
    import pythoncom
    pythoncom.CoUninitialize()


worker = STAWorker(_initialize_com, _uninitialize_com)
atexit.register(worker.close)


def install_sta_dispatch(*modules):
    """Keep MCP tool schemas/docstrings unchanged while dispatching Windows COM."""
    if sys.platform != 'win32':
        return
    for module in modules:
        for name, function in list(vars(module).items()):
            if not name.startswith('word_live_') or not inspect.iscoroutinefunction(function):
                continue
            if getattr(function, '_sta_dispatched', False):
                continue

            def wrap(original):
                @functools.wraps(original)
                async def dispatched(*args, **kwargs):
                    return await worker.call(original, *args, **kwargs)
                dispatched._sta_dispatched = True
                return dispatched

            setattr(module, name, wrap(function))
