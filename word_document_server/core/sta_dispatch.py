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
import time
import uuid
import os
import json
import math
from collections import deque
from contextlib import nullcontext
from word_document_server.core.coordination import DispatchError, live_lock


class STAWorker:
    def __init__(self, initialize=None, uninitialize=None, capacity=32,
                 guard=None, queue_wait_seconds=30, pump=None):
        if capacity < 1 or not math.isfinite(queue_wait_seconds) or queue_wait_seconds <= 0:
            raise ValueError('Capacity and queue wait must be positive')
        self.queue = queue.Queue(capacity)
        self.initialize = initialize
        self.uninitialize = uninitialize
        self.closed = threading.Event()
        self.lock = threading.Lock()
        self.thread = None
        self.guard = guard or nullcontext
        self.queue_wait_seconds = queue_wait_seconds
        self.pump = pump
        self.state_lock = threading.Lock()
        self.running = None
        self.recent = deque(maxlen=100)
        self.submitted = self.rejected = self.detached = 0
        self.fatal_error = None

    def status(self):
        """Diagnostics never include document paths, arguments or text."""
        with self.state_lock:
            running = dict(self.running) if self.running else None
            if running:
                running['elapsed_seconds'] = round(time.monotonic() - running.pop('_start'), 4)
            return {'pid': os.getpid(), 'closed': self.closed.is_set(), 'fatal_error': self.fatal_error,
                    'queue_depth': self.queue.qsize(), 'queue_capacity': self.queue.maxsize,
                    'queue_wait_limit_seconds': self.queue_wait_seconds,
                    'submitted': self.submitted, 'rejected': self.rejected,
                    'client_detachments': self.detached, 'running': running,
                    'recent': [dict(item) for item in self.recent],
                    'automatic_retry': False,
                    'coordination_scope': 'cooperating processes in this login session'}

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
                    future, function, args, kwargs, record = self.queue.get(timeout=0.1)
                except queue.Empty:
                    if initialized and self.pump:
                        self.pump()
                    continue
                start = time.monotonic()
                record['queue_wait_seconds'] = round(start - record.pop('_queued'), 4)
                try:
                    if not future.set_running_or_notify_cancel():
                        record['outcome'] = 'cancelled_before_start'
                        continue
                    if record['queue_wait_seconds'] >= self.queue_wait_seconds:
                        raise DispatchError('QUEUE_WAIT_EXPIRED', 'Queued operation expired before starting; read current document state before submitting again')
                    if init_error is not None:
                        raise DispatchError('COM_INITIALIZATION_FAILED', 'Could not initialize the Word STA worker') from init_error
                    with self.guard():
                        record['operation_started'] = True
                        with self.state_lock:
                            self.running = {'request_id': record['request_id'], 'tool': record['tool'],
                                            '_start': time.monotonic()}
                        result = function(*args, **kwargs)
                        if inspect.isawaitable(result):
                            result = loop.run_until_complete(result)
                        record['outcome'] = 'returned'
                        if isinstance(result, str):
                            try:
                                payload = json.loads(result)
                            except (ValueError, TypeError):
                                payload = None
                            if isinstance(payload, dict):
                                if payload.get('error'):
                                    record['outcome'] = 'tool_reported_error'
                                    message = str(payload['error'])
                                    code = 'WORD_OPERATION_FAILED'
                                    if message.startswith('[') and '] ' in message:
                                        candidate = message[1:message.index(']')]
                                        if candidate in ('PATH_NOT_ALLOWED', 'PATH_POLICY_INVALID'):
                                            code = candidate
                                    if any(value in message.lower() for value in
                                           ('-2147418111', '-2147417846', '0x80010001', '0x8001010a')):
                                        code = 'WORD_BUSY_DURING_OPERATION'
                                    payload.setdefault('code', code)
                                    payload.setdefault('automatic_retry', False)
                                    # The tool body was entered; it may have made
                                    # partial changes before reporting a failure.
                                    payload.setdefault('mutation_outcome', 'inspect_document_before_retry')
                                payload['_dispatch'] = {
                                    'request_id': record['request_id'],
                                    'operation_started': True,
                                    'queue_wait_seconds': record['queue_wait_seconds'],
                                    'execution_seconds': round(time.monotonic() - start, 4)}
                                result = json.dumps(payload, ensure_ascii=False)
                        future.set_result(result)
                except DispatchError as error:
                    record['outcome'] = error.code
                    error.request_id = record['request_id']
                    future.set_exception(error)
                except BaseException as error:
                    record['outcome'] = 'tool_exception'
                    future.set_exception(error)
                finally:
                    record['execution_seconds'] = round(time.monotonic() - start, 4)
                    with self.state_lock:
                        self.running = None
                        self.recent.append(record)
                        if not record['operation_started'] and record['outcome'] != 'cancelled_before_start':
                            self.rejected += 1
                    self.queue.task_done()
        except BaseException as error:
            with self.state_lock:
                self.fatal_error = type(error).__name__
        finally:
            # A failed COM pump/worker must reject new work rather than accept
            # jobs into a queue whose consumer has already exited.
            with self.lock:
                self.closed.set()
            while True:
                try:
                    future, *_, record = self.queue.get_nowait()
                except queue.Empty:
                    break
                future.cancel()
                record.pop('_queued', None)
                record['outcome'] = 'cancelled_on_shutdown'
                with self.state_lock:
                    self.recent.append(record)
                self.queue.task_done()
            loop.close()
            if initialized and self.uninitialize:
                self.uninitialize()

    async def call(self, function, *args, **kwargs):
        if threading.current_thread() is self.thread:
            result = function(*args, **kwargs)
            return await result if inspect.isawaitable(result) else result
        future = concurrent.futures.Future()
        record = {'request_id': uuid.uuid4().hex, 'tool': function.__name__,
                  '_queued': time.monotonic(), 'operation_started': False}
        with self.lock:
            if self.closed.is_set():
                raise DispatchError('WORKER_CLOSED', 'Word STA worker is closed', record['request_id'])
            if self.thread is None:
                self.thread = threading.Thread(target=self._run, name='word-mcp-STA', daemon=True)
                self.thread.start()
            with self.state_lock:
                self.submitted += 1
            try:
                self.queue.put_nowait((future, function, args, kwargs, record))
            except queue.Full:
                with self.state_lock:
                    self.rejected += 1
                raise DispatchError('QUEUE_FULL', 'Word queue is full; no operation started', record['request_id']) from None
        try:
            return await asyncio.wrap_future(future)
        except asyncio.CancelledError:
            # A started COM operation may finish after the client disconnects.
            with self.state_lock:
                self.detached += 1
            raise

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


def _pump_com():
    import pythoncom
    pythoncom.PumpWaitingMessages()


worker = STAWorker(_initialize_com, _uninitialize_com, guard=live_lock,
                   queue_wait_seconds=float(os.getenv('WORD_MCP_QUEUE_WAIT_SECONDS', '30')),
                   pump=_pump_com)
atexit.register(worker.close)


def install_sta_dispatch(*modules):
    """Keep MCP tool schemas/docstrings unchanged while dispatching Windows COM."""
    if sys.platform != 'win32':
        return
    for module in modules:
        for name, function in list(vars(module).items()):
            if not (name.startswith('word_live_') or name == 'word_screen_capture') or not inspect.iscoroutinefunction(function):
                continue
            if getattr(function, '_sta_dispatched', False):
                continue

            def wrap(original):
                @functools.wraps(original)
                async def dispatched(*args, **kwargs):
                    try:
                        return await worker.call(original, *args, **kwargs)
                    except DispatchError as error:
                        return json.dumps(error.as_dict())
                dispatched._sta_dispatched = True
                return dispatched

            setattr(module, name, wrap(function))
