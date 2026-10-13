"""Fail-fast locks shared by cooperating MCP processes; no document data."""
import hashlib
import os
import sys
import tempfile
from pathlib import Path


class DispatchError(RuntimeError):
    def __init__(self, code, message, request_id=None):
        super().__init__(message)
        self.code, self.request_id = code, request_id

    def as_dict(self):
        return {'error': str(self), 'code': self.code, 'request_id': self.request_id,
                'operation_started': False, 'automatic_retry': False}

    def __str__(self):
        return '[' + self.code + '] ' + super().__str__()


class ProcessLock:
    """Fail fast; OS releases ownership on death. Acquire/release on one thread.

    An abandoned Windows mutex refuses the next call so Word can be inspected.
    The lock covers participating clients, not arbitrary external Word add-ins.
    """
    def __init__(self, key, busy_code='APP_BUSY'):
        self.key, self.busy_code = key, busy_code
        self.handle, self.owned = None, False

    def __enter__(self):
        digest = hashlib.sha256(self.key.encode('utf-8')).hexdigest()
        if sys.platform == 'win32':
            import win32event
            import win32api
            self.handle = win32event.CreateMutex(None, False, 'Local\\word-mcp-ksnk-v2-' + digest)
            try:
                outcome = win32event.WaitForSingleObject(self.handle, 0)
                if outcome == win32event.WAIT_TIMEOUT:
                    raise DispatchError(self.busy_code, 'Another cooperating MCP operation owns this resource; no operation started')
                if outcome not in (win32event.WAIT_OBJECT_0, win32event.WAIT_ABANDONED):
                    raise DispatchError('LOCK_FAILED', 'Could not acquire the coordination lock')
                self.owned = True
                if outcome == win32event.WAIT_ABANDONED:
                    raise DispatchError('APP_STATE_UNCERTAIN', 'A previous owner exited while working; inspect Word before submitting another edit')
            except BaseException:
                if self.owned:
                    win32event.ReleaseMutex(self.handle)
                    self.owned = False
                win32api.CloseHandle(self.handle)
                self.handle = None
                raise
        else:
            import fcntl
            folder = Path(tempfile.gettempdir()) / ('word-mcp-ksnk-locks-' + str(os.getuid()))
            folder.mkdir(mode=0o700, exist_ok=True)
            self.handle = open(folder / digest, 'a+b')
            try:
                fcntl.flock(self.handle, fcntl.LOCK_EX | fcntl.LOCK_NB)
                self.owned = True
            except BlockingIOError:
                self.handle.close(); self.handle = None
                raise DispatchError(self.busy_code, 'Another cooperating MCP operation owns this resource; no operation started') from None
        return self

    def __exit__(self, *_):
        if self.handle is None:
            return
        try:
            if self.owned:
                if sys.platform == 'win32':
                    import win32event
                    win32event.ReleaseMutex(self.handle)
                else:
                    import fcntl
                    fcntl.flock(self.handle, fcntl.LOCK_UN)
        finally:
            if sys.platform == 'win32':
                import win32api
                win32api.CloseHandle(self.handle)
            else:
                self.handle.close()
            self.handle, self.owned = None, False


def live_lock():
    # Conservative application-wide key: Selection and Undo may be shared.
    return ProcessLock('interactive-word-application')
