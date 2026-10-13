"""COM connection manager for Microsoft Word on Windows.

Provides functions to connect to a running Word instance and find open documents.
Only works on Windows with pywin32 installed.
"""

import os
import ntpath
import sys
import unicodedata
from contextlib import contextmanager


def get_word_app():
    """Get a reference to the running Word application via COM.

    Returns the Word.Application COM object that has open documents.
    When multiple Word instances are running, iterates through all
    Running Object Table (ROT) entries to find one with documents.
    Raises RuntimeError if Word is not running or not on Windows.
    """
    if sys.platform != "win32":
        raise RuntimeError("Word COM automation is only available on Windows")

    import win32com.client

    try:
        app = win32com.client.GetActiveObject("Word.Application")
        if app.Documents.Count > 0:
            return app
        # GetActiveObject found an empty instance — scan ROT for others
        app_with_docs = _find_word_with_docs()
        if app_with_docs is not None:
            return app_with_docs
        # No instance has documents; return the empty one (caller may open a file)
        return app
    except Exception:
        # GetActiveObject failed entirely — try ROT scan
        app_with_docs = _find_word_with_docs()
        if app_with_docs is not None:
            return app_with_docs
        raise RuntimeError(
            "Microsoft Word is not running. Please open Word first."
        )


def _find_word_with_docs():
    """Scan the Running Object Table for a Word.Application with open docs.

    Handles Office 365 / OneDrive scenarios where GetActiveObject returns an
    empty Application proxy.  In these cases, documents are registered in the
    ROT as file monikers (.docx paths or https://d.docs.live.net/... URLs).
    We grab the Document COM object from such a moniker and reach the real
    Application via ``doc.Application``.

    Returns the Word.Application COM object if found, or None.
    """
    for app in _iter_word_apps():
        if app.Documents.Count > 0:
            return app
    return None


def find_document(app, filename: str = None):
    """Find an open document by filename.

    Args:
        app: Word.Application COM object.
        filename: Document name (basename) or full path.
                  If None or empty, returns the active document.

    Returns:
        Document COM object.

    Raises:
        ValueError: If the document is not found or no documents are open.
    """
    if not filename:
        if app.Documents.Count == 0:
            raise ValueError("No documents are open in Word")
        doc = app.ActiveDocument
        from word_document_server.core.path_policy import check_path
        check_path(doc.FullName)
        return doc
    is_path = ntpath.isabs(filename) or '/' in filename or '\\' in filename
    target = normalize_document_path(filename) if is_path else unicodedata.normalize('NFC', filename).casefold()
    matches, seen = [], set()

    def inspect(candidate):
        docs = candidate.Documents
        for i in range(1, docs.Count + 1):
            doc = docs(i)
            full = normalize_document_path(doc.FullName)
            identity = (getattr(getattr(doc, 'Windows', None), 'Count', 0) and doc.Windows(1).Hwnd, full)
            if identity in seen:
                continue
            seen.add(identity)
            key = full if is_path else unicodedata.normalize('NFC', doc.Name).casefold()
            if key == target:
                matches.append(doc)

    inspect(app)
    # A full path is authoritative: never fall back to a same-named document.
    if is_path and len(matches) == 1:
        from word_document_server.core.path_policy import check_path
        check_path(matches[0].FullName)
        return matches[0]
    for candidate in _iter_word_apps():
        inspect(candidate)
    if len(matches) > 1:
        raise ValueError(f"Ambiguous document '{filename}'. Supply its full path.")
    if matches:
        from word_document_server.core.path_policy import check_path
        check_path(matches[0].FullName)
        return matches[0]
    raise ValueError(f"Document '{filename}' is not open in Word.")


def normalize_document_path(path):
    """Case-insensitive Windows path identity, including Unicode normalization."""
    path = str(path)
    if '://' not in path:
        path = ntpath.normpath(path if ntpath.isabs(path) else os.path.abspath(path))
    return unicodedata.normalize('NFC', path).casefold()


def _iter_word_apps():
    """Bind only Word/file ROT entries; never probe every other running app."""
    if sys.platform != 'win32':
        return
    import pythoncom
    import win32com.client
    rot = pythoncom.GetRunningObjectTable(0)
    enum = rot.EnumRunning()
    seen = set()
    while True:
        batch = enum.Next(1)
        if not batch:
            break
        try:
            moniker = batch[0]
            name = moniker.GetDisplayName(pythoncom.CreateBindCtx(0), None).lower()
            is_app = '000209ff' in name
            if not is_app and not name.endswith(('.doc', '.docx', '.docm', '.dot', '.dotx', '.dotm')):
                continue
            obj = win32com.client.Dispatch(rot.GetObject(moniker).QueryInterface(pythoncom.IID_IDispatch))
            candidate = obj if is_app else obj.Application
            key = candidate.ActiveWindow.Hwnd
            if key not in seen:
                seen.add(key)
                yield candidate
        except Exception:
            continue


@contextmanager
def undo_record(app, name: str):
    """Wrap a block of COM mutations in a single Word UndoRecord.

    Groups all changes into one Ctrl+Z entry in Word's undo stack.
    The undo record name appears in Edit > Undo and in the undo history.
    Degrades gracefully on Word 2007 or earlier (no UndoRecord support).

    Args:
        app: Word.Application COM object.
        name: Label for the undo entry (truncated to 64 chars by Word).

    Usage::

        with undo_record(app, "MCP: Insert Text"):
            doc.Range(0, 0).InsertBefore("Hello")
    """
    try:
        rec = app.UndoRecord
    except AttributeError:
        rec = None  # Legacy Word without UndoRecord support.
    if rec is not None:
        if rec.IsRecordingCustomRecord:
            raise RuntimeError('Word already has an active custom Undo record. Finish it before editing.')
        rec.StartCustomRecord(name[:64])
    previous_screen = None
    try:
        try:
            previous_screen = app.ScreenUpdating
            app.ScreenUpdating = False
        except AttributeError:
            pass
        yield
    finally:
        try:
            if rec is not None:
                rec.EndCustomRecord()
        finally:
            if previous_screen is not None:
                app.ScreenUpdating = previous_screen


@contextmanager
def edit_session(app, doc, name, track_changes=False):
    """Restore author and tracking even if COM setup or the edit fails."""
    from word_document_server.defaults import DEFAULT_AUTHOR
    previous_track, previous_author = doc.TrackRevisions, app.UserName
    with undo_record(app, name):
        try:
            if track_changes:
                doc.TrackRevisions = True
                app.UserName = DEFAULT_AUTHOR
            yield
        finally:
            try:
                doc.TrackRevisions = previous_track
            finally:
                app.UserName = previous_author
