"""PDF export that owns only its private Word instance and preserves outputs."""
import os
import tempfile
from pathlib import Path
from word_document_server.utils.save_utils import atomic_write_bytes


def export_open_document(doc, output):
    """Export without SaveAs (which would change the live document identity)."""
    from word_document_server.core.path_policy import check_path
    output = check_path(output)
    fd, temporary = tempfile.mkstemp(suffix='.pdf', prefix='.word-mcp-', dir=output.parent)
    os.close(fd)
    try:
        doc.ExportAsFixedFormat(OutputFileName=temporary, ExportFormat=17,
                                OpenAfterExport=False)
        data = Path(temporary).read_bytes()
        if not data.startswith(b'%PDF-'):
            raise ValueError('Word did not produce a valid PDF')
        atomic_write_bytes(data, output)
    finally:
        Path(temporary).unlink(missing_ok=True)


def convert_in_private_word(filename, output):
    """Call on a worker thread; never Quit a user's GetActiveObject instance."""
    import pythoncom
    import win32com.client
    from word_document_server.core.path_policy import check_path
    filename = check_path(filename)
    check_path(output)
    pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
    app = doc = None
    try:
        app = win32com.client.DispatchEx('Word.Application')
        app.Visible = False
        app.DisplayAlerts = 0
        app.AutomationSecurity = 3  # Disable macros in the private converter.
        doc = app.Documents.Open(FileName=os.path.abspath(filename), ReadOnly=True,
                                 AddToRecentFiles=False, Visible=False)
        export_open_document(doc, output)
    finally:
        try:
            if doc is not None:
                doc.Close(SaveChanges=0)
        finally:
            try:
                if app is not None:
                    app.Quit(SaveChanges=0)
            finally:
                pythoncom.CoUninitialize()
