"""Atomic DOCX saves with recoverable destination backups.

python-docx >= 1.2 preserves the source OPC relationship graph, including
comments. Never inject comment parts from an unrelated destination file.
"""
import os
import shutil
import tempfile
import uuid
import zipfile
from datetime import datetime, timezone
from io import BytesIO
from pathlib import Path


def backup_existing(path):
    """Preserve an existing output; a failed backup aborts the overwrite."""
    from word_document_server.core.path_policy import check_path
    path = check_path(path)
    if not path.exists():
        return None
    folder = path.parent / '.word-mcp-backups'
    check_path(folder)
    folder.mkdir(exist_ok=True)
    stamp = datetime.now(timezone.utc).strftime('%Y%m%dT%H%M%S%fZ')
    dest = folder / f'{path.name}.{stamp}.{uuid.uuid4().hex[:8]}.bak'
    shutil.copy2(path, dest)
    return dest


def atomic_write_bytes(data, path, backup=True):
    """Replace only after serialization succeeds, on the same filesystem."""
    from word_document_server.core.path_policy import check_path
    path = check_path(path)
    fd, temporary = tempfile.mkstemp(prefix='.word-mcp-', suffix='.tmp', dir=path.parent)
    try:
        with os.fdopen(fd, 'wb') as output:
            output.write(data)
            output.flush()
            os.fsync(output.fileno())
        if backup:
            backup_existing(path)
        os.replace(temporary, path)
    finally:
        Path(temporary).unlink(missing_ok=True)


def install_save_hook():
    """Install once; leave stream outputs under their caller's control."""
    from docx.document import Document
    original = Document.save
    if getattr(original, '_custom_parts_hooked', False):
        return

    def safe_save(self, path_or_stream):
        if not isinstance(path_or_stream, (str, os.PathLike)):
            return original(self, path_or_stream)
        buffer = BytesIO()
        original(self, buffer)
        data = buffer.getvalue()
        with zipfile.ZipFile(BytesIO(data)) as package:
            bad = package.testzip()
            if bad:
                raise ValueError(f'Invalid DOCX member: {bad}')
        atomic_write_bytes(data, path_or_stream)

    safe_save._custom_parts_hooked = True
    safe_save.__wrapped__ = original
    Document.save = safe_save
