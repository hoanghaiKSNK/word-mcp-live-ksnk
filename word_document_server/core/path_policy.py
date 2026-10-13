"""Optional, canonical path containment for MCP tool inputs and derived outputs.

Unset WORD_MCP_ALLOWED_ROOTS preserves existing filesystem access. This is path
policy, not an OS sandbox: Word-linked resources and concurrent junction changes
are outside its boundary. Never pass untrusted macros to Word automation.
"""
import os
from pathlib import Path
from word_document_server.core.coordination import DispatchError

PATH_PARAMETERS = {'filename', 'source_filename', 'destination_filename',
                   'target_filename', 'source_filenames', 'directory',
                   'image_path', 'output_path', 'output_filename', 'save_as'}


def roots():
    value = os.getenv('WORD_MCP_ALLOWED_ROOTS', '').strip()
    if not value:
        return ()
    entries = [entry.strip() for entry in value.split(os.pathsep)]
    if not entries or any(not entry or not Path(entry).is_absolute() for entry in entries):
        raise DispatchError('PATH_POLICY_INVALID', 'Allowed roots must be absolute paths separated by the OS path separator')
    return tuple(Path(entry).resolve() for entry in entries)


def check_path(value):
    allowed = roots()
    if not allowed:
        return Path(value).resolve()
    raw = os.fspath(value)
    if '://' in raw or raw.startswith('\\\\.\\') or 'GLOBALROOT' in raw.upper():
        raise DispatchError('PATH_NOT_ALLOWED', 'This path is outside the configured document roots')
    # Refuse NTFS alternate streams and device aliases, not just traversal.
    path = Path(raw)
    if any(':' in part for part in path.parts[1:]) or (os.name == 'nt' and ':' in path.name):
        raise DispatchError('PATH_NOT_ALLOWED', 'Alternate streams are not document paths')
    canonical = path.resolve()
    if not any(canonical == root or root in canonical.parents for root in allowed):
        raise DispatchError('PATH_NOT_ALLOWED', 'This path is outside the configured document roots')
    return canonical


def input_paths(arguments, live=False, tool_name=''):
    paths = set()
    for name, value in arguments.items():
        if name not in PATH_PARAMETERS or not value:
            continue
        values = value if isinstance(value, (list, tuple)) else [value]
        for item in values:
            # A live basename is a Word document selector, not a cwd filename.
            # Its actual FullName is checked by find_document before use.
            if live and name == 'filename' and not os.path.isabs(item) and '/' not in item and '\\' not in item:
                continue
            paths.add(str(check_path(item)))
            if not live and name == 'output_filename' and tool_name == 'convert_to_pdf' and not item.lower().endswith('.pdf'):
                paths.add(str(check_path(item + '.pdf')))
            if not live and 'filename' in name and not item.endswith('.docx') and name != 'output_filename':
                paths.add(str(check_path(item + '.docx')))
    if not live and tool_name == 'copy_document' and not arguments.get('destination_filename'):
        source = arguments.get('source_filename')
        if source:
            source = source if source.endswith('.docx') else source + '.docx'
            base, ext = os.path.splitext(source)
            paths.add(str(check_path(base + '_copy' + ext)))
    if not live and tool_name == 'convert_to_pdf' and not arguments.get('output_filename'):
        source = arguments.get('filename')
        if source:
            source = source if source.endswith('.docx') else source + '.docx'
            paths.add(str(check_path(os.path.splitext(source)[0] + '.pdf')))
    return sorted(os.path.normcase(path) for path in paths)
