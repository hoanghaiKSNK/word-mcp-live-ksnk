"""Installed-wheel stdio acceptance with private synthetic Word and two clients."""
import argparse
import asyncio
import json
from pathlib import Path
import os


def unpack(result):
    assert not result.is_error, result
    for block in result.content:
        if hasattr(block, 'text'):
            return json.loads(block.text)
    raise AssertionError('Missing JSON result')


async def verify(doc, runtime, output):
    from fastmcp import Client
    from fastmcp.client.transports import StdioTransport
    from word_document_server.core.coordination import live_lock
    env = {'MCP_TRANSPORT': 'stdio', 'PYTHONIOENCODING': 'utf-8',
           'WORD_MCP_ALLOWED_ROOTS': str(output)}
    def client(number):
        return Client(StdioTransport(command=str(runtime/'.venv/Scripts/word_mcp_server.exe'),
                                     args=[], cwd=str(runtime), env=env,
                                     log_file=output/f'mcp-{number}.log'), timeout=45, init_timeout=45)
    calls = []
    async with client(1) as first, client(2) as second:
        tools = await first.list_tools()
        (output/'tool-schemas.json').write_text(json.dumps([t.model_dump(mode='json') for t in tools], indent=2), encoding='utf-8')
        assert len(tools) == 121
        async def invoke(name, **kwargs):
            result = unpack(await first.call_tool(name, {'filename': doc.FullName, **kwargs}))
            assert 'error' not in result, result
            calls.append(name)
            return result
        statuses = await asyncio.gather(first.call_tool('word_mcp_dispatch_status', {}),
                                        second.call_tool('word_mcp_dispatch_status', {}))
        statuses = [unpack(status) for status in statuses]
        assert all(item['version'] == '1.6.21+ksnk.2' for item in statuses)
        assert statuses[0]['live']['pid'] != statuses[1]['live']['pid']
        before = doc.Content.Text
        with live_lock():
            # Each server rejects before COM; status and file tools remain usable.
            blocked = await asyncio.gather(first.call_tool('word_live_insert_text', {'filename': doc.FullName, 'text': 'must not insert'}),
                                           second.call_tool('word_live_insert_text', {'filename': doc.FullName, 'text': 'must not insert'}))
            blocked = [unpack(item) for item in blocked]
            assert all(item['code'] == 'APP_BUSY' and item['operation_started'] is False for item in blocked)
            status = unpack(await first.call_tool('word_mcp_dispatch_status', {}))
            assert status['live']['rejected'] >= 1
            files = [output/'file A.docx', output/'file B.docx']
            created = await asyncio.gather(*(first.call_tool('create_document', {'filename': str(path), 'title': 'Synthetic KSNK test'}) for path in files))
            assert all(not item.is_error for item in created)
            assert all(path.is_file() for path in files)
            assert doc.Content.Text == before
        await invoke('word_live_get_text')
        replaced = await invoke('word_live_replace_text', find_text='KSNK', replace_text='IPC', track_changes=True, replace_all=True)
        assert replaced['replacements'] == 3
        assert replaced['_dispatch']['operation_started'] is True
        await invoke('word_live_undo', times=1)
        assert doc.Content.Text == before
        await invoke('word_live_add_comment', start=0, end=8, text='Bình luận tiếng Việt', author='KSNK Test')
        await invoke('word_live_reply_to_comment', comment_index=1, text='Phản hồi', author='KSNK Test')
        comments = await invoke('word_live_get_comments')
        assert comments['comment_count'] == 1 and comments['comments'][0]['reply_count'] == 1
        await invoke('word_live_save', save_as=str(output/'mcp-ksnk2.pdf'))
        assert (output/'mcp-ksnk2.pdf').read_bytes().startswith(b'%PDF-')
        denied = unpack(await first.call_tool('get_document_text', {'filename': str(output.parent/'outside-policy.docx')}))
        assert denied['code'] == 'PATH_NOT_ALLOWED' and denied['operation_started'] is False
        status = unpack(await first.call_tool('word_mcp_dispatch_status', {}))
        serialized = json.dumps(status, ensure_ascii=False)
        assert 'Bình luận' not in serialized and 'synthetic-ksnk2.docx' not in serialized
        assert status['live']['queue_depth'] == 0
        return {'passed': True, 'version': status['version'], 'tools': len(tools),
                'two_server_pids': [s['live']['pid'] for s in statuses],
                'cross_process_edit_rejected': True, 'offline_work_while_live_locked': True,
                'path_rejected_before_read': True, 'live_calls': calls, 'diagnostics': status}


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--runtime', type=Path, required=True)
    parser.add_argument('--output', type=Path, required=True)
    args = parser.parse_args()
    args.output.mkdir(parents=True, exist_ok=True)
    import pythoncom
    import win32com.client
    pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
    app = doc = None
    owned = False
    try:
        try: previous = win32com.client.GetActiveObject('Word.Application')._oleobj_.QueryInterface(pythoncom.IID_IUnknown)
        except Exception: previous = None
        app = win32com.client.DispatchEx('Word.Application')
        owned = app._oleobj_.QueryInterface(pythoncom.IID_IUnknown) != previous
        assert owned, 'Refusing to reuse user Word'
        app.Visible = False; app.DisplayAlerts = 0; app.AutomationSecurity = 3
        doc = app.Documents.Add()
        doc.Content.Font.Name = 'Times New Roman'; doc.Content.Font.Size = 13
        doc.Content.ParagraphFormat.LineSpacingRule = 1
        setup = doc.PageSetup
        setup.PageWidth = 21*72/2.54; setup.PageHeight = 29.7*72/2.54
        setup.TopMargin = 3.5*72/2.54; setup.BottomMargin = 3*72/2.54
        setup.LeftMargin = 3.5*72/2.54; setup.RightMargin = 2*72/2.54
        doc.Content.Text = 'Kiểm soát nhiễm khuẩn - KSNK KSNK KSNK.\rTài liệu giả lập.\r'
        doc.SaveAs2(str(args.output/'synthetic-ksnk2.docx'), FileFormat=16)
        report = asyncio.run(verify(doc, args.runtime, args.output))
        (args.output/'protocol-acceptance.json').write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')
        print(json.dumps({key: value for key, value in report.items() if key != 'diagnostics'}, ensure_ascii=False))
    finally:
        try:
            if doc is not None: doc.Close(SaveChanges=0)
        finally:
            try:
                if app is not None and owned: app.Quit(SaveChanges=0)
            finally: pythoncom.CoUninitialize()


if __name__ == '__main__': main()
