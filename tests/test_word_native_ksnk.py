"""Opt-in real Word tests. Only a private DispatchEx instance is modified.

Run WORD_MCP_NATIVE=1 python -m pytest tests/test_word_native_ksnk.py -q
No user documents are opened, saved or closed.
"""
import asyncio
import json
import os
import sys
import time
from pathlib import Path

import pytest
from word_document_server.core import word_com, table_com
from word_document_server.tools import live_tools, live_read_tools, live_layout_tools

pytestmark = pytest.mark.skipif(sys.platform != 'win32' or os.environ.get('WORD_MCP_NATIVE') != '1',
                                reason='Explicit opt-in and Windows Word required')


def call(function, **kwargs):
    result = json.loads(asyncio.run(function(**kwargs)))
    assert 'error' not in result, result
    assert result.get('success', True), result
    return result


@pytest.fixture(scope='module')
def private_word():
    import pythoncom
    import win32com.client
    pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
    try:
        try: previous = win32com.client.GetActiveObject('Word.Application')._oleobj_.QueryInterface(pythoncom.IID_IUnknown)
        except Exception: previous = None
        app = win32com.client.DispatchEx('Word.Application')
        owned = app._oleobj_.QueryInterface(pythoncom.IID_IUnknown) != previous
        assert owned, 'DispatchEx reused an existing Word instance; refusing edits'
        app.Visible = False; app.DisplayAlerts = 0; app.AutomationSecurity = 3
        yield app
    finally:
        try:
            if 'app' in locals() and locals().get('owned', False):
                app.Quit(SaveChanges=0)
        finally: pythoncom.CoUninitialize()


@pytest.fixture
def document(private_word, tmp_path, monkeypatch):
    app = private_word; doc = app.Documents.Add()
    doc.Content.Font.Name = 'Times New Roman'; doc.Content.Font.Size = 13
    setup = doc.PageSetup; setup.PageWidth = 21 * 72 / 2.54
    setup.PageHeight = 29.7 * 72 / 2.54
    setup.TopMargin = 3.5 * 72 / 2.54; setup.BottomMargin = 3 * 72 / 2.54
    setup.LeftMargin = 3.5 * 72 / 2.54; setup.RightMargin = 2 * 72 / 2.54
    doc.Content.ParagraphFormat.LineSpacingRule = 1
    doc.Content.Text = 'Kiểm soát nhiễm khuẩn KSNK KSNK KSNK.\rCuối tài liệu an toàn.\r'
    doc.SaveAs2(str(tmp_path/'KSNK synthetic.docx'), FileFormat=16)
    monkeypatch.setattr(word_com, 'get_word_app', lambda: app)
    monkeypatch.setattr(word_com, '_iter_word_apps', lambda: [])
    yield doc
    doc.Close(SaveChanges=0)


@pytest.mark.parametrize('inherited', [False, True])
def test_replace_all_tracked_and_single_undo(document, inherited):
    doc=document; before=doc.Content.Text; author=doc.Application.UserName
    doc.TrackRevisions=inherited
    result=call(live_tools.word_live_replace_text, filename=doc.FullName,
                find_text='KSNK', replace_text='IPC', replace_all=True, track_changes=not inherited)
    assert result['replacements']==3 and result['tracked']
    assert doc.Revisions.Count >= 3
    assert bool(doc.TrackRevisions)==inherited
    assert doc.Application.UserName==author and doc.Application.ScreenUpdating
    assert doc.Undo(1)
    assert doc.Content.Text==before and doc.Revisions.Count==0


@pytest.mark.parametrize('replacement', ['', 'KSNK KSNK', 'KSNK'])
def test_forward_replace_does_not_revisit_replacement(document, replacement):
    result=call(live_tools.word_live_replace_text, filename=document.FullName,
                find_text='KSNK', replace_text=replacement)
    assert result['replacements']==3


def test_comments_reply_resolve_undo_and_roundtrip(document, tmp_path):
    doc=document
    call(live_read_tools.word_live_add_comment, filename=doc.FullName, start=0, end=8,
         text='Bình luận kiểm thử tiếng Việt', author='KSNK Test')
    assert doc.Comments.Count==1
    call(live_read_tools.word_live_reply_to_comment, filename=doc.FullName, comment_index=1,
         text='Phản hồi kiểm thử', author='KSNK Test')
    assert doc.Comments(1).Replies.Count==1
    call(live_read_tools.word_live_resolve_comment, filename=doc.FullName, comment_index=1)
    assert doc.Comments(1).Done
    doc.Undo(1); assert not doc.Comments(1).Done
    doc.Save()
    from docx import Document
    from word_document_server.utils.save_utils import install_save_hook
    install_save_hook(); dest=tmp_path/'comment copy.docx'; Document(doc.FullName).save(dest)
    import zipfile
    with zipfile.ZipFile(dest) as package:
        assert 'word/comments.xml' in package.namelist()
        assert 'word/commentsExtended.xml' in package.namelist()
        assert 'Bình luận' in package.read('word/comments.xml').decode('utf-8')


def test_table_delete_range_preserves_following_text_and_undo(document):
    doc=document; table=doc.Tables.Add(doc.Range(0,0), NumRows=2, NumColumns=2)
    table_com.set_range(table, [['Một','Hai'],['Ba','Bốn']])
    before=doc.Content.Text; tail=doc.Range(table.Range.End,doc.Content.End).Text
    start,end=table.Range.Start,table.Range.End
    call(live_tools.word_live_delete_text, filename=doc.FullName,start=start,end=end)
    assert doc.Tables.Count==0 and doc.Content.Text==tail
    doc.Undo(1); assert doc.Content.Text==before and doc.Tables.Count==1


def test_bulk_cell_markers_and_single_undo(document):
    doc=document; table=doc.Tables.Add(doc.Range(0,0), NumRows=2, NumColumns=2)
    before=doc.Content.Text
    call(live_tools.word_live_modify_table,filename=doc.FullName,operation='set_range',
         cells=[['Vi sinh','Đề kháng'],['E. coli','KSNK']],start_row=1,start_col=1)
    assert table.Rows.Count==2 and table.Columns.Count==2
    assert table.Cell(1,1).Range.Text=='Vi sinh\r\x07'
    doc.Undo(1); assert doc.Content.Text==before


def test_header_numbering_keeps_existing_text_and_selection(document):
    doc=document; header=doc.Sections(1).Headers(1); header.Range.Text='KSNK - HEADER\r'
    before=doc.Application.Selection.Start
    original_header=header.Range.Text
    call(live_layout_tools.word_live_add_page_numbers,filename=doc.FullName,position='header',
         prefix='Trang ',suffix=' - KSNK',include_total=True)
    assert 'KSNK - HEADER' in header.Range.Text and 'Trang ' in header.Range.Text
    assert [f.Type for f in header.Range.Fields]==[33,26]
    assert doc.Application.Selection.Start==before
    doc.Undo(1); assert header.Range.Text==original_header


def test_pdf_export_keeps_document_identity_and_review_state(document, tmp_path):
    doc=document; doc.TrackRevisions=True
    call(live_tools.word_live_insert_text,filename=doc.FullName,text='Thay đổi thử ',position='start')
    call(live_read_tools.word_live_add_comment,filename=doc.FullName,start=0,end=3,text='Ghi chú')
    before=(doc.FullName,doc.Content.Text,doc.Revisions.Count,doc.Comments.Count,doc.Application.ActiveWindow.Hwnd)
    output=tmp_path/'output.pdf'
    call(live_tools.word_live_save,filename=doc.FullName,save_as=str(output))
    assert output.read_bytes().startswith(b'%PDF-')
    assert before==(doc.FullName,doc.Content.Text,doc.Revisions.Count,doc.Comments.Count,doc.Application.ActiveWindow.Hwnd)


def test_page_read_matches_full_collection(document):
    doc=document
    doc.Content.Text=''.join(f'Đoạn {i} KSNK kiểm thử trang.\r' for i in range(1,181))
    doc.Repaginate(); pages=doc.ComputeStatistics(2)
    for page in (1,pages):
        started=time.perf_counter()
        result=call(live_read_tools.word_live_get_page_text,filename=doc.FullName,page=page)
        new_elapsed=time.perf_counter()-started
        expected=[]
        started=time.perf_counter()
        for i in range(1,doc.Paragraphs.Count+1):
            rng=doc.Paragraphs(i).Range
            if rng.End <= result['range_start']: continue
            if rng.Start >= result['range_end']: break
            expected.append({'index':i,'text':rng.Text.rstrip('\r\x07'),'char_start':rng.Start,'char_end':rng.End})
        assert result['paragraphs']==expected
        if page==pages:
            print(json.dumps({'benchmark':'last_page','paragraphs':doc.Paragraphs.Count,
                             'pages':pages,'new_seconds':round(new_elapsed,4),
                             'legacy_scan_seconds':round(time.perf_counter()-started,4)}))


def test_invalid_format_and_table_inputs_do_not_partially_mutate(document):
    doc=document; before=doc.Content.Text; old_bold=doc.Content.Font.Bold
    result=json.loads(asyncio.run(live_tools.word_live_format_text(filename=doc.FullName,
                      start=0,end=5,bold=True,paragraph_alignment='invalid')))
    assert 'error' in result and doc.Content.Font.Bold==old_bold
    result=json.loads(asyncio.run(live_tools.word_live_add_table(filename=doc.FullName,rows=1,cols=1,
                      data=[['ok','overflow']])))
    assert 'error' in result and doc.Tables.Count==0 and doc.Content.Text==before


def test_live_save_keeps_disk_backup(document):
    doc=document; original=Path(doc.FullName).read_bytes()
    call(live_tools.word_live_insert_text,filename=doc.FullName,text='Bổ sung ',position='start')
    call(live_tools.word_live_save,filename=doc.FullName)
    backups=list((Path(doc.FullName).parent/'.word-mcp-backups').glob('*.bak'))
    assert len(backups)==1 and backups[0].read_bytes()==original


def test_linked_section_header_isolated(document):
    doc=document
    doc.Sections(1).Headers(1).Range.Text='HEADER CHUNG'
    call(live_layout_tools.word_live_add_section_break,filename=doc.FullName)
    original=doc.Sections(1).Headers(1).Range.Text
    call(live_layout_tools.word_live_add_page_numbers,filename=doc.FullName,section_index=2,
         position='header',prefix='Trang ')
    assert doc.Sections(1).Headers(1).Range.Text==original
    assert doc.Sections(1).Headers(1).Range.Fields.Count==0
    assert not doc.Sections(2).Headers(1).LinkToPrevious


def test_cross_reference_does_not_move_selection(document):
    doc=document; doc.Bookmarks.Add('KSNKTarget',doc.Range(0,8))
    before=doc.Application.Selection.Start
    call(live_tools.word_live_insert_cross_reference,filename=doc.FullName,ref_type='bookmark',
         ref_item=1,ref_kind='text',insert_position='end')
    assert doc.Application.Selection.Start==before
    assert any(field.Type==3 for field in doc.Fields)  # wdFieldRef


def test_figure_reference_uses_real_caption_ids(document):
    doc=document
    doc.Range(0,0).InsertCaption(Label=-1, Title=' - Hình kiểm thử')
    result=call(live_tools.word_live_list_cross_reference_items,filename=doc.FullName,ref_type='figure')
    assert result['count']==1
    call(live_tools.word_live_insert_cross_reference,filename=doc.FullName,ref_type='figure',
         ref_item=1,ref_kind='number',insert_position='end')
    assert any(field.Type==3 for field in doc.Fields)


def test_cursor_cannot_edit_another_document(document):
    doc=document; other=doc.Application.Documents.Add()
    try:
        other.Content.Text='OTHER DOCUMENT'
        doc_before,other_before=doc.Content.Text,other.Content.Text
        result=json.loads(asyncio.run(live_tools.word_live_insert_text(filename=doc.FullName,
                                     position='cursor',text='must not insert')))
        assert 'error' in result
        assert doc.Content.Text==doc_before and other.Content.Text==other_before
    finally:
        other.Close(SaveChanges=0)


def test_comment_anchor_survives_outside_edits_save_and_word_reopen(document):
    doc = document
    start = doc.Content.Text.index('KSNK')
    call(live_read_tools.word_live_add_comment, filename=doc.FullName, start=start,
         end=start + 4, text='Neo KSNK phải còn đúng', author='KSNK Test')
    call(live_read_tools.word_live_reply_to_comment, filename=doc.FullName, comment_index=1,
         text='Phản hồi giữ nguyên', author='KSNK Test')
    call(live_tools.word_live_insert_text, filename=doc.FullName, position='start', text='Đầu tài liệu. ')
    call(live_tools.word_live_replace_text, filename=doc.FullName,
         find_text='Cuối tài liệu', replace_text='Phần kết luận', track_changes=True)
    assert doc.Comments(1).Scope.Text == 'KSNK'
    assert doc.Comments(1).Scope.Start == start + len('Đầu tài liệu. ')
    call(live_tools.word_live_save, filename=doc.FullName)
    saved = doc.FullName
    # A separate read-only Word open checks persisted anchors and threaded replies.
    # Avoid closing the fixture's original COM document proxy.
    import shutil
    folder = Path(saved).parent
    copy = folder / 'anchor reopen.docx'
    shutil.copy2(saved, copy)
    reopened = doc.Application.Documents.Open(str(copy), ReadOnly=True, AddToRecentFiles=False)
    try:
        report = call(live_read_tools.word_live_get_comments, filename=reopened.FullName)
        assert report['comment_count'] == 1
        assert not report['ancestry_unavailable']
        parent = report['comments'][0]
        assert parent['scope'] == 'KSNK'
        assert parent['reply_count'] == 1
        assert 'Phản hồi giữ nguyên' in parent['replies'][0]['text']
        assert reopened.Comments(parent['index']).Replies.Count == 1
        assert reopened.Revisions.Count > 0
    finally:
        reopened.Close(SaveChanges=0)


def test_replace_across_format_runs_keeps_unaffected_format_and_undo(document):
    doc = document
    doc.TrackRevisions = False
    doc.Content.Text = 'Trước ABCDEF sau.\r'
    start = doc.Content.Text.index('ABCDEF')
    doc.Range(start, start + 3).Font.Bold = True
    doc.Range(start + 3, start + 6).Font.Italic = True
    tail = doc.Content.Text.index('sau')
    doc.Range(tail, tail + 3).Font.Underline = 1
    before = doc.Content.Text
    result = call(live_tools.word_live_replace_text, filename=doc.FullName,
                  find_text='ABCDEF', replace_text='KSNK', track_changes=True)
    assert result['replacements'] == 1
    # Word-native Find spans formatting runs. No promise is made that all old
    # per-character styles map to a replacement of a different length.
    search = doc.Content.Duplicate
    assert search.Find.Execute(FindText='sau')
    assert search.Font.Underline == 1
    assert doc.Undo(1) and doc.Content.Text == before
    assert doc.Range(start, start + 3).Font.Bold
    assert doc.Range(start + 3, start + 6).Font.Italic


def test_interrupted_edit_restores_settings_and_can_be_undone(document):
    doc = document; app = doc.Application
    before = doc.Content.Text
    settings = (bool(doc.TrackRevisions), app.UserName, bool(app.ScreenUpdating))
    with pytest.raises(RuntimeError, match='injected failure'):
        with word_com.edit_session(app, doc, 'Synthetic failure test', True):
            doc.Range(0, 0).InsertBefore('Thay đổi trước lỗi. ')
            raise RuntimeError('injected failure')
    assert (bool(doc.TrackRevisions), app.UserName, bool(app.ScreenUpdating)) == settings
    # A failed operation can have changed Word. We inspect and Undo explicitly.
    assert doc.Content.Text != before
    assert doc.Undo(1) and doc.Content.Text == before
    call(live_tools.word_live_insert_text, filename=doc.FullName, position='start', text='Phục hồi. ')
    assert doc.Content.Text.startswith('Phục hồi. ')
