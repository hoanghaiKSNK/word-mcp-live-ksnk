"""Synthetic 10/50/100-page benchmark; owns and closes only its Word instance.

Measures page extraction against the old full paragraph scan, not model latency
or a claim that all editing operations scale identically. JSON contains timings
and sizes only. No patient or user document is read.
"""
import argparse
import asyncio
import json
from pathlib import Path
import statistics
import time
from unittest.mock import patch


def percentile(values, fraction):
    ordered = sorted(values)
    index = (len(ordered) - 1) * fraction
    lower = int(index)
    upper = min(lower + 1, len(ordered) - 1)
    return ordered[lower] + (ordered[upper] - ordered[lower]) * (index - lower)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--output', type=Path, required=True)
    parser.add_argument('--samples', type=int, default=5)
    args = parser.parse_args()
    import pythoncom
    import win32com.client
    from word_document_server.core import word_com
    from word_document_server.tools.live_read_tools import word_live_get_page_text
    pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
    app = None
    owned = False
    rows = []
    try:
        try:
            previous = win32com.client.GetActiveObject('Word.Application')._oleobj_.QueryInterface(pythoncom.IID_IUnknown)
        except Exception:
            previous = None
        app = win32com.client.DispatchEx('Word.Application')
        owned = app._oleobj_.QueryInterface(pythoncom.IID_IUnknown) != previous
        assert owned, 'Refusing to reuse a user Word instance'
        app.Visible = False; app.DisplayAlerts = 0; app.AutomationSecurity = 3
        with patch.object(word_com, 'get_word_app', lambda: app), patch.object(word_com, '_iter_word_apps', lambda: []):
            for pages in (10, 50, 100):
                doc = app.Documents.Add()
                try:
                    setup = doc.PageSetup
                    setup.PageWidth = 21 * 72 / 2.54; setup.PageHeight = 29.7 * 72 / 2.54
                    setup.TopMargin = 3.5 * 72 / 2.54; setup.BottomMargin = 3 * 72 / 2.54
                    setup.LeftMargin = 3.5 * 72 / 2.54; setup.RightMargin = 2 * 72 / 2.54
                    content = []
                    for page in range(1, pages + 1):
                        content.append(''.join(f'Trang {page}, đoạn {para}: kiểm soát nhiễm khuẩn.\r' for para in range(1, 7)))
                        if page < pages: content.append('\f')
                    doc.Content.Text = ''.join(content)
                    doc.Content.Font.Name = 'Times New Roman'; doc.Content.Font.Size = 13
                    doc.Content.ParagraphFormat.LineSpacingRule = 1
                    doc.Repaginate()
                    assert doc.ComputeStatistics(2) == pages
                    for tracking in (False, True):
                        doc.TrackRevisions = tracking
                        if tracking:
                            doc.Range(0, 0).InsertBefore('Bổ sung có theo dõi. ')
                        doc.Repaginate()
                        actual_pages = doc.ComputeStatistics(2)
                        fast, legacy = [], []
                        for _ in range(args.samples):
                            started = time.perf_counter()
                            result = json.loads(asyncio.run(word_live_get_page_text(page=actual_pages)))
                            fast.append(time.perf_counter() - started)
                            assert 'error' not in result, result
                            started = time.perf_counter()
                            expected = []
                            for index in range(1, doc.Paragraphs.Count + 1):
                                rng = doc.Paragraphs(index).Range
                                if rng.End <= result['range_start']: continue
                                if rng.Start >= result['range_end']: break
                                expected.append({'index': index, 'text': rng.Text.rstrip('\r\x07'),
                                                 'char_start': rng.Start, 'char_end': rng.End})
                            legacy.append(time.perf_counter() - started)
                            assert result['paragraphs'] == expected
                        row = {'pages': actual_pages, 'paragraphs': doc.Paragraphs.Count,
                               'tracking': tracking, 'revision_count': doc.Revisions.Count,
                               'samples': args.samples,
                               'new_median_seconds': round(statistics.median(fast), 4),
                               'new_p95_seconds': round(percentile(fast, .95), 4),
                               'legacy_median_seconds': round(statistics.median(legacy), 4),
                               'legacy_p95_seconds': round(percentile(legacy, .95), 4),
                               'outputs_equal': True}
                        rows.append(row)
                        print(json.dumps(row), flush=True)
                finally:
                    doc.Close(SaveChanges=0)
        args.output.write_text(json.dumps({'benchmark': 'last_page_extraction',
                                         'p95_method': 'linear interpolation of small synthetic sample',
                                         'com_rpc_count_measured': False, 'rows': rows}, indent=2), encoding='utf-8')
    finally:
        try:
            if app is not None and owned: app.Quit(SaveChanges=0)
        finally: pythoncom.CoUninitialize()


if __name__ == '__main__': main()
