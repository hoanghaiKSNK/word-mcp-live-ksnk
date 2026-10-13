# KSNK maintenance fork

Version `1.6.21+ksnk.1`, based on upstream commit `12bbff97873bbce38be2dde52af4a6209a92a381`.
This is a maintenance fork of Word MCP Live, not a Microsoft/OpenAI product.
Original MIT license and contributor attribution are retained. No paid model API
or additional AI subscription is used by this MCP server.

## Safe live editing

- Supply the full path when several open documents have the same filename.
- Read current ranges immediately before editing. Word uses UTF-16 character
  positions; Python string indexes differ for characters such as emoji.
- Track Changes already enabled in the document stays enabled. Requesting
  `track_changes=True` temporarily enables it and restores the previous state.
- Tracked replace-all finds matches before editing and applies them backwards.
  Replacements do not accept/reject unrelated revisions. Limit: fewer than 50,000
  matches. A match crossing a table cell terminator aborts before replacements.
- Each live edit has a custom Undo record. An existing active custom record is
  never closed by this server; finish that record and explicitly request again.
- Bulk table writes reject overflows, unsafe control characters and aliases to
  merged cells before writing any cell. `None` skips a cell. Use individual
  operations for irregular merged tables.
- Page fields are appended while retaining existing header/footer content.
  A linked target section is detached from its predecessor. Existing PAGE
  fields must be edited explicitly rather than duplicated.
- DOCX saves and PDF exports create unique recoverable copies of existing
  outputs in the sibling `.word-mcp-backups` directory. Python-docx writes are
  atomic; native Word Save/SaveAs are native operations preceded by a backup.
  Source comments/related OPC parts are kept; destination comments are not merged.
- PDF export of an open document preserves its filename and review state.
  File conversion uses a separate macro-disabled, read-only Word instance and
  closes only that owned instance.

## Execution and cancellation

Windows live tools run serially on one COM STA thread. Read calls and edit calls
share the queue, so a long Word operation can delay other Word calls while MCP
remains responsive. Queue capacity: 32 waiting operations. Queued cancellation
prevents edits from starting; cancellation cannot safely interrupt a COM call
already running. Do not blindly retry an edit after a timeout: read the document
and its revisions first. This improves server scheduling, not AI model latency.

## Tests

```powershell
python -m pytest tests/test_ksnk_regressions.py tests/test_sta_dispatch.py -q
$env:WORD_MCP_NATIVE = '1'
python -m pytest tests/test_word_native_ksnk.py -q -p no:faulthandler
```

Native tests create and close a private Word instance with synthetic Vietnamese
documents. They check tracked replacements, a single Undo, comments/replies,
table boundaries, page fields, PDF identity and page-range extraction. The
standalone conversion test requires installed Word on Windows. Do not interpret
passing tests as proof that every Word feature or third-party add-in is fault-free.
macOS JXA behavior is retained and has not been tested on macOS in this update.

## Install and recover

Install a pinned wheel into a new virtual environment. Keep the previous runtime
unchanged. Back up the MCP configuration before switching only the `word_live`
command/cwd. Compare parsed configuration to ensure other MCP entries remain
unchanged. Start a new MCP session after switching; an existing session may still
hold the old process. Restore only the previous `word_live` entry for rollback,
so unrelated configuration changes made later are retained.

An MCP cannot guarantee atomic rollback of arbitrary COM failures. When a native
operation partially fails, inspect the document and explicitly Undo that call
if needed. Never accept all revisions or discard comments as an error recovery
shortcut. Keep original documents and use Save As for substantive revisions.

## Cross-reference implementation sources

Word enum values and bookmark ReferenceItem handling were checked against the
installed Word type library and Microsoft's documentation:
[WdReferenceKind](https://learn.microsoft.com/en-us/office/vba/api/word.wdreferencekind),
[WdCaptionLabelID](https://learn.microsoft.com/en-us/office/vba/api/word.wdcaptionlabelid),
[Range.InsertCrossReference](https://learn.microsoft.com/en-us/office/vba/api/word.range.insertcrossreference).
