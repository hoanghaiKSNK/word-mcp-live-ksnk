# KSNK incremental P0 acceptance - 13 October 2026

Baseline: upstream 12bbff97873bbce38be2dde52af4a6209a92a381, previous KSNK commit 6d52b664832322f38cdcec8137329ee371418d85. New package: 1.6.21+ksnk.2. Original MIT license retained; no KitchenSink4Word source was copied.

## Problem and final behavior

Separate Word MCP processes previously held only their own STA queues. They could enter the same interactive Word concurrently. A named Windows mutex now protects cooperating live clients in a login session. Contention returns APP_BUSY before the tool body; abandoned ownership refuses the next operation with APP_STATE_UNCERTAIN. The lock does not protect against previous releases or arbitrary add-ins. Restart old MCP sessions after switching runtime.

Closed-file tools previously returned async coroutines that could perform synchronous parsing on the MCP event loop. They now run in a bounded four-thread pool. Canonical resources are locked in sorted order; same-file contention returns FILE_BUSY. Different files can run concurrently, including while the live Word lock is held. Word itself stays on a single STA with an idle COM message pump.

Diagnostics are immediately available through word_mcp_dispatch_status. They identify this server's version/PID and report live/file queue depths, request IDs, current duration and bounded recent timings. No arguments, document paths or text enter diagnostics. Queued jobs expire before starting after 30 seconds by default. Running COM jobs are not interrupted or retried after client cancellation.

Word can include reply descendants in Document.Comments after reopening. get_comments now filters by Comment.Ancestor, reports replies only under their parent and preserves native indexes. Scope positions and resolved state are included. If ancestry cannot be read on an older Word version, the response states that limitation.

Optional WORD_MCP_ALLOWED_ROOTS gates canonical tool paths, source lists, resolved live selectors and derived output helpers. It is deliberately unset in the default installation to preserve existing document locations. It is path containment, not an OS sandbox; Word-linked resources and hostile concurrent filesystem replacement are outside its boundary.

## Validation performed

- Full suite: 50 passed, zero failed, zero skipped, on Windows with native Word enabled. Synthetic fixtures only; private Word instances are used for mutations.
- Real OS cross-process mutex contention and abandoned-owner rejection; recovery after tool exceptions; queue expiry/fullness; cancelled running operation executes once and retains ownership; dead STA refuses rather than accepting work into a dead queue.
- Independent-file overlap, same-file rejection and continued slot ownership after cancellation.
- Path traversal, sibling-directory and actual Windows junction escape refusal; denied requests do not enter the tool body; source lists and derived outputs checked.
- Native Track Changes/Undo/table/header/cross-reference/PDF regressions retained. Added comment anchor and threaded-reply persistence after Word reopen, find across formatting runs with unaffected styles retained, and interrupted-edit setting restoration followed by explicit Undo.
- Installed wheel tested over stdio with two MCP clients. Both refused an edit while another process held the live lock; diagnostics and file creation remained available. Live tracked replace, Undo, comments/reply and PDF export passed. An outside-root read was refused before entry. Existing tool schemas retained, plus one diagnostic tool (121 total).
- Synthetic 10/50/100-page extraction benchmarks, tracking off/on, five samples per case, matched old scan output. See PERFORMANCE.md.

## Remaining scope

This is not full closure of the original roadmap. Still outstanding: systematic issue-by-issue reproductions for upstream #7/#15/#16/#17/PR #18; a comprehensive policy review across every platform and every indirect file operation; persisted idempotency across server restarts; document-state/anchor preconditions for all mutations; benchmark coverage of all editing workflows and actual COM RPC counts; multi-instance live job routing; and deeper KitchenSink feature evaluation.

Failed native edits can leave changes. A custom Undo record is not an atomic transaction. Inspect the document and explicitly Undo when required. Do not blindly retry a timed-out edit.

## Design sources

- Office threading: https://learn.microsoft.com/en-us/visualstudio/vsto/threading-support-in-office?view=visualstudio
- Comment ancestry/replies: https://learn.microsoft.com/en-us/office/vba/api/word.comment.ancestor and https://learn.microsoft.com/en-us/office/vba/api/word.comment.replies
- KitchenSink4Word documentation (design comparison only; community AGPL-3.0): https://github.com/KitchenSink4AI/KitchenSink4Word
