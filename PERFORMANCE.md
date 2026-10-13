# Synthetic Word page extraction benchmark - 13 October 2026

Measures reading the final page using the KSNK range-based implementation versus the previous full paragraph scan. A private Word instance contains Vietnamese synthetic text, A4, Times New Roman 13, 1.5 line spacing and explicit page breaks. Five samples per case. Tracking-on includes one genuine insertion revision. No user document or model API is involved.

| Pages | Tracking | Paragraphs | New median (s) | Old scan median (s) | New p95 (s) | Old scan p95 (s) |
| --- | --- | --- | --- | --- | --- | --- |
| 10 | Off | 68 | 0.1354 | 0.6446 | 0.1842 | 0.7963 |
| 10 | On | 68 | 0.1219 | 0.5126 | 0.1323 | 0.7074 |
| 50 | Off | 343 | 0.1710 | 3.5655 | 0.1930 | 4.5643 |
| 50 | On | 343 | 0.1290 | 3.5467 | 0.1614 | 3.8790 |
| 100 | Off | 682 | 0.1368 | 7.3585 | 0.1842 | 7.7864 |
| 100 | On | 682 | 0.1397 | 8.2269 | 0.1642 | 9.1735 |

Every returned paragraph, index and range matched the legacy scan. The existing ksnk.1 page optimization is being validated at larger sizes; this table does not attribute all gains to the ksnk.2 coordination change. These are small synthetic samples, with interpolated p95, not production tail-latency estimates. Actual COM RPC counts, CPU saturation, model response latency and all-edit performance were not measured.

Run with installed Word on Windows:

```powershell
python scripts/benchmark_ksnk_word.py --output performance.json --samples 5
```

Live and file scheduling metrics are available through word_mcp_dispatch_status. Independent file work has four workers by default; interactive Word operations remain serialized. The concurrency tests verify actual overlap using gates, not timing guesses.
