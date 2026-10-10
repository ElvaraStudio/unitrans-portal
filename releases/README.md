# Releases

Built artifacts, kept because a build cannot always be reproduced later.

| File | Built | Notes |
|---|---|---|
| `UniTrans Waybill Sorter.exe` | 2026-06-15 09:42 | **Current build.** Matches `waybill_sorter.py` as of 09:40 and the 1.0 feature set in `../README.md`. |
| `UniTrans Waybill Sorter 1.0.exe` | 2026-06-15 08:43 | Earlier build from a source revision that no longer exists on disk. **Cannot be rebuilt** — kept as a rollback. |
| `UniTrans-Waybill-Sorter-transfer-20260615.zip` | 2026-06-15 16:31 | Hand-off bundle containing source, spec and both executables. Redundant with this folder — delete it if you do not need to re-send it. |

## Rebuilding

From the project root:

```powershell
.\build.ps1
```

`build.ps1` recreates `.venv`, installs `requirements.txt`, and writes a fresh
`dist\UniTrans Waybill Sorter.exe`.

Python resolution order: `py -3` → `python3` / `python` (skipping the Windows
Store alias stub) → `%LOCALAPPDATA%\Python\bin\python.exe`. Pin one explicitly
with `$env:UNITRANS_PYTHON`.

Two things were changed when this project moved out of `Documents\Codex`:

- The script previously preferred a **Codex-bundled** Python under
  `%USERPROFILE%\.cache\codex-runtimes\`, and fell back to the Windows Store
  `python.exe` stub — which silently opens the Store instead of running. Both are
  gone; a real interpreter is now required and validated up front.
- The two generated `.spec` files were removed. They were PyInstaller output
  (not inputs) and had absolute paths into the retired Codex runtime cache baked
  in. `build.ps1` regenerates its own.

The original 1.0 build used Python 3.12. A rebuild here resolves to Python
3.14.6 — **re-test before shipping to the client.**
