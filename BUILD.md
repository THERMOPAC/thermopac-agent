# Thermopac Drawing Structuring Agent — Build & Delivery

**Version:** v1.0.24
**Phase:** Phase 1

## What this package provides

| Item | Status |
|---|---|
| All Python source code | ✅ Complete — production-ready |
| `config.ini` template | ✅ Complete — edit with your credentials |
| `bootstrap.bat` | ✅ Complete — turns source into running agent in ~5 min on any Windows PC |
| `requirements.txt` | ✅ Complete |

## Delivery Path 1 — Bootstrap (recommended for initial deployment)

**Requirements:** Windows 10/11, Python 3.9+, internet access.

```
1. Download and extract this ZIP to any folder
2. Edit config.ini (node_id, node_token, solidworks_version, template_path)
3. Right-click bootstrap.bat → Run as administrator
4. Double-click Desktop shortcut "ThermopacStructurer" to start
```

Bootstrap takes ~3 minutes. It creates a `venv\` virtual environment,
installs `pywin32` and `requests`, registers COM DLLs, and runs a self-test.

## Delivery Path 2 — Build EXE (produces standalone binary)

**Requirements:** Windows 10/11, Python 3.11, `pip install pyinstaller`.

```
pip install -r requirements.txt
pip install pyinstaller
pyinstaller --onedir --name ThermopacStructurer --hidden-import win32com agent\main_structurer.py
```

Output: `dist\ThermopacStructurer\ThermopacStructurer.exe`

## Python dependencies

| Package | Purpose |
|---|---|
| `pywin32` | SolidWorks COM interface (`win32com.client`) |
| `requests` | HTTP client — cloud API calls |

Install: `pip install pywin32 requests`

## Supported SolidWorks Versions

| config.ini setting | ProgID |
|---|---|
| `solidworks_version = 2019` | `SldWorks.Application.27` |
| `solidworks_version = 2020` | `SldWorks.Application.28` |
| `solidworks_version = 2021` | `SldWorks.Application.29` |
| `solidworks_version = 2022` | `SldWorks.Application.30` |
| `solidworks_version = 2023` | `SldWorks.Application.31` |
| `solidworks_version = 2024` | `SldWorks.Application.32` |

## Safety contract (read before deployment)

- `DispatchEx()` only — always creates a NEW isolated SolidWorks process
- `GetActiveObject()` is **never** used — the agent never attaches to an engineer's session
- `swApp.Visible = False` always — runs headless
- `ExitApp()` in `finally` — always runs, even on error
- Orphan guard: `taskkill /F /PID` if `ExitApp()` fails
- Pre-flight runs **before** SolidWorks launches — if pre-flight fails, SW is never touched

For full audit baseline see: `docs/drawing-structuring-agent-baseline.md`
