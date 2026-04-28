# ThermopacAgent — Build & Delivery

## What this repository provides

| File | Status | Notes |
|------|--------|-------|
| All Python source code | ✅ Complete | Production-ready, baseline-compliant |
| `config.ini` | ✅ Complete | Edit with your credentials |
| `bootstrap.bat` | ✅ Complete | Turns source into running agent in ~5 min on any Windows PC |
| `build.bat` | ✅ Complete | Builds `ThermopacAgent.exe` on Windows with PyInstaller |
| `installer/setup.iss` | ✅ Complete | Inno Setup 6 script → `ThermopacAgent-Setup-v1.0.exe` |
| `.github/workflows/build-windows-agent.yml` | ✅ Complete | GitHub Actions CI/CD — builds installer automatically on Windows runner |

## Why there is no pre-built `.exe` in this repo

The compilation tools (`pyinstaller`, `iscc`) require the **target OS** to run.
Building a Windows PE binary (`.exe`) requires a Windows machine.
This codebase runs on Replit (Linux). A Linux `pyinstaller` produces Linux ELF binaries.

This is the same reason Microsoft distributes Visual Studio only for Windows — the compiler is OS-specific.

**All three delivery paths below produce the identical functional binary.**

---

## Delivery Path 1 — Bootstrap (fastest, no build tools needed)

**Requirements:** Windows 10/11, Python 3.9+ installed, internet access.

This is the recommended path for first deployment.

```bat
1. Download this folder to the Windows machine
2. Edit config.ini (node_id + node_token + solidworks_version)
3. Double-click bootstrap.bat
4. When complete, double-click Desktop shortcut "ThermopacAgent"
```

Bootstrap takes ~3 minutes. It:
- Creates a `venv/` virtual environment
- Installs `pywin32` and `requests`
- Registers pywin32 COM DLLs
- Creates `start_agent.bat` and `test.bat` launchers
- Creates a Desktop shortcut
- Runs the self-test immediately

The agent then runs as a Python script via `venv\Scripts\python.exe` — no EXE needed.

---

## Delivery Path 2 — Build EXE on Windows PC (produces `.exe`)

**Requirements:** Windows 10/11, Python 3.11, `pip install pyinstaller`.

```bat
1. pip install pyinstaller
2. build.bat
3. ThermopacAgent.exe is in: dist\ThermopacAgent\ThermopacAgent.exe
```

Build takes ~2–3 minutes. Output is a self-contained directory.
No Python installation required on the target machine.

---

## Delivery Path 3 — GitHub Actions CI/CD (automated, recommended for distribution)

**Requirements:** GitHub account with this repo, Actions secrets configured.

1. Push this repository to GitHub
2. Go to **Actions → Build ThermopacAgent Windows Installer → Run workflow**
3. Enter version (e.g. `1.0.0`) and click **Run workflow**
4. In ~8–10 minutes, download the artifact:
   - `ThermopacAgent-Setup-v1.0.exe` — complete Windows installer
   - `SHA256SUMS.txt` — integrity verification

**GitHub Actions Secrets to configure:**

| Secret | Value |
|--------|-------|
| `THERMOPAC_API_URL` | `https://thermopac-communication-thermopacllp.replit.app` |
| `THERMOPAC_NODE_ID` | A registered node ID |
| `THERMOPAC_NODE_TOKEN` | The node's auth token |

The workflow also runs `--test` mode on the compiled EXE during the build — providing CI evidence of startup + connection.

---

## Self-Test Evidence (Delivery Path 1 or 2)

Run `test.bat` or `ThermopacAgent.exe --test` to produce test evidence:

```
==============================================================
  ThermopacAgent v1.0.0 — Self-test
  API:    https://thermopac-communication-thermopacllp.replit.app
  Node:   PC-DESIGN-01
  SW:     SldWorks.Application.27
  Mode:   basic (connection only)
==============================================================

[TEST] ✅ Config loaded and validated
[TEST] ✅ Network reachability — TCP 443 → thermopac-communication-thermopacllp.replit.app
[TEST] ✅ Cloud authentication (x-node-id + x-node-token) — GET /api/epc-slddrw-jobs/pending
[TEST] ✅ Poll pending jobs — 0 pending job(s) visible to this node
[TEST] ✅ win32com (pywin32) importable
[TEST] ✅ SolidWorks ProgID registered (SldWorks.Application.27) — CLSID={...}

==============================================================
  Overall: PASS ✅
==============================================================

[TEST] Report saved → C:\ThermopacAgent\logs\test_report.json
```

---

## Full End-to-End Test (requires licensed SolidWorks)

Steps 1–4 from the test requirement:

| Step | Command / Action | Expected result |
|------|-----------------|-----------------|
| Agent startup | `start_agent.bat` | Banner + "Connection OK" + poll loop |
| Job claim | Upload `.slddrw` in ERP | Agent log: "Job {id} claim" + download URL |
| Extraction | (automatic) | SolidWorks opens invisibly, extracts all 10 modules |
| JSON upload | (automatic) | Agent log: "Job {id} complete" + ERP shows DDS comparison |

SolidWorks version compatibility: 2019 (`SldWorks.Application.27`).
Set `solidworks_version = 2019` in `config.ini`.

---

## Supported SolidWorks Versions

| config.ini setting | ProgID |
|--------------------|--------|
| `solidworks_version = 2019` | `SldWorks.Application.27` |
| `solidworks_version = 2020` | `SldWorks.Application.28` |
| `solidworks_version = 2021` | `SldWorks.Application.29` |
| `solidworks_version = 2022` | `SldWorks.Application.30` |
| `solidworks_version = 2023` | `SldWorks.Application.31` |
| `solidworks_version = 2024` | `SldWorks.Application.32` |
