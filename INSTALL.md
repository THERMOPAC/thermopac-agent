# ThermopacAgent — Installation Guide
### THERMOPAC ERP | SolidWorks Extraction Agent | Phase 1

---

## Prerequisites

Before starting, confirm the following on the Windows PC that will run the agent:

| Requirement | Details |
|---|---|
| Windows | 10 or 11, 64-bit |
| Python | 3.9 or higher — [python.org/downloads](https://python.org/downloads) |
| SolidWorks | 2019, 2020, 2021, 2022, 2023, or 2024 — licensed and working |
| Internet | Access to `thermopac-communication-thermopacllp.replit.app` (HTTPS port 443) |

Python does not need to be a special build — the standard installer from python.org is fine.
When installing Python, **check the box "Add Python to PATH"**.

---

## Step 1 — Admin Registers This Node

A Thermopac **Superuser** must create credentials for this PC in the cloud app before installation.

1. Log in to the Thermopac ERP
2. Navigate to **EPC → Drawing Controls**
3. Locate the **Agent Nodes** section and click **Register New Node**
4. Enter:
   - **Node ID** — a short unique name for this PC, e.g. `PC-DESIGN-01`
   - **Label** — friendly name, e.g. `Design Office — Prasad`
5. The system shows a **Node Token** — **copy it immediately** (displayed once only)

You now have:
- `node_id` — e.g. `PC-DESIGN-01`
- `node_token` — the secret token string

---

## Step 2 — Edit config.ini

Open `config.ini` (in this folder) in Notepad and fill in your values:

```ini
[cloud]
api_url    = https://thermopac-communication-thermopacllp.replit.app
node_id    = PC-DESIGN-01          ← change to your node ID from Step 1
node_token = REPLACE_WITH_YOUR_TOKEN   ← paste your token from Step 1

[agent]
poll_interval_sec = 10
job_timeout_sec   = 600
max_retries       = 3

[paths]
temp_dir = C:\ThermopacAgent\temp
log_dir  = C:\ThermopacAgent\logs

[solidworks]
solidworks_version = 2019          ← change to your installed SolidWorks version
visible = false
```

**SolidWorks version numbers:**

| Installed version | Set solidworks_version to |
|---|---|
| SolidWorks 2019 | `2019` |
| SolidWorks 2020 | `2020` |
| SolidWorks 2021 | `2021` |
| SolidWorks 2022 | `2022` |
| SolidWorks 2023 | `2023` |
| SolidWorks 2024 | `2024` |

Save and close config.ini.

---

## Step 3 — Run bootstrap.bat

**Right-click** `bootstrap.bat` → **Run as administrator**

The script will:

1. Locate your Python installation
2. Create an isolated Python virtual environment (`venv\` in this folder)
3. Install `pywin32` and `requests`
4. Register pywin32 COM DLLs with Windows (required for SolidWorks integration)
5. Create `start_agent.bat` — double-click to start the agent
6. Create `test.bat` — double-click to run a self-test
7. Create a **Desktop shortcut** called `ThermopacAgent`
8. Run the self-test immediately

This takes approximately **3–5 minutes** depending on download speed.

**Expected output at end of bootstrap:**

```
[TEST] ✅ Config loaded and validated
[TEST] ✅ Network reachability
[TEST] ✅ Cloud authentication (x-node-id + x-node-token)
[TEST] ✅ Poll pending jobs — 0 pending job(s)
[TEST] ✅ win32com (pywin32) importable
[TEST] ✅ SolidWorks ProgID registered (SldWorks.Application.27)

Overall: PASS ✅
```

If you see **FAIL** on any step, see the Troubleshooting section below.

---

## Step 4 — Start the Agent

Double-click the **ThermopacAgent** shortcut on the Desktop
(or double-click `start_agent.bat` in this folder).

A console window opens:

```
  _____ _
 |_   _| |__   ___ _ __ _ __ ___   ___  _ __   __ _  ___
 ...

  SolidWorks Extraction Agent  v1.0.0
  THERMOPAC ERP Integration

[Agent] Starting — api_url=https://thermopac-... node_id=PC-DESIGN-01
[Agent] Connection OK — entering poll loop
[Agent] No pending jobs
[Agent] No pending jobs
...
```

Leave this window open (you can minimise it). It polls every 10 seconds and picks up jobs automatically.

---

## Step 5 — Verify End-to-End (first real job)

1. In the Thermopac ERP, open any EPC drawing record
2. Navigate to **Drawing Controls → Drawing Verification**
3. Upload a `.slddrw` file
4. Watch the agent console — within 10 seconds:

```
[Agent] 1 pending job(s) — processing first
[Runner] Job 7 start — file=C10308-CPS-ACS-S6T-20-P28.slddrw
[Runner] Downloading to C:\ThermopacAgent\temp\...
[Extractor] Launching SolidWorks (SldWorks.Application.27)...
[Extractor] Document open
[Properties] drawing_number='C10308-...' revision='B'
[Sheets] 2 sheet(s) found
[DesignData] Found 8 row(s)
[Runner] Uploading extraction result (JSON)...
[Runner] Job 7 complete (87.4s)
```

5. Back in the ERP, the **Drawing Verification** card refreshes and shows:
   - Extracted parameters
   - DDS comparison (PASS / WARN / MISMATCH per parameter)
   - Critical mismatches highlighted in red (blocks approval)

---

## Running the Self-Test

At any time, double-click `test.bat` to verify the installation.

A full report is saved to:
```
C:\ThermopacAgent\logs\test_report.json
```

---

## Log Files

| File | Contents |
|---|---|
| `C:\ThermopacAgent\logs\agent.log` | Main agent log (daily rotation, 30 days retained) |
| `C:\ThermopacAgent\logs\test_report.json` | Latest self-test result |

---

## Troubleshooting

| Symptom | Action |
|---|---|
| `Authentication failed` | Verify `node_id` and `node_token` in config.ini match what the admin registered |
| `SolidWorks ProgID not found` | Check `solidworks_version` in config.ini matches your installed version |
| `Python not found` in bootstrap.bat | Install Python from python.org — check "Add Python to PATH" |
| `OpenDoc6 returned None` | SolidWorks licence issue — try `visible = true` to see the error on screen |
| `DesignDataNotFoundError` | The drawing must contain a table with "Design Data" in its title |
| Job stuck at `processing` | Cloud auto-resets stale jobs after 30 minutes — check agent.log for the error |
| Agent crashes immediately | Open `C:\ThermopacAgent\logs\agent.log` and check the last few lines |
| `pip install` fails | Check internet connection — port 443 must be open to `pypi.org` |

---

## Files in This Package

| File / Folder | Purpose |
|---|---|
| `bootstrap.bat` | **Start here** — sets up Python environment and launchers |
| `config.ini` | Configuration — edit with your credentials and SW version |
| `start_agent.bat` | Created by bootstrap — starts the polling agent |
| `test.bat` | Created by bootstrap — runs self-test |
| `build.bat` | Builds `ThermopacAgent.exe` (requires PyInstaller) |
| `agent/` | Core agent modules (config, logger, job client, runner, main) |
| `extractor/` | SolidWorks extraction modules (10 specialist extractors) |
| `installer/setup.iss` | Inno Setup script for compiled installer |
| `requirements.txt` | Python dependency list |
| `BUILD.md` | How to build the compiled EXE installer via GitHub Actions |
| `INSTALL.md` | This file |

---

## Getting the Compiled Installer (optional)

If you need a proper Windows installer (`.exe`) for deployment to machines without Python:

1. Push this repository to GitHub
2. Configure three repository secrets: `THERMOPAC_API_URL`, `THERMOPAC_NODE_ID`, `THERMOPAC_NODE_TOKEN`
3. Go to **Actions → Build ThermopacAgent Windows Installer → Run workflow**
4. Enter version `1.0.0` and click Run
5. Download `ThermopacAgent-Setup-v1.0.exe` from the build artifacts (~8 minutes)

See `BUILD.md` for full details.

---

*Agent baseline: `docs/slddrw-extraction-agent-baseline-v3.md`*
*Cloud API: `https://thermopac-communication-thermopacllp.replit.app`*

---

## Step 1 — Admin Registers This Node

Before installing, a Thermopac admin must register this PC in the cloud app:

1. Log into the Thermopac ERP as Superuser
2. Navigate to **EPC Drawing Controls → Agent Nodes**
3. Click **Register New Node**
4. Enter:
   - **Node ID** — a short unique name for this PC (e.g. `PC-DESIGN-01`)
   - **Label** — friendly name (e.g. `Design Office PC – Prasad`)
5. The system generates a **Node Token** — copy it immediately (shown only once)

---

## Step 2 — Install the Agent

1. Double-click `ThermopacAgent-Setup-v1.0.exe`
2. Accept the license agreement
3. Choose install location (default: `C:\Program Files\ThermopacAgent\`)
4. Choose whether to create a Desktop shortcut
5. Choose whether to auto-start on Windows login (recommended for shared design PCs)
6. Click **Install**

The installer creates:
- Application files in the install folder
- Working temp folder: `C:\ThermopacAgent\temp\`
- Log folder: `C:\ThermopacAgent\logs\`

---

## Step 3 — Configure the Agent

Edit `config.ini` in the install folder (e.g. `C:\Program Files\ThermopacAgent\config.ini`):

```ini
[cloud]
api_url    = https://thermopac-communication-thermopacllp.replit.app
node_id    = PC-DESIGN-01          ; must match what admin registered
node_token = PASTE_TOKEN_HERE      ; token from Step 1 (keep secret)

[agent]
poll_interval_sec = 10
job_timeout_sec   = 600
max_retries       = 3

[paths]
temp_dir = C:\ThermopacAgent\temp
log_dir  = C:\ThermopacAgent\logs

[solidworks]
solidworks_version = 2024          ; change to match your installed version
; solidworks_progid =              ; leave blank (auto-resolved from version)
visible            = false
```

**Supported SolidWorks versions:**

| Version | ProgID (auto-resolved) |
|---------|----------------------|
| 2019    | SldWorks.Application.27 |
| 2020    | SldWorks.Application.28 |
| 2021    | SldWorks.Application.29 |
| 2022    | SldWorks.Application.30 |
| 2023    | SldWorks.Application.31 |
| 2024    | SldWorks.Application.32 |

---

## Step 4 — Start the Agent

Double-click the **ThermopacAgent** shortcut (Start Menu or Desktop).

A console window opens showing:
```
[Agent] Starting — api_url=https://... node_id=PC-DESIGN-01 ...
[Agent] Testing connection…
[Agent] Connection OK — entering poll loop
[Agent] No pending jobs
...
```

Leave this window open (minimise to taskbar). It will pick up jobs automatically.

**If auto-start was selected during install**, the agent starts automatically 30 seconds after Windows login — no manual action required.

---

## Step 5 — Verify

1. From the Thermopac ERP, open any EPC drawing record at `/epc/drawing-controls`
2. Upload a `.slddrw` file in the **DWG Attachments** card
3. Watch the agent console — within one poll interval (10s) you should see:
   ```
   [Agent] 1 pending job(s)
   [Runner] Job 42 start — file=C10308-CPS-ACS-S6T-20-P28.slddrw
   [Runner] Downloading…
   [Extractor] Launching SolidWorks (SldWorks.Application.32)…
   [Extractor] Document open
   [Properties] drawing_number='C10308-…' revision='B'
   [DesignData] Found 10 row(s)
   [Runner] Uploading extraction result…
   [Runner] Job 42 complete (145.2s)
   ```
4. In the Thermopac ERP, the **Drawing Verification** card updates with the extraction result and DDS comparison

---

## Troubleshooting

| Symptom | Check |
|---------|-------|
| `Authentication failed` | Verify `node_id` and `node_token` in config.ini match what admin registered |
| `SolidWorks not found` | Verify `solidworks_version` matches your installed version |
| `OpenDoc6 returned None` | Ensure SolidWorks is properly licensed; try `visible = true` to see errors |
| `DesignDataNotFoundError` | Drawing must contain a table with "Design Data" in the title |
| Job stuck at `processing` | Cloud auto-resets stale jobs after 30 min; check agent logs |
| Agent crashes on startup | Check `C:\ThermopacAgent\logs\agent.log` for the full error |

---

## Log Files

Logs are written to `C:\ThermopacAgent\logs\agent.log`
Daily rotation — previous days saved as `agent.log.YYYY-MM-DD`
30 days retained automatically.

---

## Uninstall

Use **Windows Settings → Apps** or **Control Panel → Programs** and uninstall **ThermopacAgent**.
The installer removes the scheduled task automatically.
Temp and log folders are **not** deleted (they may contain useful logs).

---

## Building from Source

If you need to rebuild the EXE:

```bat
pip install -r requirements.txt
build.bat
```

Then compile the installer:
```
iscc installer\setup.iss
```

Output: `installer_output\ThermopacAgent-Setup-v1.0.exe`

---

*Baseline design: `docs/slddrw-extraction-agent-baseline-v3.md`*
