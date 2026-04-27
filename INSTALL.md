# Thermopac Drawing Structuring Agent — Installation Guide
### THERMOPAC ERP | SolidWorks WRITE Agent | v1.0.1 | Phase 1

---

> **Role:** This agent CREATES and UPDATES SolidWorks `.slddrw` files from DDS data.
> It is a **WRITE** agent — it writes to the filesystem.
> The **Extraction Agent** (separate package) is the complementary READ agent.

---

## Prerequisites

| Requirement | Details |
|---|---|
| Windows | 10 or 11, 64-bit |
| Python | 3.9 or higher — [python.org/downloads](https://python.org/downloads) — check "Add Python to PATH" |
| SolidWorks | 2019, 2020, 2021, 2022, 2023, or 2024 — licensed and working |
| Template file | A SolidWorks `.drwdot` drawing template (your standard drawing template) |
| Staging folder | A writable folder for staged `.slddrw` output (local or UNC path) |
| Internet | HTTPS access to `thermopac-communication-thermopacllp.replit.app` (port 443) |

---

## Step 1 — Admin Registers This Node

A Thermopac **Superuser** must create credentials for this PC before the agent can authenticate.

1. Log in to the Thermopac ERP
2. Navigate to **EPC → Drawing Controls → Agent Nodes**
3. Click **Register New Node**
4. Enter:
   - **Node ID** — a short unique name for this PC, e.g. `PC-STRUCTURER-01`
   - **Agent Type** — select `structurer`
5. The system shows a **Node Token** — **copy it immediately** (displayed once only)

You now have:
- `node_id` — e.g. `PC-STRUCTURER-01`
- `node_token` — the secret token string

---

## Step 2 — Edit config.ini

Open `config.ini` in this folder with Notepad and fill in your values:

```ini
[cloud]
api_url    = https://thermopac-communication-thermopacllp.replit.app
node_id    = PC-STRUCTURER-01        ; ← your node ID from Step 1
node_token = REPLACE_WITH_YOUR_TOKEN ; ← paste your token from Step 1

[solidworks]
solidworks_version = 2019            ; ← your installed SolidWorks version

[structurer]
template_path = C:\SolidWorks Templates\Standard_A1.drwdot  ; ← REQUIRED
staging_root  = C:\ThermopacStaging\drawings                 ; ← writable output folder
```

### SolidWorks version numbers

| Installed version | Set `solidworks_version` to |
|---|---|
| SolidWorks 2019 | `2019` |
| SolidWorks 2020 | `2020` |
| SolidWorks 2021 | `2021` |
| SolidWorks 2022 | `2022` |
| SolidWorks 2023 | `2023` |
| SolidWorks 2024 | `2024` |

### Structurer settings

| Setting | Description |
|---|---|
| `template_path` | Full path to your standard SolidWorks drawing template (`.drwdot`). **Required** — agent fails pre-flight if empty or missing. |
| `staging_root` | Root folder for staged output. Subdirectories are created automatically per `drawing_control_id`. Default: `C:\ThermopacStaging\drawings`. |

---

## Step 3 — Run bootstrap.bat

**Right-click** `bootstrap.bat` → **Run as administrator**

The script will:
1. Locate your Python installation
2. Create an isolated Python virtual environment (`venv\`)
3. Install `pywin32` and `requests`
4. Register pywin32 COM DLLs with Windows (required for SolidWorks COM access)
5. Create `start_structurer.bat` — double-click to start the agent
6. Create `test.bat` — double-click to run a connection self-test
7. Create a **Desktop shortcut** called `ThermopacStructurer`
8. Run the self-test immediately

Takes approximately **3–5 minutes** on first run.

---

## Step 4 — Verify Self-Test

After bootstrap, expected test output (before editing config.ini):

```
[TEST] ✅ Config loaded
[TEST] ⚠  Authentication failed — edit config.ini with node_id and node_token
```

After editing config.ini correctly:

```
[TEST] ✅ Config loaded and validated
[TEST] ✅ Network reachability — TCP 443 → thermopac-communication-thermopacllp.replit.app
[TEST] ✅ Cloud authentication (x-node-id + x-node-token)
[TEST] ✅ Poll pending structure jobs — 0 pending
[TEST] ✅ win32com (pywin32) importable
[TEST] ✅ SolidWorks ProgID registered (SldWorks.Application.27)

Overall: PASS ✅
```

---

## Step 5 — Start the Agent

Double-click the **ThermopacStructurer** shortcut on the Desktop
(or double-click `start_structurer.bat`).

A console window opens:

```
  _____ _
 |_   _| |__   ___ _ __ _ __ ___   ___  _ __   __ _  ___
   | | | '_ \ / _ \ '__| '_ ` _ \ / _ \| '_ \ / _` |/ __|
   | | | | | |  __/ |  | | | | | | (_) | |_) | (_| | (__
   |_| |_| |_|\___|_|  |_| |_| |_|\___/| .__/ \__,_|\___|
                                        |_|
  Drawing Structuring Agent  v1.0.1
  THERMOPAC ERP Integration

  api_url       : https://thermopac-communication-thermopacllp.replit.app
  node_id       : PC-STRUCTURER-01
  solidworks    : SldWorks.Application.27
  template_path : C:\SolidWorks Templates\Standard_A1.drwdot
  staging_root  : C:\ThermopacStaging\drawings

[Structurer] Connection OK — entering idle poll loop
[Structurer] No pending jobs
[Structurer] No pending jobs
...
```

Leave this window open (minimise to taskbar). It polls every 10 seconds automatically.

---

## How a Structure Job Runs

When the ERP dispatches a drawing structure job to this node:

```
[Structurer] 1 pending job(s) — processing first
[StructRunner] ══ Job 42 start ══ drawing=T-PRJ-001-GA-001 rev=A mode=create_new
[Structurer] Pre-flight passed — staging_path=C:\ThermopacStaging\drawings\DC-2024-0042\T-PRJ-001-GA-001_rev-A.slddrw
[COM] Launching dedicated SolidWorks instance: SldWorks.Application.27
[COM] Instance created via DispatchEx(SldWorks.Application.27)
[Structurer] SolidWorks hidden instance ready
[Structurer] NewDocument from template: C:\SolidWorks Templates\Standard_A1.drwdot
[Structurer] New drawing document created
[Structurer] Writing custom properties…
[Structurer] Property written: Drawing_Number = 'T-PRJ-001-GA-001'
[Structurer] Property written: Revision = 'A'
[Structurer] Property written: Tag_No = 'HX-101'
...
[Structurer] SaveAs3 → C:\ThermopacStaging\drawings\DC-2024-0042\T-PRJ-001-GA-001_rev-A.slddrw
[Structurer] Save successful
[COM] Document closed
[COM] Dedicated SolidWorks instance exited cleanly
[StructRunner] ══ Job 42 complete (12.3s) ══
```

---

## Phase 1 Scope — What This Agent Does

| Action | Phase 1 |
|---|---|
| Create new `.slddrw` from `.drwdot` template | ✅ |
| Update existing `.slddrw` custom properties | ✅ |
| Write DDS custom properties | ✅ |
| Title block population via `$PRP:` / `$PRPSHEET:` (template-side) | ✅ |
| DDS table insertion | ❌ Phase 3 |
| Heuristic title block note injection | ❌ Phase 2 |
| PDF generation | ❌ Out of scope |
| GCS upload (final release path) | ❌ Out of scope |

---

## Staging Path Format

Output files are written to:
```
{staging_root}\{drawingControlId}\{DrawingNumber}_rev-{Revision}.slddrw
```

Example:
```
C:\ThermopacStaging\drawings\DC-2024-0042\T-PRJ-001-GA-001_rev-A.slddrw
```

---

## Troubleshooting

| Symptom | Action |
|---|---|
| `Authentication failed` | Check `node_id` and `node_token` in config.ini match the admin-registered values |
| `template_path not found` | Verify the `.drwdot` file exists at the configured path — check UNC share access |
| `staging_root not writable` | Ensure the PC has write permission to the staging folder |
| `mode=create_new but file already exists` | A file for this drawing+revision already exists — delete it or use `update_existing` mode |
| `SolidWorks ProgID not found` | Set `solidworks_version` to match your installed SolidWorks version |
| `NewDocument returned None` | SolidWorks licence issue — check the licence is active |
| `OpenDoc7 returned None` | File may be open/locked by another process |
| Job stuck at `in_progress` | Cloud auto-resets stale jobs after 30 min — check `agent.log` for the error |
| Agent crashes on startup | Check `C:\ThermopacStructurer\logs\agent.log` |

---

## Log Files

| File | Contents |
|---|---|
| `C:\ThermopacStructurer\logs\agent.log` | Main structurer log (daily rotation, 30 days retained) |

---

## Files in This Package

| File / Folder | Purpose |
|---|---|
| `bootstrap.bat` | **Start here** — sets up Python environment and launchers |
| `config.ini` | Configuration — edit with your credentials, SW version, template path |
| `start_structurer.bat` | Created by bootstrap — starts the polling agent |
| `test.bat` | Created by bootstrap — runs connection self-test |
| `agent/` | Core agent modules (config, logger, job client, runner, main) |
| `extractor/` | Shared COM helpers (sw_instance.py — SolidWorks process management) |
| `structurer/` | Phase 1 SolidWorks WRITE logic (solidworks_structurer.py) |
| `requirements.txt` | Python dependency list |
| `INSTALL.md` | This file |
| `BUILD.md` | How to build the compiled EXE installer |

---

*Baseline: `docs/drawing-structuring-agent-baseline.md`*
*Cloud API: `https://thermopac-communication-thermopacllp.replit.app`*
*Extraction Agent: separate package (Thermopac Extraction Agent v1.0.70)*
