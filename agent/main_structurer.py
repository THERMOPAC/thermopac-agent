"""
main_structurer.py — Drawing Structuring Agent entry point.

Runtime sequence:
  PC boots → Windows Scheduler starts this process at login
  → load config.ini
  → connect to API
  → enter idle poll loop (no SolidWorks launched here)
  → job arrives → claim → run_structuring() → ExitApp → idle

CLI:
  --test      Verify config + connection then exit (0=ok 1=fail)
  --api-url   Override api_url from config.ini
  --node-id   Override node_id from config.ini
  --node-token Override node_token from config.ini
"""

from __future__ import annotations
import argparse
import os
import signal
import socket
import sys
import time

for _stream in (sys.stdout, sys.stderr):
    try:
        if hasattr(_stream, "reconfigure"):
            _stream.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

if getattr(sys, "frozen", False):
    _base = os.path.dirname(sys.executable)
else:
    _base = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

sys.path.insert(0, _base)

from agent.config                import AgentConfig
from agent.logger                import build_logger
from agent.structure_job_client  import StructureJobClient, STRUCTURER_VERSION
from agent.structure_job_runner  import run_structure_job

BANNER = r"""
  _____ _
 |_   _| |__   ___ _ __ _ __ ___   ___  _ __   __ _  ___
   | | | '_ \ / _ \ '__| '_ ` _ \ / _ \| '_ \ / _` |/ __|
   | | | | | |  __/ |  | | | | | | (_) | |_) | (_| | (__
   |_| |_| |_|\___|_|  |_| |_| |_|\___/| .__/ \__,_|\___|
                                        |_|
  Drawing Structuring Agent  v{version}
  THERMOPAC ERP Integration
"""

_shutdown = False


def _handle_signal(signum, frame):
    global _shutdown
    print("\n[Structurer] Shutdown signal received — finishing current job then stopping…")
    _shutdown = True


def _parse_args():
    p = argparse.ArgumentParser(
        prog="ThermopacStructurer",
        description="ThermopacStructurer — SolidWorks drawing creation agent for THERMOPAC ERP",
    )
    p.add_argument("config",       nargs="?", help="Path to config.ini (optional)")
    p.add_argument("--test",       action="store_true", help="Test config + connection then exit")
    p.add_argument("--api-url",    default="", help="Override api_url from config.ini")
    p.add_argument("--node-id",    default="", help="Override node_id from config.ini")
    p.add_argument("--node-token", default="", help="Override node_token from config.ini")
    return p.parse_args()


def _auto_register(config, logger) -> bool:
    import requests as _req
    url  = f"{config.api_url}/api/epc-agent-nodes/auto-register"
    body = {
        "node_id":       config.node_id,
        "node_token":    config.node_token,
        "machine_name":  socket.gethostname(),
        "agent_version": STRUCTURER_VERSION,
        "agent_type":    "structurer",
    }
    logger.info(f"[Structurer] Auto-registering node '{config.node_id}'…")
    try:
        r = _req.post(url, json=body, timeout=15)
        if r.status_code == 200:
            logger.info(f"[Structurer] Node '{config.node_id}' registered")
            return True
        logger.error(f"[Structurer] Auto-registration failed: HTTP {r.status_code} — {r.text[:300]}")
        return False
    except Exception as e:
        logger.error(f"[Structurer] Auto-registration request failed: {e}")
        return False


def _print_startup_config(config) -> None:
    print("-" * 62)
    print(f"  api_url       : {config.api_url}")
    print(f"  node_id       : {config.node_id}")
    print(f"  solidworks    : {config.sw_progid or 'NOT DETECTED'}")
    print(f"  template_path : {config.structurer_template_path or 'NOT SET'}")
    print(f"  staging_root  : {config.structurer_staging_root or 'NOT SET'}")
    print("-" * 62)
    print()


def main():
    global _shutdown

    args = _parse_args()
    print(BANNER.format(version=STRUCTURER_VERSION))

    config = AgentConfig(args.config)

    if args.api_url:
        config.api_url    = args.api_url.rstrip("/")
    if args.node_id:
        config.node_id    = args.node_id
    if args.node_token:
        config.node_token = args.node_token

    _print_startup_config(config)

    logger = build_logger(config.log_dir, name="thermopac_structurer")
    logger.info(f"[Structurer] Starting — {config.summary()}")
    logger.info(f"[Structurer] Agent version: {STRUCTURER_VERSION}")
    logger.info(f"[Structurer] template_path: {config.structurer_template_path}")
    logger.info(f"[Structurer] staging_root:  {config.structurer_staging_root}")

    client = StructureJobClient(
        config.api_url, config.node_id, config.node_token, logger
    )

    if config.mode == "testing":
        if not _auto_register(config, logger):
            sys.exit(1)

    if args.test:
        ok = client.test_connection()
        print(f"[Structurer] Connection test: {'PASS' if ok else 'FAIL'}")
        sys.exit(0 if ok else 1)

    signal.signal(signal.SIGINT,  _handle_signal)
    signal.signal(signal.SIGTERM, _handle_signal)

    logger.info(f"[Structurer] Testing connection to {config.api_url}…")
    if not client.test_connection():
        logger.error("[Structurer] Connection test failed — check config.ini and network")
        sys.exit(1)
    logger.info("[Structurer] Connection OK — entering idle poll loop")
    logger.info(
        f"[Structurer] Poll interval: {config.poll_interval_sec}s | "
        f"Job timeout: {config.job_timeout_sec}s"
    )

    # ── Idle poll loop (no SolidWorks at this level) ──────────────────────────
    while not _shutdown:
        try:
            jobs = client.get_pending_jobs()

            if not jobs:
                logger.debug("[Structurer] No pending jobs")
            else:
                logger.info(f"[Structurer] {len(jobs)} pending job(s) — processing first")
                run_structure_job(jobs[0], client, config, logger)

        except Exception as e:
            logger.error(f"[Structurer] Poll loop error: {e}", exc_info=True)

        if _shutdown:
            break

        for _ in range(config.poll_interval_sec * 2):
            if _shutdown:
                break
            time.sleep(0.5)

    logger.info("[Structurer] Shutdown complete")


if __name__ == "__main__":
    main()
