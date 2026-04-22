"""
job_runner.py — Orchestrates one extraction job.

Flow:
  1. Download .slddrw from signed URL → verify SHA-256
  2. Run SolidWorks extraction in a worker thread with job-level timeout
  3. POST /complete or /fail
  4. Delete temp files (always, in finally)

The job-level timeout is enforced by joining the worker thread with a timeout.
On timeout a cancellation Event is set; the extractor checks it between modules.
SolidWorks close + ExitApp always runs in the extractor's finally block.
"""

from __future__ import annotations
import hashlib
import os
import shutil
import threading
import time
from datetime import datetime, timezone

import requests

from agent.job_client import JobClient, AGENT_VERSION
from extractor.solidworks_extractor import run_extraction


def run_job(job: dict, client: JobClient, config, logger) -> None:
    job_id   = job["id"]
    filename = job.get("filename", f"drawing_{job_id}.slddrw")
    logger.info(f"[Runner] ══ Job {job_id} start ══ file={filename}")

    temp_dir = os.path.join(config.temp_dir, f"job_{job_id}")
    os.makedirs(temp_dir, exist_ok=True)

    try:
        # ── 1. Claim ─────────────────────────────────────────────────────────
        import socket
        claim = client.claim_job(job_id, AGENT_VERSION, socket.gethostname())
        if claim is None:
            logger.info(f"[Runner] Job {job_id} claim failed — skipping")
            return

        download_url    = claim["download_url"]
        expected_sha256 = claim.get("sha256", "")
        filename        = claim.get("filename", filename)

        # ── 2. Download ───────────────────────────────────────────────────────
        temp_path = os.path.join(temp_dir, filename)
        logger.info(f"[Runner] Downloading {filename}…")
        _download(download_url, temp_path, logger)

        # ── 3. Verify SHA-256 ─────────────────────────────────────────────────
        actual_sha256 = _sha256(temp_path)
        logger.info(f"[Runner] SHA-256 actual={actual_sha256[:16]}… "
                    f"expected={expected_sha256[:16]}…")
        if expected_sha256 and actual_sha256 != expected_sha256.lower():
            reason = f"SHA-256 mismatch (expected={expected_sha256}, got={actual_sha256})"
            logger.error(f"[Runner] {reason}")
            client.fail_job(job_id, reason)
            return

        # ── 4. Extract (job-level timeout) ────────────────────────────────────
        cancel_event = threading.Event()
        result_holder: dict = {}
        node_id = config.node_id

        def worker():
            try:
                data = run_extraction(temp_path, config, cancel_event, logger)
                result_holder["ok"] = data
            except Exception as ex:
                result_holder["error"] = str(ex)

        t = threading.Thread(target=worker, daemon=True)
        t_start = time.monotonic()
        t.start()
        t.join(timeout=config.job_timeout_sec)
        elapsed = time.monotonic() - t_start

        if t.is_alive():
            cancel_event.set()
            logger.error(f"[Runner] Job {job_id} timed out after {elapsed:.0f}s — "
                         f"cancellation event set, waiting 30s for SW cleanup…")
            t.join(timeout=30)
            client.fail_job(job_id,
                            f"Job timeout exceeded after {config.job_timeout_sec}s")
            return

        if "error" in result_holder:
            reason = result_holder["error"]
            logger.error(f"[Runner] Extraction failed: {reason}")
            client.fail_job(job_id, reason)
            return

        extraction_result = result_holder["ok"]

        # Stamp agent metadata
        extraction_result["agent"] = {
            "node_id":              node_id,
            "agent_version":        AGENT_VERSION,
            "machine_name":         _machine_name(),
            "extraction_timestamp": datetime.now(timezone.utc).isoformat(),
        }

        # ── 5. Upload ─────────────────────────────────────────────────────────
        _debug_payload(extraction_result, job_id, logger)
        logger.info(f"[Runner] Uploading extraction result for job {job_id}…")
        success = client.complete_job(job_id, extraction_result)
        if not success:
            client.fail_job(job_id, "Server rejected extraction JSON (Zod validation)")
            return

        logger.info(f"[Runner] ══ Job {job_id} complete ({elapsed:.1f}s) ══")

    except Exception as e:
        logger.error(f"[Runner] Unhandled error for job {job_id}: {e}", exc_info=True)
        try:
            client.fail_job(job_id, f"Unhandled agent error: {e}")
        except Exception:
            pass
    finally:
        # Always remove temp dir
        try:
            shutil.rmtree(temp_dir, ignore_errors=True)
            logger.debug(f"[Runner] Temp dir removed: {temp_dir}")
        except Exception as ex:
            logger.warning(f"[Runner] Temp cleanup failed: {ex}")


def _download(url: str, dest: str, logger) -> None:
    with requests.get(url, stream=True, timeout=120) as r:
        r.raise_for_status()
        with open(dest, "wb") as f:
            for chunk in r.iter_content(chunk_size=8192):
                f.write(chunk)
    size = os.path.getsize(dest)
    logger.info(f"[Runner] Download complete: {size:,} bytes → {dest}")


def _sha256(path: str) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for block in iter(lambda: f.read(65536), b""):
            h.update(block)
    return h.hexdigest()


def _machine_name() -> str:
    import socket
    return socket.gethostname()


def _debug_payload(payload: dict, job_id: int, logger) -> None:
    """Log a one-line summary of the final payload before upload."""
    keys   = list(payload.keys())
    cp_val = payload.get("customProperties")
    cpv_val = payload.get("customPropertyVerification")

    def _status(v):
        if isinstance(v, dict):
            return f"object(keys={list(v.keys())})"
        return "null" if v is None else type(v).__name__

    cp_fields = len(cp_val.get("fields", [])) if isinstance(cp_val, dict) else "n/a"
    cpv_status = (cp_val or {}).get("status", "n/a") if isinstance(cpv_val, dict) else "n/a"
    logger.info(
        f"[Runner] Pre-upload payload job {job_id}: "
        f"keys={keys} | "
        f"customProperties={_status(cp_val)} fields={cp_fields} | "
        f"customPropertyVerification={_status(cpv_val)} status={cpv_status}"
    )
