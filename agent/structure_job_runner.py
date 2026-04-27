"""
structure_job_runner.py — Orchestrates one drawing structuring job.

Flow:
  1. Claim job (marks it IN_PROGRESS on the server)
  2. Run run_structuring() in a worker thread with job-level timeout
  3. POST /complete or /fail with result JSON
  4. Return to idle poll loop

Key differences from extraction job_runner.py:
  - No file download (the agent creates the file, not receives it)
  - No SHA-256 verification
  - Pre-flight errors are caught here and reported via fail_job()
"""

from __future__ import annotations
import socket
import threading
import time
from datetime import datetime, timezone

from agent.structure_job_client import StructureJobClient, STRUCTURER_VERSION
from structurer.solidworks_structurer import run_structuring, PreflightError


def run_structure_job(
    job: dict,
    client: StructureJobClient,
    config,
    logger,
) -> None:
    job_id = job["id"]
    logger.info(
        f"[StructRunner] ══ Job {job_id} start ══ "
        f"drawing={job.get('drawing_number')} rev={job.get('revision')} "
        f"mode={job.get('mode', 'create_new')}"
    )

    try:
        # ── 1. Claim ──────────────────────────────────────────────────────────
        claim = client.claim_job(job_id)
        if claim is None:
            logger.info(f"[StructRunner] Job {job_id} claim failed — skipping")
            return

        # Server may return an enriched job payload on claim; merge it in
        if isinstance(claim, dict):
            job = {**job, **claim}

        # ── 2. Run structuring (with timeout) ─────────────────────────────────
        cancel_event  = threading.Event()
        result_holder: dict = {}

        def worker():
            try:
                data = run_structuring(job, config, cancel_event, logger)
                result_holder["ok"] = data
            except PreflightError as e:
                result_holder["preflight_error"] = str(e)
            except Exception as e:
                result_holder["error"] = str(e)

        t_start = time.monotonic()
        t       = threading.Thread(target=worker, daemon=True)
        t.start()
        t.join(timeout=config.job_timeout_sec)
        elapsed = time.monotonic() - t_start

        if t.is_alive():
            cancel_event.set()
            logger.error(
                f"[StructRunner] Job {job_id} timed out after {elapsed:.0f}s — "
                "cancellation event set, waiting 30s for SW cleanup…"
            )
            t.join(timeout=30)
            client.fail_job(
                job_id,
                f"Job timeout exceeded after {config.job_timeout_sec}s",
            )
            return

        # ── Pre-flight failure — clean fail, no retry ─────────────────────────
        if "preflight_error" in result_holder:
            reason = result_holder["preflight_error"]
            logger.error(f"[StructRunner] Pre-flight failed: {reason}")
            client.fail_job(job_id, f"Pre-flight: {reason}")
            return

        # ── Extraction / structuring error ────────────────────────────────────
        if "error" in result_holder:
            reason = result_holder["error"]
            logger.error(f"[StructRunner] Structuring failed: {reason}")
            client.fail_job(job_id, reason)
            return

        structure_result = result_holder["ok"]

        # ── Stamp agent metadata ──────────────────────────────────────────────
        structure_result["agent"] = {
            "node_id":              config.node_id,
            "agent_version":        STRUCTURER_VERSION,
            "machine_name":         socket.gethostname(),
            "structure_timestamp":  datetime.now(timezone.utc).isoformat(),
        }

        # ── 3. Upload result ──────────────────────────────────────────────────
        logger.info(f"[StructRunner] Uploading result for job {job_id}…")
        success = client.complete_job(job_id, structure_result)
        if not success:
            client.fail_job(job_id, "Server rejected structuring result (validation error)")
            return

        logger.info(f"[StructRunner] ══ Job {job_id} complete ({elapsed:.1f}s) ══")

    except Exception as e:
        logger.error(
            f"[StructRunner] Unhandled error for job {job_id}: {e}", exc_info=True
        )
        try:
            client.fail_job(job_id, f"Unhandled agent error: {e}")
        except Exception:
            pass
