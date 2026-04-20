"""
job_client.py — HTTP client for all cloud API calls.
All requests carry x-node-id + x-node-token headers.
"""

from __future__ import annotations
import json
from typing import Optional
import requests

AGENT_VERSION = "1.0.16"


class JobClient:
    def __init__(self, api_url: str, node_id: str, node_token: str, logger):
        self._base   = api_url.rstrip("/")
        self._logger = logger
        self._headers = {
            "x-node-id":    node_id,
            "x-node-token": node_token,
            "Content-Type": "application/json",
            "User-Agent":   f"ThermopacAgent/{AGENT_VERSION}",
        }

    # ── Poll ──────────────────────────────────────────────────────────────────

    def get_pending_jobs(self) -> list[dict]:
        """GET /api/epc-slddrw-jobs/pending — returns list of pending job objects."""
        try:
            r = self._get("/api/epc-slddrw-jobs/pending", timeout=15)
            return r.get("jobs", [])
        except Exception as e:
            self._logger.warning(f"[Client] poll error: {e}")
            return []

    # ── Claim ─────────────────────────────────────────────────────────────────

    def claim_job(self, job_id: int, agent_version: str,
                  machine_name: str) -> Optional[dict]:
        """
        POST /api/epc-slddrw-jobs/:id/claim
        Returns dict with { download_url, filename, sha256 } on success.
        Returns None on 409 (race lost) or error.
        """
        body = {"agent_version": agent_version, "machine_name": machine_name}
        try:
            r = self._post(f"/api/epc-slddrw-jobs/{job_id}/claim", body, timeout=20)
            return r
        except _ConflictError:
            self._logger.info(f"[Client] Job {job_id} already claimed (409)")
            return None
        except Exception as e:
            self._logger.warning(f"[Client] claim error job {job_id}: {e}")
            return None

    # ── Complete ──────────────────────────────────────────────────────────────

    def complete_job(self, job_id: int, extraction_result: dict) -> bool:
        """
        POST /api/epc-slddrw-jobs/:id/complete
        Returns True on success, False on validation error or network error.
        """
        try:
            self._post(f"/api/epc-slddrw-jobs/{job_id}/complete",
                       {"extraction_result": extraction_result}, timeout=30)
            return True
        except _ValidationError as e:
            self._logger.error(f"[Client] complete rejected (422): {e}")
            return False
        except Exception as e:
            self._logger.error(f"[Client] complete error job {job_id}: {e}")
            return False

    # ── Fail ──────────────────────────────────────────────────────────────────

    def fail_job(self, job_id: int, reason: str) -> None:
        """POST /api/epc-slddrw-jobs/:id/fail — best effort."""
        try:
            self._post(f"/api/epc-slddrw-jobs/{job_id}/fail",
                       {"reason": reason[:1000]}, timeout=15)
        except Exception as e:
            self._logger.warning(f"[Client] fail report error job {job_id}: {e}")

    # ── Ping / auth test ──────────────────────────────────────────────────────

    def test_connection(self) -> bool:
        """GET /api/epc-slddrw-jobs/pending to verify auth. Returns True if ok."""
        try:
            self._get("/api/epc-slddrw-jobs/pending", timeout=10)
            return True
        except _AuthError:
            self._logger.error("[Client] Authentication failed — check node_id and node_token")
            return False
        except Exception as e:
            self._logger.error(f"[Client] Connection test failed: {e}")
            return False

    # ── Internal ──────────────────────────────────────────────────────────────

    def _get(self, path: str, timeout: int = 15) -> dict:
        url = self._base + path
        r = requests.get(url, headers=self._headers, timeout=timeout)
        return self._handle(r)

    def _post(self, path: str, body: dict, timeout: int = 30) -> dict:
        url = self._base + path
        r = requests.post(url, headers=self._headers,
                          data=json.dumps(body, ensure_ascii=False), timeout=timeout)
        return self._handle(r)

    @staticmethod
    def _handle(r: requests.Response) -> dict:
        if r.status_code == 401:
            raise _AuthError(f"401 Unauthorized from {r.url}")
        if r.status_code == 409:
            raise _ConflictError("409 Conflict")
        if r.status_code == 422:
            raise _ValidationError(r.text[:500])
        r.raise_for_status()
        if r.content:
            return r.json()
        return {}


class _AuthError(Exception):
    pass

class _ConflictError(Exception):
    pass

class _ValidationError(Exception):
    pass
