"""
config.py — Load and validate config.ini.
Resolves SolidWorks version to COM ProgID.
"""

import configparser
import os
import sys

SW_VERSION_PROGID = {
    2019: "SldWorks.Application.27",
    2020: "SldWorks.Application.28",
    2021: "SldWorks.Application.29",
    2022: "SldWorks.Application.30",
    2023: "SldWorks.Application.31",
    2024: "SldWorks.Application.32",
}


class AgentConfig:
    def __init__(self, path: str = None):
        if path is None:
            path = self._default_path()
        if not os.path.exists(path):
            print(f"[CONFIG] ERROR: config.ini not found at {path}")
            sys.exit(1)

        cfg = configparser.ConfigParser()
        cfg.read(path, encoding="utf-8")

        # ── Cloud ─────────────────────────────────────────────────────────────
        self.api_url    = self._require(cfg, "cloud", "api_url").rstrip("/")
        self.node_id    = self._require(cfg, "cloud", "node_id").strip()
        self.node_token = self._require(cfg, "cloud", "node_token").strip()

        # ── Agent ─────────────────────────────────────────────────────────────
        self.poll_interval_sec = cfg.getint("agent", "poll_interval_sec", fallback=10)
        self.job_timeout_sec   = cfg.getint("agent", "job_timeout_sec",   fallback=600)
        self.max_retries       = cfg.getint("agent", "max_retries",       fallback=3)

        # ── Paths ─────────────────────────────────────────────────────────────
        self.temp_dir = cfg.get("paths", "temp_dir",
                                fallback=r"C:\ThermopacAgent\temp")
        self.log_dir  = cfg.get("paths", "log_dir",
                                fallback=r"C:\ThermopacAgent\logs")

        os.makedirs(self.temp_dir, exist_ok=True)
        os.makedirs(self.log_dir,  exist_ok=True)

        # ── SolidWorks ────────────────────────────────────────────────────────
        self.sw_visible = cfg.getboolean("solidworks", "visible", fallback=False)

        explicit_progid = cfg.get("solidworks", "solidworks_progid", fallback="").strip()
        if explicit_progid:
            self.sw_progid = explicit_progid
        else:
            ver = cfg.getint("solidworks", "solidworks_version", fallback=0)
            if ver not in SW_VERSION_PROGID:
                print(f"[CONFIG] ERROR: solidworks_version={ver} not supported. "
                      f"Supported: {sorted(SW_VERSION_PROGID.keys())}")
                sys.exit(1)
            self.sw_progid = SW_VERSION_PROGID[ver]

        self._config_path = path

    def summary(self) -> str:
        return (
            f"api_url={self.api_url} | node_id={self.node_id} | "
            f"sw_progid={self.sw_progid} | poll={self.poll_interval_sec}s | "
            f"timeout={self.job_timeout_sec}s"
        )

    @staticmethod
    def _require(cfg: configparser.ConfigParser, section: str, key: str) -> str:
        if not cfg.has_option(section, key):
            print(f"[CONFIG] ERROR: Missing required config: [{section}] {key}")
            sys.exit(1)
        val = cfg.get(section, key).strip()
        if not val:
            print(f"[CONFIG] ERROR: [{section}] {key} must not be empty")
            sys.exit(1)
        return val

    @staticmethod
    def _default_path() -> str:
        exe_dir = os.path.dirname(sys.executable if getattr(sys, "frozen", False)
                                  else os.path.abspath(__file__))
        candidates = [
            os.path.join(exe_dir, "config.ini"),
            os.path.join(exe_dir, "..", "config.ini"),
            r"C:\ThermopacAgent\config.ini",
        ]
        for p in candidates:
            if os.path.exists(p):
                return p
        return candidates[0]
