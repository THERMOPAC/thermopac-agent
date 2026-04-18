"""
config.py — Load, auto-create, and validate config.ini.

MODES
─────
testing    (default)
  • If node_token is missing/placeholder, a random token is auto-generated
    and saved into config.ini.
  • Agent then calls POST /api/epc-agent-nodes/auto-register so the cloud
    accepts the self-issued token.  The cloud endpoint must have
    AGENT_AUTO_REGISTER=true set in its environment.

production
  • node_token must be cloud/admin-issued.  Missing token → hard exit with
    step-by-step instructions.  No token is ever auto-generated.

OTHER AUTO-FILLS (both modes)
  api_url             → http://localhost:3000  (edit for production)
  node_id             → socket.gethostname()   (Windows machine name)
  solidworks_version  → highest version found in Windows registry (if set to 0)
"""

from __future__ import annotations
import configparser
import os
import secrets
import socket
import sys

SW_VERSION_PROGID = {
    2019: "SldWorks.Application.27",
    2020: "SldWorks.Application.28",
    2021: "SldWorks.Application.29",
    2022: "SldWorks.Application.30",
    2023: "SldWorks.Application.31",
    2024: "SldWorks.Application.32",
}

_TOKEN_PLACEHOLDER = "REPLACE_WITH_YOUR_TOKEN"
_DEFAULT_API_URL   = "https://thermopac-communication-thermopacllp.replit.app"


class AgentConfig:
    def __init__(self, path: str = None):
        if path is None:
            path = self._default_path()

        # Auto-create config.ini if missing at the single canonical path
        if not os.path.exists(path):
            _create_default_config(path)

        print(f"[CONFIG] Loaded from: {path}")

        cfg = configparser.ConfigParser()
        cfg.read(path, encoding="utf-8")

        # Print api_url immediately after reading so it's visible before any other logic
        _early_api_url = (
            cfg.get("cloud", "api_url", fallback="").strip().rstrip("/")
            or _DEFAULT_API_URL
        )
        print(f"[CONFIG] api_url:      {_early_api_url}")

        # ── Mode ──────────────────────────────────────────────────────────────
        raw_mode = cfg.get("agent", "mode", fallback="testing").strip().lower()
        if raw_mode not in ("testing", "production"):
            print(f"[CONFIG] ERROR: [agent] mode must be 'testing' or 'production', got '{raw_mode}'")
            sys.exit(1)
        self.mode = raw_mode

        # ── Cloud ─────────────────────────────────────────────────────────────
        self.api_url = (
            cfg.get("cloud", "api_url", fallback="").strip().rstrip("/")
            or _DEFAULT_API_URL
        )

        self.node_id = (
            cfg.get("cloud", "node_id", fallback="").strip()
            or socket.gethostname()
        )

        raw_token = cfg.get("cloud", "node_token", fallback="").strip()
        self.token_auto_generated = False

        if not raw_token or raw_token == _TOKEN_PLACEHOLDER:
            if self.mode == "production":
                _abort_missing_token(path, self.node_id)
            else:
                # testing mode — generate, persist, flag for auto-registration
                raw_token = secrets.token_hex(32)
                _save_token(cfg, path, raw_token)
                self.token_auto_generated = True
                print(f"[CONFIG] Testing mode: auto-generated node_token and saved to config.ini")
                print(f"[CONFIG] Agent will self-register with the cloud on startup.")
                print()

        self.node_token = raw_token

        # ── Agent ─────────────────────────────────────────────────────────────
        self.poll_interval_sec = cfg.getint("agent", "poll_interval_sec", fallback=10)
        self.job_timeout_sec   = cfg.getint("agent", "job_timeout_sec",   fallback=600)
        self.max_retries       = cfg.getint("agent", "max_retries",       fallback=3)

        # ── Paths ─────────────────────────────────────────────────────────────
        self.temp_dir = cfg.get("paths", "temp_dir", fallback=r"C:\ThermopacAgent\temp")
        self.log_dir  = cfg.get("paths", "log_dir",  fallback=r"C:\ThermopacAgent\logs")
        os.makedirs(self.temp_dir, exist_ok=True)
        os.makedirs(self.log_dir,  exist_ok=True)

        # ── SolidWorks ────────────────────────────────────────────────────────
        self.sw_visible = cfg.getboolean("solidworks", "visible", fallback=False)

        explicit_progid = cfg.get("solidworks", "solidworks_progid", fallback="").strip()
        if explicit_progid:
            self.sw_progid       = explicit_progid
            self.sw_version      = 0
            self.sw_autodetected = False
        else:
            ver_str = cfg.get("solidworks", "solidworks_version", fallback="0").strip()
            ver     = int(ver_str) if ver_str.isdigit() else 0

            if ver and ver not in SW_VERSION_PROGID:
                print(f"[CONFIG] ERROR: solidworks_version={ver} is not supported.")
                print(f"[CONFIG]   Supported: {sorted(SW_VERSION_PROGID.keys())}")
                print(f"[CONFIG]   Edit [solidworks] solidworks_version in: {path}")
                sys.exit(1)

            if ver:
                self.sw_version      = ver
                self.sw_progid       = SW_VERSION_PROGID[ver]
                self.sw_autodetected = False
            else:
                detected = _detect_solidworks_version()
                self.sw_version      = detected
                self.sw_progid       = SW_VERSION_PROGID[detected] if detected else ""
                self.sw_autodetected = True

        self._config_path = path

    # ── Public ────────────────────────────────────────────────────────────────

    def summary(self) -> str:
        sw = (f"{self.sw_progid}" if self.sw_progid else "NOT DETECTED")
        if self.sw_autodetected and self.sw_progid:
            sw += " [auto-detected]"
        return (
            f"mode={self.mode} | api_url={self.api_url} | node_id={self.node_id} | "
            f"sw={sw} | poll={self.poll_interval_sec}s | timeout={self.job_timeout_sec}s"
        )

    @staticmethod
    def _default_path() -> str:
        """
        Single source of truth for config location.

        Frozen (installed EXE):
          → same folder as ThermopacAgent.exe
          → e.g. C:\\Program Files\\ThermopacAgent\\config.ini

        Source / dev:
          → project root (parent of the agent/ package)
        """
        if getattr(sys, "frozen", False):
            # sys.executable = C:\Program Files\ThermopacAgent\ThermopacAgent.exe
            base = os.path.dirname(sys.executable)
        else:
            # __file__ = <project>/agent/config.py  →  parent = <project>/
            base = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        return os.path.normpath(os.path.join(base, "config.ini"))


# ── Module-level helpers ──────────────────────────────────────────────────────

def _detect_solidworks_version() -> int:
    """Return highest installed SolidWorks version from Windows registry, or 0."""
    try:
        import winreg
        for ver in sorted(SW_VERSION_PROGID.keys(), reverse=True):
            try:
                winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, SW_VERSION_PROGID[ver])
                return ver
            except OSError:
                continue
    except ImportError:
        pass
    return 0


def _save_token(cfg: configparser.ConfigParser, path: str, token: str) -> None:
    """Write the generated token back into config.ini under [cloud] node_token."""
    if not cfg.has_section("cloud"):
        cfg.add_section("cloud")
    cfg.set("cloud", "node_token", token)
    with open(path, "w", encoding="utf-8") as f:
        cfg.write(f)


def _create_default_config(path: str) -> None:
    """Create a starter config.ini with auto-filled values."""
    machine_name = socket.gethostname()
    detected_ver = _detect_solidworks_version()
    sw_ver_str   = str(detected_ver) if detected_ver else "0"
    sw_comment   = (
        f"; Auto-detected SolidWorks {detected_ver}"
        if detected_ver
        else "; SolidWorks not detected — set manually (2019–2024)"
    )

    os.makedirs(os.path.dirname(path) or ".", exist_ok=True)
    content = f"""\
; ThermopacAgent configuration
; Auto-created on first run

[cloud]
api_url    = {_DEFAULT_API_URL}
node_id    = {machine_name}
node_token = {_TOKEN_PLACEHOLDER}

[agent]
; testing | production
; testing  — auto-generates token and self-registers with cloud
; production — requires cloud/admin-issued token, no auto-registration
mode = testing

poll_interval_sec = 10
job_timeout_sec   = 600
max_retries       = 3

[paths]
temp_dir = C:\\ThermopacAgent\\temp
log_dir  = C:\\ThermopacAgent\\logs

[solidworks]
{sw_comment}
solidworks_version = {sw_ver_str}
; solidworks_progid =
visible = false
"""
    with open(path, "w", encoding="utf-8") as f:
        f.write(content)

    print(f"[CONFIG] Created default config.ini at: {path}")
    print(f"[CONFIG]   node_id    = {machine_name}  (from machine name)")
    print(f"[CONFIG]   api_url    = {_DEFAULT_API_URL}  (default)")
    if detected_ver:
        print(f"[CONFIG]   solidworks = {detected_ver}  (auto-detected)")
    else:
        print(f"[CONFIG]   solidworks = not detected — set solidworks_version manually")
    print()


def _abort_missing_token(config_path: str, node_id: str) -> None:
    """Production mode: print clear instructions and exit."""
    print()
    print("=" * 62)
    print("  ERROR: node_token is not set  [production mode]")
    print("=" * 62)
    print()
    print("  Production mode requires a cloud/admin-issued token.")
    print("  Tokens cannot be auto-generated in production mode.")
    print()
    print("  Steps to get your token:")
    print("    1. Log in to the Thermopac ERP as Superuser")
    print("    2. Go to EPC -> Drawing Controls -> Agent Nodes")
    print(f"    3. Register this node  (suggested ID: {node_id})")
    print("    4. Copy the token — displayed ONCE only")
    print("    5. Paste it into config.ini under [cloud] node_token")
    print()
    print(f"  Config file: {config_path}")
    print()
    print("  To use testing mode instead, set [agent] mode = testing")
    print("=" * 62)
    sys.exit(1)
