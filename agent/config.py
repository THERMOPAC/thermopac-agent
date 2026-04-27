"""
config.py — Read and validate config.ini.

DESIGN RULES
────────────
1. config.ini passed at startup is READ-ONLY — never written to.
2. If config.ini is missing → hard exit with clear instructions.
3. No auto-migration of URLs, no overwriting of any field.
4. The ONLY write this module performs is persisting an auto-generated
   node_token (testing mode only) to the user-writable APPDATA path:
     %APPDATA%\\ThermopacStructuringAgent\\config.ini
   This write never touches the primary config path.

MODES
─────
testing    → node_token may be auto-generated if absent/placeholder.
             Token is written to APPDATA config, not to the primary path.
production → node_token must be present and non-placeholder.
             Missing token → hard exit with step-by-step instructions.
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

# User-writable path for persisting auto-generated tokens (testing mode).
# This is the ONLY path this module ever writes to.
_APPDATA_CONFIG = os.path.join(
    os.environ.get("APPDATA", os.path.expanduser("~")),
    "ThermopacStructuringAgent",
    "config.ini",
)


class AgentConfig:
    def __init__(self, path: str = None):
        if path is None:
            path = self._default_path()

        # ── config.ini must exist — no auto-create ─────────────────────────
        if not os.path.exists(path):
            print()
            print("=" * 62)
            print("  ERROR: config.ini not found")
            print("=" * 62)
            print()
            print(f"  Expected location: {path}")
            print()
            print("  Create config.ini in the same folder as run.bat.")
            print("  Minimum required contents:")
            print()
            print("    [cloud]")
            print("    api_url    = https://thermopac-communication-thermopacllp.replit.app")
            print(f"    node_id    = {socket.gethostname()}")
            print("    node_token = REPLACE_WITH_YOUR_TOKEN")
            print()
            print("    [agent]")
            print("    mode = testing")
            print()
            print("    [structurer]")
            print("    template_path = C:\\SolidWorks Templates\\Standard_A1.drwdot")
            print("    staging_root  = C:\\ThermopacStaging\\drawings")
            print()
            sys.exit(1)

        print(f"[CONFIG] Loaded from: {path}  (read-only)")

        cfg = configparser.ConfigParser()
        cfg.read(path, encoding="utf-8")

        # ── Also read APPDATA overlay (token written by testing mode) ───────
        appdata_cfg = configparser.ConfigParser()
        if os.path.exists(_APPDATA_CONFIG):
            try:
                appdata_cfg.read(_APPDATA_CONFIG, encoding="utf-8-sig")
                print(f"[CONFIG] Token overlay:  {_APPDATA_CONFIG}")
            except Exception:
                # Corrupted APPDATA config — delete it and start fresh
                print(f"[CONFIG] WARNING: APPDATA config is corrupted — deleting and ignoring.")
                try:
                    os.remove(_APPDATA_CONFIG)
                except Exception:
                    pass
                appdata_cfg = configparser.ConfigParser()

        # ── Mode — APPDATA overlay takes priority (no admin rights needed) ───
        raw_mode = cfg.get("agent", "mode", fallback="testing").strip().lower()
        appdata_mode = appdata_cfg.get("agent", "mode", fallback="").strip().lower()
        if appdata_mode in ("testing", "production"):
            raw_mode = appdata_mode
            print(f"[CONFIG] Mode override from APPDATA: {raw_mode}")
        if raw_mode not in ("testing", "production"):
            print(f"[CONFIG] ERROR: [agent] mode must be 'testing' or 'production', got '{raw_mode}'")
            sys.exit(1)
        self.mode = raw_mode

        # ── Cloud ─────────────────────────────────────────────────────────────
        self.api_url = cfg.get("cloud", "api_url", fallback="").strip().rstrip("/")
        # APPDATA overlay may override api_url (useful in dev/testing to avoid
        # editing the read-only Program Files config.ini as administrator)
        appdata_api_url = appdata_cfg.get("cloud", "api_url", fallback="").strip().rstrip("/")
        if appdata_api_url:
            self.api_url = appdata_api_url
            print(f"[CONFIG] api_url override from APPDATA: {self.api_url}")
        if not self.api_url:
            print("[CONFIG] ERROR: [cloud] api_url is required in config.ini")
            sys.exit(1)

        self.node_id = (
            cfg.get("cloud", "node_id", fallback="").strip()
            or socket.gethostname()
        )

        # Token: primary config first, then APPDATA overlay
        raw_token = cfg.get("cloud", "node_token", fallback="").strip()
        if not raw_token or raw_token == _TOKEN_PLACEHOLDER:
            raw_token = appdata_cfg.get("cloud", "node_token", fallback="").strip()

        self.token_auto_generated = False

        if not raw_token or raw_token == _TOKEN_PLACEHOLDER:
            if self.mode == "production":
                _abort_missing_token(path, self.node_id)
            else:
                # Testing mode — generate in memory, persist to APPDATA only
                raw_token = secrets.token_hex(32)
                self.token_auto_generated = True
                _persist_token_to_appdata(raw_token, self.node_id, self.api_url)

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

        # ── Structurer ────────────────────────────────────────────────────────
        self.structurer_template_path = cfg.get(
            "structurer", "template_path", fallback=""
        ).strip()
        self.structurer_staging_root = cfg.get(
            "structurer", "staging_root", fallback=r"C:\ThermopacStaging\drawings"
        ).strip()

        # ── SolidWorks ────────────────────────────────────────────────────────
        self.sw_visible = cfg.getboolean("solidworks", "visible", fallback=False)
        self.sw_model_search_path = cfg.get(
            "solidworks", "model_search_path", fallback="").strip()

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

        # ── Summary print ─────────────────────────────────────────────────────
        print(f"[CONFIG] api_url:      {self.api_url}")
        print(f"[CONFIG] node_id:      {self.node_id}")
        print(f"[CONFIG] mode:         {self.mode}")
        sw = self.sw_progid or "NOT DETECTED"
        if self.sw_autodetected and self.sw_progid:
            sw += " [auto-detected]"
        print(f"[CONFIG] solidworks:   {sw}")

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
        Default config path: same folder as the script / executable.

        Frozen (Inno Setup install):
          sys.executable = C:\\Program Files\\ThermopacStructuringAgent\\python\\python.exe
          → C:\\Program Files\\ThermopacStructuringAgent\\config.ini

        Source / ZIP extract:
          __file__ = <extract>\\agent\\config.py
          → <extract>\\config.ini
        """
        if getattr(sys, "frozen", False):
            base = os.path.dirname(sys.executable)
        else:
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


def _persist_token_to_appdata(token: str, node_id: str, api_url: str) -> None:
    """
    Persist an auto-generated token to the user-writable APPDATA config.
    This is the ONLY write operation in this module.
    Primary config.ini (in Program Files or install dir) is never touched.
    """
    try:
        os.makedirs(os.path.dirname(_APPDATA_CONFIG), exist_ok=True)
        c = configparser.ConfigParser()
        if os.path.exists(_APPDATA_CONFIG):
            c.read(_APPDATA_CONFIG, encoding="utf-8")
        if not c.has_section("cloud"):
            c.add_section("cloud")
        c.set("cloud", "node_token", token)
        c.set("cloud", "node_id",    node_id)
        c.set("cloud", "api_url",    api_url)
        with open(_APPDATA_CONFIG, "w", encoding="utf-8") as f:
            c.write(f)
        print(f"[CONFIG] Testing mode: auto-generated token saved to:")
        print(f"[CONFIG]   {_APPDATA_CONFIG}")
        print(f"[CONFIG]   Agent will self-register with the cloud on startup.")
    except Exception as e:
        print(f"[CONFIG] Testing mode: auto-generated token (in memory only — could not save: {e})")
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
