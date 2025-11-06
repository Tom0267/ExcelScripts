# pathHelper.py
from pathlib import Path
import os, sys, shutil

def is_frozen() -> bool:
    return getattr(sys, "frozen", False)

def app_dir() -> Path:
    return Path(sys.executable).parent if is_frozen() else Path(__file__).resolve().parent

def resource_path(rel: str) -> Path:
    base = Path(getattr(sys, "_MEIPASS", app_dir()))
    return (base / rel).resolve()

def user_csv_path() -> Path:
    appdata = os.getenv("APPDATA") or str(Path.home() / "AppData" / "Roaming")
    cfg_dir = Path(appdata) / "StationMatcher"
    cfg_dir.mkdir(parents=True, exist_ok=True)
    return cfg_dir / "stationMap.csv"

def _find_template_csv() -> Path | None:
    for p in (
        resource_path("Resources/stationMap.csv"),
        app_dir() / "Resources" / "stationMap.csv",
        app_dir() / "stationMap.csv",
    ):
        if p.exists():
            return p
    return None

def ensure_user_csv() -> str:
    r"""
    Ensure an editable CSV exists at %APPDATA%\StationMatcher\stationMap.csv.
    Creates it with a header if missing.
    """
    dst = user_csv_path()
    if not dst.exists():
        # 1) Try to copy a template
        src = _find_template_csv()
        try:
            if src:
                shutil.copyfile(src, dst)
            else:
                dst.write_text("Code,Name\n", encoding="utf-8")
        except Exception:
            # Last resort: make sure file exists with a header
            dst.parent.mkdir(parents=True, exist_ok=True)
            if not dst.exists():
                dst.write_text("Code,Name\n", encoding="utf-8")

    # If file exists but is empty, add a header so Excel opens it cleanly
    try:
        if dst.stat().st_size == 0:
            dst.write_text("Code,Name\n", encoding="utf-8")
    except Exception:
        pass

    return str(dst)