from __future__ import annotations

import os
import sys
from pathlib import Path


def _set_if_exists(env_name: str, path: Path) -> bool:
    if path.exists():
        os.environ[env_name] = str(path)
        return True
    return False


def configure_tk_env() -> None:
    candidates: list[tuple[Path, Path]] = []

    public_root = Path(r"C:\Users\Public\codex_tcl")
    candidates.append((public_root / "tcl8.6", public_root / "tk8.6"))

    meipass = Path(getattr(sys, "_MEIPASS", Path(sys.executable).resolve().parent))
    for root in (meipass, meipass / "_internal", Path(sys.executable).resolve().parent, Path(sys.executable).resolve().parent / "_internal"):
        candidates.append((root / "_tcl_data", root / "_tk_data"))
        candidates.append((root / "tcl8.6", root / "tk8.6"))

    for tcl_dir, tk_dir in candidates:
        tcl_ok = _set_if_exists("TCL_LIBRARY", tcl_dir)
        tk_ok = _set_if_exists("TK_LIBRARY", tk_dir)
        if tcl_ok or tk_ok:
            return


configure_tk_env()
