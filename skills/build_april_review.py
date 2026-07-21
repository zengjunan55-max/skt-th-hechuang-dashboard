from __future__ import annotations

import importlib.util
from pathlib import Path
from types import ModuleType


_THIS_FILE = Path(__file__).resolve()
_CLEAN_PATH = _THIS_FILE.with_name("build_april_review_clean.py")


def _load_clean_module() -> ModuleType:
    spec = importlib.util.spec_from_file_location("build_april_review_clean_entry", _CLEAN_PATH)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"Failed to load clean generator from {_CLEAN_PATH}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


_CLEAN = _load_clean_module()


def __getattr__(name: str):
    if hasattr(_CLEAN, name):
        return getattr(_CLEAN, name)
    base = getattr(_CLEAN, "_BASE", None)
    if base is not None and hasattr(base, name):
        return getattr(base, name)
    raise AttributeError(f"module {__name__!r} has no attribute {name!r}")


def main() -> None:
    _CLEAN.main()


if __name__ == "__main__":
    main()
