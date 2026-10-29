"""检查依赖并用与当前解释器同目录的 pythonw 启动界面。"""

from __future__ import annotations

import os
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
REQ = ROOT / "requirements.txt"


def _deps_ok() -> bool:
    try:
        import customtkinter  # noqa: F401
        import lxml  # noqa: F401
        import win32com.client  # noqa: F401
    except ImportError:
        return False
    return True


def _ensure_deps() -> None:
    if _deps_ok():
        return
    print("[local-redact] 缺少依赖，正在安装 requirements.txt ...", flush=True)
    code = subprocess.call([sys.executable, "-m", "pip", "install", "-r", str(REQ)])
    if code != 0 or not _deps_ok():
        print("依赖安装失败。请手动执行：python -m pip install -r requirements.txt", flush=True)
        input("按回车键退出...")
        raise SystemExit(1)


def main() -> int:
    os.chdir(ROOT)
    _ensure_deps()
    pythonw = Path(sys.executable).with_name("pythonw.exe")
    if not pythonw.is_file():
        print(f"找不到 pythonw：{pythonw}", flush=True)
        input("按回车键退出...")
        return 1
    subprocess.Popen(
        [str(pythonw), "-m", "local_redact.gui", *sys.argv[1:]],
        cwd=str(ROOT),
        close_fds=True,
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
