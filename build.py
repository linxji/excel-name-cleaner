# -*- coding: utf-8 -*-
"""一键打包：生成 dist/Excel名称清理工具.exe

用法：
    pip install pyinstaller pywin32
    python build.py
"""
import os
import shutil
import subprocess
import sys

ROOT = os.path.dirname(os.path.abspath(__file__))
NAME = "Excel名称清理工具"


def main():
    try:
        import PyInstaller  # noqa: F401
    except ImportError:
        print("缺少 PyInstaller，请先执行：pip install pyinstaller pywin32")
        return 1

    icon = os.path.join(ROOT, "assets", "app.ico")
    ver = os.path.join(ROOT, "build", "version_info.txt")
    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--noconfirm", "--clean", "--onefile", "--windowed",
        "--name", NAME,
        "--icon", icon,
        "--add-data", icon + ";.",
        "--paths", os.path.join(ROOT, "src"),
        "--hidden-import", "win32com.client",
        "--hidden-import", "pythoncom",
        "--hidden-import", "pywintypes",
        "--distpath", os.path.join(ROOT, "dist"),
        "--workpath", os.path.join(ROOT, "build", "work"),
        "--specpath", os.path.join(ROOT, "build"),
    ]
    if os.path.exists(ver):
        cmd += ["--version-file", ver]
    cmd.append(os.path.join(ROOT, "src", "main.py"))

    print("正在打包 …")
    rc = subprocess.call(cmd)
    if rc != 0:
        print("打包失败，退出码 %d" % rc)
        return rc
    out = os.path.join(ROOT, "dist", NAME + ".exe")
    print("\n完成：%s（%.1f MB）" % (out, os.path.getsize(out) / 1048576.0))
    return 0


if __name__ == "__main__":
    sys.exit(main())
