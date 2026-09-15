#!/usr/bin/env python3
"""macOS 一键打包：考次链接大魔王 → dist/考次链接大魔王.app

使用：直接 python3 mac_build.py（或 bash mac_build.sh 包装）

前置：Mac 上 Python 3.11+、Chrome 已装。Apple Silicon (M1/M2/M3/M4) 与 Intel 都可。
"""
import os
import shutil
import subprocess
import sys

BASE = os.path.dirname(os.path.abspath(__file__))
PARENT = os.path.dirname(BASE)   # outputs/（uom_kaoci_export.py 在这里）
MAIN = os.path.join(BASE, "main.py")
APP_NAME = "考次链接大魔王"
VENV_DIR = os.path.join(BASE, ".mac_venv")
PY = sys.executable  # 当前 Python


def run(cmd, **kw):
    print("$", " ".join(cmd))
    r = subprocess.run(cmd, **kw)
    if r.returncode != 0:
        sys.exit(r.returncode)


def setup_venv():
    if not os.path.isdir(VENV_DIR):
        run([PY, "-m", "venv", VENV_DIR])
    pip = os.path.join(VENV_DIR, "bin", "pip")
    if not os.path.isfile(pip):
        sys.exit("虚拟环境 pip 不存在：{}".format(pip))
    # 依赖：界面用 PySide6，请求用 requests，Excel 用 openpyxl
    # PyInstaller 必须装在 venv 里（不要用全局）
    run([pip, "install", "--upgrade", "pip"])
    run([pip, "install", "PySide6==6.7.*", "requests", "openpyxl", "pyinstaller>=6.0"])


def build():
    venv_py = os.path.join(VENV_DIR, "bin", "python")
    pyinstaller = os.path.join(VENV_DIR, "bin", "pyinstaller")
    if not os.path.isfile(pyinstaller):
        sys.exit("pyinstaller 未装：{}".format(pyinstaller))
    # 清理旧产物
    for d in ("build", "dist"):
        p = os.path.join(BASE, d)
        if os.path.isdir(p):
            shutil.rmtree(p)
    # PyInstaller 参数：
    #   --windowed       = macOS 上等同 --noconsole（GUI 不弹终端）
    #   --name           产物名（会生成 APP_NAME.app）
    #   --paths          让 pyinstaller 能找到同级 outputs/ 下的 core
    #   --osx-bundle-identifier  Bundle ID（公网必备，本机自用也建议填）
    #   --hidden-import  强制打包 uom_kaoci_export（避免丢）
    #   --exclude-module 排除不必要的依赖减小体积
    args = [
        pyinstaller,
        "--noconfirm", "--clean",
        "--windowed",
        "--name", APP_NAME,
        "--paths", PARENT,
        "--osx-bundle-identifier", "com.kacida.kaoci",
        "--hidden-import", "uom_kaoci_export",
        # 显式带上标准库（token 解析用到 base64/json，避免极端情况下被裁掉）
        "--hidden-import", "base64",
        "--hidden-import", "json",
        "--exclude-module", "playwright",
        "--exclude-module", "pytest",
        "--exclude-module", "tkinter",
        "--exclude-module", "matplotlib",
        "--exclude-module", "PyQt5",
        "--exclude-module", "PyQt6",
        "--exclude-module", "numpy",
        "--exclude-module", "scipy",
        "--exclude-module", "PIL",
        MAIN,
    ]
    run(args, cwd=BASE)
    # 产物路径
    app = os.path.join(BASE, "dist", "{}.app".format(APP_NAME))
    print("\n✅ 打包完成")
    print("   应用包:", app)
    print("   启动方式: open '{}'".format(app))
    print("   或直接双击打开（首次可能 Gatekeeper 拦截，需右键→打开）")


def main():
    if sys.platform != "darwin":
        print("⚠️ 此脚本仅在 macOS 上运行（当前平台: {}）。".format(sys.platform))
        print("   请把整个 kaoci_dmw 目录拷贝到 MacBook 后再运行本脚本。")
        sys.exit(1)
    setup_venv()
    build()


if __name__ == "__main__":
    main()