# -*- coding: utf-8 -*-
"""一键打包脚本：把考次链接大魔王 打包成单文件 exe"""
import os
import sys

BASE = os.path.dirname(os.path.abspath(__file__))
PARENT = os.path.dirname(BASE)   # outputs/
MAIN = os.path.join(BASE, "main.py")

if getattr(sys, "frozen", False):
    raise SystemExit("请在源码环境运行本脚本")

# 使用 PyInstaller 命令行打包（避免 spec 复杂化）
# 关键点：
#   1) --paths 让 PyInstaller 能找到 uom_kaoci_export（在上级 outputs/）
#   2) --collect-all playwright 会太大；代码里 playwright 是可选依赖
#      （ImportError 时自动降级 websocket），所以排除它减小体积
#   3) --noconsole: 纯 GUI 不弹黑框
#   4) --onefile: 单文件，方便分发给同事
#   5) main.py 已适配 sys.frozen，token.txt / 调试profile 落在 exe 旁
import subprocess
args = [
    sys.executable, "-m", "PyInstaller",
    "--noconfirm", "--clean",
    "--onefile",
    "--noconsole",
    "--name", "考次链接大魔王",
    "--paths", PARENT,
    "--exclude-module", "playwright",
    "--exclude-module", "pytest",
    "--exclude-module", "tkinter",
    "--exclude-module", "matplotlib",
    "--exclude-module", "PyQt5",
    "--exclude-module", "PyQt6",
    "--exclude-module", "numpy",
    "--exclude-module", "scipy",
    "--exclude-module", "PIL",
    "--hidden-import", "uom_kaoci_export",
    # token 解析用到 Base64(JSON)，显式带上防裁剪
    "--hidden-import", "base64",
    "--hidden-import", "json",
    MAIN,
]
print("RUN:", " ".join(args))
r = subprocess.run(args, cwd=BASE)
sys.exit(r.returncode)
