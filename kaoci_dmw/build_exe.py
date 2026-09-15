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

# ---- 打包前依赖自检 ----------------------------------------------------
# 血的教训：曾把 numpy 加进 --exclude-module 压体积，结果 core 模块里的
# pandas 在运行时 ImportError（pandas 强依赖 numpy），exe 一启动就崩。
# 这里强制校验「运行时真正用到的库」都能导入，缺一个就直接中止打包。
REQUIRED = [
    "requests",      # 网络请求
    "pandas",        # core: uom_kaoci_export 导出 xlsx
    "numpy",         # pandas 的硬依赖，禁止排除
    "openpyxl",      # pandas 写 xlsx 的引擎
    "PySide6.QtWidgets",
]
_missing = []
for _m in REQUIRED:
    try:
        __import__(_m)
    except Exception as _e:
        _missing.append(f"{_m} -> {_e}")
if _missing:
    print("[依赖自检失败] 以下运行时依赖不可导入，打包会产出坏 exe：")
    for _x in _missing:
        print("   -", _x)
    raise SystemExit(1)
print("[依赖自检通过]", ", ".join(REQUIRED))

args = [
    sys.executable, "-m", "PyInstaller",
    "--noconfirm", "--clean",
    "--onefile",
    "--noconsole",
    "--name", "考次链接大魔王",
    "--paths", PARENT,
    # 仅排除「确认用不到的重包」。注意：不要排除 numpy / pandas / openpyxl，
    # core 模块 uom_kaoci_export 依赖它们写 Excel。
    "--exclude-module", "playwright",
    "--exclude-module", "pytest",
    "--exclude-module", "tkinter",
    "--exclude-module", "matplotlib",
    "--exclude-module", "PyQt5",
    "--exclude-module", "PyQt6",
    "--exclude-module", "scipy",
    "--hidden-import", "uom_kaoci_export",
    "--hidden-import", "pandas",
    "--hidden-import", "numpy",
    "--hidden-import", "openpyxl",
    # token 解析用到 Base64(JSON)，显式带上防裁剪
    "--hidden-import", "base64",
    "--hidden-import", "json",
    MAIN,
]
print("RUN:", " ".join(args))
r = subprocess.run(args, cwd=BASE)
sys.exit(r.returncode)
