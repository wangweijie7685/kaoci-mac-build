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
    "requests",      # 网络请求（兜底）
    "curl_cffi",     # v1.7: Chrome TLS/JA3 指纹模拟，绕 UOM 风控「限制登录」
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
    "--hidden-import", "curl_cffi",
    # token 解析用到 Base64(JSON)，显式带上防裁剪
    "--hidden-import", "base64",
    "--hidden-import", "json",
    MAIN,
]
print("RUN:", " ".join(args))
r = subprocess.run(args, cwd=BASE)
if r.returncode != 0:
    print("❌ PyInstaller 构建失败")
    sys.exit(r.returncode)

# ---- 构建后验证（硬门槛，不通过就报错退出）--------------------------
# 让产物自己 import 一遍全部运行时依赖。这是唯一能真正证明
# 「依赖被打进包了」的办法：产物能生成、进程能起来，都不代表能用。
exe = os.path.join(BASE, "dist", "考次链接大魔王.exe")
print("\n[验证] 运行产物自检：", exe, "--self-test")
try:
    st = subprocess.run([exe, "--self-test"], cwd=os.path.dirname(exe),
                        capture_output=True, timeout=180)
    out = (st.stdout or b"").decode("utf-8", "ignore")
    err = (st.stderr or b"").decode("utf-8", "ignore")
    print("  退出码:", st.returncode)
    for line in (out + err).splitlines():
        print("  |", line)
    if st.returncode != 0:
        print("❌ 产物自检失败：包内缺少依赖，请勿分发！")
        sys.exit(1)
    print("✅ 产物自检通过：冻结环境依赖完整")
except subprocess.TimeoutExpired:
    print("❌ 产物自检超时")
    sys.exit(1)

# 再做一次归档 + 启动窗口检查（依赖精确 TOC + 区分主窗口/错误对话框）
verifier = os.path.join(BASE, "verify_exe.py")
if os.path.isfile(verifier):
    print("\n[验证] 归档与启动窗口检查")
    v = subprocess.run([sys.executable, verifier, exe], cwd=BASE)
    if v.returncode != 0:
        print("❌ 归档/启动窗口检查未通过")
        sys.exit(v.returncode)

