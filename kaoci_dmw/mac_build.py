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
    # 依赖：界面用 PySide6，请求用 requests，Excel 用 pandas + openpyxl
    #   ！！pandas 是 core 模块 uom_kaoci_export 的硬依赖，必须装，
    #     否则打出来的 .app 一启动就 ModuleNotFoundError: pandas
    #   ！！pandas 又强依赖 numpy，打包时绝对不能把 numpy 排除掉
    # PyInstaller 必须装在 venv 里（不要用全局）
    run([pip, "install", "--upgrade", "pip"])
    run([pip, "install", "PySide6==6.7.*", "requests", "pandas", "numpy",
         "openpyxl", "curl_cffi", "pyinstaller>=6.0"])


def selfcheck_venv():
    """打包前确认运行时依赖都能在 venv 里导入，缺一个就中止。"""
    venv_py = os.path.join(VENV_DIR, "bin", "python")
    required = ["requests", "curl_cffi", "pandas", "numpy", "openpyxl",
                "PySide6.QtWidgets"]
    code = (
        "import sys\n"
        "bad=[]\n"
        "for m in %r:\n"
        "    try: __import__(m)\n"
        "    except Exception as e: bad.append('%%s -> %%s' %% (m, e))\n"
        "print('[依赖自检]', 'FAIL' if bad else 'OK', *bad, sep=' ')\n"
        "sys.exit(1 if bad else 0)\n" % (required,)
    )
    run([venv_py, "-c", code])


def verify_bundle(app_path):
    """构建后自检：让 .app 里的可执行文件自己 import 一遍依赖。

    这是唯一可靠的判定方式。曾经的做法是「在 .app 目录里找 numpy/pandas
    文件夹」，但 macOS 上纯 Python 包（requests/openpyxl/core 模块）被装进
    可执行文件内部的 PYZ 归档，磁盘上并不存在对应目录，会误判成缺失。
    现在改为直接运行产物：<app>/Contents/MacOS/<name> --self-test
    退出码 0 = 依赖完整；非 0 = 缺依赖，立即中止发布。
    """
    print("\n[构建后自检] 运行 .app 内可执行文件 --self-test")
    exe_name = os.path.basename(app_path)[:-4]  # 去掉 .app
    inner = os.path.join(app_path, "Contents", "MacOS", exe_name)
    if not os.path.isfile(inner):
        print("  ! 找不到可执行文件：", inner)
        return False
    try:
        st = subprocess.run([inner, "--self-test"], capture_output=True, timeout=180)
    except subprocess.TimeoutExpired:
        print("  ! 自检超时")
        return False
    out = (st.stdout or b"").decode("utf-8", "ignore")
    err = (st.stderr or b"").decode("utf-8", "ignore")
    print("  退出码:", st.returncode)
    for line in (out + err).splitlines():
        print("  |", line)
    return st.returncode == 0


def verify_bundle_files(app_path):
    """辅助检查：确认带二进制扩展的 numpy / pandas / PySide6 确实在磁盘上。

    纯 Python 包不在磁盘上属正常（在 PYZ 里），所以这里只查二进制包，
    最终结论仍以 verify_bundle 的 --self-test 结果为准。
    """
    hits = []
    for root, _dirs, files in os.walk(app_path):
        rel = os.path.relpath(root, app_path)
        for f in files:
            hits.append(os.path.join(rel, f))
    blob = "\n".join(hits).lower()
    checks = {
        "numpy(二进制)": lambda s: "numpy/core" in s or "/numpy/" in s,
        "pandas(二进制)": lambda s: "pandas/_libs" in s or "/pandas/" in s,
        "PySide6": lambda s: "pyside6" in s,
    }
    print("\n[构建后自检] 二进制依赖落盘检查")
    ok = True
    for name, fn in checks.items():
        hit = fn(blob)
        print("  {} {}".format("OK  " if hit else "MISS", name))
        ok = ok and hit
    return ok


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
    #   --hidden-import  强制打包 uom_kaoci_export / pandas / numpy（避免丢）
    #   --exclude-module 只排除「确认用不到」的包减小体积
    #       ！！切勿排除 numpy / pandas / openpyxl：core 模块靠它们写 Excel，
    #         排除后产物能生成但一启动就 ImportError（已踩过坑）
    args = [
        pyinstaller,
        "--noconfirm", "--clean",
        "--windowed",
        "--name", APP_NAME,
        "--paths", PARENT,
        "--osx-bundle-identifier", "com.kacida.kaoci",
        "--hidden-import", "uom_kaoci_export",
        "--hidden-import", "pandas",
        "--hidden-import", "numpy",
        "--hidden-import", "openpyxl",
        "--hidden-import", "curl_cffi",
        # 显式带上标准库（token 解析用到 base64/json，避免极端情况下被裁掉）
        "--hidden-import", "base64",
        "--hidden-import", "json",
        "--exclude-module", "playwright",
        "--exclude-module", "pytest",
        "--exclude-module", "tkinter",
        "--exclude-module", "matplotlib",
        "--exclude-module", "PyQt5",
        "--exclude-module", "PyQt6",
        "--exclude-module", "scipy",
        MAIN,
    ]
    run(args, cwd=BASE)
    # 产物路径
    app = os.path.join(BASE, "dist", "{}.app".format(APP_NAME))
    print("\n✅ 打包完成")
    print("   应用包:", app)
    print("   启动方式: open '{}'".format(app))
    print("   或直接双击打开（首次可能 Gatekeeper 拦截，需右键→打开）")
    # 落地自检：让产物自己 import 一遍依赖，不合格直接失败退出
    verify_bundle_files(app)          # 仅供参考的落盘检查
    if not verify_bundle(app):        # 决定性判定
        print("\n❌ 构建后自检未通过：产物缺少关键依赖，请勿分发！")
        sys.exit(1)
    print("\n✅ 构建后自检通过：冻结环境依赖完整")


def main():
    if sys.platform != "darwin":
        print("⚠️ 此脚本仅在 macOS 上运行（当前平台: {}）。".format(sys.platform))
        print("   请把整个 kaoci_dmw 目录拷贝到 MacBook 后再运行本脚本。")
        sys.exit(1)
    setup_venv()
    selfcheck_venv()
    build()


if __name__ == "__main__":
    main()