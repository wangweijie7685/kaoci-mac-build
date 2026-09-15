# -*- coding: utf-8 -*-
"""
exe 产物验证脚本（打包后必跑）
===============================
两道检查，缺一不可：

  1) 归档检查 —— 扫描单文件 exe 内部的 PyInstaller TOC，
     确认 numpy / pandas / openpyxl 等运行时依赖真的被收进去了。
     （教训：只把 numpy 加进 --exclude-module，exe 能生成、进程也能起，
       但一启动就弹 "Unable to import required dependency numpy"，
       光看"进程存活"是发现不了的。）

  2) 窗口检查 —— 启动 exe，用 Win32 EnumWindows 枚举顶层窗口标题：
       命中 "考次链接大魔王"            -> 成功
       命中 "Unhandled exception"/"Traceback" -> 失败（把错误框文本也抓出来）
     这样能真正区分「GUI 正常」与「弹了错误对话框」。

用法：
    python verify_exe.py                    # 默认校验 dist/考次链接大魔王.exe
    python verify_exe.py 发布exe/xxx.exe
"""
import os
import re
import subprocess
import sys
import time

APP_NAME = "考次链接大魔王"

# 必须在 exe 归档里能找到（读 PyInstaller 真实 TOC，不用字节子串——
# 子串会被 pandas 源码里出现的 "numpy" 字样误命中，测不出真问题）
# 注意：不能用 ^numpy\. 这种宽泛规则，因为 "numpy.libs\xxx.dll" 只是随包 DLL，
# 不代表 numpy 包本体被打进去了。这里一律用「签名条目」精确匹配。
REQUIRED_ARCHIVE = [
    ("numpy", (r"numpy[\\/]core[\\/]_multiarray_umath",
               r"^numpy\.core$", r"^numpy\.version$", r"^numpy\._core$")),
    ("pandas", (r"pandas[\\/]_libs[\\/]",
                r"^pandas\.core$", r"^pandas\._libs$", r"^pandas\.util\._decorators$")),
    ("openpyxl", (r"^openpyxl\.workbook\.workbook$", r"^openpyxl\.utils$",
                  r"^openpyxl\.reader\.excel$")),
    ("requests", (r"^requests\.api$", r"^requests\.sessions$")),
    ("PySide6.QtWidgets", (r"PySide6[\\/]QtWidgets", r"^PySide6\.QtWidgets$")),
    ("uom_kaoci_export", (r"^uom_kaoci_export$", r"uom_kaoci_export\.")),
]
# 出现这些字样说明弹的是错误框
ERROR_HINTS = ["Unhandled exception", "Traceback", "ImportError", "Error"]

_MAGIC = b"MEI\014\013\012\013\016"   # PyInstaller 归档 cookie 魔数


def _archive_names(exe_path):
    """返回 exe 内全部归档条目名（外层 TOC + PYZ 内纯 Python 模块）。

    注意：PyInstaller 会把纯 Python 模块装进 PYZ 子归档，只读外层 TOC
    会漏掉 requests / openpyxl / numpy 的包本体，必须一并解析。
    失败返回 None。
    """
    try:
        from PyInstaller.archive.readers import CArchiveReader
    except Exception as e:
        print(f"  ! 无法加载 PyInstaller 归档读取器: {e}")
        return None
    try:
        reader = CArchiveReader(exe_path)
    except Exception as e:
        print(f"  ! 归档读取失败: {e}")
        return None

    names = set(reader.toc.keys())
    for entry in list(reader.toc.keys()):
        if entry.lower().endswith(".pyz"):
            try:
                sub = reader.open_embedded_archive(entry)
                names.update(sub.toc.keys())
            except Exception as e:
                print(f"  ! 读取子归档 {entry} 失败: {e}")
    return names


def check_archive(exe_path):
    print("\n[1/2] 归档依赖检查")
    size_mb = os.path.getsize(exe_path) / 1048576
    print(f"  产物: {exe_path}")
    print(f"  体积: {size_mb:.1f} MB")
    names = _archive_names(exe_path)
    if names is None:
        print("  归档无法读取，跳过内容检查（仅按体积判断）")
        return size_mb >= 30
    print(f"  归档条目数: {len(names)}")
    ok = True
    for label, pats in REQUIRED_ARCHIVE:
        hit = next((n for n in names if any(re.search(p, n) for p in pats)), None)
        print(f"  {'OK  ' if hit else 'MISS'} {label:20} {hit or ''}")
        ok = ok and bool(hit)
    if size_mb < 30:
        print("  ! 体积异常偏小，可能依赖被裁掉了")
        ok = False
    return ok


def check_window(exe_path, wait_s=25):
    """启动 exe 并判定启动是否成功。

    两个必须处理的坑：
      1) PyInstaller 单文件包会派生 bootloader 父进程 + 真正运行的应用子进程。
         只 terminate 父进程会留下子进程（及其错误对话框），污染下一次检测。
      2) 因此检测时必须「只认本次启动新出现的窗口」——启动前先给现有窗口拍快照，
         否则会把上一轮残留的错误对话框误判成本次失败（已踩过）。
    """
    print("\n[2/2] 启动窗口检查")
    import ctypes
    from ctypes import wintypes

    user32 = ctypes.windll.user32
    EnumWindows = user32.EnumWindows
    WNDENUMPROC = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
    GetWindowTextW = user32.GetWindowTextW
    GetWindowTextLengthW = user32.GetWindowTextLengthW

    def snapshot():
        """返回 {hwnd: title}"""
        out = {}

        def cb(hwnd, _):
            n = GetWindowTextLengthW(hwnd)
            if n:
                buf = ctypes.create_unicode_buffer(n + 1)
                GetWindowTextW(hwnd, buf, n + 1)
                out[hwnd] = buf.value
            return True

        EnumWindows(WNDENUMPROC(cb), 0)
        return out

    def kill_tree(pid):
        """杀掉整个进程树（单文件包会派生父子两个进程）。"""
        try:
            subprocess.run(["taskkill", "/F", "/T", "/PID", str(pid)],
                           capture_output=True)
        except Exception:
            pass

    before = set(snapshot().keys())
    print(f"  启动前已有窗口 {len(before)} 个（作为基线，忽略）")
    proc = subprocess.Popen([exe_path], cwd=os.path.dirname(exe_path) or ".")
    print(f"  已启动 PID {proc.pid}")

    result = None
    for i in range(wait_s):
        time.sleep(1)
        rc = proc.poll()
        if rc is not None:
            # 父进程退出不代表失败（也可能真的崩了），再给子进程 2 秒亮相机会
            time.sleep(2)
        new = {h: t for h, t in snapshot().items() if h not in before}
        for t in new.values():
            if t == APP_NAME or APP_NAME in t:
                print(f"  [{i+1}s] 命中主窗口标题: {t!r}")
                result = True
                break
            if any(h in t for h in ERROR_HINTS):
                print(f"  [{i+1}s] 命中错误对话框: {t!r}  -> 启动失败")
                result = False
                break
        if result is not None:
            break
        if rc is not None and not new:
            print(f"  [{i+1}s] 进程已退出 rc={rc} 且无新窗口 -> 启动失败")
            result = False
            break

    kill_tree(proc.pid)
    print("  测试进程已关闭（含子进程）")
    if result is None:
        print("  ! 超时未捕获到窗口，无法判定")
        return False
    return result


def main():
    rel = sys.argv[1] if len(sys.argv) > 1 else os.path.join("dist", APP_NAME + ".exe")
    exe = os.path.abspath(rel)
    if not os.path.isfile(exe):
        print("找不到产物:", exe)
        return 2
    a = check_archive(exe)
    b = check_window(exe)
    print("\n" + "=" * 46)
    print(f"归档依赖: {'通过' if a else '失败'}   启动窗口: {'通过' if b else '失败'}")
    print("结论:", "验证通过，可发布" if (a and b) else "验证失败，不要发布")
    print("=" * 46)
    return 0 if (a and b) else 1


if __name__ == "__main__":
    sys.exit(main())
