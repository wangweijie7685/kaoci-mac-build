# -*- coding: utf-8 -*-
"""
考次链接大魔王  v1.7
=====================
UOM 民航局无人机考试系统「考次链接」批量导出桌面工具

v1.7 变更：
  * 新增：Chrome 浏览器指纹 —— UOM 风控开始按 TLS/JA3 指纹识别脚本
    客户端（表现为 token 频繁被「限制登录」），网络层改用 curl_cffi
    模拟真 Chrome 握手指纹，从协议层与真浏览器无差别；requests 保留兜底。
  * v1.6：考点过滤框 + 服务端单页 limit<=100 硬上限自动分页修复。

v1.6 变更：
  * 新增：考点过滤框 —— 考试点太多，下拉翻页找不到。
    在「考点过滤」框输入关键字（如“吉林”），下拉列表实时只显示
    匹配的考点；清空恢复全部。查询时选中的考点仍精确传给接口。
  * 修复：查询考次一律 500 —— UOM 服务端给列表接口加了单页
    limit<=100 硬上限（超过直接 500），core 已改为自动分页拉取。

v1.5 变更：
  * 新增：--self-test 自检开关 —— 打包脚本打完后会让产物自己 import 一遍
    全部运行时依赖，用退出码判定「包是否完整」。避免再次出现「能构建、
    能启动、一用就崩」的坏包（此前曾因排除 numpy 导致启动即报
    Unable to import required dependency numpy）。

v1.4 变更：
  * 修复：token 解析支持 Base64 嵌套 —— UOM 登录态实际以 Base64(JSON)
    形式存在 localStorage 的 session_token 键里，旧版按裸 UUID 匹配导致
    始终扫不到新 token（表现为「扫到的登录态均已过期」）。现在自动解码。
  * 新增：切换 Tab 时分别记忆各自的输出目录，来回切换不互相覆盖。
  * 变更：默认输出目录改为用户 Downloads 目录。

v1.3 变更：
  * 修复：token 过期后「一键获取登录态」失效 —— 之前扫到 localStorage 里
    缓存的过期旧 token 会直接保存并提示成功；现在每个候选 token 都先经
    接口校验，过期的列入黑名单继续等待，直到重新登录拿到新 token。
  * 新增：考试点下拉支持输入关键字模糊搜索（包含匹配）。

功能：
  * Tab 切换：按机构查看 / 按考点查看
  * 筛选：考试日期区间 + 考试点（模糊搜索）
  * 参数：线程数 / 等待时间 / 导出条数 / 导出全部
  * 输出：每个考次一个 xlsx，第1行合并居中抬头，第2行 21 列表头，第3行起考生明细
"""
import os
import sys
import time
import threading
import datetime
import subprocess

# 源码运行模式：core 模块在上级 outputs/ 目录（与 kaoci_dmw/ 平级）
if not getattr(sys, "frozen", False):
    sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))


# ---- 打包自检开关 ------------------------------------------------------
# 打包脚本在构建完成后执行：<产物> --self-test
# 在「冻结环境」里真刀真枪 import 一遍关键依赖，全部通过则退出码 0。
# 这样「依赖到底有没有被打进包」就是可验证的事实，而不是靠体积或字符串猜。
# 教训：曾把 numpy 加进 PyInstaller 的 --exclude-module 压体积，产物照样
# 生成、也没有任何报错，用户一双击就弹
# "Unable to import required dependency numpy"。
SELF_TEST_MODULES = [
    "requests",           # 网络请求
    "curl_cffi",          # v1.7: Chrome TLS 指纹模拟（绕 UOM 风控）
    "pandas",             # core 模块导 Excel
    "numpy",              # pandas 的硬依赖
    "openpyxl",           # pandas 写 xlsx 的引擎
    "PySide6.QtCore",     # 界面
    "PySide6.QtWidgets",
    "uom_kaoci_export",   # 核心业务模块
]

if "--self-test" in sys.argv:
    import importlib

    def _emit(msg):
        # --noconsole / --windowed 打包后 sys.stdout/stderr 可能是 None，
        # 直接 write 会抛 AttributeError，这里做空值兜底。
        for _s in (sys.stdout, sys.stderr):
            if _s is not None:
                try:
                    _s.write(msg + "\n")
                    _s.flush()
                except Exception:
                    pass

    _bad = []
    for _m in SELF_TEST_MODULES:
        try:
            importlib.import_module(_m)
        except Exception as _e:
            _bad.append("{} -> {}: {}".format(_m, type(_e).__name__, _e))
    if _bad:
        _emit("[self-test] FAIL 冻结环境缺少以下依赖：")
        for _x in _bad:
            _emit("    - " + _x)
        sys.exit(1)
    _emit("[self-test] OK 已导入: " + ", ".join(SELF_TEST_MODULES))
    sys.exit(0)

# PyInstaller 冻结模式与源码模式都从包内/上级目录导入 core
import uom_kaoci_export as core  # noqa: E402

# ---- Qt 绑定自动适配：优先 PySide6，其次 PyQt5 / PyQt6 ----
QT_BINDING = None
try:
    from PySide6.QtCore import Qt, Signal, QThread, QDate, QTimer  # noqa: E402
    from PySide6.QtWidgets import (  # noqa: E402
        QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QGridLayout,
        QTabWidget, QLabel, QLineEdit, QPushButton, QComboBox, QDateEdit,
        QSpinBox, QDoubleSpinBox, QCheckBox, QTableWidget, QTableWidgetItem,
        QProgressBar, QTextEdit, QFileDialog, QGroupBox, QHeaderView, QMessageBox,
        QCompleter,
    )
    QT_BINDING = "PySide6"
    # PySide6 顶层不带 Qt.MatchContains 等枚举别名，这里补一层
    QComboBox.NoInsert = QComboBox.InsertPolicy.NoInsert
    Qt.MatchContains = Qt.MatchFlag.MatchContains
    Qt.CaseInsensitive = Qt.CaseSensitivity.CaseInsensitive
except ImportError:
    try:
        from PyQt5.QtCore import Qt, pyqtSignal as Signal, QThread, QDate, QTimer  # noqa: E402
        from PyQt5.QtWidgets import (  # noqa: E402
            QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QGridLayout,
            QTabWidget, QLabel, QLineEdit, QPushButton, QComboBox, QDateEdit,
            QSpinBox, QDoubleSpinBox, QCheckBox, QTableWidget, QTableWidgetItem,
            QProgressBar, QTextEdit, QFileDialog, QGroupBox, QHeaderView, QMessageBox,
            QCompleter,
        )
        QT_BINDING = "PyQt5"
    except ImportError:
        from PyQt6.QtCore import Qt, pyqtSignal as Signal, QThread, QDate, QTimer  # noqa: E402
        from PyQt6.QtWidgets import (  # noqa: E402
            QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QGridLayout,
            QTabWidget, QLabel, QLineEdit, QPushButton, QComboBox, QDateEdit,
            QSpinBox, QDoubleSpinBox, QCheckBox, QTableWidget, QTableWidgetItem,
            QProgressBar, QTextEdit, QFileDialog, QGroupBox, QHeaderView, QMessageBox,
            QCompleter,
        )
        QT_BINDING = "PyQt6"
    # PyQt6 枚举换到子命名空间，这里做一层兼容
    QTableWidget.NoEditTriggers = QTableWidget.EditTrigger.NoEditTriggers
    QTableWidget.SelectRows = QTableWidget.SelectionBehavior.SelectRows
    QHeaderView.ResizeToContents = QHeaderView.ResizeMode.ResizeToContents
    QHeaderView.Stretch = QHeaderView.ResizeMode.Stretch
    QMessageBox.Yes = QMessageBox.StandardButton.Yes
    QLineEdit.Password = QLineEdit.EchoMode.Password
    QLineEdit.Normal = QLineEdit.EchoMode.Normal
    Qt.AlignCenter = Qt.AlignmentFlag.AlignCenter
    QComboBox.NoInsert = QComboBox.InsertPolicy.NoInsert
    Qt.MatchContains = Qt.MatchFlag.MatchContains
    Qt.CaseInsensitive = Qt.CaseSensitivity.CaseInsensitive

APP_NAME = "考次链接大魔王  有问题请联系田老师"
if getattr(sys, "frozen", False):
    if sys.platform == "darwin":
        # macOS .app 内部只读 → 数据文件放 ~/Library/Application Support/APP_NAME/
        APP_DIR = os.path.join(
            os.path.expanduser("~/Library/Application Support"), APP_NAME
        )
    else:
        # Windows/Linux 可执行文件旁的目录可写
        APP_DIR = os.path.dirname(os.path.abspath(sys.executable))
else:
    APP_DIR = os.path.dirname(os.path.abspath(__file__))
try:
    os.makedirs(APP_DIR, exist_ok=True)
except Exception:
    pass
TOKEN_FILE = os.path.join(APP_DIR, "token.txt")   # 登录态持久化文件
DEFAULT_TOKEN = core.TOKEN
DEFAULT_OUT = r"C:\Users\ThinkPad\Downloads"
DEFAULT_OUT_PLACE = r"C:\Users\ThinkPad\Downloads"
CDP_PORT = 9223   # Chrome 远程调试端口（CDP 提取登录态）
# 独立的调试浏览器配置目录：用它启动 Chrome，绝不碰用户日常 Chrome 的 profile
# （日常浏览器继续正常用，登录态/标签都不受影响；此目录内的 UOM 登录态会持久化保存）
DEBUG_PROFILE_DIR = os.path.join(APP_DIR, "chrome_uom_profile")

MODE_AGENCY = "agency"
MODE_PLACE = "place"

TABLE_HEADERS = ["序号", "考试日期", "考试点", "服务方", "考试员", "考次编号", "状态"]


# ==================== 工具 ====================
def _t(v):
    if isinstance(v, dict):
        return v.get("title") or v.get("showTitle") or v.get("name") or ""
    return v or ""


def _safe_start_worker(owner, attr, worker):
    """安全地启动一个新的 QThread 并替换 owner.<attr>。

    反复点击「刷新/查询」时旧线程对象若被 Python GC 回收，而底层 C++
    线程仍在运行，PySide6 会直接段错误崩溃（源码跑 PyQt5 宽容不崩，
    打包后的 PySide6 极敏感）。因此替换前必须：
      1) 断开旧 worker 的信号（避免旧槽误触发）；
      2) 等旧线程结束（wait）；
      3) 手动 deleteLater 释放。
    """
    old = getattr(owner, attr, None)
    if old is not None and old is not worker:
        # 先断开，避免旧 worker 的 finished 信号误伤新状态
        for sig in getattr(old, "_sig_names", ()):
            try:
                sig.disconnect()
            except Exception:
                pass
        if old.isRunning():
            old.wait(30000)   # 等旧线程收尾（网络请求最多 60s 超时，这里给 30s 宽限）
        old.deleteLater()
    # finished 后：C++ 对象被 deleteLater 删除，必须把 owner 的引用清空，
    # 否则下次 isRunning() 会对已删除对象调用 → "Internal C++ object already
    # deleted" 崩溃
    def _clear_ref():
        if getattr(owner, attr, None) is worker:
            setattr(owner, attr, None)
    worker.finished.connect(_clear_ref)
    worker.finished.connect(worker.deleteLater)
    setattr(owner, attr, worker)
    worker.start()
    return worker


def row_date(r):
    return core.fmt_date(r.get("KAOSRJ"))


# ==================== 查询线程 ====================
class QueryWorker(QThread):
    sig_ok = Signal(list)
    sig_fail = Signal(str)
    _sig_names = ()   # 由 __init__ 填充为实际信号元组，供 _safe_start_worker 断开

    def __init__(self, mode, start, end, exam_place, parent=None):
        super().__init__(parent)
        # 注意：别用 self.start —— 会覆盖 QThread.start()
        self.mode = mode
        self.start_date = start
        self.end_date = end
        self.exam_place = exam_place
        self._sig_names = (self.sig_ok, self.sig_fail)

    def run(self):
        try:
            rows = core.fetch_list(self.start_date, self.end_date, mode=self.mode,
                                   exam_place=self.exam_place)
            self.sig_ok.emit(rows or [])
        except Exception as e:
            self.sig_fail.emit(str(e)[:300])


class PlaceWorker(QThread):
    """拉取考试点字典，填充下拉框（按 title 去重）"""
    sig_ok = Signal(dict)
    sig_fail = Signal(str)
    _sig_names = ()

    def __init__(self, parent=None):
        super().__init__(parent)
        self._sig_names = (self.sig_ok, self.sig_fail)

    def run(self):
        # title -> name（保留首个 name 作为发送值）
        places = {}
        try:
            for mode in (MODE_AGENCY, MODE_PLACE):
                try:
                    for r in core.fetch_list("", "", mode=mode) or []:
                        p = r.get("EXAM_PLACE")
                        if not isinstance(p, dict):
                            continue
                        title = p.get("title")
                        if not title:
                            continue
                        # UOM 历史遗留：同名考点有多 UUID，按 title 去重
                        places.setdefault(title, p.get("name") or title)
                except Exception:
                    pass
                time.sleep(1.5)
            self.sig_ok.emit(places)
        except Exception as e:
            self.sig_fail.emit(str(e)[:200])


# ==================== 导出线程 ====================
class ExportWorker(QThread):
    sig_log = Signal(str)
    sig_progress = Signal(int, int)
    sig_done = Signal(int, int, int)
    _sig_names = ()

    def __init__(self, items, outdir, threads, delay, parent=None):
        super().__init__(parent)
        self.items = items          # [(fname, billcode, verifycode, date, place)]
        self.outdir = outdir
        self.threads = max(1, int(threads))
        self.delay = max(0.0, float(delay))
        self._stop = threading.Event()
        self._lock = threading.Lock()
        self._last = 0.0
        self._sig_names = (self.sig_log, self.sig_progress, self.sig_done)

    def stop(self):
        self._stop.set()

    def wait_slot(self):
        """全局限速：保证两次请求间隔 >= delay"""
        if self.delay <= 0:
            return
        with self._lock:
            now = time.time()
            delta = now - self._last
            if delta < self.delay:
                time.sleep(self.delay - delta)
            self._last = time.time()

    def job(self, item):
        fname, bill, vc, date_str, place = item
        if self._stop.is_set():
            return None
        try:
            self.wait_slot()
            master, students, examiners = core.fetch_detail(bill, vc)
            d = core.fmt_date((master or {}).get("KAOSRJ")) or date_str
            p = _t((master or {}).get("EXAM_PLACE")) or place or "未知考点"
            title = core.build_title(d, p, examiners)
            df = core.build_rows(students)
            core.write_xlsx(os.path.join(self.outdir, fname + ".xlsx"), title, df)
            return (fname, len(students), len(examiners))
        except Exception as e:
            return (fname, -1, str(e)[:80])

    def run(self):
        from concurrent.futures import ThreadPoolExecutor, as_completed
        ok = fail = students = 0
        done = 0
        total = len(self.items)
        try:
            with ThreadPoolExecutor(max_workers=self.threads) as ex:
                futs = [ex.submit(self.job, it) for it in self.items]
                for fu in as_completed(futs):
                    if self._stop.is_set():
                        self.sig_log.emit("已停止")
                        break
                    try:
                        res = fu.result()
                    except Exception as e:
                        res = None
                        self.sig_log.emit(f"  异常: {str(e)[:80]}")
                    done += 1
                    self.sig_progress.emit(done, total)
                    if res:
                        name, n, extra = res
                        if n < 0:
                            fail += 1
                            self.sig_log.emit(f"  [{done}/{total}] 失败 {name}: {extra}")
                        else:
                            ok += 1
                            students += n
                            self.sig_log.emit(f"  [{done}/{total}] {name}.xlsx  考生 {n} 人")
        finally:
            self.sig_done.emit(ok, fail, students)


# ==================== 单个维度面板 ====================
class ModePanel(QWidget):
    def __init__(self, mode, log_fn, token_getter=None):
        super().__init__()
        self.mode = mode
        self.log = log_fn
        self.token_getter = token_getter or (lambda: core.TOKEN)
        self.rows = []
        self._places = {}
        self._build()

    # 应用当前 Token（查询/刷新前调用）
    def apply_token(self):
        tok = (self.token_getter() or "").strip()
        if tok and tok != core.TOKEN:
            core.set_token(tok)

    def _build(self):
        lay = QVBoxLayout(self)

        # ---- 筛选 ----
        g = QGroupBox("筛选条件")
        gl = QGridLayout(g)
        gl.addWidget(QLabel("考试日期"), 0, 0)
        self.d_begin = QDateEdit()
        self.d_begin.setCalendarPopup(True)
        self.d_begin.setDisplayFormat("yyyy-MM-dd")
        self.d_begin.setDate(QDate.currentDate().addDays(-30))
        self.d_end = QDateEdit()
        self.d_end.setCalendarPopup(True)
        self.d_end.setDisplayFormat("yyyy-MM-dd")
        self.d_end.setDate(QDate.currentDate())
        hb = QHBoxLayout()
        hb.addWidget(self.d_begin)
        hb.addWidget(QLabel("~"))
        hb.addWidget(self.d_end)
        hw = QWidget()
        hw.setLayout(hb)
        gl.addWidget(hw, 0, 1)

        gl.addWidget(QLabel("考试点"), 0, 2)
        self.cb_place = QComboBox()
        self.cb_place.setMinimumWidth(180)
        self.cb_place.addItem("全部", "")
        self._setup_place_fuzzy()   # v1.3：考试点支持输入关键字模糊过滤
        gl.addWidget(self.cb_place, 0, 3)

        # v1.6：考点太多，下拉翻页找不到 → 加一个实时过滤框
        # 输入关键字后，下拉列表只剩匹配的考点；清空恢复全部
        gl.addWidget(QLabel("考点过滤"), 1, 2)
        self.ed_place_filter = QLineEdit()
        self.ed_place_filter.setPlaceholderText("🔍 输入关键字，实时过滤下方考点列表")
        self.ed_place_filter.setClearButtonEnabled(True)
        self.ed_place_filter.textChanged.connect(self._filter_places)
        gl.addWidget(self.ed_place_filter, 1, 3)
        self.btn_query = QPushButton("查询考次")
        self.btn_query.setMinimumWidth(110)
        self.btn_query.clicked.connect(self.do_query)
        gl.addWidget(self.btn_query, 0, 4)
        gl.setColumnStretch(1, 1)
        lay.addWidget(g)

        # ---- 表格 ----
        self.table = QTableWidget(0, len(TABLE_HEADERS))
        self.table.setHorizontalHeaderLabels(TABLE_HEADERS)
        self.table.verticalHeader().setVisible(False)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setAlternatingRowColors(True)
        # 列宽：用户可拖动（覆盖默认自适应）
        hdr = self.table.horizontalHeader()
        hdr.setSectionResizeMode(QHeaderView.Interactive)
        hdr.setStretchLastSection(True)
        hdr.setMinimumSectionSize(60)
        # 列拖拽重排：按下表头后拖动可改变列顺序
        hdr.setSectionsMovable(True)
        hdr.setSectionsClickable(True)
        hdr.setHighlightSections(True)
        hdr.setDefaultAlignment(Qt.AlignCenter)
        # 右键表头：恢复默认列顺序
        hdr.setContextMenuPolicy(Qt.CustomContextMenu)
        hdr.customContextMenuRequested.connect(self.on_header_menu)
        lay.addWidget(self.table, 1)
        # 给列设个初始宽度（避免太挤）
        init_widths = {0: 50, 1: 100, 2: 110, 3: 180, 4: 240, 5: 170, 6: 70}
        for c, w in init_widths.items():
            if c < self.table.columnCount():
                self.table.setColumnWidth(c, w)


        self.lbl_info = QLabel("尚未查询")
        lay.addWidget(self.lbl_info)

        self.qworker = None
        self.pworker = None

    # ---------- 考试点模糊搜索（v1.3） ----------
    def _setup_place_fuzzy(self):
        """考试点下拉可编辑 + 按关键字包含匹配（考试点多，手动翻页太累）"""
        self.cb_place.setEditable(True)
        self.cb_place.setInsertPolicy(QComboBox.NoInsert)  # 手输文本不新增列表项
        comp = QCompleter(self.cb_place.model(), self.cb_place)
        comp.setCaseSensitivity(Qt.CaseInsensitive)
        try:
            comp.setFilterMode(Qt.MatchContains)  # Qt>=5.14：包含即命中（真模糊）
        except Exception:
            pass  # 更旧版本 Qt 自动退化为前缀匹配
        self.cb_place.setCompleter(comp)

    # ---------- 考试点实时过滤（v1.6） ----------
    def _filter_places(self, text):
        """按关键字实时过滤下拉列表：输入“吉林”只留含吉林的考点，清空恢复全部。

        与 QCompleter 的区别：completer 只在输入时给建议；这里直接重建下拉列表，
        点开下拉就能看到过滤后的全部命中项，考试点多时比翻下拉好用得多。
        """
        kw = (text or "").strip().lower()
        cur_name = self.cb_place.currentData()
        self.cb_place.blockSignals(True)
        self.cb_place.clear()
        self.cb_place.addItem("全部", "")
        for title in sorted(self._places.keys()):
            if kw and kw not in title.lower():
                continue
            self.cb_place.addItem(title, self._places[title])
        # 尽量保住用户已选中的考点（若它仍匹配过滤条件）
        idx = self.cb_place.findData(cur_name)
        if idx >= 0:
            self.cb_place.setCurrentIndex(idx)
        self.cb_place.blockSignals(False)

    def _resolve_place(self):
        """把考试点输入解析成接口要发的 name 值；支持关键字模糊匹配"""
        text = self.cb_place.currentText().strip()
        if not text or text == "全部":
            return ""
        idx = self.cb_place.currentIndex()
        if idx >= 0 and self.cb_place.itemText(idx) == text:
            return self.cb_place.currentData() or ""
        # 输入文本没对上列表项 → 按「包含」模糊匹配考点 title
        if text in self._places:
            return self._places[text]
        low = text.lower()
        hits = [t for t in self._places if low in t.lower()]
        if not hits:
            self.log(f"考试点关键字“{text}”未命中任何考点，按【全部】查询")
            return ""
        if len(hits) > 1:
            extra = "、".join(hits[1:5])
            self.log(f"考试点关键字“{text}”命中 {len(hits)} 个考点，按第一个"
                     f"“{hits[0]}”查询（其余：{extra}…）")
        return self._places[hits[0]]

    # ---------- 考试点 ----------
    def set_places(self, places):
        """places: {title: name}"""
        self._places = dict(places or {})
        cur = self.cb_place.currentData()
        # 按当前过滤关键字重建列表（无关键字 = 全量）
        self._filter_places(self.ed_place_filter.text() if hasattr(self, "ed_place_filter") else "")
        if cur:
            idx = self.cb_place.findData(cur)
            if idx >= 0:
                self.cb_place.setCurrentIndex(idx)
        cnt = len(self._places)
        self.lbl_place_count = getattr(self, "lbl_place_count", None)
        self.log(f"考试点列表已加载：{cnt} 个考点"
                 + (f"（过滤后下拉 {self.cb_place.count() - 1} 项）"
                    if hasattr(self, "ed_place_filter") and self.ed_place_filter.text().strip() else ""))

    def refresh_places(self):
        if self.pworker and self.pworker.isRunning():
            return
        self.apply_token()
        w = PlaceWorker(parent=self)
        w.sig_ok.connect(self.set_places)
        w.sig_fail.connect(lambda m: self.log(f"考点列表拉取失败: {m}"))
        _safe_start_worker(self, "pworker", w)

    # ---------- 查询 ----------
    def do_query(self):
        if self.qworker and self.qworker.isRunning():
            return
        self.apply_token()
        self.btn_query.setEnabled(False)
        self.btn_query.setText("查询中...")
        self.log(f"[{self.mode_name()}] 查询中...")
        w = QueryWorker(
            self.mode,
            self.d_begin.date().toString("yyyy-MM-dd"),
            self.d_end.date().toString("yyyy-MM-dd"),
            self._resolve_place(),
            parent=self,
        )
        w.sig_ok.connect(self.on_query_ok)
        w.sig_fail.connect(self.on_query_fail)
        _safe_start_worker(self, "qworker", w)

    def mode_name(self):
        return "按机构" if self.mode == MODE_AGENCY else "按考点"

    def on_header_menu(self, pos):
        """表头右键：恢复默认列顺序 + 自动列宽"""
        # 三种绑定的 exec_/exec 兼容
        try:
            from PySide6.QtWidgets import QMenu  # noqa: F401
            _exec = lambda m, p: m.exec_(p)
        except Exception:
            try:
                from PyQt5.QtWidgets import QMenu  # noqa: F401
                _exec = lambda m, p: m.exec_(p)
            except Exception:
                from PyQt6.QtWidgets import QMenu  # noqa: F401
                _exec = lambda m, p: m.exec(p)

        menu = QMenu(self)
        act_reset = menu.addAction("恢复默认列顺序")
        menu.addSeparator()
        act_fit = menu.addAction("自动列宽（按内容）")
        act_sel = _exec(menu, self.table.horizontalHeader().mapToGlobal(pos))
        if act_sel == act_reset:
            hdr = self.table.horizontalHeader()
            # 顺序：先记录当前 visualIndex 顺序，再倒过来 move 回去
            current = [hdr.logicalIndex(v) for v in range(hdr.count())]
            for new_vis, old_log in enumerate(current):
                if new_vis == old_log:
                    continue
                # 把 logical=old_log 移到 new_vis
                hdr.moveSection(new_vis, old_log)
            # 还原宽度
            init_widths = [50, 100, 110, 180, 240, 170, 70]
            for c, w in enumerate(init_widths):
                if c < self.table.columnCount():
                    self.table.setColumnWidth(c, w)
        elif act_sel == act_fit:
            self.table.resizeColumnsToContents()
            self.table.horizontalHeader().setStretchLastSection(True)

    def on_query_ok(self, rows):
        self.btn_query.setEnabled(True)
        self.btn_query.setText("查询考次")
        self.rows = [r for r in rows if (r or {}).get("BILLCODE")]
        self.fill_table()
        self.log(f"[{self.mode_name()}] 查询完成：{len(self.rows)} 个考次")

    def on_query_fail(self, msg):
        self.btn_query.setEnabled(True)
        self.btn_query.setText("查询考次")
        self.log(f"[{self.mode_name()}] 查询失败: {msg}")
        QMessageBox.warning(self, "查询失败", msg)

    def fill_table(self):
        self.table.setRowCount(len(self.rows))
        for i, r in enumerate(self.rows, 1):
            vals = [
                str(i),
                row_date(r),
                _t(r.get("EXAM_PLACE")),
                _t(r.get("KAOSFWTGF")),
                _t(r.get("COMMISSION")),
                r.get("BILLCODE") or "",
                _t(r.get("BILLSTATE")),
            ]
            for j, v in enumerate(vals):
                item = QTableWidgetItem(v)
                item.setTextAlignment(Qt.AlignCenter)
                self.table.setItem(i - 1, j, item)
        self.lbl_info.setText(f"共 {len(self.rows)} 个考次")

    # ---------- 取待导出项 ----------
    def build_items(self, limit=0):
        rows = self.rows if not limit else self.rows[:limit]
        used = {}
        items = []
        for r in rows:
            d = row_date(r)
            p = _t(r.get("EXAM_PLACE")) or "未知考点"
            base = f"{d}_{core.safe_name(p)}"
            cnt = used.get(base, 0)
            used[base] = cnt + 1
            fname = base if cnt == 0 else f"{base}({cnt + 1})"
            items.append((fname, r.get("BILLCODE"), r.get("verifyCode"), d, p))
        return items


# ==================== 主窗口 ====================
class MainWindow(QMainWindow):
    sig_token = Signal(str)   # CDP 提取到 token 后回填主线程
    sig_log = Signal(str)     # 后台线程写日志（自动转主线程，保证 Qt 线程安全）
    sig_login_done = Signal() # 后台线程结束通知主线程恢复按钮（不要用 QTimer.singleShot）
    _uom_new_ts = 0.0         # _ensure_uom_tab 节流：距上次真正新开 UOM 标签的秒数

    def __init__(self):
        super().__init__()
        self.setWindowTitle(f"{APP_NAME}  v1.7")
        self.resize(1180, 760)
        self.worker = None
        self._login_cancel = threading.Event()  # 一键登录取消信号
        # 按 tab 分别记忆输出目录：{0: 按机构, 1: 按考点}，切 tab 时不互相覆盖
        self._out_memo = {}
        self._last_tab = None
        self._build()
        self._last_tab = self.tabs.currentIndex()
        self.sig_token.connect(self._cdp_apply)
        self.sig_log.connect(self._append_log)
        self.sig_login_done.connect(self._login_ui_done)
        # 启动时自动加载已保存的登录态
        self.load_token()

    def _build(self):
        root = QWidget()
        lay = QVBoxLayout(root)
        lay.setSpacing(8)

        # ---- 顶部：登录态 ----
        hb = QHBoxLayout()
        hb.addWidget(QLabel("登录态:"))
        self.ed_token = QLineEdit(DEFAULT_TOKEN)
        self.ed_token.setEchoMode(QLineEdit.Password)
        self.ed_token.setToolTip("Authorization Token（UUID）。\n"
                                "不用手动管这里：点「一键获取登录态」会自动从 Chrome 提取并保存")
        hb.addWidget(self.ed_token, 1)
        self.btn_tok = QPushButton("显示")
        self.btn_tok.setCheckable(True)
        self.btn_tok.setFixedWidth(50)
        self.btn_tok.toggled.connect(
            lambda c: self.ed_token.setEchoMode(QLineEdit.Normal if c else QLineEdit.Password))
        hb.addWidget(self.btn_tok)
        # —— 主推按钮：一键搞定「开调试Chrome + 提取 + 保存」——
        self.btn_login = QPushButton("一键获取登录态")
        self.btn_login.setToolTip("自动完成：\n"
                                 "1) 若 Chrome 未开调试模式，自动重启 Chrome（会先关掉当前 Chrome）\n"
                                 "2) 打开 UOM 民航局页面\n"
                                 "3) 读取登录状态 → 自动保存。\n"
                                 "若弹出登录页，登录一次即可。")
        self.btn_login.setMinimumWidth(150)
        self.btn_login.setStyleSheet(
            "QPushButton{background:#2d6cdf;color:white;font-weight:bold;"
            "border-radius:5px;padding:5px 16px;}"
            "QPushButton:hover{background:#1e57bd;}"
            "QPushButton:disabled{background:#8faee0;color:#e8eef9;}")
        self.btn_login.clicked.connect(self.onekey_login)
        hb.addWidget(self.btn_login)
        # —— 取消按钮：一键获取卡住时可中断 ——
        self.btn_cancel_login = QPushButton("取消")
        self.btn_cancel_login.setToolTip("一键获取登录态卡住时点这个强制中断")
        self.btn_cancel_login.setEnabled(False)
        self.btn_cancel_login.setFixedWidth(50)
        self.btn_cancel_login.clicked.connect(self.cancel_login)
        hb.addWidget(self.btn_cancel_login)
        self.btn_save = QPushButton("保存登录态")
        self.btn_save.setToolTip("把当前 Token 存到本地 token.txt，下次启动自动加载")
        self.btn_save.clicked.connect(self.save_token)
        hb.addWidget(self.btn_save)
        self.btn_refresh = QPushButton("刷新考试点")
        self.btn_refresh.clicked.connect(self.refresh_places)
        hb.addWidget(self.btn_refresh)
        lay.addLayout(hb)

        # ---- Tab ----
        self.tabs = QTabWidget()
        self.panel_agency = ModePanel(MODE_AGENCY, self.log,
                                      token_getter=lambda: self.ed_token.text())
        self.panel_place = ModePanel(MODE_PLACE, self.log,
                                     token_getter=lambda: self.ed_token.text())
        self.tabs.addTab(self.panel_agency, "按机构查看")
        self.tabs.addTab(self.panel_place, "按考点查看")
        self.tabs.currentChanged.connect(self.on_tab_changed)
        lay.addWidget(self.tabs, 1)

        # ---- 导出参数 ----
        g = QGroupBox("导出参数")
        gl = QGridLayout(g)
        gl.addWidget(QLabel("线程数"), 0, 0)
        self.sp_threads = QSpinBox()
        self.sp_threads.setRange(1, 16)
        self.sp_threads.setValue(4)
        self.sp_threads.setFixedWidth(80)
        gl.addWidget(self.sp_threads, 0, 1)

        gl.addWidget(QLabel("等待时间(秒)"), 0, 2)
        self.sp_delay = QDoubleSpinBox()
        self.sp_delay.setRange(0.0, 30.0)
        self.sp_delay.setSingleStep(0.5)
        self.sp_delay.setValue(1.5)
        self.sp_delay.setFixedWidth(80)
        self.sp_delay.setToolTip("全局请求最小间隔，越小越容易触发 429 限流")
        gl.addWidget(self.sp_delay, 0, 3)

        gl.addWidget(QLabel("导出条数"), 0, 4)
        self.sp_limit = QSpinBox()
        self.sp_limit.setRange(1, 9999)
        self.sp_limit.setValue(20)
        self.sp_limit.setFixedWidth(80)
        gl.addWidget(self.sp_limit, 0, 5)

        self.chk_all = QCheckBox("导出全部")
        self.chk_all.setChecked(False)
        self.chk_all.toggled.connect(lambda c: self.sp_limit.setEnabled(not c))
        gl.addWidget(self.chk_all, 0, 6)

        gl.addWidget(QLabel("输出目录"), 1, 0)
        self.ed_out = QLineEdit(DEFAULT_OUT)
        self.ed_out.setToolTip("每个 Tab 各记一份输出目录，切换时自动保留上次用的")
        # 手动编辑完（失焦/回车）立即记忆到当前 tab
        self.ed_out.editingFinished.connect(self._remember_cur_out)
        gl.addWidget(self.ed_out, 1, 1, 1, 4)
        self.btn_browse = QPushButton("浏览...")
        self.btn_browse.clicked.connect(self.browse)
        gl.addWidget(self.btn_browse, 1, 5, 1, 2)
        lay.addWidget(g)

        # ---- 操作按钮 ----
        hb2 = QHBoxLayout()
        self.btn_export = QPushButton("开始导出")
        self.btn_export.setFixedHeight(34)
        self.btn_export.clicked.connect(self.do_export)
        hb2.addWidget(self.btn_export)
        self.btn_stop = QPushButton("停止")
        self.btn_stop.setFixedHeight(34)
        self.btn_stop.setEnabled(False)
        self.btn_stop.clicked.connect(self.do_stop)
        hb2.addWidget(self.btn_stop)
        self.btn_open = QPushButton("打开输出目录")
        self.btn_open.setFixedHeight(34)
        self.btn_open.clicked.connect(self.open_out)
        hb2.addWidget(self.btn_open)
        hb2.addStretch(1)
        lay.addLayout(hb2)

        self.bar = QProgressBar()
        self.bar.setTextVisible(True)
        lay.addWidget(self.bar)

        self.txt_log = QTextEdit()
        self.txt_log.setReadOnly(True)
        self.txt_log.setMaximumHeight(180)
        lay.addWidget(self.txt_log)

        self.setCentralWidget(root)
        self.log(f"Qt 绑定: {QT_BINDING}")
        self.log("使用流程：① 若登录态为空/已过期，先点顶部蓝色「一键获取登录态」；")
        self.log("          ② 选 Tab -> 设日期 -> 点「查询考次」；③ 点「开始导出」。")
        QTimer.singleShot(300, self.refresh_places)

    # ---------- 辅助 ----------
    def log(self, msg):
        """日志（线程安全：后台线程调用会自动转主线程再写 UI）"""
        if threading.current_thread() is not threading.main_thread():
            self.sig_log.emit(msg)
        else:
            self._append_log(msg)

    def _append_log(self, msg):
        ts = datetime.datetime.now().strftime("%H:%M:%S")
        line = f"[{ts}] {msg}"
        self.txt_log.append(line)
        # 同步打到 console：源码运行直接看终端；exe 冻结模式下
        # stdout 已在 main() 兜底重定向到 kaoci_run.log，不会丢
        out = sys.stdout if sys.stdout is not None else sys.stderr
        if out is not None:
            try:
                out.write(line + "\n")
                out.flush()
            except Exception:
                pass

    def cur_panel(self):
        return self.panel_agency if self.tabs.currentIndex() == 0 else self.panel_place

    def _remember_cur_out(self):
        """把当前输入框里的输出目录，按当前 tab 分别留存。
        按机构(tab0) / 按考点(tab1) 各记一份，互不覆盖。"""
        try:
            key = 0 if self.tabs.currentIndex() == 0 else 1
            self._out_memo[key] = self.ed_out.text().strip()
        except Exception:
            pass

    def on_tab_changed(self, idx):
        """切换 tab：先把「旧 tab」的目录存下来，再恢复「新 tab」上次用过的目录。
        首次进入某 tab 时用其默认目录。"""
        # 先记录离开的那个 tab 的目录（_last_tab 记录上一个 tab 索引）
        last = getattr(self, "_last_tab", None)
        if last is not None and last != idx:
            try:
                if self.ed_out.text().strip():
                    self._out_memo[last] = self.ed_out.text().strip()
            except Exception:
                pass
        # 恢复目标 tab 的目录（没有记录就用默认值）
        default = DEFAULT_OUT if idx == 0 else DEFAULT_OUT_PLACE
        self.ed_out.setText(self._out_memo.get(idx) or default)
        self._last_tab = idx

    def refresh_places(self):
        self.log("正在拉取考试点列表...")
        self.panel_agency.refresh_places()
        self.panel_place.refresh_places()

    # ==================== 登录态（Token）管理 ====================
    def _probe_rows(self):
        """token 有效性探测：不限日期拉考次列表（与考点刷新同参数）。
        注意别只用“当天”——当天没考次时空列表会把有效 token 误判成过期。"""
        try:
            return core.fetch_list("", "", mode=MODE_AGENCY) or []
        except Exception:
            return []

    def _token_is_valid(self, tok):
        """校验候选 token 是否真的有效（调一次接口验证）"""
        try:
            core.set_token(tok)
            return len(self._probe_rows()) > 0
        except Exception:
            return False

    def load_token(self):
        """启动时从 token.txt 读取已保存的登录态"""
        try:
            if os.path.isfile(TOKEN_FILE):
                tok = open(TOKEN_FILE, encoding="utf-8").read().strip()
                if tok:
                    self.ed_token.setText(tok)
                    core.set_token(tok)
                    self.log(f"已从 {os.path.basename(TOKEN_FILE)} 加载登录态")
                    return True
        except Exception:
            pass
        return False

    def save_token(self):
        tok = self.ed_token.text().strip()
        if not tok:
            QMessageBox.warning(self, "提示", "登录态为空，请先填写")
            return
        try:
            with open(TOKEN_FILE, "w", encoding="utf-8") as f:
                f.write(tok)
            core.set_token(tok)
            # 顺带做一次有效性探测（拉最近 1 天的考次数）
            ok = False
            try:
                core.set_token(tok)
                rows = self._probe_rows()
                ok = len(rows) > 0
                self.log(f"登录态已保存，有效性探测：{len(rows)} 条考次可访问")
            except Exception as e:
                self.log(f"登录态已保存，但探测失败: {str(e)[:120]}")
            QMessageBox.information(
                self, "保存完成",
                f"登录态已保存到：\n{TOKEN_FILE}\n\n"
                + ("探测通过：Token 当前有效" if ok else
                   "探测未通过：Token 可能已过期，请点「一键获取登录态」"))
        except Exception as e:
            QMessageBox.warning(self, "错误", f"保存失败: {e}")

    # ==================== 一键获取登录态（CDP 全自动） ====================
    # 面向使用者：只点一个蓝色按钮，后台自动完成——
    #   1) 探测 9223 调试端口是否已开
    #   2) 未开 → 自动重启 Chrome/Edge（调试模式）并打开 UOM
    #   3) 轮询等待登录态 → 提取 UUID Token → 自动保存 token.txt
    # ------------------------------------------------------------------
    @staticmethod
    def _find_chrome():
        """找浏览器可执行文件，优先 Chrome（登录态在其 profile），Edge 仅作兜底。
        返回 (exe路径, 进程名)；都找不到返回 (None, None)。
        支持 Windows（注册表 + 常见路径）和 macOS（/Applications）。"""
        system = sys.platform
        if system == "darwin":
            # macOS：固定路径 / mdfind
            for app, name in (
                ("/Applications/Google Chrome.app/Contents/MacOS/Google Chrome", "Google Chrome"),
                ("/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge", "Microsoft Edge"),
            ):
                if os.path.isfile(app):
                    return app, name
            # mdfind 兜底
            try:
                r = subprocess.run(
                    ["mdfind", "kMDItemCFBundleIdentifier='com.google.Chrome'"],
                    capture_output=True, text=True, timeout=5,
                )
                for line in r.stdout.splitlines():
                    bundle = line.strip()
                    if not bundle:
                        continue
                    exe_path = os.path.join(bundle, "Contents/MacOS/Google Chrome")
                    if os.path.isfile(exe_path):
                        return exe_path, "Google Chrome"
            except Exception:
                pass
            return None, None

        # Windows 分支
        envs = (os.environ.get("LOCALAPPDATA", ""),
                os.environ.get("PROGRAMFILES", ""),
                os.environ.get("PROGRAMFILES(X86)", ""))
        chrome_files, edge_files = [], []
        for base in envs:
            if base:
                chrome_files.append(os.path.join(base, "Google", "Chrome", "Application", "chrome.exe"))
                edge_files.append(os.path.join(base, "Microsoft", "Edge", "Application", "msedge.exe"))

        def _reg(sub):
            """读注册表 App Paths，返回 exe 或 None"""
            try:
                import winreg  # type: ignore
                for root in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
                    try:
                        with winreg.OpenKey(root, sub) as k:
                            exe, _ = winreg.QueryValueEx(k, None)
                            if exe and os.path.isfile(exe):
                                return exe
                    except OSError:
                        pass
            except Exception:
                pass
            return None

        for sub in (r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe",
                    r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe"):
            exe = _reg(sub)
            if exe:
                return exe, os.path.basename(exe)
        for p in chrome_files:
            if os.path.isfile(p):
                return p, os.path.basename(p)
        for sub in (r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\msedge.exe",
                    r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\msedge.exe"):
            exe = _reg(sub)
            if exe:
                return exe, os.path.basename(exe)
        for p in edge_files:
            if os.path.isfile(p):
                return p, os.path.basename(p)
        return None, None

    @staticmethod
    def _cdp_alive(port=CDP_PORT):
        """探测调试端口是否已就绪"""
        try:
            import requests as _rq
            _rq.get(f"http://127.0.0.1:{port}/json/version", timeout=2)
            return True
        except Exception:
            return False

    @staticmethod
    def _chrome_alive(pname):
        """检查浏览器进程是否在跑。Windows 用 tasklist，macOS 用 pgrep。返回 bool"""
        if sys.platform == "darwin":
            try:
                # macOS Chrome/Edge 进程名带空格 → 用 -f 匹配完整命令行
                r = subprocess.run(
                    ["pgrep", "-fl", pname],
                    capture_output=True, text=True, timeout=5,
                )
                return any(pname in line for line in r.stdout.splitlines())
            except Exception:
                return False
        try:
            r = subprocess.run(["tasklist", "/FI", f"IMAGENAME eq {pname}",
                                "/NH", "/FO", "CSV"],
                               capture_output=True, text=True, timeout=5)
            return pname.lower() in r.stdout.lower()
        except Exception:
            return False

    def _launch_debug_browser(self):
        """用「独立调试配置」启动 Chrome，并打开 UOM。

        关键设计（根治方案）：
        * 绝不 taskkill 用户日常 Chrome —— 日常浏览器完全不受影响；
        * 用 --user-data-dir=DEBUG_PROFILE_DIR 单独起一个 Chrome 实例，
          没有 profile 锁冲突，9223 调试端口 100% 能起来；
        * 首次使用需在弹出的 Chrome 里登录一次 UOM（手机号/扫码），
          登录态会持久化在该配置目录，下次自动带登录状态。
        """
        if self._login_cancel.is_set():
            return False
        exe, pname = self._find_chrome()
        if not exe:
            self.log("未找到 Chrome/Edge，无法自动开启。请先安装 Chrome 后重试。")
            return False
        profile = DEBUG_PROFILE_DIR
        try:
            os.makedirs(profile, exist_ok=True)
        except Exception as e:
            self.log(f"无法创建调试配置目录 {profile}: {e}")
            return False
        self.log(f"[1/3] 使用独立调试配置启动 {pname}（不干扰你正在用的浏览器）")
        self.log(f"[2/3] 以调试模式(端口 {CDP_PORT})打开 UOM 页面 ...")
        try:
            # --no-first-run 跳过首次启动欢迎页；--no-default-browser-check 跳过设默认浏览器询问
            # 所有参数都先做 NUL 校验，避免再次触发
            # "source code string cannot contain null bytes" / "embedded null character" 等异常
            args = [exe,
                    f"--user-data-dir={profile}",
                    f"--remote-debugging-port={CDP_PORT}",
                    "--no-first-run",
                    "--no-default-browser-check",
                    "https://uom.caac.gov.cn/"]
            for _a in args:
                if "\x00" in _a:
                    self.log(f"启动参数含 NUL 字符，拒绝启动: {_a!r}")
                    return False
            subprocess.Popen(args)
        except Exception as e:
            self.log(f"启动浏览器失败: {e}")
            return False
        self.log("[3/3] 等待 9223 端口就绪 ...")
        for i in range(25):
            if self._login_cancel.is_set():
                self.log("已取消")
                return False
            if self._cdp_alive():
                self.log(f"  9223 端口已就绪（用时 {i+1} 秒）")
                return True
            if i > 0 and i % 5 == 0:
                self.log(f"  仍在等待 9223 ...（已 {i} 秒）")
            time.sleep(1)
        self.log("  25 秒内端口 9223 未就绪。")
        self.log("  请检查：上次的调试 Chrome 是否还开着？（关掉它再重试）")
        return self._cdp_alive()

    @staticmethod
    def _ensure_uom_tab(port=CDP_PORT):
        """确保调试 Chrome 里有一个 UOM 页面。

        三种情况：
        * 已有 uom 页面 → 不动；
        * 有 caac.gov.cn 相关页面（说明正停在 SSO/登录流程）→ 不重复开标签，
          等用户登录完成后页面自己跳回 UOM；
        * 完全没有任何 caac 页面 → 才新开 UOM 标签，且 10 秒节流防重复。
        """
        import requests as _rq
        try:
            tabs = _rq.get(f"http://127.0.0.1:{port}/json", timeout=3).json()
        except Exception:
            return
        for t in tabs:
            if t.get("type") == "page" and "uom.caac.gov.cn" in (t.get("url") or ""):
                return
        for t in tabs:
            if t.get("type") == "page" and "caac.gov.cn" in (t.get("url") or ""):
                return
        now = time.time()
        if now - MainWindow._uom_new_ts < 10:
            return
        MainWindow._uom_new_ts = now
        try:
            _rq.put(f"http://127.0.0.1:{port}/json/new?https://uom.caac.gov.cn/", timeout=5)
        except Exception:
            pass

    @staticmethod
    def _extract_tokens_from_value(value):
        """从单个 localStorage 值里提取所有 UUID 形态 token。

        UOM 的登录态不是裸 UUID，而是 Base64(JSON) 存在 localStorage 的
        `session_token` 键里，形如：
            W3sibXNnIjoi...  ->  [{"msg":"操作成功","code":0,
                "token":"c20ad1d4-....","username":"xxx"}, ...]
        因此必须：① 值本身就是 UUID → 直接用；
                 ② 值能 Base64 解码 → 递归扫解码后的 JSON 文本里的 UUID。
        """
        import re as _re
        import base64 as _b64
        import json as _json
        uuid_re = _re.compile(r"[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-"
                              r"[0-9a-fA-F]{4}-[0-9a-fA-F]{12}")

        found = []

        def _add(s):
            for m in uuid_re.findall(str(s)):
                if m not in found:
                    found.append(m)

        def _walk(obj, depth=0):
            """递归遍历任意 JSON 结构，抽取所有 UUID 字符串。"""
            if depth > 6:
                return
            if isinstance(obj, str):
                _add(obj)
            elif isinstance(obj, dict):
                for v in obj.values():
                    _walk(v, depth + 1)
            elif isinstance(obj, (list, tuple)):
                for v in obj:
                    _walk(v, depth + 1)

        raw = str(value or "")
        # ① 裸 UUID
        if uuid_re.fullmatch(raw.strip()):
            _add(raw.strip())
            return found
        # ② 尝试 Base64 解码（含 URL-safe / 缺 padding 的情况）
        s = raw.strip().strip('"')
        for cand in {s, s.replace("-", "+").replace("_", "/")}:
            if not cand:
                continue
            pad = cand + "=" * (-len(cand) % 4)
            try:
                dec = _b64.b64decode(pad, validate=False).decode("utf-8", "ignore")
            except Exception:
                continue
            if not dec:
                continue
            try:
                _walk(_json.loads(dec))
            except Exception:
                # 解码后不是合法 JSON，也再正则捞一遍
                _add(dec)
        # ③ 原串里直接正则兜底（可能 token 明文混在其它结构里）
        _add(raw)
        return found

    @staticmethod
    def _scan_token_candidates(port=CDP_PORT):
        """连接调试端口，扫现有 UOM 页面的 localStorage/cookies，
        返回所有 UUID 形态的候选 token 列表（其中可能混有页面缓存的过期旧值，
        由调用方逐个校验后取有效的用）。

        优先用 playwright 的 storage_state() 直接拉取（避免 Runtime.evaluate
        的 JS 注入在某些 Chrome 版本里因 hidden NUL 字符抛
        "source code string cannot contain null bytes"）；拉不到再降级到
        原生 websocket（但带 NUL 校验）。
        """
        import re as _re
        uuid_re = _re.compile(r"^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-"
                              r"[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$")

        def _pick_uuids(d):
            """从 dict（cookies+origins）里收集所有 UUID 格式的 token。
            支持 Base64(JSON) 嵌套（UOM 的 session_token 就是这种）。"""
            found = []
            if not isinstance(d, dict):
                return found
            extract = MainWindow._extract_tokens_from_value
            for src in d.get("origins", []):
                for k, v in (src.get("localStorage") or []):
                    for tok in extract(v):
                        if tok not in found:
                            found.append(tok)
            # 再看 cookies 里有没有 UUID 形态的值
            for c in d.get("cookies", []):
                for tok in extract(c.get("value", "")):
                    if tok not in found:
                        found.append(tok)
            return found

        # 方式1：playwright connect_over_cdp + storage_state（推荐）
        try:
            from playwright.sync_api import sync_playwright
            with sync_playwright() as p:
                try:
                    browser = p.chromium.connect_over_cdp(f"http://127.0.0.1:{port}")
                except Exception:
                    return None
                found = []
                for ctx in browser.contexts:
                    try:
                        st = ctx.storage_state()
                    except Exception:
                        st = None
                    for tok in _pick_uuids(st):
                        if tok not in found:
                            found.append(tok)
                    # storage_state 没命中 → 退到该 ctx 下每个 uom 页面里 evaluate
                    for pg in ctx.pages:
                        if "uom.caac.gov.cn" in (pg.url or ""):
                            try:
                                js = ("(function(){var out={};"
                                      "for(var i=0;i<localStorage.length;i++){"
                                      "var k=localStorage.key(i);"
                                      "out[k]=localStorage.getItem(k);}"
                                      "return out;})()")
                                obj = pg.evaluate(js)
                                if isinstance(obj, dict):
                                    for _k, v in obj.items():
                                        for tok in MainWindow._extract_tokens_from_value(v):
                                            if tok not in found:
                                                found.append(tok)
                            except Exception:
                                continue
                return found
        except ImportError:
            pass

        # 方式2：原生 websocket CDP（带 NUL 校验，避免再次触发该异常）
        try:
            import requests as _rq
            import websocket
        except ImportError:
            return None
        try:
            tabs = _rq.get(f"http://127.0.0.1:{port}/json", timeout=3).json()
        except Exception:
            return None
        target = None
        for t in tabs:
            if t.get("type") == "page" and "uom.caac.gov.cn" in (t.get("url") or ""):
                target = t
                break
        if not target:
            return None
        ws = None
        try:
            import json as _json
            ws = websocket.create_connection(target["webSocketDebuggerUrl"], timeout=8)
            js = ("(function(){var out={};"
                  "for(var i=0;i<localStorage.length;i++){var k=localStorage.key(i);"
                  "out[k]=localStorage.getItem(k);}"
                  "return out;})()")
            # 防御：js 里万一出现 NUL 就直接放弃（避免触发 CDP 端的
            # "source code string cannot contain null bytes" 异常）
            if "\x00" in js:
                return None
            ws.send(_json.dumps({"id": 1, "method": "Runtime.evaluate",
                                 "params": {"expression": js, "returnByValue": True}}))
            deadline = time.time() + 5
            while time.time() < deadline:
                raw = ws.recv()
                if "\x00" in raw:
                    continue
                msg = _json.loads(raw)
                if msg.get("id") != 1:
                    continue
                val = (msg.get("result", {})
                       .get("result", {}).get("value"))
                if isinstance(val, dict):
                    found = []
                    for _k, v in val.items():
                        for tok in MainWindow._extract_tokens_from_value(v):
                            if tok not in found:
                                found.append(tok)
                    return found
                return []
        except Exception:
            return None
        finally:
            if ws:
                try:
                    ws.close()
                except Exception:
                    pass

    # ---------- 一键入口（UI 线程触发，后台线程干活） ----------
    def onekey_login(self):
        if getattr(self, "_login_thread", None) and self._login_thread.isRunning():
            return
        if not self._cdp_alive():
            r = QMessageBox.question(
                self, "开启登录态自动获取",
                "接下来程序会打开一个「独立调试 Chrome」窗口并进入 UOM 页面。\n\n"
                "• 你正在用的日常浏览器完全不受影响\n"
                "• 首次使用请在调试 Chrome 里用手机号/扫码登录一次 UOM\n"
                "• 登录状态会自动保存，之后点查询/导出不再需要手动管\n\n"
                "是否继续？",
                QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
            if r != QMessageBox.Yes:
                self.log("已取消一键获取")
                return
        self._login_cancel.clear()
        self.btn_login.setEnabled(False)
        self.btn_login.setText("获取中...")
        self.btn_cancel_login.setEnabled(True)
        self._login_thread = threading.Thread(target=self._login_worker, daemon=True)
        self._login_thread.start()

    def cancel_login(self):
        """用户点取消按钮：让 worker 在下一个检查点退出"""
        self.log("收到取消指令，正在中断 ...（下次循环/等待时立即退出）")
        self._login_cancel.set()
        self.btn_cancel_login.setEnabled(False)

    def _login_worker(self):
        try:
            if self._login_cancel.is_set():
                return
            # 第 0 步：先看 token.txt 里现有的 token 还活不活着，活着就秒过
            # （用户场景：登录态已经保存过，没必要再启动 Chrome）
            try:
                if os.path.isfile(TOKEN_FILE):
                    existing = open(TOKEN_FILE, encoding="utf-8").read().strip()
                    if existing:
                        core.set_token(existing)
                        rows = self._probe_rows()
                        if rows:
                            self.log(f"现有 token.txt 仍有效（{len(rows)} 条考次可访问），"
                                     f"无需启动 Chrome")
                            # 直接 sig_token，_cdp_apply 里会判断"无变化"只 log 不弹窗
                            self.sig_token.emit(existing)
                            return
                        else:
                            self.log("现有 token.txt 已失效（401/空响应），将从 Chrome 重新获取")
                            # 立即把失效的 token 删掉，避免下次启动还拿这个失败 token 调 API
                            try:
                                os.remove(TOKEN_FILE)
                            except Exception:
                                pass
            except Exception as e:
                self.log(f"现有 token 探测失败: {str(e)[:100]}，将走 Chrome 提取流程")
                # 探测时如果发生意外异常（如连接错误），也把 token.txt 清掉，避免污染
                try:
                    if os.path.isfile(TOKEN_FILE):
                        os.remove(TOKEN_FILE)
                except Exception:
                    pass

            # 第 1 步：现有 token 不可用 → 启动 Chrome + 扫描
            if not self._cdp_alive():
                if not self._launch_debug_browser():
                    return
            if self._login_cancel.is_set():
                return
            self.log("Chrome 调试模式已就绪")
            self.log("  若调试 Chrome 里弹出 UOM 登录页，请用手机号/扫码登录一次（仅首次需要）")
            try:
                self._ensure_uom_tab()
                time.sleep(2)
            except Exception as e:
                self.log(f"打开 UOM 页面失败: {str(e)[:100]}")
            self.log("正在读取登录状态……")
            deadline = time.time() + 90   # 90 秒足够；超时直接放弃避免按钮长时间 disabled
            last_hint = 0.0
            ensure_attempts = 0
            stale = set()   # v1.3 修复：已校验过期的 token 进黑名单，不再盲存
            while time.time() < deadline:
                if self._login_cancel.is_set():
                    self.log("已取消，不再扫描")
                    return
                # 若循环内 UOM 标签丢失（例如重启后第一轮没起），再次确保
                self._ensure_uom_tab()
                ensure_attempts += 1
                cands = self._scan_token_candidates()
                tok = None
                for cand in cands or []:
                    if cand in stale:
                        continue
                    if self._token_is_valid(cand):
                        tok = cand
                        break
                    stale.add(cand)
                if tok:
                    try:
                        with open(TOKEN_FILE, "w", encoding="utf-8") as f:
                            f.write(tok)
                    except Exception as e:
                        self.log(f"保存 token.txt 失败: {e}")
                    self.sig_token.emit(tok)
                    return
                if cands:
                    # 能扫到 token 但全部校验失败 → 是页面/localStorage 里缓存的
                    # 过期旧值，必须在调试 Chrome 里重新登录一次才会产生新 token
                    if time.time() - last_hint >= 20:
                        last_hint = time.time()
                        self.log("扫到的登录态均已过期（页面缓存了旧 token）：请在调试"
                                 "Chrome 的 UOM 页面【退出账号后重新登录一次】，"
                                 "登录成功瞬间程序会自动抓到新 token 并保存 ...")
                else:
                    if time.time() - last_hint >= 20:
                        last_hint = time.time()
                        self.log("还没读到登录态：若 Chrome 里弹出了 UOM 登录页，"
                                 "请先登录一次，程序会自动识别 ...")
                time.sleep(2)
            self.log("90 秒内未读取到有效登录态：请确认 Chrome 里 UOM 已重新登录成功，"
                     "再点一次「一键获取登录态」")
        except Exception as e:
            self.log(f"一键获取异常: {str(e)[:200]}")
        finally:
            # 用 Signal 而非 QTimer.singleShot：singleShot 会在 worker 线程
            # 创建一次性定时器，但 worker 线程没有 QEventLoop，timer 永远不触发，
            # 按钮会卡在"获取中..."。Signal.emit() 走 Qt queued connection，
            # 跨线程投递到主线程执行，安全可靠。
            self.sig_login_done.emit()

    def _login_ui_done(self):
        self.btn_login.setEnabled(True)
        self.btn_login.setText("一键获取登录态")
        self.btn_cancel_login.setEnabled(False)

    def _cdp_apply(self, tok):
        changed = self.ed_token.text().strip() != tok
        self.ed_token.setText(tok)
        core.set_token(tok)
        if changed:
            self.log(f"CDP: 已提取登录态并保存（前8位 {tok[:8]}...）")
            QMessageBox.information(self, "获取成功",
                                    "已从 Chrome 提取登录态并自动保存。\n"
                                    "下次启动本程序将自动加载，无需重复操作。")
        else:
            self.log("CDP: 登录态无变化，仍可继续使用")

    def browse(self):
        d = QFileDialog.getExistingDirectory(self, "选择输出目录", self.ed_out.text())
        if d:
            self.ed_out.setText(d)
            self._remember_cur_out()   # 立刻记到当前 tab，切走再切回不丢

    def open_out(self):
        d = self.ed_out.text().strip()
        if not os.path.isdir(d):
            QMessageBox.information(self, "提示", f"目录不存在：{d}")
            return
        try:
            if sys.platform == "darwin":
                subprocess.Popen(["open", d])
            elif sys.platform == "win32":
                os.startfile(d)  # noqa: PLR5501
            else:
                subprocess.Popen(["xdg-open", d])
        except Exception as e:
            QMessageBox.warning(self, "打开失败", f"无法打开目录：{e}")

    # ---------- 导出 ----------
    def do_export(self):
        panel = self.cur_panel()
        if not panel.rows:
            QMessageBox.information(self, "提示", "请先点「查询考次」")
            return
        token = self.ed_token.text().strip()
        if not token:
            QMessageBox.warning(self, "提示", "请填写 Authorization")
            return
        core.set_token(token)

        outdir = self.ed_out.text().strip()
        try:
            os.makedirs(outdir, exist_ok=True)
        except Exception as e:
            QMessageBox.warning(self, "错误", f"无法创建输出目录: {e}")
            return

        limit = 0 if self.chk_all.isChecked() else self.sp_limit.value()
        items = panel.build_items(limit)
        if not items:
            QMessageBox.information(self, "提示", "没有可导出的考次")
            return

        self.btn_export.setEnabled(False)
        self.btn_stop.setEnabled(True)
        self.bar.setRange(0, len(items))
        self.bar.setValue(0)
        self.log(f"开始导出：{panel.mode_name()} | {len(items)} 个考次 | "
                 f"线程 {self.sp_threads.value()} | 间隔 {self.sp_delay.value()}s")
        self.log(f"输出目录：{outdir}")

        w = ExportWorker(items, outdir, self.sp_threads.value(),
                         self.sp_delay.value(), parent=self)
        w.sig_log.connect(self.log)
        w.sig_progress.connect(
            lambda c, t: (self.bar.setValue(c),
                          self.bar.setFormat(f"{c}/{t}  ({c * 100 // max(t, 1)}%)")))
        w.sig_done.connect(self.on_done)
        _safe_start_worker(self, "worker", w)

    def do_stop(self):
        if self.worker and self.worker.isRunning():
            self.worker.stop()
            self.btn_stop.setEnabled(False)
            self.log("正在停止，等待当前请求结束...")

    def on_done(self, ok, fail, students):
        self.btn_export.setEnabled(True)
        self.btn_stop.setEnabled(False)
        self.log(f"完成：成功 {ok} 个，失败 {fail} 个，考生合计 {students} 人")
        self.bar.setFormat(f"完成：成功 {ok} / 失败 {fail}")
        QMessageBox.information(self, "导出完成",
                                f"成功 {ok} 个\n失败 {fail} 个\n考生合计 {students} 人\n\n"
                                f"目录：{self.ed_out.text()}")

    def closeEvent(self, e):
        if self.worker and self.worker.isRunning():
            r = QMessageBox.question(self, "确认", "导出任务正在运行，确定退出？")
            if r != QMessageBox.Yes:
                e.ignore()
                return
            self.worker.stop()
            self.worker.wait(3000)
        e.accept()


def main():
    # ---- exe 兜底：windowed 模式下 sys.stdout/stderr 为 None，
    #      core 模块里的 print() 会直接抛 AttributeError —— 重定向到日志文件 ----
    if getattr(sys, "frozen", False) and (sys.stdout is None or sys.stderr is None):
        try:
            _logpath = os.path.join(APP_DIR, "kaoci_run.log")
            _fh = open(_logpath, "a", encoding="utf-8", buffering=1)
            sys.stdout = _fh
            sys.stderr = _fh
        except Exception:
            pass

    # ---- 全局异常钩子：把未捕获异常写进日志，避免窗口闪退后无痕 ----
    def _hook(tp, val, tb):
        try:
            import traceback
            with open(os.path.join(APP_DIR, "kaoci_run.log"), "a",
                      encoding="utf-8") as f:
                f.write("\n[未捕获异常]\n")
                traceback.print_exception(tp, val, tb, file=f)
        except Exception:
            pass
        try:
            sys.__excepthook__(tp, val, tb)
        except Exception:
            pass
    sys.excepthook = _hook

    QApplication.setStyle("Fusion")
    app = QApplication(sys.argv)
    w = MainWindow()
    w.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
