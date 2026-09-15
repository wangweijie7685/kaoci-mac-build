# -*- coding: utf-8 -*-
"""
UOM 考次链接批量导出工具 v4
============================
功能：按日期范围拉取 UOM 考次列表 -> 逐考次拉取考生明细 -> 每个考次导出一个 xlsx

输出格式（严格对齐 2026-07-23_吉林01.xlsx）：
    第1行：合并单元格 = 年月日_考点_考试员(去重)
    第2行：21 列考生明细表头
    第3行起：考生明细数据

用法：
    python uom_kaoci_export.py                          # 默认最近30天
    python uom_kaoci_export.py 2026-08-07 2026-09-07    # 指定日期范围
    python uom_kaoci_export.py 2026-08-07 2026-09-07 E:\\xz\\out
    python uom_kaoci_export.py 2026-08-07 2026-09-07 E:\\xz\\out --sample 3     # 只导前3个（测试）
    python uom_kaoci_export.py 2026-08-07 2026-09-07 E:\\xz\\out --delay 1.5    # 请求间隔秒

Token 过期：F12 -> Network -> 复制新的 Authorization 替换下方 TOKEN
"""
import os
import sys
import time
import argparse
import datetime

import requests
import pandas as pd

# ==================== 配置 ====================
TOKEN = "e5cad3fb-236c-4be5-a841-09cefc5d868f"
BASE = "https://uom.caac.gov.cn/api"
HEADERS = {
    "Authorization": TOKEN,
    "Content-Type": "application/json",
    "Accept": "application/json, text/plain, */*",
    "Origin": "https://uom.caac.gov.cn",
    "Referer": "https://uom.caac.gov.cn/",
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/152.0.0.0 Safari/537.36",
    "X-Requested-With": "XMLHttpRequest",
    "language": "zh-CN",
    "timezone": "Asia/Shanghai",
}

# 21 列表头（与参考表完全一致）
DETAIL_COLUMNS = [
    "培训机构", "姓名", "性别", "证件号码", "执照种类", "类别等级", "级别等级",
    "超视距等级", "教员等级", "理论签注类型", "理论考试日期 ", "理论成绩",
    "理论成绩是否通过", "实践签注类型", "综合问答考试日期", "综合问答分数",
    "综合问答是否通过", "飞行考试日期", "飞行是否通过", "地面站考试日期", "地面站是否通过",
]

# 表头 -> 明细接口字段映射
FIELD_MAP = {
    "培训机构":         "AGENCY",
    "姓名":             "XINGM",
    "性别":             "XINGB",
    "证件号码":         "ZHENGJHM",
    "执照种类":         "ZHIZZL",
    "类别等级":         "LEIBDJ",
    "级别等级":         "JIBDJ",
    "超视距等级":       "CHAOSJDJ",
    "教员等级":         "JIAOYDJ",
    "理论签注类型":     "THEORY_TYPE",
    "理论考试日期 ":    "THEORY_DATE",
    "理论成绩":         "THEORY_SCORE",
    "理论成绩是否通过": "THEORY_PASS",
    "实践签注类型":     "PRACTICE_TYPE",
    "综合问答考试日期": "ZONGHWDKSRQ",
    "综合问答分数":     "COMPREHENSIVE_SCORE",
    "综合问答是否通过": "COMPREHENSIVE_PASS",
    "飞行考试日期":     "FEIXKSRQ",
    "飞行是否通过":     "FLIGHT_PASS",
    "地面站考试日期":   "DIMZKSRQ",
    "地面站是否通过":   "GROUND_PASS",
}

def set_token(token):
    """更新 Authorization（GUI 里可改）"""
    global TOKEN
    TOKEN = (token or "").strip()
    HEADERS["Authorization"] = TOKEN


DATE_COLUMNS = {"理论考试日期 ", "综合问答考试日期", "飞行考试日期", "地面站考试日期"}
SCORE_COLUMNS = {"理论成绩", "综合问答分数"}


# ==================== 工具函数 ====================
def _t(v):
    """取枚举字段的 title；普通值原样返回"""
    if isinstance(v, dict):
        return v.get("title") or v.get("showTitle") or v.get("name") or ""
    return v


def fmt_date(v):
    if v is None or v == "":
        return ""
    return str(v).split(" ")[0]


def fmt_score(v):
    if v is None or v == "":
        return ""
    try:
        return f"{float(v):.2f}"
    except (TypeError, ValueError):
        return str(v)


def fmt_value(col, v):
    if col in DATE_COLUMNS:
        return fmt_date(v)
    if col in SCORE_COLUMNS:
        return fmt_score(v)
    if col == "证件号码":
        if isinstance(v, dict):
            return v.get("showTitle") or v.get("title") or ""
        return v or ""
    s = _t(v)
    return "" if s is None else s


def safe_name(s, default="未知"):
    """文件名安全化"""
    s = str(s or default).strip()
    for ch in '\\/:*?"<>|':
        s = s.replace(ch, "_")
    return s or default


# ==================== 接口 ====================
# 两种查看维度：
#   agency = 按机构查看（默认）  -> /billList/data/master/query
#   place  = 按考点查看          -> /billList/excute
LIST_MODES = {
    "agency": {
        "path": "billList/data/master/query",
        "defineName": "UOM_L_UOM_CAOKY_KAOCLJ",
    },
    "place": {
        "path": "billList/excute",
        "defineName": "UOM_L_UOM_CAOKY_KAOCLJ_FWF",
    },
}


def fetch_list(start_date, end_date, page_size=100, mode="agency", exam_place=""):
    """拉取考次列表（自动分页，带 verifyCode）。
    mode: agency=按机构 / place=按考点；exam_place: 考试点 ID（空=全部）

    2026-09-15 重要：UOM 服务端对单页条数加了硬上限 —— limit>100 一律返回
    500（"服务器开小差了"，之前一直用 500 单页拉全量，现已必崩）。
    这里改为 limit=100 + offset 循环翻页，直到取完 total。
    """
    cfg = LIST_MODES.get(mode)
    if not cfg:
        raise ValueError(f"未知模式 {mode}，可选: {list(LIST_MODES)}")

    PAGE_MAX = 100          # 服务端硬上限（实测 101 即 500）
    page_size = min(page_size, PAGE_MAX)

    all_rows = []
    offset = 0
    total = None
    while True:
        payload = {
            "UOM_CAOKY_KAOCLJ_BILLCODE": "",
            "UOM_CAOKY_KAOCLJ_KAOSRJ": [start_date or "", end_date or ""],
            "UOM_CAOKY_KAOCLJ_EXAM_PLACE": exam_place or "",
            "UOM_CAOKY_KAOCLJ_AGENCY": "",
            "UOM_CAOKY_KAOCLJ_UNITCODE": "",
            "sortFields": [],
            "defineName": cfg["defineName"],
            "offset": offset,
            "limit": page_size,
        }
        if mode == "agency":
            # 按机构查看额外带这两个筛选项
            payload["UOM_CAOKY_KAOCLJ_PEIXXX_KAOSFWTGF"] = ""
            payload["UOM_CAOKY_KAOCLJ_REGION"] = ""

        r = requests.post(f"{BASE}/{cfg['path']}", headers=HEADERS,
                          json=payload, timeout=60)
        if r.status_code == 500:
            # 单页超限等服务器侧 500：降半页大小重试一次，仍 500 则报错
            if page_size > 10:
                page_size = max(10, page_size // 2)
                print(f"  服务端 500，单页降到 {page_size} 重试 ...")
                continue
            r.raise_for_status()
        r.raise_for_status()
        d = r.json()
        if not isinstance(d, dict) or "rows" not in d:
            raise RuntimeError(f"列表接口返回异常（Token 可能已过期）: {str(d)[:200]}")
        all_rows.extend(d["rows"])
        if total is None:
            total = d.get("total") or 0
        offset += len(d["rows"])
        print(f"  查看维度: {'按机构' if mode == 'agency' else '按考点'}"
              f" | 进度: {len(all_rows)}/{total}")
        # 取完或本页不足一页时结束
        if len(d["rows"]) < page_size or (total and len(all_rows) >= total) or not d["rows"]:
            break
    # 服务端翻页偶有漂移（取回数 > total），按 total 裁掉尾部多余行
    if total and len(all_rows) > total:
        all_rows = all_rows[:total]
    print(f"  查看维度: {'按机构' if mode == 'agency' else '按考点'}"
          f" | 考次总数: {total}，本次取回: {len(all_rows)}")
    return all_rows


def fetch_detail(bill_code, verify_code, retries=4):
    """拉取单个考次的考生明细（429 限流自动退避重试）"""
    payload = {
        "defineCode": "UOM_B_UOM_CAOKY_KAOCLJ",
        "billCode": bill_code,
        "viewName": "view",
        "triggerOrigin": "PC",
        "verifyCode": verify_code,
        "schemeCode": "",
    }
    last = None
    for attempt in range(retries):
        try:
            r = requests.post(f"{BASE}/bill/view/get", headers=HEADERS,
                              json=payload, timeout=120)
            if r.status_code == 429:
                wait = 10 * (attempt + 1)
                print(f"      429 限流，等待 {wait}s 后重试 ({attempt + 1}/{retries})")
                time.sleep(wait)
                continue
            r.raise_for_status()
            d = r.json()
            break
        except requests.HTTPError as e:
            last = e
            if getattr(e.response, "status_code", None) == 429:
                wait = 10 * (attempt + 1)
                print(f"      429 限流，等待 {wait}s 后重试 ({attempt + 1}/{retries})")
                time.sleep(wait)
                continue
            raise
    else:
        raise last or RuntimeError("明细请求失败（重试耗尽）")
    d = r.json()
    if not isinstance(d, dict) or not d.get("data"):
        return {}, [], []
    data = d["data"].get("data") or {}
    master = (data.get("UOM_CAOKY_KAOCLJ") or [{}])[0]
    students = data.get("UOM_CAOKY_KAOCLJ_PEIXXX") or []
    examiners = []
    for e in (data.get("UOM_CAOKY_KAOCLJ_M") or []):
        name = _t((e or {}).get("BINDINGVALUE"))
        name = str(name).strip() if name is not None else ""
        if name and name not in examiners:
            examiners.append(name)
    return master, students, examiners


# ==================== 导出 ====================
def build_rows(students):
    """明细 -> DataFrame"""
    rows = []
    for s in students:
        s = s or {}
        rows.append({col: fmt_value(col, s.get(FIELD_MAP[col])) for col in DETAIL_COLUMNS})
    return pd.DataFrame(rows, columns=DETAIL_COLUMNS)


def build_title(date_str, place, examiners):
    """第1行抬头：年月日_考点_考试员(去重)"""
    title = f"{date_str}_{place}"
    if examiners:
        title += "_" + ",".join(examiners)
    return title


def write_xlsx(path, title, df):
    """写文件：第1行合并抬头，第2行表头，第3行起数据"""
    from openpyxl.styles import Font, Alignment
    from openpyxl.utils import get_column_letter

    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        # header=False + startrow=2(0基) -> 数据写在 Excel 第3行
        df.to_excel(writer, sheet_name="考次链接", index=False, header=False, startrow=2)
        ws = writer.sheets["考次链接"]

        for j, name in enumerate(DETAIL_COLUMNS, 1):
            ws.cell(row=2, column=j, value=name)

        ws.cell(row=1, column=1, value=title)
        ws.merge_cells(start_row=1, start_column=1,
                       end_row=1, end_column=len(DETAIL_COLUMNS))

        c1 = ws.cell(row=1, column=1)
        c1.font = Font(bold=True, size=12)
        c1.alignment = Alignment(horizontal="center", vertical="center")
        for j in range(1, len(DETAIL_COLUMNS) + 1):
            c = ws.cell(row=2, column=j)
            c.font = Font(bold=True)
            c.alignment = Alignment(horizontal="center", vertical="center")
            ws.column_dimensions[get_column_letter(j)].width = 16
        ws.row_dimensions[1].height = 24
        ws.freeze_panes = "A3"


# ==================== 主流程 ====================
def main():
    ap = argparse.ArgumentParser(description="UOM 考次链接批量导出")
    ap.add_argument("start", nargs="?", help="开始日期 YYYY-MM-DD")
    ap.add_argument("end", nargs="?", help="结束日期 YYYY-MM-DD")
    ap.add_argument("outdir", nargs="?", help="输出目录")
    ap.add_argument("--sample", type=int, default=0, help="只导前 N 个考次（测试用）")
    ap.add_argument("--delay", type=float, default=1.0, help="每次明细请求间隔秒")
    ap.add_argument("--bill", help="只导出指定 BILLCODE（补跑用，逗号分隔多个）")
    ap.add_argument("--mode", default="agency", choices=["agency", "place"],
                    help="查看维度：agency=按机构(默认) / place=按考点")
    args = ap.parse_args()

    today = datetime.date.today()
    end = args.end or today.strftime("%Y-%m-%d")
    start = args.start or (today - datetime.timedelta(days=30)).strftime("%Y-%m-%d")
    outdir = args.outdir or os.path.join(
        os.path.dirname(os.path.abspath(__file__)), "kaoci_files")
    os.makedirs(outdir, exist_ok=True)

    print(f"日期范围: {start} ~ {end}")
    print(f"输出目录: {outdir}")

    print("\n[1/3] 拉取考次列表...")
    rows = fetch_list(start, end, mode=args.mode)
    if args.bill:
        wanted = {b.strip() for b in args.bill.split(",") if b.strip()}
        rows = [r for r in rows if (r or {}).get("BILLCODE") in wanted]
        missing = wanted - {(r or {}).get("BILLCODE") for r in rows}
        if missing:
            print(f"  警告：列表中未找到 {missing}")
    elif args.sample:
        rows = rows[:args.sample]
    print(f"  待处理考次: {len(rows)}")

    print("\n[2/3] 逐考次拉取明细并导出...")
    ok, fail, total_students = 0, 0, 0
    used = {}

    for i, r in enumerate(rows, 1):
        r = r or {}
        bill, vc = r.get("BILLCODE"), r.get("verifyCode")
        if not bill:
            print(f"  [{i}/{len(rows)}] 跳过：无 BILLCODE")
            continue
        try:
            master, students, examiners = fetch_detail(bill, vc)
            date_str = fmt_date((master or {}).get("KAOSRJ") or r.get("KAOSRJ"))
            place = _t((master or {}).get("EXAM_PLACE")) or _t(r.get("EXAM_PLACE")) or "未知考点"

            title = build_title(date_str, place, examiners)
            df = build_rows(students)

            base = f"{date_str}_{safe_name(place)}"
            cnt = used.get(base, 0)
            used[base] = cnt + 1
            fname = base if cnt == 0 else f"{base}({cnt + 1})"

            write_xlsx(os.path.join(outdir, fname + ".xlsx"), title, df)
            ok += 1
            total_students += len(students)
            print(f"  [{i}/{len(rows)}] {fname}.xlsx  "
                  f"考生 {len(students)} 人 / 考试员 {len(examiners)} 人")
        except Exception as e:
            fail += 1
            print(f"  [{i}/{len(rows)}] 失败 {bill}: {str(e)[:100]}")
        time.sleep(args.delay)

    print("\n[3/3] 完成")
    print(f"  成功 {ok} 个，失败 {fail} 个，考生合计 {total_students} 人")
    print(f"  输出目录: {outdir}")


if __name__ == "__main__":
    main()
