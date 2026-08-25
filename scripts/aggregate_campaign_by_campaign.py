#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""广告活动维度分组聚合脚本（独立实现，仅标准库；供大模型数据分析用）

流程：
  1. 读取 筛选结果文件夹 下最新一份 广告活动报告_*.xlsx（流式解析，表头行自动定位）
  2. 按 [广告组合ID + 广告活动ID + 定位类型] 分组，按指标列 SUM / AVERAGE 聚合
  3. 读取 赛狐原表数据 下最新一份 广告组合_*.csv，portfolioId=广告组合ID → name，
     作为新列 [广告组合名称] 插入在 [广告组合ID] 之后（聚合完成后新增）
  4. 输出 CSV（UTF-8 BOM）到 广告分析聚合表 文件夹

数据完整性约束（不删除任何数据）：
  - 所有输入行全部参与聚合（各组行数总和 = 输入行数），不按任何条件丢弃行或组
  - 空值照口径处理：空单元格跳过该格累加（该行仍参与其他列）；空ID/空定位归入
    专用组；名称未匹配保留该组、名称留空
  - 输入文件只读；遇到非数值文本等异常只报错退出 1，绝不静默跳过

输出形式（喂大模型口径）：
  - 点击率 / 转化率 / ACoS → 小数×100 保留 2 位，带 %（如 4.35%）
  - ROAS → 倍数保留 2 位，不带 %（如 23.01，行业惯例）
  - 其余 SUM/AVERAGE → 纯数字保留 2 位

用法：
  python aggregate_campaign_by_campaign.py
  python aggregate_campaign_by_campaign.py --input <广告活动报告.xlsx> --portfolio <广告组合.csv> --out <输出目录>

退出码 0 = 成功；1 = 文件缺失/表头缺失/聚合列缺失/非数值文本等（fail-fast）
"""

import argparse
import csv
import re
import sys
import zipfile
from datetime import datetime, timedelta
from pathlib import Path
from xml.etree import ElementTree as ET

# 嵌入式发行版（存在 python311._pth，isolated+safe_path 模式）下脚本所在目录
# 不会自动进入 sys.path，需手动加入才能 import 同目录模块。
sys.path.insert(0, str(Path(__file__).resolve().parent))

from fetch_sellfox_daily_ad_reports import (
    _read_shared_strings,
    _first_worksheet_path,
    _cell_col_letter,
    _local_name,
    _cell_value,
)

# ── 路径常量 ────────────────────────────────────────────────────────
BASE_DIR = Path.home() / "Desktop" / "赛狐广告报表明细"
DEFAULT_INPUT_DIR = BASE_DIR / "赛狐原表筛选结果文件夹"
DEFAULT_PORTFOLIO_DIR = BASE_DIR / "赛狐原表数据"
DEFAULT_OUT_DIR = BASE_DIR / "广告分析聚合表"

# 分组键列（内部三键分组；广告活动名称作为保留唯一值列一并记录）
KEY_PORTFOLIO = "广告组合ID"
KEY_CAMPAIGN_ID = "广告活动ID"
KEY_TARGETING = "定位类型"
KEY_CAMPAIGN_NAME = "广告活动"
NAME_COL = "广告组合名称"
DATE_COL = "日期"               # 周维度分组的日期源列
MONTH_COL = "月份"              # 周维度表新增列: 周归属月 = 所在周周一的年月（如 2026年8月）
WEEK_COL = "周维度"             # 周维度表新增列: 连续周序号 w1、w2…（从数据最早一周累加，跨月不重置）
_EPOCH = datetime(1899, 12, 30)  # Excel 序列号起点

# 指标列定义: 列名 -> ("sum"|"avg", 输出类型)
# 输出类型: num=纯数字2位, pct=百分比(×100带%), roas=倍数(不带%)
METRIC_COLS = [
    ("广告花费", "sum", "num"),
    ("广告曝光量", "sum", "num"),
    ("广告点击量", "sum", "num"),
    ("CPC", "avg", "num"),
    ("广告点击率", "avg", "pct"),
    ("广告转化率", "avg", "pct"),
    ("ACoS", None, "acos"),      # SUM(花费)/SUM(销售额)
    ("ROAS", None, "roas"),      # SUM(销售额)/SUM(花费)
    ("本广告产品订单量", "sum", "num"),
    ("其他产品广告订单量", "sum", "num"),
    ("广告销售额", "sum", "num"),
    ("本广告产品销售额", "sum", "num"),
    ("其他产品广告销售额", "sum", "num"),
    ("广告销量", "sum", "num"),
    ("本广告产品销量", "sum", "num"),
    ("其他产品广告销量", "sum", "num"),
    ("广告订单量", "sum", "num"),
]

WEEK_NOTES = """
周维度表 (广告活动聚合_周维度_*.csv) 补充说明:
  分组键: 广告组合ID + 广告活动ID + 定位类型 + 日期所在周（周一为一周开始，周一~周日）
  日期列: 显示该组实际出现的最早~最晚日期（yyyy-mm-dd ~ yyyy-mm-dd），
    非整周边界；日期为空的行归入空周组，日期列留空（不删行）
  月份/周维度列: 月份 = 所在周周一的年月，跨月周归周一所在月，中文格式 2026年8月
    （避免纯 2026-08 被 Excel 打开时误转成 Jul-26 类日期显示）；
    周维度 = 从输入数据最早一周起的连续周序号（首周 w1，向下累加 w2、w3…，
    跨月不重新计数，月份切换由「月份」列区分）；空周组两列留空
  指标口径与全量聚合表完全一致（SUM/AVERAGE/ACoS/ROAS 同规则）
"""


def fail(msg):
    print(f"错误: {msg}", file=sys.stderr)
    sys.exit(1)


def latest_file(directory, prefix, suffix):
    """取目录下最新的 前缀*后缀 文件；找不到返回 None。"""
    if not directory.is_dir():
        return None
    files = sorted(
        (f for f in directory.iterdir()
         if f.name.startswith(prefix) and f.name.endswith(suffix) and f.is_file()),
        key=lambda f: f.stat().st_mtime,
    )
    return files[-1] if files else None


def pick_input(directory, prefix, suffix, fallback_dir=None):
    """报告输入文件选择：固定名优先（最新一次拉取覆盖旧文件），
    缺失则回退目录内最新 前缀*后缀（旧版带日期后缀文件兜底），再回退 fallback_dir。"""
    fixed = Path(directory) / f"{prefix}{suffix}"
    if fixed.exists():
        return fixed
    latest = latest_file(directory, prefix, suffix)
    if latest is not None:
        return latest
    if fallback_dir is not None:
        fixed2 = Path(fallback_dir) / f"{prefix}{suffix}"
        if fixed2.exists():
            return fixed2
        latest2 = latest_file(fallback_dir, prefix, suffix)
        if latest2 is not None:
            return latest2
    return None


# ── xlsx 读取（表头自动定位 + 行→dict）────────────────────────────
def load_xlsx_rows(path):
    """读取 xlsx 第一个工作表，返回 (表头dict{列名:列字母}, 数据行list[dict{列名:文本}])。

    表头行不固定在第1行：扫描前5行，取包含 KEY_PORTFOLIO 的第一行作为表头。
    fail-fast: 打不开/无工作表/前5行无表头 → 退出1。
    """
    try:
        zf = zipfile.ZipFile(path)
    except Exception as e:
        fail(f"无法打开输入文件 {path}: {e}")
    with zf:
        try:
            shared = _read_shared_strings(zf)
            sheet_path = _first_worksheet_path(zf)
            xml_text = zf.read(sheet_path).decode("utf-8")
        except Exception as e:
            fail(f"读取工作表失败 {path}: {e}")

    m = re.search(r"(<sheetData[^>]*>)(.*?)(</sheetData>)", xml_text, re.S)
    if not m:
        fail(f"工作表缺少 <sheetData> 数据区: {path}")
    rows = re.findall(r"<row\b[^>]*>.*?</row>", m.group(2), re.S)
    if len(rows) < 2:
        fail(f"数据行不足（含表头仅 {len(rows)} 行）: {path}")

    # 表头定位: 前5行中找包含分组键列名的行
    header_idx = None
    header = {}
    for i, row in enumerate(rows[:5]):
        h = {}
        for c in ET.fromstring(row):
            if _local_name(c) != "c":
                continue
            col = _cell_col_letter(c)
            val = _cell_value(c, shared)
            if col and val:
                h[val] = col
        if KEY_PORTFOLIO in h:
            header_idx = i
            header = h
            break
    if header_idx is None:
        fail(f"前5行未找到含「{KEY_PORTFOLIO}」的表头行: {path}")

    # 校验分组键列与聚合所需列全部存在（ACoS/ROAS 由花费/销售额推导，不要求列本身存在）
    need = [KEY_PORTFOLIO, KEY_CAMPAIGN_ID, KEY_TARGETING, KEY_CAMPAIGN_NAME, DATE_COL] + \
           [n for n, _, _ in METRIC_COLS if n not in ("ACoS", "ROAS")]
    absent = [c for c in need if c not in header]
    if absent:
        fail(f"表头缺少聚合所需列: {absent}\n实际表头: {sorted(header)}")

    # 数据行 → {列名: 文本}
    data = []
    for row in rows[header_idx + 1:]:
        vals = {}
        for c in ET.fromstring(row):
            if _local_name(c) != "c":
                continue
            col = _cell_col_letter(c)
            if col:
                vals[col] = _cell_value(c, shared)
        data.append({name: vals.get(letter, "") for name, letter in header.items()})
    return header, data


def to_float(text):
    """文本→float；空/非数值返回 None（空值跳过聚合，非数值交由调用方处理）。"""
    s = (text or "").strip().replace(",", "")
    if not s or s == "-":
        return None
    try:
        return float(s)
    except ValueError:
        return None


# ── 聚合 ───────────────────────────────────────────────────────────
def parse_date(text):
    """[日期]单元格 → datetime.date。

    支持 Excel 序列号文本（如 46252.0 → 2026-08-18）与 yyyy-mm-dd / yyyy/m/d。
    空/"-" → None（空日期归空周组）；无法解析 → 返回 "BAD"（调用方 fail-fast）。
    """
    s = (text or "").strip()
    if not s or s == "-":
        return None
    try:
        f = float(s)
        if 20000 < f < 80000:   # 1954~2089 年范围，排除普通数值误判
            return (_EPOCH + timedelta(days=int(f))).date()
    except ValueError:
        pass
    for fmt in ("%Y-%m-%d", "%Y/%m/%d"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return "BAD"


def month_week_label(wk_iso, first_wk):
    """周键(所在周周一的 yyyy-mm-dd) → (月份, 周维度) 标签；first_wk 为数据最早周键。

    规则: 周归属月份 = 周一所在月（跨月周归周一所在月），中文格式 2026年8月
    （避免被 Excel 误识别为日期）；
    w序号 = 从数据最早一周起的连续序号（首周 w1，每周 +1，跨月不重新计数）。
    空串 → ("", "")。
    """
    if not wk_iso:
        return "", ""
    d = datetime.fromisoformat(wk_iso).date()
    first = datetime.fromisoformat(first_wk).date()
    return f"{d.year}年{d.month}月", f"w{(d - first).days // 7 + 1}"


def aggregate(data, weekly=False):
    """按 (广告组合ID, 广告活动ID, 定位类型) 分组聚合（weekly=True 时追加
    "日期所在周"为附加分组键，周一为一周开始）。

    返回结果列表[(portfolio_id, campaign_name, campaign_id, targeting, 日期范围,
    指标dict, 行数)]；日期范围仅周维度有意义（全量模式恒为空串）。
    数据完整性: 所有行全部计入（空键归入空字符串键组），零删行。
    fail-fast 非数值/非日期。
    """
    groups = {}   # key -> {"names": set, "sum": {}, "avg_sum": {}, "avg_n": {}, "n": int, ...}
    empty_portfolio_rows = 0
    empty_targeting_rows = 0
    empty_date_rows = 0
    bad_cells = []
    for r_idx, row in enumerate(data, start=1):
        pid = (row.get(KEY_PORTFOLIO) or "").strip()
        cid = (row.get(KEY_CAMPAIGN_ID) or "").strip()
        tgt = (row.get(KEY_TARGETING) or "").strip()
        cname = (row.get(KEY_CAMPAIGN_NAME) or "").strip()
        if not pid:
            empty_portfolio_rows += 1
        if not tgt:
            empty_targeting_rows += 1
        wk = ""
        d = None
        if weekly:
            d = parse_date(row.get(DATE_COL, ""))
            if d == "BAD":
                fail(f"第{r_idx}行列「{DATE_COL}」无法解析日期: "
                     f"{(row.get(DATE_COL) or '').strip()!r}")
            if d is None:
                empty_date_rows += 1
            else:
                wk = (d - timedelta(days=d.weekday())).isoformat()  # 所在周周一
        key = (pid, cid, tgt, wk) if weekly else (pid, cid, tgt)
        g = groups.setdefault(
            key, {"names": set(), "sum": {}, "avg_sum": {}, "avg_n": {},
                  "dmin": None, "dmax": None, "n": 0})
        if d is not None:
            if g["dmin"] is None or d < g["dmin"]:
                g["dmin"] = d
            if g["dmax"] is None or d > g["dmax"]:
                g["dmax"] = d
        if cname:
            g["names"].add(cname)
        g["n"] += 1
        for name, agg, _fmt in METRIC_COLS:
            if agg is None:       # ACoS / ROAS 推导列
                continue
            v = to_float(row.get(name, ""))
            if v is None:
                s = (row.get(name) or "").strip()
                if s and s != "-":
                    bad_cells.append((r_idx, name, s))
                continue
            if agg == "sum":
                g["sum"][name] = g["sum"].get(name, 0.0) + v
            else:  # avg
                g["avg_sum"][name] = g["avg_sum"].get(name, 0.0) + v
                g["avg_n"][name] = g["avg_n"].get(name, 0) + 1
    if bad_cells:
        r, name, s = bad_cells[0]
        fail(f"第{r}行列「{name}」存在非数值文本 {s!r}（共{len(bad_cells)}处），请先清洗数据")
    if empty_portfolio_rows:
        print(f"提示: {empty_portfolio_rows} 行广告组合ID为空，已归入空ID组（不删行）")
    if empty_targeting_rows:
        print(f"提示: {empty_targeting_rows} 行定位类型为空，已归入空定位组（不删行）")
    if weekly and empty_date_rows:
        print(f"提示: {empty_date_rows} 行「{DATE_COL}」为空，已归入空周组（不删行）")

    # 周维度连续序号锚点: 全部数据中最早一周的周一（空周组不参与）
    first_wk = ""
    if weekly:
        week_keys = [k[-1] for k in groups if k[-1]]
        first_wk = min(week_keys) if week_keys else ""

    results = []
    for gkey, g in groups.items():
        out = {}
        for name, agg, fmt in METRIC_COLS:
            if fmt == "acos":
                cost = g["sum"].get("广告花费", 0.0)
                sales = g["sum"].get("广告销售额", 0.0)
                out[name] = None if sales == 0 else cost / sales * 100
            elif fmt == "roas":
                cost = g["sum"].get("广告花费", 0.0)
                sales = g["sum"].get("广告销售额", 0.0)
                out[name] = None if cost == 0 else sales / cost
            elif agg == "sum":
                out[name] = g["sum"].get(name, 0.0)
            else:  # avg
                n = g["avg_n"].get(name, 0)
                v = (g["avg_sum"].get(name, 0.0) / n) if n else None
                out[name] = (v * 100) if (v is not None and fmt == "pct") else v
        campaign_name = " | ".join(sorted(g["names"])) if g["names"] else ""
        if weekly:
            pid, cid, tgt, wk = gkey
            drange = "" if g["dmin"] is None else \
                f"{g['dmin'].isoformat()} ~ {g['dmax'].isoformat()}"
            mlabel, wlabel = month_week_label(wk, first_wk)
            results.append((pid, campaign_name, cid, tgt, drange, out, g["n"],
                            mlabel, wlabel))
        else:
            pid, cid, tgt = gkey
            results.append((pid, campaign_name, cid, tgt, "", out, g["n"]))

    # 排序: 空组合ID组最后，其余按(组合ID, 活动名称, 定位类型)；周维度再按日期范围
    if weekly:
        results.sort(key=lambda t: (t[0] == "", t[0], t[1], t[3], t[4]))
    else:
        results.sort(key=lambda t: (t[0] == "", t[0], t[1], t[3]))
    return results


def load_portfolio_names(path):
    """读广告组合CSV → {portfolioId: name}。"""
    names = {}
    with open(path, encoding="utf-8-sig", newline="") as f:
        for row in csv.DictReader(f):
            pid = (row.get("portfolioId") or "").strip()
            if pid:
                names[pid] = (row.get("name") or "").strip()
    return names


def fmt_num(v, fmt):
    if v is None:
        return ""
    if fmt in ("pct", "acos"):   # 百分比列（值已是×100后的数，只补%号）
        return f"{round(v, 2):.2f}%"
    return f"{round(v, 2):.2f}"


# ── 入口 ───────────────────────────────────────────────────────────
def parse_args():
    ap = argparse.ArgumentParser(
        description="广告活动报告按 [广告组合ID+广告活动ID+定位类型] 分组聚合 → CSV")
    ap.add_argument("--input", default=None,
                    help=f"广告活动报告xlsx（默认取 {DEFAULT_INPUT_DIR} 下最新一份）")
    ap.add_argument("--portfolio", default=None,
                    help=f"广告组合CSV（默认取 {DEFAULT_PORTFOLIO_DIR} 下最新一份）")
    ap.add_argument("--out", default=str(DEFAULT_OUT_DIR), help=f"输出目录（默认 {DEFAULT_OUT_DIR}）")
    return ap.parse_args()


def main():
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass
    args = parse_args()

    in_file = Path(args.input) if args.input else pick_input(DEFAULT_INPUT_DIR, "广告活动报告", ".xlsx")
    pf_file = Path(args.portfolio) if args.portfolio else latest_file(DEFAULT_PORTFOLIO_DIR, "广告组合_", ".csv")
    if in_file is None:
        fail(f"{DEFAULT_INPUT_DIR} 下未找到 广告活动报告.xlsx（或历史 广告活动报告_*.xlsx）")
    if pf_file is None:
        fail(f"{DEFAULT_PORTFOLIO_DIR} 下未找到 广告组合_*.csv")

    print("=" * 60)
    print("广告活动维度分组聚合")
    print(f"  输入表   : {in_file}")
    print(f"  组合表   : {pf_file}")
    print(f"  输出目录 : {args.out}")
    print("=" * 60)

    header, data = load_xlsx_rows(in_file)
    print(f"✓ 读取完成: {len(data)} 行数据, {len(header)} 列")

    results = aggregate(data)
    total_group_rows = sum(t[6] for t in results)
    if total_group_rows != len(data):
        fail(f"完整性校验失败: 各组行数合计 {total_group_rows} ≠ 输入行数 {len(data)}")
    print(f"✓ 分组完成: {len(results)} 组（行数合计 {total_group_rows}/{len(data)}，零删行）")

    # 周维度分组（原分组键 + 日期所在周，周一为一周开始）
    results_w = aggregate(data, weekly=True)
    total_w = sum(t[6] for t in results_w)
    if total_w != len(data):
        fail(f"周维度完整性校验失败: 各组行数合计 {total_w} ≠ 输入行数 {len(data)}")
    print(f"✓ 周维度分组完成: {len(results_w)} 组（行数合计 {total_w}/{len(data)}，零删行）")

    names = load_portfolio_names(pf_file)
    matched = sum(1 for t in results if t[0] in names)
    print(f"✓ 名称匹配: {matched}/{len(results)}")

    out_dir = Path(args.out)
    out_dir.mkdir(parents=True, exist_ok=True)
    # 输出命名 = 固定名（同类型仅一个文件；同名存在时直接覆盖，聚合表目录不累积旧文件）
    out_file = out_dir / "广告活动聚合.csv"
    if out_file.exists():
        print(f"提示: 输出文件已存在，覆盖写入 → {out_file.name}")

    header_cols = [KEY_PORTFOLIO, NAME_COL, KEY_CAMPAIGN_NAME, KEY_CAMPAIGN_ID, KEY_TARGETING] + \
                  [n for n, _, _ in METRIC_COLS]
    with open(out_file, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(header_cols)
        for pid, cname, cid, tgt, _dr, metrics, _n in results:
            w.writerow([pid, names.get(pid, ""), cname, cid, tgt] +
                       [fmt_num(metrics.get(m), fmt) for m, _, fmt in METRIC_COLS])

    out_file_w = out_dir / "广告活动聚合_周维度.csv"
    header_cols_w = [KEY_PORTFOLIO, NAME_COL, KEY_CAMPAIGN_NAME, KEY_CAMPAIGN_ID,
                     KEY_TARGETING, DATE_COL, MONTH_COL, WEEK_COL] + \
                    [n for n, _, _ in METRIC_COLS]
    with open(out_file_w, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(header_cols_w)
        for pid, cname, cid, tgt, drange, metrics, _n, mlabel, wlabel in results_w:
            w.writerow([pid, names.get(pid, ""), cname, cid, tgt, drange,
                        mlabel, wlabel] +
                       [fmt_num(metrics.get(m), fmt) for m, _, fmt in METRIC_COLS])

    print("-" * 60)
    print(f"✓ 聚合结果 {len(results)} 组 → {out_file}")
    print(f"✓ 周维度聚合结果 {len(results_w)} 组 → {out_file_w}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
