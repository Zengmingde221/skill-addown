#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""广告购买商品维度分组聚合脚本（独立实现，仅标准库；供大模型数据分析用）

流程：
  1. 读取 赛狐原表筛选结果文件夹 下 广告购买商品报告.xlsx（固定名优先，缺失回退最新一份；下载阶段已按活动ID关联筛选，口径与其他维度一致）
  2. 读取 广告分析聚合表 下最新一份 广告活动聚合*.csv（排除周维度变体），构建
     广告活动ID → (广告组合ID, 广告组合名称) 映射（购买商品报告无组合列，经活动ID关联）
  3. 读取 赛狐原表数据 下 在线产品明细.csv（fetch_sellfox_online_products.py 导出），
     构建 ASIN → SKU(sku) 映射（供其他ASIN列翻译）
  4. 按 [广告活动ID + ASIN + SKU + 投放 + 定位类型] 分组聚合：
       - 广告活动      : 组内保留唯一值（多值按 | 合并，不删行）
       - 其他ASIN      : 组内去重（含格内逗号拆分），输出 ASIN→SKU 映射字典（紧凑JSON，
                         键序=首次出现顺序）；SKU 查在线产品明细的 sku 列（接口 rows.sku），
                         一个ASIN对应多个SKU时值按 ; 拼接全部去重，明细未收录的ASIN值为空串
       - 其他SKU销量   : SUM，保留 2 位小数
       - 其他SKU销售额 : SUM，保留 2 位小数
  5. 广告组合ID / 广告组合名称 = 关联新增列（活动ID未匹配 → 留空，不删组）
  6. 输出 CSV（UTF-8 BOM）到 广告分析聚合表 文件夹

数据完整性约束（不删除任何数据）：
  - 所有输入行全部参与聚合（各组行数总和 = 输入行数），不按任何条件丢弃行或组
  - 空值照口径处理：空单元格跳过该格累加（该行仍参与分组与计数）；
    空投放/空定位类型等空键归入对应空值组
  - 输入文件只读；遇到非数值文本等异常只报错退出 1，绝不静默跳过
  - Σ各组其他SKU销量 == 原表逐行合计，否则退出 1
  - Σ各组其他SKU销售额 == 原表逐行合计，否则退出 1

用法：
  python aggregate_purchaseditem_by_campaign.py
  python aggregate_purchaseditem_by_campaign.py --input <购买商品报告.xlsx> --campaign-agg <广告活动聚合.csv> --products <在线产品明细.csv> --out <输出目录>

退出码 0 = 成功；1 = 文件缺失/表头缺失/聚合列缺失/非数值文本等（fail-fast）
"""

import argparse
import csv
import json
import re
import sys
import zipfile
from datetime import datetime
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
DEFAULT_INPUT_DIR = BASE_DIR / "赛狐原表筛选结果文件夹"   # 购买商品报告筛选版（活动ID关联）
DEFAULT_RAW_DIR = BASE_DIR / "赛狐原表数据"               # 旧版无筛选原表兜底
DEFAULT_CAMPAIGN_AGG_DIR = BASE_DIR / "广告分析聚合表"
DEFAULT_OUT_DIR = BASE_DIR / "广告分析聚合表"

# 分组键列
KEY_CAMPAIGN_ID = "广告活动ID"
KEY_ASIN = "ASIN"
KEY_SKU = "SKU"
KEY_TARGET = "投放"
KEY_TARGETING = "定位类型"

# 组内保留唯一值列
KEY_CAMPAIGN_NAME = "广告活动"

# 聚合指标列
COL_OTHER_ASIN = "其他ASIN"            # 去重后输出 ASIN→SKU(sku) 映射字典（紧凑JSON，查在线产品明细）
COL_OTHER_QTY = "其他SKU销量"          # SUM
COL_OTHER_SALES = "其他SKU销售额"      # SUM

# 关联新增列
NAME_COL = "广告组合名称"
PID_COL = "广告组合ID"

# 广告活动聚合CSV中的关联列
AGG_KEY_CAMPAIGN_ID = "广告活动ID"
AGG_KEY_PID = "广告组合ID"
AGG_KEY_PNAME = "广告组合名称"

# 在线产品明细CSV（fetch_sellfox_online_products.py 导出，固定名）的列
PRODUCTS_FILE_NAME = "在线产品明细.csv"
PROD_COL_SHOP = "shopId"
PROD_COL_ASIN = "asin"
PROD_COL_SKU = "sku"


def fail(msg):
    print(f"错误: {msg}", file=sys.stderr)
    sys.exit(1)


def latest_file(directory, prefix, suffix, exclude_sub=()):
    """取目录下最新的 前缀*后缀 文件；找不到返回 None。

    exclude_sub: 文件名包含任一片段则排除（如 "_周维度_"）。
    """
    if not directory.is_dir():
        return None
    files = sorted(
        (f for f in directory.iterdir()
         if f.name.startswith(prefix) and f.name.endswith(suffix) and f.is_file()
         and not any(sub in f.name for sub in exclude_sub)),
        key=lambda f: f.stat().st_mtime,
    )
    return files[-1] if files else None


def pick_input(directory, prefix, suffix, fallback_dir=None, exclude_sub=()):
    """报告输入文件选择：固定名优先（最新一次拉取覆盖旧文件），
    缺失则回退目录内最新 前缀*后缀（旧版带日期后缀文件兜底），再回退 fallback_dir。"""
    fixed = Path(directory) / f"{prefix}{suffix}"
    if fixed.exists():
        return fixed
    latest = latest_file(directory, prefix, suffix, exclude_sub)
    if latest is not None:
        return latest
    if fallback_dir is not None:
        fixed2 = Path(fallback_dir) / f"{prefix}{suffix}"
        if fixed2.exists():
            return fixed2
        latest2 = latest_file(fallback_dir, prefix, suffix, exclude_sub)
        if latest2 is not None:
            return latest2
    return None


# ── xlsx 读取（表头自动定位 + 行→dict）────────────────────────────
def load_xlsx_rows(path, header_marker, need_cols):
    """读取 xlsx 第一个工作表，返回 (表头dict{列名:列字母}, 数据行list[dict{列名:文本}])。

    表头行不固定在第1行：扫描前5行，取包含 header_marker 列名的第一行作为表头。
    need_cols 为必须存在的列（缺失 → fail-fast 退出1）。
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
        if header_marker in h:
            header_idx = i
            header = h
            break
    if header_idx is None:
        fail(f"前5行未找到含「{header_marker}」的表头行: {path}")

    absent = [c for c in need_cols if c not in header]
    if absent:
        fail(f"{path.name} 表头缺少所需列: {absent}\n实际表头: {sorted(header)}")

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


# ── 活动ID → (组合ID, 组合名称) 映射（来自广告活动聚合CSV）──────────
def build_campaign_map(campaign_agg_file):
    """读广告活动聚合CSV，返回 {广告活动ID: (广告组合ID, 广告组合名称)}。

    同一活动ID对应多个组合 → fail-fast 退出 1（不应发生，需人工核查）。
    """
    mapping = {}
    with open(campaign_agg_file, encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        cols = reader.fieldnames or []
        for need in (AGG_KEY_CAMPAIGN_ID, AGG_KEY_PID, AGG_KEY_PNAME):
            if need not in cols:
                fail(f"{campaign_agg_file.name} 缺少列「{need}」，实际列: {cols}")
        for row in reader:
            cid = (row.get(AGG_KEY_CAMPAIGN_ID) or "").strip()
            pid = (row.get(AGG_KEY_PID) or "").strip()
            pname = (row.get(AGG_KEY_PNAME) or "").strip()
            if not cid:
                continue
            prev = mapping.get(cid)
            if prev is None:
                mapping[cid] = (pid, pname)
            elif prev != (pid, pname):
                fail(f"活动ID {cid} 在活动聚合表中出现多个组合: {prev} vs {(pid, pname)}，"
                     f"请先核查数据")
    print(f"✓ 活动ID→(组合ID, 组合名称) 映射: {len(mapping)} 个活动")
    return mapping


# ── ASIN→SKU 映射（来自在线产品明细CSV）─────────────────────────────
def build_asin_sku_map(products_file):
    """读在线产品明细CSV，返回 {asin: "sku1;sku2"}（全部去重SKU按首次出现顺序 ; 拼接）。

    同一ASIN在明细表中可对应多个SKU（多变体共用ASIN等，保留不武断取一）；空asin行跳过；
    SKU为空的行不参与拼接，该ASIN值为空串。
    """
    skus_by_asin = {}
    with open(products_file, encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        cols = reader.fieldnames or []
        for need in (PROD_COL_SHOP, PROD_COL_ASIN, PROD_COL_SKU):
            if need not in cols:
                fail(f"{products_file.name} 缺少列「{need}」，实际列: {cols}")
        for row in reader:
            asin = (row.get(PROD_COL_ASIN) or "").strip()
            sku = (row.get(PROD_COL_SKU) or "").strip()
            if not asin:
                continue
            skus = skus_by_asin.setdefault(asin, [])
            if sku and sku not in skus:
                skus.append(sku)
    multi = sum(1 for v in skus_by_asin.values() if len(v) > 1)
    mapping = {a: ";".join(v) for a, v in skus_by_asin.items()}
    print(f"✓ ASIN→SKU(sku) 映射: {len(mapping)} 个ASIN"
          f"（其中 {multi} 个对应多个SKU，值内以 ; 拼接）")
    return mapping


# ── 聚合 ───────────────────────────────────────────────────────────
def aggregate(data):
    """按 (广告活动ID, ASIN, SKU, 投放, 定位类型) 分组聚合。

    返回 (results, raw_sales_total, raw_qty_total)：
      results = [(campaign_id, asin, sku, target, targeting, campaign_name,
                  other_asin_list, sales_sum, n_rows, qty_sum)]
      （other_asin_list = 组内有序去重ASIN列表，ASIN→SKU字典在写CSV时查映射生成）
    所有行全部计入（空键归入空字符串键组），零删行；非数值 fail-fast。
    """
    groups = {}   # key -> {"names": set, "asins": list(有序去重), "qty": float, "sales": float, "n": int}
    empty_counts = {KEY_TARGET: 0, KEY_TARGETING: 0}
    bad_cells = []
    raw_sales_total = 0.0
    raw_qty_total = 0.0
    for r_idx, row in enumerate(data, start=1):
        cid = (row.get(KEY_CAMPAIGN_ID) or "").strip()
        asin = (row.get(KEY_ASIN) or "").strip()
        sku = (row.get(KEY_SKU) or "").strip()
        target = (row.get(KEY_TARGET) or "").strip()
        tgt = (row.get(KEY_TARGETING) or "").strip()
        cname = (row.get(KEY_CAMPAIGN_NAME) or "").strip()
        oasin = (row.get(COL_OTHER_ASIN) or "").strip()
        for k, v in ((KEY_TARGET, target), (KEY_TARGETING, tgt)):
            if not v:
                empty_counts[k] += 1
        key = (cid, asin, sku, target, tgt)
        g = groups.setdefault(
            key, {"names": set(), "asins": [], "asins_seen": set(),
                  "qty": 0.0, "sales": 0.0, "n": 0})
        if cname:
            g["names"].add(cname)
        # 单元格内可含逗号分隔的多个ASIN（如 B094FZMYGF,B0BQJ1NDXQ），先拆分再去重
        for one_asin in (a.strip() for a in oasin.split(",")):
            if one_asin and one_asin not in g["asins_seen"]:
                g["asins_seen"].add(one_asin)
                g["asins"].append(one_asin)
        for col, bucket in ((COL_OTHER_QTY, "qty"), (COL_OTHER_SALES, "sales")):
            v = to_float(row.get(col, ""))
            if v is None:
                s = (row.get(col) or "").strip()
                if s and s != "-":
                    bad_cells.append((r_idx, col, s))
            else:
                g[bucket] += v
                if col == COL_OTHER_QTY:
                    raw_qty_total += v
                else:
                    raw_sales_total += v
        g["n"] += 1
    if bad_cells:
        r, col, s = bad_cells[0]
        fail(f"第{r}行列「{col}」存在非数值文本 {s!r}"
             f"（共{len(bad_cells)}处），请先清洗数据")
    for k, n in empty_counts.items():
        if n:
            print(f"提示: {n} 行「{k}」为空，已归入对应空值组（不删行）")

    results = []
    multi_name_groups = 0
    for gkey, g in groups.items():
        cid, asin, sku, target, tgt = gkey
        if len(g["names"]) > 1:
            multi_name_groups += 1
        cname = " | ".join(sorted(g["names"])) if g["names"] else ""
        results.append((cid, asin, sku, target, tgt, cname,
                        g["asins"], g["sales"], g["n"], g["qty"]))
    if multi_name_groups:
        print(f"提示: {multi_name_groups} 组内「广告活动」存在多个值，已按 | 合并")

    # 排序: 组合名称不在组内，按 活动名称、活动ID、ASIN、SKU、投放、定位类型
    results.sort(key=lambda t: (t[5], t[0], t[1], t[2], t[3], t[4]))
    return results, raw_sales_total, raw_qty_total


# ── 入口 ───────────────────────────────────────────────────────────
def parse_args():
    ap = argparse.ArgumentParser(
        description="广告购买商品报告按 [广告活动ID+ASIN+SKU+投放+定位类型] 分组聚合 → CSV")
    ap.add_argument("--input", default=None,
                    help=f"购买商品报告xlsx（默认取 {DEFAULT_INPUT_DIR} 下最新一份）")
    ap.add_argument("--campaign-agg", default=None,
                    help=f"广告活动聚合CSV，用于活动ID→组合ID/名称关联（默认取最新一份）")
    ap.add_argument("--products", default=None,
                    help=f"在线产品明细CSV，用于其他ASIN→SKU(sku)映射"
                         f"（默认 {DEFAULT_RAW_DIR / PRODUCTS_FILE_NAME}）")
    ap.add_argument("--out", default=str(DEFAULT_OUT_DIR),
                    help=f"输出目录（默认 {DEFAULT_OUT_DIR}）")
    return ap.parse_args()


def main():
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass
    args = parse_args()

    in_file = Path(args.input) if args.input else \
        pick_input(DEFAULT_INPUT_DIR, "广告购买商品报告", ".xlsx",
                   fallback_dir=DEFAULT_RAW_DIR)
    ca_file = Path(args.campaign_agg) if args.campaign_agg else \
        pick_input(DEFAULT_CAMPAIGN_AGG_DIR, "广告活动聚合", ".csv",
                   exclude_sub=("_周维度",))
    prod_file = Path(args.products) if args.products else \
        DEFAULT_RAW_DIR / PRODUCTS_FILE_NAME
    if in_file is None:
        fail(f"{DEFAULT_INPUT_DIR} 下未找到 广告购买商品报告.xlsx（或历史 广告购买商品报告_*.xlsx）")
    if ca_file is None:
        fail(f"{DEFAULT_CAMPAIGN_AGG_DIR} 下未找到 广告活动聚合*.csv（组合ID/名称关联需要）")
    if not prod_file.exists():
        fail(f"未找到 {prod_file}（其他ASIN→SKU映射需要）。"
             f"请先运行: python fetch_sellfox_online_products.py --shop-ids <与报表同批的店铺ID>")

    print("=" * 60)
    print("广告购买商品维度分组聚合")
    print(f"  输入表       : {in_file}")
    print(f"  活动聚合表   : {ca_file}")
    print(f"  在线产品明细 : {prod_file}")
    print(f"  输出目录     : {args.out}")
    print("=" * 60)

    need = [KEY_CAMPAIGN_ID, KEY_ASIN, KEY_SKU, KEY_TARGET, KEY_TARGETING,
            KEY_CAMPAIGN_NAME, COL_OTHER_ASIN, COL_OTHER_QTY, COL_OTHER_SALES]
    header, data = load_xlsx_rows(in_file, KEY_CAMPAIGN_ID, need)
    print(f"✓ 读取完成: {len(data)} 行数据, {len(header)} 列")

    cp_map = build_campaign_map(ca_file)
    asin_sku_map = build_asin_sku_map(prod_file)

    results, raw_sales_total, raw_qty_total = aggregate(data)

    # 完整性校验1: 零删行
    total_group_rows = sum(t[8] for t in results)
    if total_group_rows != len(data):
        fail(f"完整性校验失败: 各组行数合计 {total_group_rows} ≠ 输入行数 {len(data)}")
    print(f"✓ 分组完成: {len(results)} 组（行数合计 {total_group_rows}/{len(data)}，零删行）")

    # 完整性校验2: Σ组内其他SKU销量 == 原表逐行合计
    group_qty_total = sum(t[9] for t in results)
    if abs(group_qty_total - raw_qty_total) > 0.005:
        fail(f"销量校验失败: 各组SUM合计 {group_qty_total:.2f} ≠ "
             f"原表逐行合计 {raw_qty_total:.2f}")

    # 完整性校验3: Σ组内其他SKU销售额 == 原表逐行合计
    group_sales_total = sum(t[7] for t in results)
    if abs(group_sales_total - raw_sales_total) > 0.005:
        fail(f"销售额校验失败: 各组SUM合计 {group_sales_total:.2f} ≠ "
             f"原表逐行合计 {raw_sales_total:.2f}")
    print(f"✓ 其他SKU销量校验: 各组合计 {group_qty_total:.2f} == "
          f"原表合计 {raw_qty_total:.2f}")
    print(f"✓ 其他SKU销售额校验: 各组合计 {group_sales_total:.2f} == "
          f"原表合计 {raw_sales_total:.2f}")

    # 组合ID/名称关联
    unmatched = [t for t in results if t[0] not in cp_map]
    if unmatched:
        print(f"提示: {len(unmatched)} 组的活动ID在活动聚合表中未找到，"
              f"组合ID/组合名称留空（不删组；多为已暂停活动）")
    print(f"✓ 组合关联: {len(results) - len(unmatched)}/{len(results)} 组已匹配")

    # 其他ASIN未出现在在线产品明细 → 字典值留空
    out_asins = {a for t in results for a in t[6]}
    unmatched_asins = sorted(a for a in out_asins if not asin_sku_map.get(a))
    if unmatched_asins:
        print(f"提示: {len(unmatched_asins)} 个其他ASIN未出现在在线产品明细，字典中留空"
              f"（如 {unmatched_asins[:3]}）")

    out_dir = Path(args.out)
    out_dir.mkdir(parents=True, exist_ok=True)
    # 输出命名 = 固定名（同类型仅一个文件；同名存在时直接覆盖，聚合表目录不累积旧文件）
    out_file = out_dir / "购买商品聚合.csv"
    if out_file.exists():
        print(f"提示: 输出文件已存在，覆盖写入 → {out_file.name}")

    header_cols = [NAME_COL, PID_COL, KEY_CAMPAIGN_NAME, KEY_CAMPAIGN_ID,
                   KEY_TARGETING, KEY_ASIN, KEY_SKU, KEY_TARGET,
                   COL_OTHER_ASIN, COL_OTHER_QTY, COL_OTHER_SALES]
    with open(out_file, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(header_cols)
        for cid, asin, sku, target, tgt, cname, oasins, sales, _n, qty in results:
            pid, pname = cp_map.get(cid, ("", ""))
            if oasins:
                oasin_cell = json.dumps(
                    {a: asin_sku_map.get(a, "") for a in oasins},
                    ensure_ascii=False, separators=(",", ":"))
            else:
                oasin_cell = ""
            w.writerow([pname, pid, cname, cid, tgt, asin, sku, target,
                        oasin_cell, f"{round(qty, 2):.2f}", f"{round(sales, 2):.2f}"])

    print("-" * 60)
    print(f"✓ 聚合结果 {len(results)} 组 → {out_file}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
