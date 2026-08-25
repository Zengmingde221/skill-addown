#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""赛狐开放平台 — SP 广告天维度(daily)报告下载脚本（独立实现，仅标准库）

流程：
  1. GET  /api/oauth/v2/token.json                    获取 access_token（本地缓存，24h 有效期内复用）
  2. POST /api/cpc/download/createTask.json           为每种报告类型创建下载任务
  3. POST /api/cpc/download/pageList.json             轮询任务进度，生成完成后下载 Excel 文件
  4. POST /api/cpc/manageData/portfolio.json          导出广告组合基础数据（nextToken 分页 → CSV）

用法：
  python fetch_sellfox_daily_ad_reports.py --shop-ids 128154,128155 --start 2026-08-01 --end 2026-08-12

默认输出（原表与筛选结果分别保留在根目录下的两个子文件夹）：
  <桌面>/赛狐广告报表明细/赛狐原表数据/           原表（7 个下载文件，固定文件名，最新一次拉取覆盖旧文件）
  <桌面>/赛狐广告报表明细/赛狐原表筛选结果文件夹/  筛选结果表（7 个：5 种按运行状态筛选 + 搜索词/购买商品按活动ID关联筛选）

说明：
  - client_id / client_secret 默认内置，可通过环境变量 SF_CLIENT_ID / SF_CLIENT_SECRET 覆盖
  - 文件名 = 报告中文名.xlsx（固定名，不再带日期后缀；本次拉取区间见运行日志）
  - 退出码 0 = 全部报告下载完成；1 = 存在失败/超时
"""

import argparse
import csv
import hashlib
import hmac
import json
import os
import random
import re
import sys
import tempfile
import time
import zipfile
from datetime import datetime
from pathlib import Path
from urllib import error as urlerror
from urllib import parse as urlparse
from urllib import request as urlrequest
from xml.etree import ElementTree as ET

# ── 常量 ────────────────────────────────────────────────────────────
API_BASE = "https://openapi.sellfox.com"
TOKEN_PATH = "/api/oauth/v2/token.json"
CREATE_TASK_PATH = "/api/cpc/download/createTask.json"
PAGE_LIST_PATH = "/api/cpc/download/pageList.json"
PORTFOLIO_PATH = "/api/cpc/manageData/portfolio.json"
GRANT_TYPE = "client_credentials"

# 开发者凭证（默认内置，可用环境变量覆盖）
DEFAULT_CLIENT_ID = "368900"
DEFAULT_CLIENT_SECRET = "9eaec859-fb6a-4709-a98d-7432bcb93c02"

# 7 种报告类型 → 中文名（用于文件命名）
REPORT_TYPES = [
    "adCampaignReport",
    "adGroupReport",
    "adProductReport",
    "adSpaceReport",
    "adSearchTermReport",
    "adPurchasedItemReport",
    "adTargeringReport",
]
REPORT_NAMES = {
    "adCampaignReport": "广告活动报告",
    "adGroupReport": "广告组报告",
    "adProductReport": "广告产品报告",
    "adSpaceReport": "广告位报告",
    "adSearchTermReport": "搜索词报告",
    "adPurchasedItemReport": "广告购买商品报告",
    "adTargeringReport": "投放报告",
}

BASE_OUTPUT_DIR = Path.home() / "Desktop" / "赛狐广告报表明细"                # 根目录
DEFAULT_OUTPUT_DIR = BASE_OUTPUT_DIR / "赛狐原表数据"                        # 原表目录
DEFAULT_FILTERED_DIR = BASE_OUTPUT_DIR / "赛狐原表筛选结果文件夹"             # 筛选结果目录
TOKEN_CACHE_FILE = Path(tempfile.gettempdir()) / "sellfox_ad_report_token.json"
TOKEN_REFRESH_BUFFER = 60       # token 提前 60s 视为过期
HTTP_TIMEOUT = 60               # 接口请求超时（秒）
DOWNLOAD_TIMEOUT = 300          # 报告下载超时（秒）

# 公共错误码（见文档 1759053 公共报错）
CODE_TOKEN_INVALID = 40001
CODE_RATE_LIMIT = 40019

# 重试参数
TASK_INTERVAL = 2.0     # 创建任务之间的间隔（秒）
POLL_INTERVAL = 2.0     # 轮询间隔（秒）
MAX_POLLS = 90          # 最大轮询次数（约 3 分钟）
MAX_RATE_RETRIES = 3    # 限流 40019 时的退避重试次数

# 广告组合（portfolio）导出参数（实测: body 传 shopId(数字)+pageSize, data.nextToken 游标分页）
PORTFOLIO_PAGE_SIZE = 1000      # 每页条数
PORTFOLIO_MAX_PAGES = 100       # 分页安全上限（10 万行），防 nextToken 死循环
PORTFOLIO_PAGE_INTERVAL = 1.0   # 翻页间隔（秒），防 40019
PORTFOLIO_COLUMNS = [
    "shopId", "portfolioId", "name", "servingStatus", "inBudget", "amount",
    "policy", "startDate", "endDate", "creationDate", "lastUpdatedDate",
]


class ApiError(RuntimeError):
    """赛狐接口返回 code != 0 时抛出。"""

    def __init__(self, code, msg, request_id=""):
        self.code = code
        self.msg = msg
        self.request_id = request_id
        detail = f"code={code} msg={msg}"
        if request_id:
            detail += f" requestId={request_id}"
        super().__init__(detail)


# ── 凭证 ────────────────────────────────────────────────────────────
def client_id():
    return os.environ.get("SF_CLIENT_ID", DEFAULT_CLIENT_ID)


def client_secret():
    return os.environ.get("SF_CLIENT_SECRET", DEFAULT_CLIENT_SECRET)


# ── Token 获取与缓存 ────────────────────────────────────────────────
def load_cached_token():
    """读取本地缓存的 token；未过期则返回，否则返回 None。"""
    try:
        data = json.loads(TOKEN_CACHE_FILE.read_text(encoding="utf-8"))
        if data.get("expires_at", 0) - TOKEN_REFRESH_BUFFER > time.time():
            return data.get("access_token")
    except Exception:
        pass
    return None


def save_token_cache(token, expires_in_ms):
    TOKEN_CACHE_FILE.write_text(
        json.dumps(
            {
                "access_token": token,
                "expires_at": time.time() + expires_in_ms / 1000.0,
                "fetched_at": time.time(),
            },
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )


def fetch_token():
    """GET token 接口获取 access_token（凭证默认内置）。"""
    query = urlparse.urlencode(
        {
            "client_id": client_id(),
            "client_secret": client_secret(),
            "grant_type": GRANT_TYPE,
        }
    )
    url = f"{API_BASE}{TOKEN_PATH}?{query}"
    req = urlrequest.Request(url, method="GET")
    try:
        with urlrequest.urlopen(req, timeout=HTTP_TIMEOUT) as resp:
            payload = json.loads(resp.read().decode("utf-8"))
    except urlerror.HTTPError as e:
        body = e.read().decode("utf-8", errors="replace")[:300]
        raise RuntimeError(f"获取 token 失败 HTTP {e.code}: {body}")
    except Exception as e:
        raise RuntimeError(f"获取 token 网络错误: {e}")

    if payload.get("code") != 0:
        raise RuntimeError(
            f"获取 token 失败 code={payload.get('code')} msg={payload.get('msg')}"
        )
    data = payload.get("data") or {}
    token = data.get("access_token")
    if not token:
        raise RuntimeError("token 响应中缺少 access_token")
    save_token_cache(token, data.get("expires_in", 86400000))
    return token


def get_access_token(force_refresh=False):
    """优先复用缓存 token，缓存缺失/过期或强制刷新时重新获取。"""
    if not force_refresh:
        cached = load_cached_token()
        if cached:
            return cached
    return fetch_token()


# ── 签名与请求 ──────────────────────────────────────────────────────
def build_signed_url(access_token, method, path):
    """按文档 1749562 生成 HmacSHA256 签名，返回带签名参数的完整 URL。

    参与签名的参数：access_token/client_id/method/nonce/timestamp/url，
    按键名排序后用 & 拼接，以 client_secret 为密钥做 HmacSHA256（hex 小写）。
    requestBody 业务参数不参与签名。
    """
    params = {
        "access_token": access_token,
        "client_id": client_id(),
        "method": method,
        "nonce": str(random.randint(10000, 99999)),
        "timestamp": str(int(time.time() * 1000)),
        "url": path,
    }
    raw = "&".join(f"{k}={params[k]}" for k in sorted(params))
    signature = hmac.new(
        client_secret().encode("utf-8"), raw.encode("utf-8"), hashlib.sha256
    ).hexdigest()

    query = urlparse.urlencode(
        {
            "access_token": access_token,
            "client_id": client_id(),
            "nonce": params["nonce"],
            "sign": signature,
            "timestamp": params["timestamp"],
        }
    )
    return f"{API_BASE}{path}?{query}"


def api_post(path, body):
    """带签名的 POST JSON 请求，返回响应的 data。

    内部处理：
      - token 失效（40001 或 HTTP 401）→ 强制刷新 token 后重试
      - 限流 40019 → 递增退避重试
      - 网络异常 → 短暂等待重试
    """
    access_token = get_access_token()
    last_err = None

    for attempt in range(MAX_RATE_RETRIES + 1):
        url = build_signed_url(access_token, "post", path)
        data = json.dumps(body, ensure_ascii=False).encode("utf-8")
        req = urlrequest.Request(
            url, data=data, headers={"Content-Type": "application/json"}, method="POST"
        )
        try:
            with urlrequest.urlopen(req, timeout=HTTP_TIMEOUT) as resp:
                text = resp.read().decode("utf-8")
        except urlerror.HTTPError as e:
            text = e.read().decode("utf-8", errors="replace")
            if e.code == 401:
                print("    token 失效(HTTP 401)，刷新后重试...", flush=True)
                access_token = get_access_token(force_refresh=True)
                continue
            raise RuntimeError(f"HTTP {e.code}: {text[:300]}")
        except Exception as e:
            last_err = e
            wait = 2 * (attempt + 1)
            print(f"    网络异常({e})，{wait}s 后重试...", flush=True)
            time.sleep(wait)
            continue

        try:
            result = json.loads(text)
        except json.JSONDecodeError:
            raise RuntimeError(f"响应不是合法 JSON: {text[:300]}")

        code = result.get("code")
        if code == 0:
            return result.get("data")
        if code == CODE_RATE_LIMIT:
            wait = 2 * (attempt + 1)
            print(
                f"    限流 40019，{wait}s 后重试 ({attempt + 1}/{MAX_RATE_RETRIES + 1})...",
                flush=True,
            )
            time.sleep(wait)
            continue
        if code == CODE_TOKEN_INVALID:
            print("    token 失效(40001)，刷新后重试...", flush=True)
            access_token = get_access_token(force_refresh=True)
            continue
        raise ApiError(code, result.get("msg", ""), result.get("requestId", ""))

    raise RuntimeError(f"接口调用失败: {path} — {last_err or '重试次数用尽'}")


# ── 创建下载任务 ────────────────────────────────────────────────────
def create_download_tasks(shop_ids, start, end, report_types, interval=TASK_INTERVAL):
    """逐个创建报告下载任务，返回任务列表。

    返回项: {"report_type", "name", "task_id", "error"}
    单个任务失败不中断其余任务；创建间隔默认 2s（限流 1 次/秒）。
    """
    tasks = []
    total = len(report_types)

    for idx, rtype in enumerate(report_types, start=1):
        body = {
            "shopIds": [str(s) for s in shop_ids],
            "adTypeCode": "sp",
            "reportTypeCode": rtype,
            "timeUnit": "daily",
            "reportStartDate": start,
            "reportEndDate": end,
        }
        name = REPORT_NAMES[rtype]
        print(f"[{idx}/{total}] 创建任务: {name} ...", end=" ", flush=True)
        try:
            data = api_post(CREATE_TASK_PATH, body)
        except Exception as e:
            print(f"✗ 失败: {e}")
            tasks.append({"report_type": rtype, "name": name, "task_id": None, "error": str(e)})
            continue

        task_id = (data or {}).get("id")
        if not task_id:
            print("✗ 失败: 响应缺少任务 id")
            tasks.append({"report_type": rtype, "name": name, "task_id": None, "error": "响应缺少任务 id"})
        else:
            print(f"✓ 任务ID={task_id}")
            tasks.append({"report_type": rtype, "name": name, "task_id": str(task_id), "error": None})

        if idx < total:
            time.sleep(interval)

    return tasks


# ── 轮询与下载 ──────────────────────────────────────────────────────
def fetch_task_status():
    """POST pageList 获取全部下载任务状态，返回 {任务ID: 行} 映射。"""
    data = api_post(PAGE_LIST_PATH, {})
    rows = (data or {}).get("rows") or []
    return {str(r.get("id")): r for r in rows if r.get("id") is not None}


def poll_and_download(tasks, out_dir, start, end, poll_interval=POLL_INTERVAL, max_polls=MAX_POLLS):
    """轮询任务进度并下载文件。

    返回 (已下载列表, 未完成列表)，未完成项为 (task, 最后状态)。
    文件名 = {报告中文名}.xlsx（固定名，不再带日期后缀）；日期范围取请求传入的 start~end。
    """
    pending = [t for t in tasks if t.get("task_id")]
    finished = []
    seen_state = {}

    for poll_no in range(1, max_polls + 1):
        time.sleep(poll_interval)

        try:
            status_map = fetch_task_status()
        except Exception as e:
            print(f"  [{poll_no}] 轮询失败: {e}，继续重试...", flush=True)
            continue

        still_pending = []
        for t in pending:
            row = status_map.get(t["task_id"])
            if row is None:
                still_pending.append(t)
                continue

            state = row.get("reportState", "")
            seen_state[t["task_id"]] = state
            if state != "已生成":
                still_pending.append(t)
                continue

            url = row.get("downloadUrl")
            if isinstance(url, list):
                url = url[0] if url else None
            if not url:
                still_pending.append(t)
                continue

            fname = f"{t['name']}.xlsx"
            dest = out_dir / fname
            if dest.exists():
                print(f"  提示: 文件已存在，覆盖下载 → {fname}")
            print(f"  [{poll_no}] 下载: {fname} ...", end=" ", flush=True)
            try:
                size = download_file(str(url), dest)
                print(f"✓ {size:,} 字节")
                t["file"] = str(dest)
                finished.append(t)
            except Exception as e:
                print(f"✗ {e}，稍后重试")
                still_pending.append(t)

        pending = still_pending
        if not pending:
            break
        summary = ", ".join(f"{t['name']}({t['task_id']})" for t in pending)
        print(f"  [{poll_no}] 剩余 {len(pending)} 个任务等待中: {summary}", flush=True)

    unfinished = [(t, seen_state.get(t["task_id"], "未出现在任务列表中")) for t in pending]
    return finished, unfinished


def encode_url(url):
    """只编码 URL 中的非 ASCII 字符（避免对已有百分号编码二次编码）。"""
    parts = urlparse.urlsplit(url)

    def _encode_non_ascii(text):
        return "".join(urlparse.quote(ch) if ord(ch) > 127 else ch for ch in text)

    path = "/".join(_encode_non_ascii(seg) for seg in parts.path.split("/"))
    query = _encode_non_ascii(parts.query)
    return urlparse.urlunsplit((parts.scheme, parts.netloc, path, query, parts.fragment))


def download_file(url, dest):
    """下载报告文件到 dest，返回字节数；空文件视为失败。"""
    encoded = encode_url(url)
    req = urlrequest.Request(encoded, method="GET")
    total = 0
    with urlrequest.urlopen(req, timeout=DOWNLOAD_TIMEOUT) as resp, open(dest, "wb") as fh:
        while True:
            chunk = resp.read(64 * 1024)
            if not chunk:
                break
            fh.write(chunk)
            total += len(chunk)
    if total <= 0:
        dest.unlink(missing_ok=True)
        raise RuntimeError("下载内容为空")
    return total


# ── 数据筛选（下载后按运行状态过滤，fail-fast + 显示保真）───────────
# 只删除整行 <row>，保留单元格 XML 逐字节不变 → 显示/格式/精度不改变，
# 不会引入科学计数法。状态列缺失等异常立即抛出（fail-fast）。
FILTER_RULES = {
    "adCampaignReport": ("广告活动运行状态", "已开启"),
    "adGroupReport": ("广告组运行状态", "已开启"),
    "adProductReport": ("广告产品运行状态", "已开启"),
    "adSpaceReport": ("广告活动运行状态", "已开启"),
    "adTargeringReport": ("投放运行状态", "已开启"),
    # adSearchTermReport / adPurchasedItemReport: 无"运行状态"列，改按活动ID关联筛选（见 JOIN_FILTER_RULES）
}

# 搜索词/购买商品报告无"运行状态"列，无法按状态筛选；
# 改为按 广告活动ID ∈ 筛选后活动报告(adCampaignReport) 关联剔除已暂停活动行，
# 使 5 个聚合维度口径一致（都只统计"开启中"活动）。
JOIN_FILTER_RULES = {
    "adSearchTermReport": "广告活动ID",
    "adPurchasedItemReport": "广告活动ID",
}

_XLSX_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"


def _read_shared_strings(zf):
    """读取 xl/sharedStrings.xml 的共享字符串列表（文件缺失返回空表）。"""
    if "xl/sharedStrings.xml" not in zf.namelist():
        return []
    root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
    out = []
    for si in root.iter(f"{{{_XLSX_NS}}}si"):
        out.append("".join(t.text or "" for t in si.iter(f"{{{_XLSX_NS}}}t")))
    return out


def _first_worksheet_path(zf):
    """定位 workbook 引用的第一个工作表 XML 路径；找不到抛异常。"""
    names = zf.namelist()
    if "xl/workbook.xml" in names and "xl/_rels/workbook.xml.rels" in names:
        try:
            wb = ET.fromstring(zf.read("xl/workbook.xml"))
            rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
            rns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
            sheet = wb.find(f"{{{_XLSX_NS}}}sheets/{{{_XLSX_NS}}}sheet")
            if sheet is not None:
                rid = sheet.get(f"{{{rns}}}id")
                for rel in rels:
                    if rel.get("Id") == rid:
                        target = rel.get("Target") or ""
                        if target.startswith("/"):
                            return target.lstrip("/")
                        return "xl/" + target.lstrip("/")
        except Exception:
            pass
    matches = sorted(n for n in names if re.match(r"xl/worksheets/sheet\d+\.xml$", n))
    if not matches:
        raise RuntimeError("xlsx 中未找到工作表 (xl/worksheets/sheet*.xml)")
    return matches[0]


def _cell_col_letter(cell_el):
    """取单元格列字母（由 r 属性，如 X2 → X）。"""
    m = re.match(r"([A-Z]+)", cell_el.get("r") or "")
    return m.group(1) if m else ""


def _local_name(el):
    """取元素本地名（兼容有无命名空间前缀）。"""
    return el.tag.split("}")[-1]


def _cell_value(cell_el, shared):
    """取单元格显示文本（命名空间无关；行片段单独解析时无默认命名空间）。

    支持 t="s"（共享字符串）/ t="inlineStr" / t="str" / 数字。
    """
    t = cell_el.get("t")
    v = None
    for child in cell_el:
        if _local_name(child) == "v":
            v = child
            break
    if t == "s" and v is not None and v.text is not None:
        idx = int(v.text)
        return shared[idx] if idx < len(shared) else ""
    if t == "inlineStr":
        return "".join(x.text or "" for x in cell_el.iter() if _local_name(x) == "t")
    if t == "str" and v is not None and v.text is not None:
        return v.text
    if v is not None and v.text is not None:
        return v.text
    return ""


def _cell_text_from_row(row_xml, col_letter, shared):
    """取一行 XML 中指定列字母单元格的显示文本；该列无单元格返回 ""。"""
    root = ET.fromstring(row_xml)
    for c in root:
        if _local_name(c) != "c" or _cell_col_letter(c) != col_letter:
            continue
        return _cell_value(c, shared)
    return ""


def _renumber_row(row_xml, new_row_no):
    """把行与所有单元格的 r 属性改为新行号（删除行后需连续编号，否则 Excel 显示空行）。"""
    out = re.sub(
        r'^<row\b([^>]*?)\br="\d+"', f'<row\\1 r="{new_row_no}"', row_xml, count=1
    )
    out = re.sub(
        r'\br="([A-Z]+)\d+"', lambda m: f'r="{m.group(1)}{new_row_no}"', out
    )
    return out


def _filter_sheet(xml_text, shared, col_name, keep):
    """在 sheet XML 文本上按"某列取值是否满足 keep(值)"过滤（只删除整行，保留行除行号外逐字节不变）。

    返回 (新XML文本, 保留数据行数, 剔除行数)。
    列缺失 / 无表头 / 无数据区 → 抛 RuntimeError（fail-fast）。
    """
    m = re.search(r"(<sheetData[^>]*>)(.*?)(</sheetData>)", xml_text, re.S)
    if not m:
        raise RuntimeError("工作表缺少 <sheetData> 数据区")
    # 保留 <sheetData> 开标签与 </sheetData> 闭标签（Excel 依赖该包裹，缺失则整表显示为空）
    prefix = xml_text[: m.start()] + m.group(1)
    body = m.group(2)
    suffix = m.group(3) + xml_text[m.end():]

    rows = re.findall(r"<row\b[^>]*>.*?</row>", body, re.S)
    if not rows:
        raise RuntimeError("工作表没有数据行（表头缺失）")

    # 表头：按列字母定位筛选列
    header_row = rows[0]
    header_col = ""
    for c in ET.fromstring(header_row):
        if _local_name(c) != "c":
            continue
        col = _cell_col_letter(c)
        if col and _cell_value(c, shared) == col_name:
            header_col = col
            break
    if not header_col:
        raise RuntimeError(f"表头中未找到筛选列: {col_name}")

    kept = [header_row]
    kept_count = 0
    dropped = 0
    new_row_no = 2
    for row in rows[1:]:
        if keep(_cell_text_from_row(row, header_col, shared)):
            kept.append(_renumber_row(row, new_row_no))
            new_row_no += 1
            kept_count += 1
        else:
            dropped += 1

    new_body = "".join(kept)
    max_row = new_row_no - 1

    # 同步更新 dimension / autoFilter 的行范围
    def _fix_ref(tag):
        def repl(mm):
            parts = mm.group(1).split(":")
            m3 = re.match(r"^([A-Z]+)(\d+)$", parts[0])
            if not m3:
                return mm.group(0)
            col1, r1 = m3.group(1), int(m3.group(2))
            col2 = col1
            if len(parts) > 1:
                m4 = re.match(r"^([A-Z]+)(\d+)$", parts[1])
                if m4:
                    col2 = m4.group(1)
            return f'<{tag} ref="{col1}{r1}:{col2}{max(max_row, r1)}"/>'
        return repl

    new_prefix = re.sub(
        r'<dimension\b[^>]*ref="([^"]+)"[^>]*/>', _fix_ref("dimension"), prefix, count=1
    )
    new_prefix = re.sub(
        r'<autoFilter\b[^>]*ref="([^"]+)"[^>]*/>',
        _fix_ref("autoFilter"), new_prefix, count=1,
    )
    return new_prefix + new_body + suffix, kept_count, dropped


def filter_worksheet(xml_text, shared, status_name, keyword):
    """按状态列过滤：保留 状态列取值 含 keyword 的行（只删除整行，保留行除行号外逐字节不变）。

    返回 (新XML文本, 保留数据行数, 剔除行数)。
    状态列缺失 / 无表头 / 无数据区 → 抛 RuntimeError（fail-fast）。
    """
    return _filter_sheet(xml_text, shared, status_name, lambda v: keyword in v)


def filter_worksheet_by_set(xml_text, shared, col_name, allowed):
    """按集合过滤：保留 指定列取值 ∈ allowed 的行（只删除整行）。

    用于搜索词/购买商品报告按 广告活动ID ∈ 筛选后活动报告 剔除已暂停活动行。
    """
    return _filter_sheet(xml_text, shared, col_name, lambda v: v in allowed)


def _parse_cellxfs_numfmt(styles_xml):
    """解析 styles.xml 的 cellXfs，返回 [numFmtId, ...]（单元格 s 索引 → 数字格式）。"""
    root = ET.fromstring(styles_xml)
    xfs = root.find(f"{{{_XLSX_NS}}}cellXfs")
    if xfs is None:
        return []
    return [int(xf.get("numFmtId", "0")) for xf in xfs]


def _scientific_notation_hits(xml_text, cellxfs):
    """防御性巡检：数字单元格在 **General 格式** 下会显示为科学计数法的值。

    仅拦截：值含 e/E 或 ≥11 位整数，且有效数字格式为 General（numFmtId 0）。
    带百分比/货币/日期等格式的值显示正常，不拦截。
    """
    hits = []
    for m in re.finditer(r"<c\b([^>]*?)>(?:<v>([^<]*)</v>)?", xml_text):
        attrs, v = m.group(1), m.group(2)
        if v is None:
            continue
        # 文本类单元格（s/str/inlineStr/b/e/d）不检查；数字单元格（无 t 或 t="n"）检查
        if re.search(r'\bt="(?:s|str|inlineStr|b|e|d)"', attrs):
            continue
        risky = re.match(r"^-?\d{11,}$", v) or re.search(r"[eE]", v)
        if not risky:
            continue
        s_idx = None
        ms = re.search(r'\bs="(\d+)"', attrs)
        if ms:
            s_idx = int(ms.group(1))
        numfmt = cellxfs[s_idx] if s_idx is not None and s_idx < len(cellxfs) else 0
        if numfmt == 0:
            hits.append(f"值={v} (numFmtId={numfmt})")
    return hits


def _filter_report(path, sheet_filter, out_path=None):
    """把 sheet 过滤函数应用到 xlsx 并输出（保真：只改工作表内容，其余原样拷贝）。

    - out_path 为空 → 原位覆盖 path（写 .tmp 后原子替换）
    - out_path 非空 → 写入 out_path，原表 path 保持不动

    返回 (保留行数, 剔除行数, 耗时秒)。任何异常抛 RuntimeError，
    原文件保持下载原样。
    """
    t0 = time.time()
    target = str(out_path) if out_path else str(path)
    with zipfile.ZipFile(path, "r") as zin:
        infos = zin.infolist()
        shared = _read_shared_strings(zin)
        sheet_path = _first_worksheet_path(zin)
        xml_text = zin.read(sheet_path).decode("utf-8")
        new_text, kept, dropped = sheet_filter(xml_text, shared)

        cellxfs = []
        if "xl/styles.xml" in zin.namelist():
            cellxfs = _parse_cellxfs_numfmt(zin.read("xl/styles.xml"))
        hits = _scientific_notation_hits(new_text, cellxfs)
        if hits:
            raise RuntimeError(
                f"筛选结果存在科学计数法风险值（General 格式数字）: {hits[:5]}"
            )

        tmp = target + ".tmp"
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
            for info in infos:
                if info.filename == sheet_path:
                    data = new_text.encode("utf-8")
                else:
                    data = zin.read(info.filename)
                zout.writestr(info, data)
    os.replace(tmp, target)
    return kept, dropped, time.time() - t0


def filter_report_xlsx(path, status_name, keyword, out_path=None):
    """筛选 xlsx 并输出：保留 状态列取值 含 keyword 的行（列缺失等异常抛 RuntimeError）。"""
    return _filter_report(
        path,
        lambda xml_text, shared: filter_worksheet(xml_text, shared, status_name, keyword),
        out_path,
    )


def filter_report_xlsx_by_set(path, col_name, allowed, out_path=None):
    """筛选 xlsx 并输出：保留 指定列取值 ∈ allowed 的行（活动ID关联筛选）。"""
    return _filter_report(
        path,
        lambda xml_text, shared: filter_worksheet_by_set(xml_text, shared, col_name, allowed),
        out_path,
    )


def read_column_set(path, col_name):
    """读取 xlsx 指定列的全部非空值（去重集合）；用于搜索词/购买商品按活动ID关联筛选。"""
    with zipfile.ZipFile(path, "r") as zin:
        shared = _read_shared_strings(zin)
        sheet_path = _first_worksheet_path(zin)
        xml_text = zin.read(sheet_path).decode("utf-8")
        m = re.search(r"(<sheetData[^>]*>)(.*?)(</sheetData>)", xml_text, re.S)
        if not m:
            raise RuntimeError("工作表缺少 <sheetData> 数据区")
        body = m.group(2)
        rows = re.findall(r"<row\b[^>]*>.*?</row>", body, re.S)
        if not rows:
            raise RuntimeError("工作表没有数据行（表头缺失）")
        header_row = rows[0]
        col_letter = ""
        for c in ET.fromstring(header_row):
            if _local_name(c) != "c":
                continue
            col = _cell_col_letter(c)
            if col and _cell_value(c, shared) == col_name:
                col_letter = col
                break
        if not col_letter:
            raise RuntimeError(f"表头中未找到列: {col_name}")
        out = set()
        for row in rows[1:]:
            v = _cell_text_from_row(row, col_letter, shared)
            if v:
                out.add(v)
        return out


def filter_downloaded_reports(tasks, filtered_dir=None, no_filter=False):
    """对已下载的报告执行筛选，两类：

    1) 状态筛选：命中 FILTER_RULES 的 5 种报告，按 运行状态 含 '已开启' 保留；
    2) 活动ID关联筛选：搜索词/购买商品报告无运行状态列，保留 广告活动ID
       ∈ 筛选后活动报告(adCampaignReport) 的行，剔除已暂停活动数据
       （使 5 个聚合维度口径一致）。

    - filtered_dir 非空 → 筛选结果写入该目录（同名文件），原表保留
    - filtered_dir 为空 → 原位覆盖

    任一文件筛选失败 → 立即返回 False（fail-fast，退出码 1）。
    """
    if no_filter:
        print("（已通过 --no-filter 跳过数据筛选）")
        return True
    if filtered_dir:
        Path(filtered_dir).mkdir(parents=True, exist_ok=True)
    filtered = 0
    for t in tasks:
        rule = FILTER_RULES.get(t.get("report_type"))
        if not rule or not t.get("file"):
            continue
        status_name, keyword = rule
        dst = None
        if filtered_dir:
            dst = str(Path(filtered_dir) / os.path.basename(t["file"]))
        print(f"筛选 {t['name']}（{status_name} 含 '{keyword}' 保留）...", end=" ", flush=True)
        try:
            kept, dropped, secs = filter_report_xlsx(t["file"], status_name, keyword, out_path=dst)
        except Exception as e:
            print(f"✗ {e}")
            return False
        if dst:
            t["filtered_file"] = dst
        print(f"✓ 保留 {kept:,} 行, 剔除 {dropped:,} 行（{secs:.1f}s）")
        filtered += 1

    # 活动ID关联筛选（搜索词/购买商品）：活动ID集合取自筛选后的活动报告
    camp = next((t for t in tasks if t.get("report_type") == "adCampaignReport"), None)
    camp_file = None
    if camp:
        camp_file = camp.get("filtered_file") or camp.get("file")
    if camp_file and os.path.exists(camp_file):
        try:
            allowed = read_column_set(camp_file, "广告活动ID")
        except Exception as e:
            print(f"✗ 读取活动报告活动ID集合失败: {e}")
            return False
        for t in tasks:
            col = JOIN_FILTER_RULES.get(t.get("report_type"))
            if not col or not t.get("file"):
                continue
            dst = None
            if filtered_dir:
                dst = str(Path(filtered_dir) / os.path.basename(t["file"]))
            print(
                f"筛选 {t['name']}（{col} ∈ 筛选后活动报告 {len(allowed):,} 个活动 保留）...",
                end=" ", flush=True,
            )
            try:
                kept, dropped, secs = filter_report_xlsx_by_set(
                    t["file"], col, allowed, out_path=dst
                )
            except Exception as e:
                print(f"✗ {e}")
                return False
            if dst:
                t["filtered_file"] = dst
            print(f"✓ 保留 {kept:,} 行, 剔除 {dropped:,} 行（{secs:.1f}s）")
            filtered += 1
    else:
        print("提示: 无筛选后活动报告，跳过 搜索词/购买商品 活动ID关联筛选")

    print(f"筛选完成: {filtered} 个报告已筛选")
    return True


# ── 广告组合导出（同步接口，nextToken 游标分页 → CSV）────────────────
def _portfolio_row(item):
    """把接口返回的一条组合数据规整为 CSV 行（全部按字符串原样输出，None→空串）。"""
    return ["" if item.get(c) is None else str(item.get(c, "")) for c in PORTFOLIO_COLUMNS]


def fetch_portfolios(shop_ids, out_dir):
    """按店铺导出广告组合（portfolio）全量数据到 CSV。

    分页规则（实测）：
      - 首页 body: {"shopId": <数字>, "pageSize": 1000}
      - data.nextToken 非空 → 下一页 body = 首页条件 + {"nextToken": <游标>}（其余条件沿用第一页）
      - nextToken 为空/缺失 → 结束
    返回 [{shop_id, file, count}, ...]；任一店铺失败则打印错误并返回 None（fail-fast）。
    """
    results = []
    for shop_id in shop_ids:
        rows = []
        next_token = None
        seen_tokens = set()
        try:
            for page in range(1, PORTFOLIO_MAX_PAGES + 1):
                body = {"shopId": int(shop_id), "pageSize": PORTFOLIO_PAGE_SIZE}
                if next_token:
                    body["nextToken"] = next_token
                data = api_post(PORTFOLIO_PATH, body)
                items = data.get("itemList") or []
                rows.extend(_portfolio_row(it) for it in items)
                next_token = data.get("nextToken") or ""
                print(f"  广告组合[{shop_id}] 第{page}页: {len(items)} 条, nextToken={'有' if next_token else '无'}")
                if not next_token:
                    break
                if next_token in seen_tokens:
                    raise ApiError(-1, f"nextToken 出现重复游标，疑似服务端异常: {next_token[:50]}")
                seen_tokens.add(next_token)
                time.sleep(PORTFOLIO_PAGE_INTERVAL)
            else:
                raise ApiError(-1, f"广告组合[{shop_id}] 翻页超过 {PORTFOLIO_MAX_PAGES} 页上限，疑似死循环")
        except Exception as e:
            print(f"✗ 广告组合[{shop_id}] 导出失败: {e}", file=sys.stderr)
            return None

        out_file = out_dir / f"广告组合_{shop_id}.csv"
        if out_file.exists():
            print(f"  提示: 文件已存在，覆盖写入 → {out_file.name}")
        # UTF-8 BOM：Excel 双击打开中文不乱码；ID 列为文本字符串，无科学计数法风险
        with open(out_file, "w", encoding="utf-8-sig", newline="") as f:
            writer = csv.writer(f)
            writer.writerow(PORTFOLIO_COLUMNS)
            writer.writerows(rows)
        print(f"✓ 广告组合[{shop_id}] 共 {len(rows)} 条 → {out_file}")
        results.append({"shop_id": shop_id, "file": str(out_file), "count": len(rows)})
    return results


# ── 入口 ────────────────────────────────────────────────────────────
def parse_args(argv):
    parser = argparse.ArgumentParser(
        description="赛狐 SP 广告天维度(daily)报告下载（7 种报告；原表与筛选结果分别保存）"
    )
    parser.add_argument("--shop-ids", required=True,
                        help="店铺ID，多个用逗号分隔，如 128154,128155")
    parser.add_argument("--start", required=True, help="开始日期 YYYY-MM-DD")
    parser.add_argument("--end", required=True, help="结束日期 YYYY-MM-DD")
    parser.add_argument("--out", default=str(DEFAULT_OUTPUT_DIR),
                        help=f"原表输出目录（默认 {DEFAULT_OUTPUT_DIR}）")
    parser.add_argument("--filtered-out", default=str(DEFAULT_FILTERED_DIR),
                        help=f"筛选结果目录（默认 {DEFAULT_FILTERED_DIR}）")
    parser.add_argument("--report-types", default=",".join(REPORT_TYPES),
                        help="报告类型，逗号分隔（默认全部 7 种）")
    parser.add_argument("--interval", type=float, default=TASK_INTERVAL,
                        help=f"创建任务间隔秒（默认 {TASK_INTERVAL}）")
    parser.add_argument("--poll-interval", type=float, default=POLL_INTERVAL,
                        help=f"轮询间隔秒（默认 {POLL_INTERVAL}）")
    parser.add_argument("--max-polls", type=int, default=MAX_POLLS,
                        help=f"最大轮询次数（默认 {MAX_POLLS}）")
    parser.add_argument("--no-filter", action="store_true",
                        help="下载后不执行数据筛选（默认按运行状态筛选）")
    parser.add_argument("--no-portfolio", action="store_true",
                        help="跳过广告组合(portfolio)基础数据导出（默认导出）")
    return parser.parse_args(argv)


def validate_args(args):
    """校验参数，返回 (shop_ids, report_types)；非法时打印错误并返回 None。"""
    shop_ids = [s.strip() for s in args.shop_ids.split(",") if s.strip()]
    if not shop_ids:
        print("错误: --shop-ids 不能为空", file=sys.stderr)
        return None
    for d in (args.start, args.end):
        try:
            datetime.strptime(d, "%Y-%m-%d")
        except ValueError:
            print(f"错误: 日期格式应为 YYYY-MM-DD，实际为: {d}", file=sys.stderr)
            return None
    if args.start > args.end:
        print("错误: --start 不能晚于 --end", file=sys.stderr)
        return None
    report_types = [t.strip() for t in args.report_types.split(",") if t.strip()]
    unknown = [t for t in report_types if t not in REPORT_NAMES]
    if unknown:
        print(f"错误: 未知报告类型 {unknown}，可选: {REPORT_TYPES}", file=sys.stderr)
        return None
    return shop_ids, report_types


def main(argv=None):
    # Windows 控制台/管道下 UTF-8 输出兜底（避免 cp936 编码错误）
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

    args = parse_args(argv)
    checked = validate_args(args)
    if checked is None:
        return 1
    shop_ids, report_types = checked

    out_dir = Path(args.out)
    out_dir.mkdir(parents=True, exist_ok=True)

    print("=" * 60)
    print("赛狐 SP 广告天维度报告下载")
    print(f"  店铺ID    : {', '.join(shop_ids)}")
    print(f"  日期范围  : {args.start} ~ {args.end}")
    print(f"  报告类型  : {', '.join(report_types)}")
    print(f"  输出目录  : {out_dir}")
    print("=" * 60)

    # 1. 获取 token（提前暴露凭证/网络问题）
    try:
        get_access_token()
        print("✓ token 获取/复用成功")
    except Exception as e:
        print(f"✗ {e}", file=sys.stderr)
        return 1

    # 2. 创建下载任务
    tasks = create_download_tasks(
        shop_ids, args.start, args.end, report_types, interval=args.interval
    )
    ok_tasks = [t for t in tasks if t.get("task_id")]
    fail_tasks = [t for t in tasks if not t.get("task_id")]
    print(f"任务创建完成: {len(ok_tasks)}/{len(tasks)} 成功")
    if not ok_tasks:
        print("错误: 没有可轮询的下载任务", file=sys.stderr)
        return 1

    # 3. 轮询并下载
    finished, unfinished = poll_and_download(
        ok_tasks, out_dir, args.start, args.end,
        poll_interval=args.poll_interval, max_polls=args.max_polls,
    )

    for t, state in unfinished:
        print(
            f"✗ 超时未完成: {t['name']} (任务ID={t['task_id']}, 最后状态={state})",
            file=sys.stderr,
        )

    failed = fail_tasks + [t for t, _ in unfinished]
    if failed:
        names = "、".join(t["name"] for t in failed)
        print(f"✗ 失败 {len(failed)} 个报告: {names}", file=sys.stderr)
        print("-" * 60)
        return 1

    # 4. 数据表存在性断言（fail-fast：任一文件缺失立即停止）
    for t in finished:
        if not os.path.exists(t["file"]):
            print(f"✗ 数据表不存在: {t['name']} -> {t['file']}", file=sys.stderr)
            return 1

    # 5. 数据筛选（fail-fast：文件打不开/无工作表/表头缺失/筛选列不存在/筛选异常 → 停止）
    #    原表保留在 out_dir，筛选结果写入 filtered_dir（并列文件夹）
    if not filter_downloaded_reports(
        finished, filtered_dir=Path(args.filtered_out), no_filter=args.no_filter
    ):
        print("错误: 数据筛选失败，已停止执行", file=sys.stderr)
        return 1

    # 6. 广告组合（portfolio）基础数据导出（同步接口 + nextToken 分页 → CSV）
    portfolios = None
    if not args.no_portfolio:
        print("-" * 60)
        print("开始导出广告组合(portfolio)基础数据 ...")
        portfolios = fetch_portfolios(shop_ids, out_dir)
        if portfolios is None:
            print("错误: 广告组合导出失败，已停止执行", file=sys.stderr)
            return 1

    print("-" * 60)
    print(f"✓ 原表 {len(finished)} 个 → {out_dir}")
    for t in finished:
        print(f"  {t['file']}")
    if not args.no_filter:
        flt = [t for t in finished if t.get("filtered_file")]
        print(f"✓ 筛选结果 {len(flt)} 个 → {args.filtered_out}")
        for t in flt:
            print(f"  {t['filtered_file']}")
    if portfolios:
        print(f"✓ 广告组合 {len(portfolios)} 个 → {out_dir}")
        for p in portfolios:
            print(f"  {p['file']}  ({p['count']} 条)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
