#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""赛狐开放平台 — 店铺列表下载与筛选脚本（独立实现，仅标准库）

流程：
  1. GET  /api/oauth/v2/token.json                     获取 access_token（本地缓存 24h 复用）
  2. POST /api/shop/pageList.json                     分页拉取全量店铺
       （实测：pageNo/pageSize 走 requestBody，签名只签裸路径；放 query 会 40011 签名错误）
  3. 筛选：广告授权状态 == auth(已授权) 且 店铺状态 == "0"(启用中) 的行保留
  4. 输出仅 [店铺ID, 店铺名称] 两列 CSV（UTF-8 BOM）到 赛狐原表数据 文件夹

数据完整性约束（fail-fast，不静默跳过）：
  - 累计拉取行数 == 接口返回 totalSize，不等则退出 1
  - 翻页期间 totalSize 变化时以最后一页为准重校验，仍不符退出 1

用法：
  python fetch_sellfox_shop_list.py
  python fetch_sellfox_shop_list.py --out <输出目录> --page-size 100

输出文件（固定名，存在即覆盖）：
  <桌面>/赛狐广告报表明细/赛狐原表数据/店铺列表.csv

退出码 0 = 成功；1 = 参数/网络/校验失败
"""

import argparse
import csv
import hashlib
import hmac
import json
import os
import random
import sys
import tempfile
import time
from pathlib import Path
from urllib import error as urlerror
from urllib import parse as urlparse
from urllib import request as urlrequest

# ── 常量 ────────────────────────────────────────────────────────────
API_BASE = "https://openapi.sellfox.com"
TOKEN_PATH = "/api/oauth/v2/token.json"
SHOP_LIST_PATH = "/api/shop/pageList.json"
GRANT_TYPE = "client_credentials"

# 开发者凭证（默认内置，可用环境变量覆盖）
DEFAULT_CLIENT_ID = "368900"
DEFAULT_CLIENT_SECRET = "9eaec859-fb6a-4709-a98d-7432bcb93c02"

# 筛选口径（实测返回为编码值）：
#   adStatus: "auth"=已授权 / "unauth"=未授权 / "expire"=已过期
#   status  : "0"=启用中 / 其他=非启用（枚举文档缺失，经用户确认按 0 口径）
KEEP_AD_STATUS = "auth"
KEEP_STATUS = "0"

DEFAULT_PAGE_SIZE = 100
MAX_PAGES = 100            # 分页安全上限，防死循环
PAGE_INTERVAL = 1.0        # 翻页间隔秒，防 40019 限流
HTTP_TIMEOUT = 30
MAX_RATE_RETRIES = 5
CODE_RATE_LIMIT = 40019
CODE_TOKEN_INVALID = 40001
TOKEN_REFRESH_BUFFER = 60  # 秒，提前刷新余量

DEFAULT_OUT_DIR = Path.home() / "Desktop" / "赛狐广告报表明细" / "赛狐原表数据"
OUT_FILE_NAME = "店铺列表.csv"
OUT_COLUMNS = ["店铺ID", "店铺名称"]

TOKEN_CACHE_FILE = Path(tempfile.gettempdir()) / "sellfox_token_cache.json"


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


# ── Token 获取与缓存（与 fetch_sellfox_daily_ad_reports.py 共享缓存）──
def load_cached_token():
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
    分页等业务参数走 requestBody，不参与签名（实测放 query 会导致 40011 签名错误）。
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
            # 部分业务错误码（如 40019 限流）以 HTTP 400 + JSON body 返回，需进重试逻辑
            try:
                err = json.loads(text)
            except json.JSONDecodeError:
                err = {}
            code = err.get("code")
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


# ── 店铺列表拉取与筛选 ─────────────────────────────────────────────
def fetch_all_shops(page_size):
    """分页拉取全量店铺，返回 (rows, totalSize)。

    完整性校验：累计行数 == totalSize（以最后见到的 totalSize 为准），
    不符抛 RuntimeError（fail-fast）。
    """
    rows = []
    page_no = 1
    total_size = None
    total_page = None
    while True:
        data = api_post(SHOP_LIST_PATH, {"pageNo": page_no, "pageSize": page_size})
        if not isinstance(data, dict) or "rows" not in data:
            raise RuntimeError(f"响应缺少 rows 字段: {str(data)[:300]}")
        page_rows = data.get("rows") or []

        # 翻页期间数据增减 → totalSize 会变化，以最新为准重校验
        new_total = data.get("totalSize")
        new_tp = data.get("totalPage")
        if total_size is not None and new_total != total_size:
            print(f"    提示: 翻页期间 totalSize 变化 {total_size} → {new_total}")
        total_size, total_page = new_total, new_tp

        rows.extend(page_rows)
        print(f"  第{page_no}页: {len(page_rows)} 条 (totalSize={total_size}, totalPage={total_page})")

        if total_page is not None and page_no >= total_page:
            break
        if not page_rows:
            raise RuntimeError(f"第{page_no}页返回 0 行但未到 totalPage={total_page}")
        if page_no >= MAX_PAGES:
            raise RuntimeError(f"翻页超过 {MAX_PAGES} 页上限，疑似死循环")
        page_no += 1
        time.sleep(PAGE_INTERVAL)

    if total_size is None or len(rows) != total_size:
        raise RuntimeError(f"完整性校验失败: 累计 {len(rows)} 行 ≠ totalSize {total_size}")
    return rows, total_size


def filter_shops(rows):
    """筛选 广告授权状态=auth(已授权) 且 店铺状态="0"(启用中)。

    返回 (保留行, 统计dict)；只做行级筛选，不修改行内容。
    """
    kept = []
    drop_auth = drop_status = 0
    for r in rows:
        if r.get("adStatus") != KEEP_AD_STATUS:
            drop_auth += 1
            continue
        if str(r.get("status", "")) != KEEP_STATUS:
            drop_status += 1
            continue
        kept.append(r)
    stats = {
        "drop_not_auth": drop_auth,
        "drop_not_enabled": drop_status,
        "kept": len(kept),
    }
    return kept, stats


# ── 入口 ────────────────────────────────────────────────────────────
def parse_args(argv):
    parser = argparse.ArgumentParser(
        description="赛狐店铺列表下载（筛选: 已授权+启用中；输出 店铺ID/店铺名称 CSV）"
    )
    parser.add_argument("--out", default=str(DEFAULT_OUT_DIR),
                        help=f"输出目录（默认 {DEFAULT_OUT_DIR}）")
    parser.add_argument("--page-size", type=int, default=DEFAULT_PAGE_SIZE,
                        help=f"分页大小（默认 {DEFAULT_PAGE_SIZE}）")
    return parser.parse_args(argv)


def main(argv=None):
    args = parse_args(argv)
    print("=" * 60)
    print("赛狐店铺列表下载")
    print(f"  输出目录 : {args.out}")
    print("=" * 60)

    rows, total = fetch_all_shops(args.page_size)
    print(f"✓ 拉取完成: {len(rows)} 行（totalSize={total} 对账一致）")

    kept, stats = filter_shops(rows)
    print(f"✓ 筛选完成: 保留 {stats['kept']} 家"
          f"（剔除 未授权/已过期 {stats['drop_not_auth']} 家、非启用中 {stats['drop_not_enabled']} 家）")

    out_dir = Path(args.out)
    out_dir.mkdir(parents=True, exist_ok=True)
    out_file = out_dir / OUT_FILE_NAME
    if out_file.exists():
        print(f"提示: 文件已存在，覆盖写入 → {out_file.name}")

    # UTF-8 BOM：Excel 双击打开中文不乱码；ID 为文本字符串原样输出
    with open(out_file, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(OUT_COLUMNS)
        for r in kept:
            w.writerow([str(r.get("id", "")), str(r.get("name", ""))])

    print("-" * 60)
    print(f"✓ 店铺列表 {len(kept)} 行 → {out_file}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
