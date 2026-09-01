#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""赛狐开放平台 — 在线产品明细导出脚本（独立实现，仅标准库）

流程：
  1. GET  /api/oauth/v2/token.json                     获取 access_token（本地缓存 24h 复用）
  2. POST /api/order/api/product/v2/pageList.json      分页拉取指定店铺的在线产品
       入参: shopIdList=[用户提供的店铺ID], onlineStatusList=["active"]
       （业务参数走 requestBody，签名只签裸路径；放 query 会 40011 签名错误）
  3. 输出前清洗：① 剔除 sku 含 "amzn" 的行（亚马逊自动生成SKU噪音）；
       ② 按 asin 去重（保留首次出现，后续重复行丢弃）
     清洗后输出 [shopId, asin, sku] 三列 CSV（UTF-8 BOM）到 赛狐原表数据 文件夹
       供 购买商品聚合 做 其他ASIN → SKU(sku) 映射

数据完整性约束（fail-fast，不静默跳过）：
  - 累计拉取行数 == 接口返回 totalSize，不等则退出 1
  - 翻页期间 totalSize 变化时以最后一页为准重校验，仍不符退出 1

用法：
  python fetch_sellfox_online_products.py --shop-ids 128154
  python fetch_sellfox_online_products.py --shop-ids 128154,128155 --out <输出目录> --page-size 200

输出文件（固定名，存在即覆盖）：
  <桌面>/赛狐广告报表明细/赛狐原表数据/在线产品明细.csv

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
PRODUCT_LIST_PATH = "/api/order/api/product/v2/pageList.json"
GRANT_TYPE = "client_credentials"

# 开发者凭证（默认内置，可用环境变量覆盖）
DEFAULT_CLIENT_ID = "368900"
DEFAULT_CLIENT_SECRET = "9eaec859-fb6a-4709-a98d-7432bcb93c02"

# 只导在线（上架中）产品
ONLINE_STATUS_LIST = ["active"]

DEFAULT_PAGE_SIZE = 200
MAX_PAGES = 500            # 分页安全上限，防死循环
PAGE_INTERVAL = 1.0        # 翻页间隔秒，防 40019 限流
HTTP_TIMEOUT = 30
MAX_RATE_RETRIES = 5
CODE_RATE_LIMIT = 40019
CODE_TOKEN_INVALID = 40001
TOKEN_REFRESH_BUFFER = 60  # 秒，提前刷新余量

DEFAULT_OUT_DIR = Path.home() / "Desktop" / "赛狐广告报表明细" / "赛狐原表数据"
OUT_FILE_NAME = "在线产品明细.csv"
OUT_COLUMNS = ["shopId", "asin", "sku"]

TOKEN_CACHE_FILE = Path(tempfile.gettempdir()) / "sellfox_online_products_token.json"


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


# ── 在线产品拉取 ────────────────────────────────────────────────────
def parse_shop_ids(raw):
    """解析逗号分隔的店铺ID参数：纯数字转 int，其余原样保留，去空去重。"""
    ids = []
    seen = set()
    for part in raw.replace("，", ",").split(","):
        s = part.strip()
        if not s:
            continue
        if s in seen:
            continue
        seen.add(s)
        try:
            ids.append(int(s))
        except ValueError:
            ids.append(s)
    return ids


def fetch_all_products(shop_ids, page_size):
    """分页拉取全部在线产品，返回 (rows, totalSize)。

    完整性校验：累计行数 == totalSize（以最后见到的 totalSize 为准），
    不符抛 RuntimeError（fail-fast）。
    """
    rows = []
    page_no = 1
    total_size = None
    total_page = None
    while True:
        body = {
            "shopIdList": shop_ids,
            "onlineStatusList": ONLINE_STATUS_LIST,
            "pageNo": page_no,
            "pageSize": page_size,
        }
        data = api_post(PRODUCT_LIST_PATH, body)
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


# ── 入口 ────────────────────────────────────────────────────────────
def parse_args(argv):
    parser = argparse.ArgumentParser(
        description="赛狐在线产品明细导出（onlineStatus=active；输出 shopId/asin/sku CSV，"
                    "供购买商品聚合做 其他ASIN→SKU 映射）"
    )
    parser.add_argument("--shop-ids", required=True,
                        help="店铺ID，多个用逗号分隔，如 128154,128155（与广告报表下载同批店铺）")
    parser.add_argument("--out", default=str(DEFAULT_OUT_DIR),
                        help=f"输出目录（默认 {DEFAULT_OUT_DIR}）")
    parser.add_argument("--page-size", type=int, default=DEFAULT_PAGE_SIZE,
                        help=f"分页大小（默认 {DEFAULT_PAGE_SIZE}）")
    return parser.parse_args(argv)


def main(argv=None):
    args = parse_args(argv)
    shop_ids = parse_shop_ids(args.shop_ids)
    if not shop_ids:
        print("✗ 未解析到任何店铺ID，请用 --shop-ids 128154[,128155...] 指定", file=sys.stderr)
        return 1

    print("=" * 60)
    print("赛狐在线产品明细导出")
    print(f"  店铺ID   : {shop_ids}")
    print(f"  在线状态 : {ONLINE_STATUS_LIST}")
    print(f"  输出目录 : {args.out}")
    print("=" * 60)

    rows, total = fetch_all_products(shop_ids, args.page_size)
    print(f"✓ 拉取完成: {len(rows)} 行（totalSize={total} 对账一致）")

    # 数据清洗（拉取对账之后、写CSV之前）：① 剔除 sku 含 amzn 的行；② 按 asin 去重保首次
    clean_rows = []
    seen_asins = set()
    amzn_dropped = dup_dropped = 0
    for r in rows:
        if "amzn" in str(r.get("sku", "") or "").lower():
            amzn_dropped += 1
            continue
        asin = str(r.get("asin", "") or "")
        if asin in seen_asins:
            dup_dropped += 1
            continue
        seen_asins.add(asin)
        clean_rows.append(r)
    print(f"✓ 数据清洗: 剔除 sku 含 amzn 行 {amzn_dropped}，按 asin 去重剔除 {dup_dropped}，"
          f"剩余 {len(clean_rows)}/{len(rows)} 行")

    out_dir = Path(args.out)
    out_dir.mkdir(parents=True, exist_ok=True)
    out_file = out_dir / OUT_FILE_NAME
    if out_file.exists():
        print(f"提示: 文件已存在，覆盖写入 → {out_file.name}")

    # UTF-8 BOM：Excel 双击打开中文不乱码；字段为文本字符串原样输出
    with open(out_file, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(OUT_COLUMNS)
        for r in clean_rows:
            w.writerow([
                str(r.get("shopId", r.get("shop_id", ""))),
                str(r.get("asin", "")),
                str(r.get("sku", "")),
            ])

    print("-" * 60)
    print(f"✓ 在线产品明细 {len(clean_rows)} 行 → {out_file}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
