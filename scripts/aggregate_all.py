#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""赛狐广告聚合编排：按依赖顺序一次跑完 5 个聚合脚本。

固定顺序（购买商品必须最后，依赖活动聚合产出的 广告活动聚合.csv）：
  1. aggregate_campaign_by_portfolio.py    组合维度
  2. aggregate_campaign_by_campaign.py     活动维度
  3. aggregate_placement_by_campaign.py    广告位维度
  4. aggregate_searchterm_by_campaign.py   搜索词维度
  5. aggregate_purchaseditem_by_campaign.py 购买商品维度

每个子脚本无参数调用：各自自动挑选输入目录固定名文件，输出到 广告分析聚合表。

退出码：0 = 全部成功；1 = 某一步失败（失败即停，报告失败步骤与退出码）。
"""
import os
import subprocess
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent

STEPS = [
    "aggregate_campaign_by_portfolio.py",
    "aggregate_campaign_by_campaign.py",
    "aggregate_placement_by_campaign.py",
    "aggregate_searchterm_by_campaign.py",
    "aggregate_purchaseditem_by_campaign.py",
]


def main():
    os.environ.setdefault("PYTHONUTF8", "1")
    print("=" * 60, flush=True)
    print("赛狐广告聚合编排：按依赖顺序执行 5 个聚合脚本", flush=True)
    print("=" * 60, flush=True)
    for i, name in enumerate(STEPS, 1):
        script = HERE / name
        if not script.exists():
            print(f"\n[{i}/{len(STEPS)}] 脚本缺失: {name}", flush=True)
            return 1
        print(f"\n[{i}/{len(STEPS)}] {name}", flush=True)
        result = subprocess.run([sys.executable, str(script)])
        if result.returncode != 0:
            print(f"\n✗ 失败于第 {i} 步 {name}（退出码 {result.returncode}），已停止后续步骤", flush=True)
            return 1
    print("\n✓ 5 个聚合脚本全部执行完成", flush=True)
    return 0


if __name__ == "__main__":
    sys.exit(main())
