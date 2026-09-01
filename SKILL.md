---
name: skill-addown
description: 赛狐ERP广告数据明细下载与聚合技能。下载赛狐SP广告报表并聚合为清晰的CSV聚合表，便于Agent进行广告数据分析。当用户提到赛狐、Sellfox、广告报表下载、广告数据聚合（组合/活动/广告位/搜索词/购买商品维度）、ACoS/ROAS汇总，或要求基于 赛狐广告报表明细 文件夹准备广告分析数据时使用；用户说"跑一下广告报表""聚合广告数据"也应触发。
---

# <广告数据明细导出与处理技能包>

此技能用于下载赛狐ERP广告数据明细数据并聚合清晰,便于用户使用Agent工具进行广告数据分析

## 何时使用（触发场景）

> 说明：详细的触发信息写在 frontmatter 的 description 中；此节补充正文级的使用边界。

- 技能包引用触发 : 用户主动引用技能才发起

## 前置条件

- Python 3（仅标准库，零第三方依赖）；Windows 下运行需设 `PYTHONUTF8=1`
- 凭证：client_id / client_secret 内置于抓取脚本，可用环境变量 `SF_CLIENT_ID` / `SF_CLIENT_SECRET` 覆盖
- 目录结构（桌面 `赛狐广告报表明细\`）：
  - `赛狐原表数据\` — API 下载的原表 xlsx（固定文件名，最新一次拉取覆盖旧文件）+ 组合CSV + 店铺列表 + 在线产品明细
  - `赛狐原表筛选结果文件夹\` — 筛选后的表（6 个：4 种按运行状态筛选 + 搜索词/购买商品按活动ID关联筛选；固定文件名）
  - `广告分析聚合表\` — 聚合输出 CSV（固定文件名，覆盖写）

## 工作流程

1. **执行前必确认参数（不得在指令中写死）**：运行 `fetch_sellfox_daily_ad_reports.py` 前必须与用户确认两项：
   - **日期范围**：按指定日期格式 `YYYY-MM-DD~YYYY-MM-DD`（`--start` / `--end`）询问用户本次要拉取的区间，不得默认沿用上一次或硬编码
   - **店铺ID**：向用户询问本次要拉取的店铺ID；**用户无法提供时**，先按需运行 `fetch_sellfox_shop_list.py`（见脚本模块表），给用户 `店铺列表.csv` 查看链接,辅助用户确认目标店铺ID
2. **下载原表**：参数确认后运行抓取脚本，token 本地缓存 24h 复用，任务轮询完成后 xlsx 自动落入原表数据文件夹；同批店铺ID 还需运行 `fetch_sellfox_online_products.py`（见脚本模块表）导出 在线产品明细.csv —— 购买商品聚合做其他ASIN→SKU(sku)映射的依赖，产品上架/下架变动后应重跑
3. **运行聚合**：执行 `aggregate_all.py`，一次性按依赖顺序跑完 5 个聚合脚本（组合→活动→广告位→搜索词→购买商品，购买商品自动最后跑）；各脚本自动挑选输入目录固定名文件，无需传路径
4. **运行聚合并验收**：`aggregate_all.py` 退出码非 0 即失败（失败即停并报告是哪一步），读报错定位，绝不静默跳过或手工删数据
5. **同步说明文档**：生成结果文件后，将 `references/结果文件说明.md` 复制覆盖到 `桌面\赛狐广告报表明细\结果文件说明.md`（一级目录，供非技能场景查阅）

## 脚本模块（scripts/）

`scripts/` 存放可执行脚本，用于确定性的、重复性的任务，模型不必把代码读入上下文即可直接运行。

| 脚本 | 用途 | 用法 |
| --- | --- | --- |
| `fetch_sellfox_daily_ad_reports.py` | 下载6种SP报告原表+筛选结果+广告组合CSV | `python fetch_sellfox_daily_ad_reports.py --shop-ids <店铺ID> --start <YYYY-MM-DD> --end <YYYY-MM-DD>` |
| `fetch_sellfox_shop_list.py` | 拉取已授权启用店铺列表，辅助确认店铺ID | `python fetch_sellfox_shop_list.py` |
| `fetch_sellfox_online_products.py` | 导出店铺在线产品明细（shopId/asin/sku；剔除sku含amzn行、按asin去重保首次），供购买商品聚合翻译其他ASIN | `python fetch_sellfox_online_products.py --shop-ids <店铺ID>` |
| `aggregate_all.py` | 按依赖顺序一次跑完 5 个聚合脚本 | `python aggregate_all.py` |
| `aggregate_campaign_by_portfolio.py` | 组合维度聚合 | （由 aggregate_all.py 调用） |
| `aggregate_campaign_by_campaign.py` | 活动维度聚合（产出 广告活动聚合.csv） | （由 aggregate_all.py 调用） |
| `aggregate_placement_by_campaign.py` | 广告位维度聚合 | （由 aggregate_all.py 调用） |
| `aggregate_searchterm_by_campaign.py` | 搜索词维度聚合 | （由 aggregate_all.py 调用） |
| `aggregate_purchaseditem_by_campaign.py` | 购买商品维度聚合（依赖活动聚合，须最后跑） | （由 aggregate_all.py 调用） |

使用规则：
- 优先直接调用脚本，不要在对话里重写同等逻辑
- 参数、退出码、输出格式以脚本文件头部的 docstring 为准
- 脚本失败（非零退出）时先读报错定位原因，不要盲目重试

## 参考模块（references/）

`references/` 存放的参考文档不要加载上下文，控制上下文用量。

| 文档 | 内容 | 何时读取 |
| --- | --- | --- |
| `references/结果文件说明.md` | 全部结果文件速查：原表功能简述、聚合表生成逻辑与字段口径（面向广告分析 Agent） | 不用读取,主要是给广告分析Agent提供参考 |

使用规则：
- 不要读入
- 正文若需更多细节，必须指向这里的文件而不是把内容复制进 SKILL.md

## 输出数据口径（读取聚合表做分析前必读）

输出位置与命名规则已在脚本中固定，此处不重复；此处只声明**读结果时**需要的口径：

- 编码：UTF-8 BOM（utf-8-sig），读取时按此解码
- 数值列格式（均为字符串文本，做计算前注意转换）：
  - 点击率/转化率/ACoS → `4.35%`（已乘100带%号，可直接数值比较，参与运算需去%除以100）
  - ROAS → `23.01` 倍数，不带%
  - 其余 SUM/AVG 纯数字2位；销售额 `f"{round(v,2):.2f}"`
- 空组合列（组合ID/组合名称为空）= 活动ID关联不到组合，多为已暂停活动，非数据错误；搜索词/购买商品聚合的输入已在下载阶段按活动ID关联剔除已暂停活动行，这两张表一般不再出现空组合
- 各文件生成逻辑与字段明细见 `references/结果文件说明.md`（口径细节不在本文件重复）

## 注意事项与边界

- 数据安全约束：处理过程不允许随意删除数据；异常非数值文本只报错退出 1，绝不静默跳过
- 空值口径：空单元格跳过该格累加（行仍参与分组）；空键归专用空值组
- Windows 已知坑：pwsh 中 `&` 分隔命令会报 ParserError（用 `;`）；含引号/逗号的 python 代码写临时 .py 再跑；统一 `PYTHONUTF8=1`
- 有疑问时先与用户确认再执行
