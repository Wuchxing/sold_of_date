# 日销计算

需要 Python 3.10 或更新版本，无第三方依赖。

```powershell
python C:\codex_projects\sold_of_date\daily_sales.py $PATH
```

```powershell
python daily_sales.py "太阳能喷泉数据汇总2026-09-08.xlsx"
```

结果另存为原文件名加 `_日销已计算.xlsx`，重名时自动加编号。源文件不变。

不传文件路径时打开文件选择器（需要 Python 安装包含 tkinter）：

```powershell
python daily_sales.py
```

也可以在自己的 Python 程序中调用：

```python
from pathlib import Path
from daily_sales import process

file_path = Path("太阳能喷泉数据汇总2026-09-08.xlsx")
result = process(file_path)
print(result)
```

其他参数：`--sheet 数据源` 指定工作表，`--year 2026` 指定最新销量年份，`-o 新文件.xlsx` 指定尚不存在的输出路径。

表头支持 `9.08销量`、`9.08总销量`、`9.08日销` 等日期列，按时间顺序排列，最新日销和销量日期必须一致。最新年份默认从文件名日期读取，历史跨年自动回推。多个符合条件的工作表需要显式选择。

销量中的“断货”和 `#N/A`（含公式缓存错误值）按缺失记录跳过，向前寻找有效数字销量，并使用实际日期间隔。最新销量为这些标记时跳过该行，历史标记保持原样；其他未知非数字值仍报错。

只处理最新销量有值的行，历史不足的日销留空；完全相同的历史累计销量且无历史日销时按观察到的零增量计算。10万及以上若没有可用的历史增长区间，则留空。

历史公式使用 Excel/WPS 保存的缓存数值，不调用外部链接；所需公式缺少缓存时提示重算后再处理。程序不刷新透视表或其他工作表，需查看最新汇总时在 Excel/WPS 内刷新。

程序逐项校验非目标 ZIP 内容不变，包括嵌入图片、图片关系和样式定义。仅修改最新日销，以及最新销量下降时的最新销量值。日销写入含 `MAX(0,…)` 非负保护的 Excel 公式，不使用 `ROUND`，保留小数计算结果，同时保存计算缓存以便打开时显示结果。最新日销整列（含表头和空白单元格）复制上一列日销的字体、填充、边框、对齐、数字格式等样式，以及默认样式和列宽；保留目标列可见性。最新销量下降时更正累计销量数值，并写入日销公式 `=0`。

运行测试：

```powershell
python -m unittest discover -s tests -v
```

公式引用 Python 本次识别出的有效历史单元格，日期间隔以实际天数写入公式。增删日期列、改变缺失记录或销量档位后，请重新运行程序以重新选择计算区间。公式结果可为小数，显示方式沿用上一列的数字格式；若上一列只显示整数，显示值会受格式影响，但实际数值保留小数。

## 同时计算价格

默认运行方式仍只计算日销，无需数据库或第三方依赖。加 `-a` 才读取 MySQL 汇率并计算人民币：

```powershell
python -m pip install -r requirements-price.txt
python daily_sales.py "太阳能喷泉数据汇总2026-09-08.xlsx" -a
```

结果默认另存为 `_价格与日销已计算.xlsx`，重名自动加编号。人民币写为类似 `=BL2*7.63` 的公式并保留原格式；预估供货价不改。

按最新价格列的日期及“站点”精确查询 `exchange_rates` 的 `record_date`、`country`，使用 `exchange_rate`；`currency_pair` 必须为外币/CNY。无当天汇率或价格为空时保留原人民币内容，不改用其他日期。零价格是有效数值。相同站点日期有多条汇率时明确报错。

默认配置路径为 `C:\codex_projects\log_working\.venv\password.md`，通过 `--db-config 路径` 可覆盖。配置支持“主机、用户名、密码”，可选“端口、数据库”，以冒号或等号分隔，允许 Markdown 列表。未指定数据库时，只在可访问范围内唯一存在 `exchange_rates` 时自动选择。不要将真实配置放入项目或提交 Git。

API 调用：`process(file_path, all_calculations=True)`，可通过 `db_config=Path(...)` 指定外部配置。汇率按日期批量读取，程序不写数据库；数据库读取失败时不会发布部分计算的结果文件。

共享公式仅在整组都需要更新时转换为独立公式；部分共享公式组和数组公式会明确报错，避免破坏保留行。
