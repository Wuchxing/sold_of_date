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

程序逐项校验非目标 ZIP 内容不变，包括嵌入图片、图片关系和样式。仅修改最新日销，以及最新销量下降时的最新销量值。写入整数数值并保留单元格原有格式。

运行测试：

```powershell
python -m unittest discover -s tests -v
```
