"""累计销量日销计算。仅依赖 Python 标准库，保留 XLSX 包中非目标内容。"""
from __future__ import annotations

import argparse
from copy import copy
from dataclasses import dataclass
from datetime import date
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from pathlib import Path
import posixpath
import re
import tempfile
from xml.etree import ElementTree as ET
from zipfile import ZipFile

NS = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
REL = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'
HEADER = re.compile(r'^(?:(\d{4})[./-])?(\d{1,2})[./-](\d{1,2})(总销量|销量|日销)$')


@dataclass(frozen=True)
class Point:
    day: date
    sales: Decimal
    daily: Decimal | None = None


def rounded(value: Decimal | None) -> int | None:
    return None if value is None else int(max(Decimal(0), value).quantize(Decimal(1), rounding=ROUND_HALF_UP))


def below_cap(points: list[Point]) -> Decimal | None:
    """计算末次日销；缺少历史日销时，按历史销量逐期回推。"""
    rate = None
    start = 0
    for i, current in enumerate(points):
        if i:
            previous = points[i - 1]
            elapsed = Decimal((current.day - previous.day).days)
            if elapsed <= 0:
                raise ValueError('销量日期必须严格递增')
            if current.sales < previous.sales:
                rate = Decimal(0)
            elif current.sales < 10000:
                rate = (current.sales - previous.sales) / elapsed
            elif current.sales != previous.sales:
                rate = (current.sales - previous.sales) / Decimal((current.day - points[start].day).days)
            else:
                days = Decimal((current.day - points[start].day).days)
                if rate is not None:
                    rate = min(max(Decimal(0), rate), Decimal(1000) / days)
                # 一直只有相同累计销量时，观察区间内增量为零。
                else:
                    rate = Decimal(0)
            if current.sales != previous.sales:
                start = i
        if i < len(points) - 1 and current.daily is not None:
            rate = max(Decimal(0), current.daily)
    return rate


def cell_value(cell: ET.Element, strings: list[str]) -> str | None:
    kind = cell.get('t')
    if kind == 'inlineStr':
        return ''.join(t.text or '' for t in cell.findall('.//s:t', NS))
    value = cell.find('s:v', NS)
    if value is None or value.text is None:
        return None
    if kind == 's':
        return strings[int(value.text)]
    return value.text


def number(cell: ET.Element | None, strings: list[str], *, sales: bool = False) -> Decimal | None:
    if cell is None:
        return None
    value = cell_value(cell, strings)
    if sales and value is not None and value.strip() in {'断货', '#N/A'}:
        return None
    if value is None or not value.strip():
        if cell.find('s:f', NS) is not None:
            raise ValueError(f"{cell.get('r')} 公式没有缓存值，请先用 Excel/WPS 重算并保存")
        return None
    try:
        result = Decimal(value.strip().replace(',', ''))
    except InvalidOperation as exc:
        raise ValueError(f"{cell.get('r')} 不是有效数字：{value}") from exc
    if not result.is_finite():
        raise ValueError(f"{cell.get('r')} 不是有限数字")
    return result


def discover_columns(row: ET.Element, strings: list[str], latest_year: int):
    groups = {'sales': [], 'daily': []}
    for cell in row:
        match = HEADER.fullmatch((cell_value(cell, strings) or '').strip())
        if match:
            y, m, d, kind = match.groups()
            groups['daily' if kind == '日销' else 'sales'].append(
                (re.sub(r'\d', '', cell.get('r')), int(y) if y else None, int(m), int(d)))
    result = {}
    for kind, columns in groups.items():
        resolved = []
        year, previous = latest_year, None
        for col, explicit, month, day in reversed(columns):
            if explicit:
                year = explicit
            elif previous and (month, day) > (previous.month, previous.day):
                year -= 1
            current = date(year, month, day)
            if previous and current >= previous:
                raise ValueError('表头日期重复或未按时间递增排列')
            resolved.append((current, col))
            previous = current
        result[kind] = list(reversed(resolved))
    return result


CELL_PATTERN = re.compile(rb'<c\b[^>]*?(?:/>|>.*?</c>)', re.DOTALL)
ROW_PATTERN = re.compile(rb'<row\b[^>]*>.*?</row>', re.DOTALL)


def column_index(ref: str) -> int:
    result = 0
    for char in re.sub(r'\d', '', ref):
        result = result * 26 + ord(char) - 64
    return result


def patch_xml(data: bytes, edits: dict[str, int | Decimal | None]) -> bytes:
    """只替换目标单元格值；其余 XML 字节（包括样式与图片公式）保持原样。"""
    remaining = dict(edits)

    def replace_cell(match):
        raw = match.group()
        ref_match = re.search(rb'\br="([A-Z]+\d+)"', raw)
        if not ref_match or ref_match[1].decode() not in remaining:
            return raw
        ref = ref_match[1].decode()
        value = remaining.pop(ref)
        opening = raw[:raw.index(b'>') + 1].replace(b'/>', b'>')
        opening = re.sub(rb'\s+t="[^"]*"', b'', opening)
        body = b'' if raw.endswith(b'/>') else raw[raw.index(b'>') + 1:-4]
        if re.search(rb'<f\b[^>]*\bt="(?:shared|array)"', body):
            raise ValueError(f'{ref} 含共享或数组公式，不支持局部替换')
        body = re.sub(rb'<(?:v|f|is)\b[^>]*?(?:/>|>.*?</(?:v|f|is)>)', b'', body, flags=re.DOTALL)
        val = b'' if value is None else b'<v>' + str(value).encode('ascii') + b'</v>'
        return opening + val + body + b'</c>'

    data = CELL_PATTERN.sub(replace_cell, data)
    by_row = {}
    for ref, val in remaining.items():
        if val is not None:
            by_row.setdefault(re.search(r'\d+', ref)[0], []).append((ref, val))

    def insert_cells(match):
        raw = match.group()
        row_id = re.search(rb'\br="(\d+)"', raw)[1].decode()
        for ref, value in sorted(by_row.pop(row_id, []), key=lambda item: column_index(item[0])):
            cell = f'<c r="{ref}"><v>{value}</v></c>'.encode()
            position = len(raw) - len(b'</row>')
            for existing in CELL_PATTERN.finditer(raw):
                other = re.search(rb'\br="([A-Z]+\d+)"', existing.group())[1].decode()
                if column_index(other) > column_index(ref):
                    position = existing.start()
                    break
            raw = raw[:position] + cell + raw[position:]
        return raw

    data = ROW_PATTERN.sub(insert_cells, data)
    if by_row:
        raise ValueError('待写入的行不存在')
    ET.fromstring(data)
    return data


def workbook_parts(z: ZipFile):
    strings = []
    if 'xl/sharedStrings.xml' in z.namelist():
        strings = [''.join(t.text or '' for t in item.findall('.//s:t', NS))
                   for item in ET.fromstring(z.read('xl/sharedStrings.xml'))]
    rels = {r.get('Id'): r.get('Target') for r in ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))}
    sheets = {}
    for sheet in ET.fromstring(z.read('xl/workbook.xml')).findall('s:sheets/s:sheet', NS):
        target = rels[sheet.get(REL)]
        sheets[sheet.get('name')] = target.lstrip('/') if target.startswith('/') else posixpath.normpath('xl/' + target)
    return strings, sheets


def calculate(points: list[Point]) -> tuple[int | None, Decimal | None]:
    if len(points) < 2:
        return None, None
    if points[-1].sales < points[-2].sales:
        return 0, points[-2].sales
    if points[-1].sales < 100000:
        return rounded(below_cap(points)), None
    rate = None
    start = 0
    for i in range(1, len(points)):
        current, previous = points[i], points[i - 1]
        if current.day <= previous.day:
            raise ValueError('销量日期必须严格递增')
        if current.sales > previous.sales and current.sales != 100000:
            days = (current.day - points[start].day).days
            if current.sales < 10000:
                days = (current.day - previous.day).days
            rate = (current.sales - previous.sales) / Decimal(days)
        if current.sales != previous.sales:
            start = i
    return rounded(rate), None


def process(source: Path, output: Path | None = None, *, sheet: str | None = None,
            year: int | None = None) -> dict:
    source = Path(source).resolve()
    if source.suffix.lower() != '.xlsx':
        raise ValueError('仅支持 .xlsx 文件')
    if year is None:
        match = re.search(r'(20\d{2})[-_.年]\d{1,2}[-_.月]\d{1,2}', source.stem)
        if not match:
            raise ValueError('表头没有年份：请用 --year 指定最新销量所在年份')
        year = int(match[1])
    if output is None:
        output = source.with_name(source.stem + '_日销已计算.xlsx')
        i = 2
        while output.exists():
            output = source.with_name(source.stem + f'_日销已计算_{i}.xlsx')
            i += 1
    output = Path(output).resolve()
    if source == output or output.exists():
        raise ValueError('输出路径必须是尚不存在的新文件，不能覆盖源文件或现有结果')
    with ZipFile(source) as zin:
        strings, sheets = workbook_parts(zin)
        candidates = []
        for name, part in sheets.items():
            if sheet and name != sheet:
                continue
            raw = zin.read(part)
            root = ET.fromstring(raw)
            rows = root.findall('s:sheetData/s:row', NS)
            for header in rows[:10]:
                columns = discover_columns(header, strings, year)
                if columns['sales'] and columns['daily']:
                    candidates.append((name, part, raw, rows, header, columns))
                    break
        if len(candidates) != 1:
            raise ValueError('无法唯一识别数据工作表，请用 --sheet 指定含日期销量及日销表头的工作表')
        name, part, raw, rows, header, columns = candidates[0]
        latest_day, sales_col = columns['sales'][-1]
        if columns['daily'][-1][0] != latest_day:
            raise ValueError('最新销量和最新日销日期不一致，请检查表头')
        daily_col = columns['daily'][-1][1]
        daily_map = dict(columns['daily'])
        edits = {}
        stats = dict(sheet=name, date=latest_day.isoformat(), sales_column=sales_col,
                     daily_column=daily_col, calculated=0, blank=0, skipped=0, corrected=0)
        for row in rows:
            row_id = row.get('r')
            if int(row_id) <= int(header.get('r')):
                continue
            cells = {re.sub(r'\d', '', c.get('r')): c for c in row}
            if number(cells.get(sales_col), strings, sales=True) is None:
                stats['skipped'] += 1
                continue
            points = []
            for day, col in columns['sales']:
                value = number(cells.get(col), strings, sales=True)
                if value is not None:
                    historical_daily = number(cells.get(daily_map.get(day)), strings) if day < latest_day else None
                    points.append(Point(day, value, historical_daily))
            value, corrected = calculate(points)
            edits[daily_col + row_id] = value
            stats['blank' if value is None else 'calculated'] += 1
            if corrected is not None:
                edits[sales_col + row_id] = corrected
                stats['corrected'] += 1
        patched = patch_xml(raw, edits)
        output.parent.mkdir(parents=True, exist_ok=True)
        with tempfile.NamedTemporaryFile(dir=output.parent, suffix='.xlsx', delete=False) as temp:
            temp_path = Path(temp.name)
        try:
            with ZipFile(temp_path, 'w') as zout:
                zout.comment = zin.comment
                for info in zin.infolist():
                    zout.writestr(copy(info), patched if info.filename == part else zin.read(info.filename))
            with ZipFile(temp_path) as check:
                if check.testzip() is not None:
                    raise ValueError('输出文件完整性校验失败')
                for info in zin.infolist():
                    if info.filename != part and check.read(info.filename) != zin.read(info.filename):
                        raise ValueError(f'非目标内容发生变化：{info.filename}')
            # Windows rename 不覆盖已存在文件。
            if output.exists():
                raise ValueError('输出文件已存在')
            temp_path.rename(output)
        finally:
            temp_path.unlink(missing_ok=True)
    stats['output'] = str(output)
    return stats


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('file', nargs='?', help='输入 XLSX 路径；省略时弹出选择窗口')
    parser.add_argument('-o', '--output', type=Path)
    parser.add_argument('--sheet', help='数据工作表名称，默认自动识别')
    parser.add_argument('--year', type=int, help='最新销量年份；默认读取文件名日期中的年份')
    args = parser.parse_args()
    if not args.file:
        try:
            import tkinter as tk
            from tkinter.filedialog import askopenfilename
            root = tk.Tk()
            root.withdraw()
            args.file = askopenfilename(title='选择需要计算日销的文件', filetypes=[('Excel', '*.xlsx')])
            root.destroy()
        except Exception as exc:
            parser.error(f'无法打开文件选择器，请通过命令行传入文件路径：{exc}')
        if not args.file:
            return
    try:
        result = process(Path(args.file), args.output, sheet=args.sheet, year=args.year)
    except (ValueError, OSError, ET.ParseError) as exc:
        parser.exit(1, f'处理失败：{exc}\n')
    print(f"完成：{result['calculated']} 行日销，{result['blank']} 行历史不足留空，"
          f"{result['skipped']} 行最新销量为空跳过，{result['corrected']} 行销量下降已更正。")
    print(result['output'])


if __name__ == '__main__':
    main()
