"""真实样例验收；没有本地样例或结果时跳过，不依赖外部服务。"""
from pathlib import Path
import re
import ast
from decimal import Decimal, ROUND_HALF_UP
import unittest
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from daily_sales import NS, CELL_PATTERN, number, workbook_parts

ROOT = Path(__file__).resolve().parents[1]
SOURCE = ROOT / '太阳能喷泉数据汇总2026-09-08.xlsx'
OUTPUT = ROOT / 'outputs/太阳能喷泉数据汇总2026-09-08_日销公式小数.xlsx'


@unittest.skipUnless(SOURCE.exists() and OUTPUT.exists(), '本地样例结果尚不存在')
class SampleOutputTests(unittest.TestCase):
    def test_only_authorized_cells_changed_and_images_preserved(self):
        with ZipFile(SOURCE) as source, ZipFile(OUTPUT) as output:
            strings, sheets = workbook_parts(source)
            part = sheets['数据源']
            self.assertEqual(source.namelist(), output.namelist())
            for name in source.namelist():
                if name != part:
                    self.assertEqual(source.read(name), output.read(name), name)
            old, new = source.read(part), output.read(part)
            old_cells = {c.get('r'): c for c in ET.fromstring(old).findall('.//s:sheetData/s:row/s:c', NS)}
            new_cells = {c.get('r'): c for c in ET.fromstring(new).findall('.//s:sheetData/s:row/s:c', NS)}
            old_raw = {re.search(rb'\br="([A-Z]+\d+)"', m.group())[1].decode(): m.group() for m in CELL_PATTERN.finditer(old)}
            new_raw = {re.search(rb'\br="([A-Z]+\d+)"', m.group())[1].decode(): m.group() for m in CELL_PATTERN.finditer(new)}
            def evaluate(expression):
                def visit(node):
                    if isinstance(node, ast.Constant):
                        return Decimal(str(node.value))
                    if isinstance(node, ast.Name):
                        value = number(new_cells.get(node.id), strings)
                        self.assertIsNotNone(value, node.id)
                        return value
                    if isinstance(node, ast.BinOp):
                        left, right = visit(node.left), visit(node.right)
                        if isinstance(node.op, ast.Sub):
                            return left - right
                        if isinstance(node.op, ast.Div):
                            return left / right
                    if isinstance(node, ast.Call):
                        args = [visit(arg) for arg in node.args]
                        if node.func.id == 'MAX':
                            return max(args)
                        if node.func.id == 'MIN':
                            return min(args)
                        if node.func.id == 'ROUND':
                            return args[0].quantize(Decimal(1).scaleb(-int(args[1])), rounding=ROUND_HALF_UP)
                    raise AssertionError(ast.dump(node))
                return visit(ast.parse(expression, mode='eval').body)

            for ref, cell in new_cells.items():
                if re.fullmatch(r'GB\d+', ref) and cell.find('s:f', NS) is not None:
                    self.assertEqual(evaluate(cell.find('s:f', NS).text), number(cell, strings), ref)

            changed = {ref for ref in old_raw.keys() | new_raw.keys() if old_raw.get(ref) != new_raw.get(ref)}
            for ref in changed:
                col, row = re.fullmatch(r'([A-Z]+)(\d+)', ref).groups()
                self.assertIn(col, {'DW', 'GB'}, ref)
                if col == 'GB':
                    source_style = old_cells.get('GA' + row)
                    if source_style is not None and source_style.get('s') is not None:
                        self.assertEqual(new_cells[ref].get('s'), source_style.get('s'), ref)
                if row == '1':
                    self.assertEqual(new_cells[ref].find('s:v', NS).text, old_cells[ref].find('s:v', NS).text)
                    continue
                latest = number(old_cells.get('DW' + row), strings, sales=True)
                if latest is None:
                    before = old_cells.get(ref)
                    after = new_cells.get(ref)
                    self.assertEqual(None if before is None else before.findtext('s:v', namespaces=NS), after.findtext('s:v', namespaces=NS))
                    continue
                value = number(new_cells.get(ref), strings, sales=True)
                if col == 'GB' and value is not None:
                    self.assertIsNotNone(new_cells[ref].find('s:f', NS))
                    self.assertGreaterEqual(value, 0)
                    self.assertNotIn('ROUND(', new_cells[ref].find('s:f', NS).text)
                if col == 'DW':
                    previous = None
                    for cell in ET.fromstring(old).findall(f'.//s:row[@r="{row}"]/s:c', NS):
                        column = re.sub(r'\d', '', cell.get('r'))
                        # 销量历史固定为本样例 BO:DV。
                        if len(column) == 2 and 'BO' <= column <= 'DV':
                            candidate = number(cell, strings, sales=True)
                            if candidate is not None:
                                previous = candidate
                    self.assertIsNotNone(previous)
                    self.assertLess(latest, previous)
                    self.assertEqual(value, previous)
                    self.assertEqual(number(new_cells['GB' + row], strings), 0)

            def remove_changed(match):
                ref = re.search(rb'\br="([A-Z]+\d+)"', match.group())[1].decode()
                return b'' if ref in changed else match.group()

            # 除确切获准变更的单元格外，整张工作表 XML 字节保持不变。
            def without_cols(data):
                return re.sub(rb'<cols\b[^>]*>.*?</cols>', b'', data, flags=re.DOTALL)
            self.assertEqual(without_cols(CELL_PATTERN.sub(remove_changed, old)), without_cols(CELL_PATTERN.sub(remove_changed, new)))
            def column_attrs(data):
                result = {}
                for col in ET.fromstring(data).find('s:cols', NS):
                    for index in range(int(col.get('min')), int(col.get('max')) + 1):
                        result[index] = {key: value for key, value in col.attrib.items() if key not in ('min', 'max')}
                return result
            before_cols, after_cols = column_attrs(old), column_attrs(new)
            for index in before_cols.keys() | after_cols.keys():
                self.assertEqual(after_cols[index], before_cols[183 if index == 184 else index])


if __name__ == '__main__':
    unittest.main()
