"""真实样例验收；没有本地样例或结果时跳过，不依赖外部服务。"""
from pathlib import Path
import re
import unittest
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from daily_sales import NS, CELL_PATTERN, number, workbook_parts

ROOT = Path(__file__).resolve().parents[1]
SOURCE = ROOT / '太阳能喷泉数据汇总2026-09-08.xlsx'
OUTPUT = ROOT / 'outputs/太阳能喷泉数据汇总2026-09-08_日销已计算.xlsx'


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
            changed = {ref for ref in old_raw.keys() | new_raw.keys() if old_raw.get(ref) != new_raw.get(ref)}
            for ref in changed:
                col, row = re.fullmatch(r'([A-Z]+)(\d+)', ref).groups()
                self.assertIn(col, {'DW', 'GB'}, ref)
                self.assertNotEqual(row, '1')
                latest = number(old_cells.get('DW' + row), strings, sales=True)
                self.assertIsNotNone(latest, ref)
                if ref in old_cells and ref in new_cells:
                    self.assertEqual(old_cells[ref].get('s'), new_cells[ref].get('s'), ref)
                value = number(new_cells.get(ref), strings, sales=True)
                if col == 'GB' and value is not None:
                    self.assertGreaterEqual(value, 0)
                    self.assertEqual(value, value.to_integral_value())
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
            self.assertEqual(CELL_PATTERN.sub(remove_changed, old), CELL_PATTERN.sub(remove_changed, new))


if __name__ == '__main__':
    unittest.main()
