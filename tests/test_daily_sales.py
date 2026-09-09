from datetime import date, timedelta
from decimal import Decimal as D
from pathlib import Path
import tempfile
import unittest
from zipfile import ZipFile
from xml.etree import ElementTree as ET

from daily_sales import Point, calculate, patch_xml, process, NS, discover_columns, number


def points(values, days=None, daily=None):
    return [Point(date(2026, 1, 1) + timedelta(days=(days or list(range(len(values))))[i]),
                  D(value), None if daily is None or daily[i] is None else D(daily[i]))
            for i, value in enumerate(values)]


class CalculationTests(unittest.TestCase):
    def test_missing_history(self):
        self.assertEqual(calculate(points([100])), (None, None))

    def test_small_sales_round_half_up(self):
        self.assertEqual(calculate(points([100, 105], [0, 2])), (3, None))

    def test_drop(self):
        self.assertEqual(calculate(points([12000, 11000])), (0, D(12000)))

    def test_increase_uses_first_plateau_date(self):
        self.assertEqual(calculate(points([12000, 12000, 13000], [0, 4, 10])), (100, None))

    def test_unchanged_retain(self):
        self.assertEqual(calculate(points([12000, 12000], [0, 5], [100, None])), (100, None))

    def test_unchanged_cap(self):
        self.assertEqual(calculate(points([12000, 12000, 12000], [0, 5, 20], [100, 100, None])), (50, None))

    def test_missing_daily_reconstruct(self):
        self.assertEqual(calculate(points([11000, 12000, 12000], [0, 10, 30])), (50, None))

    def test_no_growth_no_daily(self):
        self.assertEqual(calculate(points([12000, 12000], [0, 10])), (0, None))

    def test_negative_daily_clamp(self):
        self.assertEqual(calculate(points([12000, 12000], [0, 10], [-10, None])), (0, None))

    def test_cap_crossing_uses_previous_jump(self):
        self.assertEqual(calculate(points([97000, 99000, 100000], [0, 10, 11])), (200, None))

    def test_cap_plateau_keeps_jump(self):
        self.assertEqual(calculate(points([97000, 99000, 100000, 100000], [0, 10, 11, 100])), (200, None))

    def test_above_cap_new_jump(self):
        self.assertEqual(calculate(points([99000, 100000, 100000, 104000, 104000], [0, 1, 5, 21, 80])), (200, None))

    def test_above_cap_drop(self):
        self.assertEqual(calculate(points([104000, 100000])), (0, D(104000)))

    def test_no_usable_jump(self):
        self.assertEqual(calculate(points([100000, 100000])), (None, None))

    def test_zero_sales_is_value(self):
        self.assertEqual(calculate(points([0, 0])), (0, None))


class FileTests(unittest.TestCase):
    def test_missing_sales_markers_and_actual_dates(self):
        for marker in ['<is><t>断货</t></is>', '<is><t>#N/A</t></is>']:
            cell = ET.fromstring(f'<c xmlns="{NS["s"]}" r="B2" t="inlineStr">{marker}</c>')
            self.assertIsNone(number(cell, [], sales=True))
        error = ET.fromstring(f'<c xmlns="{NS["s"]}" r="C2" t="e"><f>NA()</f><v>#N/A</v></c>')
        shared = ET.fromstring(f'<c xmlns="{NS["s"]}" r="D2" t="s"><v>0</v></c>')
        self.assertIsNone(number(error, [], sales=True))
        self.assertIsNone(number(shared, ['断货'], sales=True))
        history = []
        for offset, cell in [(0, ET.fromstring(f'<c xmlns="{NS["s"]}"><v>100</v></c>')),
                             (3, error), (7, shared),
                             (10, ET.fromstring(f'<c xmlns="{NS["s"]}"><v>130</v></c>'))]:
            value = number(cell, ['断货'], sales=True)
            if value is not None:
                history.append(Point(date(2026, 1, 1) + timedelta(days=offset), value))
        self.assertEqual(calculate(history), (3, None))

    def test_unknown_sales_text_still_rejected(self):
        cell = ET.fromstring(f'<c xmlns="{NS["s"]}" r="A2" t="inlineStr"><is><t>未知</t></is></c>')
        with self.assertRaises(ValueError):
            number(cell, [], sales=True)

    def test_patch_preserves_everything_else(self):
        raw = b'<worksheet><sheetData><row r="2"><c r="A2" s="7" t="n"><f>1+1</f><v>2</v></c><c r="C2" s="9"/></row></sheetData></worksheet>'
        result = patch_xml(raw, {'A2': 3, 'B2': 4, 'C2': None})
        self.assertEqual(result, b'<worksheet><sheetData><row r="2"><c r="A2" s="7"><v>3</v></c><c r="B2"><v>4</v></c><c r="C2" s="9"></c></row></sheetData></worksheet>')

    def test_shared_formula_refused(self):
        with self.assertRaises(ValueError):
            patch_xml(b'<row r="1"><c r="A1"><f t="shared" si="0"/><v>1</v></c></row>', {'A1': 2})

    def test_cross_year(self):
        row = ET.fromstring('<row xmlns="' + NS['s'] + '"><c r="A1" t="inlineStr"><is><t>12.31销量</t></is></c><c r="B1" t="inlineStr"><is><t>1.2销量</t></is></c></row>')
        cols = discover_columns(row, [], 2026)
        self.assertEqual(cols['sales'], [(date(2025, 12, 31), 'A'), (date(2026, 1, 2), 'B')])

    def test_full_package_and_scope(self):
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / 'test2026-09-08.xlsx'
            headers = ['9.02销量', '9.08销量', '9.02日销', '9.08日销']
            header = ''.join(f'<c r="{col}1" t="inlineStr"><is><t>{label}</t></is></c>' for col, label in zip('ABCD', headers))
            data = ('<row r="2"><c r="A2"><v>100</v></c><c r="B2"><v>115</v></c><c r="C2" s="4"><v>8</v></c><c r="D2" s="4"/></row>'
                    '<row r="3"><c r="A3"><v>100</v></c><c r="B3"><v>90</v></c></row>'
                    '<row r="4"><c r="A4"><v>100</v></c><c r="D4"><v>77</v></c></row>'
                    '<row r="5"><c r="B5"><v>100</v></c><c r="D5"><v>88</v></c></row>')
            sheet = f'<worksheet xmlns="{NS["s"]}"><sheetData><row r="1">{header}</row>{data}</sheetData></worksheet>'.encode()
            parts = {'xl/workbook.xml': f'<workbook xmlns="{NS["s"]}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="数据源" sheetId="1" r:id="rId1"/></sheets></workbook>'.encode(),
                     'xl/_rels/workbook.xml.rels': b'<Relationships><Relationship Id="rId1" Target="worksheets/sheet1.xml"/></Relationships>',
                     'xl/worksheets/sheet1.xml': sheet,
                     'xl/media/image1.png': b'original-image-bytes',
                     'xl/cellimages.xml': b'original-cell-image-metadata',
                     'xl/styles.xml': b'original-styles'}
            with ZipFile(source, 'w') as z:
                for key, value in parts.items():
                    z.writestr(key, value)
            original = source.read_bytes()
            result = process(source)
            self.assertEqual((result['calculated'], result['corrected'], result['skipped'], result['blank']), (2, 1, 1, 1))
            self.assertEqual(source.read_bytes(), original)
            with ZipFile(result['output']) as z:
                for key in parts:
                    if key != 'xl/worksheets/sheet1.xml':
                        self.assertEqual(z.read(key), parts[key])
                root = ET.fromstring(z.read('xl/worksheets/sheet1.xml'))
                cells = {c.get('r'): c for c in root.findall('.//s:c', NS)}
                self.assertEqual(cells['D2'].find('s:v', NS).text, '3')
                self.assertEqual(cells['D2'].get('s'), '4')
                self.assertEqual(cells['B3'].find('s:v', NS).text, '100')
                self.assertEqual(cells['D3'].find('s:v', NS).text, '0')
                self.assertEqual(cells['D4'].find('s:v', NS).text, '77')
                self.assertIsNone(cells['D5'].find('s:v', NS))
            with self.assertRaises(ValueError):
                process(source, source)
            with self.assertRaises(ValueError):
                process(source, Path(result['output']))


if __name__ == '__main__':
    unittest.main()
