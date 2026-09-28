from datetime import date
from decimal import Decimal
from pathlib import Path
import tempfile
import re
import unittest
from unittest.mock import patch, MagicMock
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from daily_sales import process, NS, CELL_PATTERN, patch_xml, Formula
from exchange_rates import fetch_rates, read_config


class PriceTests(unittest.TestCase):
    def make_source(self, directory):
        path = Path(directory) / 'prices2026-09-09.xlsx'
        headers = ['站点', '9.08价格', '人民币', '预估供货价', '9.02销量', '9.09销量', '9.02日销', '9.09日销']
        header = ''.join(f'<c r="{chr(65+i)}1" t="inlineStr"><is><t>{value}</t></is></c>' for i,value in enumerate(headers))
        rows = []
        for row, site, price, sales in [(2,'德国','10','107'), (3,'英国','20','107'), (4,'德国',None,'107'), (5,'德国','0','107'), (6,'德国','2',None)]:
            cells = f'<c r="A{row}" t="inlineStr"><is><t>{site}</t></is></c>'
            if price is not None:
                cells += f'<c r="B{row}"><v>{price}</v></c>'
            cells += f'<c r="C{row}" s="4"><f>B{row}*9</f><v>99</v></c><c r="D{row}" s="5"><v>88</v></c><c r="E{row}"><v>100</v></c>'
            if sales is not None:
                cells += f'<c r="F{row}"><v>{sales}</v></c>'
            cells += f'<c r="G{row}" s="3"><v>1</v></c>'
            rows.append(f'<row r="{row}">{cells}</row>')
        sheet = f'<worksheet xmlns="{NS["s"]}"><sheetData><row r="1">{header}</row>{"".join(rows)}</sheetData></worksheet>'
        with ZipFile(path,'w') as z:
            z.writestr('xl/workbook.xml', f'<workbook xmlns="{NS["s"]}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="数据源" r:id="r1"/></sheets></workbook>')
            z.writestr('xl/_rels/workbook.xml.rels','<Relationships><Relationship Id="r1" Target="worksheets/sheet1.xml"/></Relationships>')
            z.writestr('xl/worksheets/sheet1.xml',sheet)
            z.writestr('xl/media/image.png',b'preserve-image')
            z.writestr('xl/styles.xml',b'preserve-styles')
        return path

    def test_all_mode_date_skip_formula_and_scope(self):
        with tempfile.TemporaryDirectory() as directory:
            source = self.make_source(directory)
            original = source.read_bytes()
            with patch('exchange_rates.fetch_rates',return_value={'德国':Decimal('7.63')}) as fetch:
                result = process(source,all_calculations=True)
            self.assertEqual(fetch.call_args.args[:2], (date(2026,9,8), {'德国','英国'}))
            self.assertEqual(result['price_calculated'],3)
            self.assertEqual(result['price_skipped_empty'],1)
            self.assertEqual(result['price_missing_sites'],['英国'])
            with ZipFile(source) as before, ZipFile(result['output']) as after:
                for name in before.namelist():
                    if name != 'xl/worksheets/sheet1.xml':
                        self.assertEqual(before.read(name),after.read(name))
                root=ET.fromstring(after.read('xl/worksheets/sheet1.xml'))
                cells={c.get('r'):c for c in root.findall('.//s:c',NS)}
                self.assertEqual(cells['C2'].findtext('s:f',namespaces=NS),'B2*7.63')
                self.assertEqual(Decimal(cells['C2'].findtext('s:v',namespaces=NS)),Decimal('76.30'))
                self.assertEqual(cells['C2'].get('s'),'4')
                self.assertEqual(cells['C5'].findtext('s:v',namespaces=NS),'0.00')
                self.assertEqual(cells['C6'].findtext('s:f',namespaces=NS),'B6*7.63')
                old={c.get('r'):ET.tostring(c) for c in ET.fromstring(before.read('xl/worksheets/sheet1.xml')).findall('.//s:c',NS)}
                for ref in old:
                    if not ref.startswith('H') and ref not in ('C2','C5','C6'):
                        self.assertEqual(old[ref],ET.tostring(cells[ref]),ref)
            self.assertEqual(original,source.read_bytes())

    def test_default_does_not_read_config_or_database(self):
        with tempfile.TemporaryDirectory() as directory:
            with patch('exchange_rates.fetch_rates',side_effect=AssertionError('must not connect')):
                result=process(self.make_source(directory),db_config=Path('missing'))
            self.assertNotIn('price_calculated',result)

    def test_database_failure_does_not_publish_partial_file(self):
        with tempfile.TemporaryDirectory() as directory:
            output=Path(directory)/'result.xlsx'
            with patch('exchange_rates.fetch_rates',side_effect=ValueError('unavailable')):
                with self.assertRaises(ValueError):
                    process(self.make_source(directory),output,all_calculations=True)
            self.assertFalse(output.exists())

    def test_shared_price_group_complete_only(self):
        raw=b'<row r="2"><c r="C2" s="4"><f t="shared" si="1" ref="C2:D2">A2*7</f><v>7</v></c><c r="D2"><f t="shared" si="1"/><v>14</v></c></row>'
        with self.assertRaises(ValueError):
            patch_xml(raw, {'C2':Formula('A2*8',8)}, replace_shared_groups=True)
        result=ET.fromstring(patch_xml(raw, {'C2':Formula('A2*8',8),'D2':Formula('B2*8',16)}, replace_shared_groups=True))
        self.assertEqual([c.find('f').text for c in result],['A2*8','B2*8'])
        self.assertEqual(result[0].get('s'),'4')

    def test_cli_a_switch(self):
        from daily_sales import main
        stats=dict(calculated=0,blank=0,skipped=0,corrected=0,output='out.xlsx',price_calculated=0,price_skipped_rate=0,price_skipped_empty=0,price_missing_sites=[])
        for switch,expected in [([],False),(['-a'],True)]:
            with patch('sys.argv',['daily_sales.py','input.xlsx',*switch]), patch('daily_sales.process',return_value=stats) as run, patch('builtins.print'):
                main()
                self.assertEqual(run.call_args.kwargs['all_calculations'],expected)


class RateTests(unittest.TestCase):
    def test_config_markdown(self):
        with tempfile.TemporaryDirectory() as directory:
            path=Path(directory)/'config.md'
            path.write_text('- 主机： localhost\n- 用户名： test\n- 密码： dummy\n',encoding='utf-8')
            self.assertEqual(read_config(path)['port'],3306)
            self.assertEqual(read_config(path)['password'],'dummy')

    def fake_connection(self,rows):
        connection=MagicMock()
        cursor=connection.cursor.return_value.__enter__.return_value
        cursor.fetchall.return_value=rows
        return connection,cursor

    def test_exact_date_and_parameterized_country(self):
        connection,cursor=self.fake_connection([('德国','EUR/CNY',Decimal('7.63'))])
        with patch('exchange_rates.read_config',return_value={'database':'test'}),patch('pymysql.connect',return_value=connection):
            self.assertEqual(fetch_rates(date(2026,9,8),{'德国'}),{'德国':Decimal('7.63')})
        self.assertEqual(cursor.execute.call_args.args[1],(date(2026,9,8),'德国'))
        self.assertIn('record_date=%s',cursor.execute.call_args.args[0])
        connection.close.assert_called_once()

    def test_invalid_direction_and_duplicate_rejected(self):
        for rows in [[('德国','CNY/EUR',Decimal('1'))],[('德国','EUR/CNY',Decimal('7'))]*2,[('德国','EUR/CNY',Decimal('0'))]]:
            connection,_=self.fake_connection(rows)
            with patch('exchange_rates.read_config',return_value={'database':'test'}),patch('pymysql.connect',return_value=connection):
                with self.assertRaises(ValueError):
                    fetch_rates(date(2026,9,8),{'德国'})

    def test_no_rate_is_empty(self):
        connection,_=self.fake_connection([])
        with patch('exchange_rates.read_config',return_value={'database':'test'}),patch('pymysql.connect',return_value=connection):
            self.assertEqual(fetch_rates(date(2026,9,8),{'德国'}),{})


ROOT = Path(__file__).resolve().parents[1]
COMBINED = ROOT / 'outputs/太阳能喷泉数据汇总2026-09-08_价格与日销.xlsx'
DAILY = ROOT / 'outputs/太阳能喷泉数据汇总2026-09-08_日销公式小数.xlsx'


@unittest.skipUnless(COMBINED.exists() and DAILY.exists(), '需要本地联合计算样例')
class RealPriceOutputTests(unittest.TestCase):
    def test_only_rmb_changed_against_daily_output(self):
        with ZipFile(DAILY) as old, ZipFile(COMBINED) as new:
            self.assertEqual(old.namelist(),new.namelist())
            for name in old.namelist():
                if name != 'xl/worksheets/sheet3.xml':
                    self.assertEqual(old.read(name),new.read(name),name)
            before=old.read('xl/worksheets/sheet3.xml'); after=new.read('xl/worksheets/sheet3.xml')
            def strip_rmb(match):
                ref=re.search(rb'\br="([A-Z]+\d+)"',match.group())[1].decode()
                return b'' if re.fullmatch(r'BM\d+',ref) else match.group()
            self.assertEqual(CELL_PATTERN.sub(strip_rmb,before),CELL_PATTERN.sub(strip_rmb,after))
            before_cells={c.get('r'):c for c in ET.fromstring(before).findall('.//s:c',NS)}
            after_cells={c.get('r'):c for c in ET.fromstring(after).findall('.//s:c',NS)}
            count=0
            for ref,cell in after_cells.items():
                if not re.fullmatch(r'BM\d+',ref):
                    continue
                old_cell=before_cells.get(ref)
                if old_cell is not None and ET.tostring(cell)==ET.tostring(old_cell):
                    continue
                self.assertEqual(cell.get('s'),None if old_cell is None else old_cell.get('s'))
                formula=cell.findtext('s:f',namespaces=NS)
                match=re.fullmatch(r'(BL\d+)\*([0-9.]+)',formula)
                self.assertIsNotNone(match,ref)
                price=Decimal(after_cells[match[1]].findtext('s:v',namespaces=NS))
                self.assertEqual(Decimal(cell.findtext('s:v',namespaces=NS)),price*Decimal(match[2]))
                count+=1
            self.assertGreater(count,0)
