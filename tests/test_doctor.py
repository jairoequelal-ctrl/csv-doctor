import base64
from copy import deepcopy
from http.server import HTTPServer
from io import BytesIO
import json
from pathlib import Path
import subprocess
import sys
import tempfile
import threading
import unittest
from urllib.request import Request, urlopen
from urllib.error import HTTPError
import zipfile
from csv_doctor.core import load_table, diagnose, apply, export, number, date_value, validate_recipe
from csv_doctor.server import make_handler

ROOT = Path(__file__).resolve().parents[1]

class DoctorTests(unittest.TestCase):
    def table(self, text):
        return load_table(text.encode(), delimiter=';')

    def test_leading_zeros(self):
        table = self.table('id;amount\n00123;5\n')
        self.assertEqual(apply(table, {'version': 1})['rows'][0]['values'][0], '00123')
        self.assertIn('leading_zero_identifier', diagnose(table)['issue_counts'])

    def test_locale(self):
        self.assertEqual(number('1.234,56', 'es'), '1234.56')
        self.assertEqual(number('1,234.56', 'en'), '1234.56')
        self.assertEqual(number('-50,00', 'es'), '-50.00')
        with self.assertRaises(ValueError): number('1,234.56', 'es')

    def test_ambiguous_number(self):
        self.assertIn('ambiguous_number', diagnose(self.table('amount\n1.234\n'))['issue_counts'])

    def test_dates(self):
        self.assertEqual(date_value('03/04/2026', 'DMY'), '2026-04-03')
        self.assertEqual(date_value('03/04/2026', 'MDY'), '2026-03-04')
        with self.assertRaises(ValueError): date_value('31/02/2026', 'DMY')

    def test_cp1252_preview_warning(self):
        table = load_table('name;id\nJosé;001\n'.encode('cp1252'), delimiter=';')
        self.assertEqual(table['rows'][0]['values'][0], 'José')
        self.assertTrue(table['warnings'])

    def test_malformed_requires_optin(self):
        table = self.table('a;b\n1;2;3\n')
        with self.assertRaises(ValueError): apply(table, {'version': 1})
        result = apply(table, {'version': 1, 'skip_malformed': True})
        self.assertEqual(result['quarantine'][0]['values'], ['1', '2', '3'])

    def test_broken_quoting_rejected(self):
        with self.assertRaises(ValueError): self.table('a;b\n1;"unfinished')

    def test_header_collisions(self):
        table = self.table(' a ;a;a_2;\n1;2;3;4\n')
        header = apply(table, {'version': 1, 'unique_headers': True, 'trim_headers': True})['header']
        self.assertEqual(len(set(header)), 4)
        self.assertEqual(header, ['a', 'a_2', 'a_2_2', 'column_4'])

    def test_failed_conversion_retained(self):
        table = self.table('a\n1,234.56\n')
        result = apply(table, {'version': 1, 'numeric_columns': [0], 'numeric_locale': 'es'})
        self.assertEqual(result['rows'][0]['values'], ['1,234.56'])
        self.assertEqual(len(result['unresolved']), 1)

    def test_empty_identity_never_deduplicated(self):
        result = apply(self.table('id;v\n;x\n;x\n'), {'version': 1, 'deduplicate_by': [0]})
        self.assertEqual(len(result['rows']), 2)
        self.assertEqual(len(result['quarantine']), 0)

    def test_source_immutable_and_rows_accounted(self):
        table = load_table((ROOT/'examples/broken.csv').read_bytes(), delimiter=';')
        original = deepcopy(table)
        recipe = json.loads((ROOT/'examples/recipe.json').read_text())
        result = apply(table, recipe)
        self.assertEqual(table, original)
        self.assertEqual(len(result['rows']), 3)
        self.assertEqual(len(result['unresolved']), 2)
        ids = [r['source_row'] for r in result['rows'] + result['quarantine']]
        self.assertEqual(sorted(ids), list(range(2, 7)))

    def test_recipe_rejects_invalid_and_overlapping_columns(self):
        for recipe in ({'version':1,'numeric_columns':[True]}, {'version':1,'date_columns':[{}]},
                       {'version':1,'numeric_columns':[0],'date_columns':[0]}, {'version':1,'code':'exec'}):
            with self.assertRaises(ValueError): validate_recipe(recipe, 2)

    def test_formula_protection_is_audited(self):
        table = self.table('name;amount\n=SUM(A1:A2);-5\n')
        with tempfile.TemporaryDirectory() as tmp:
            export(tmp, table, apply(table, {'version':1}))
            self.assertIn("'=SUM(A1:A2),-5", (Path(tmp)/'clean.csv').read_text(encoding='utf-8-sig'))
            report = json.loads((Path(tmp)/'changes.json').read_text())
            self.assertEqual(len(report['csv_export_changes']), 1)

    def test_replay_matches_export(self):
        source = ROOT/'examples/broken.csv'
        table = load_table(source.read_bytes(), source.name, delimiter=';')
        recipe = json.loads((ROOT/'examples/recipe.json').read_text())
        with tempfile.TemporaryDirectory() as tmp:
            first, second = Path(tmp)/'first', Path(tmp)/'second'
            export(first, table, apply(table, recipe))
            subprocess.run([sys.executable, str(first/'replay.py'), str(source), str(second)], check=True)
            for filename in ('clean.csv', 'changes.json', 'quarantine.json', 'recipe.json'):
                self.assertEqual((first/filename).read_bytes(), (second/filename).read_bytes())

    def test_xlsx_preserves_underlying_values_and_formulas(self):
        try: from openpyxl import Workbook
        except ImportError: self.skipTest('Optional XLSX dependency')
        book = Workbook(); sheet = book.active
        sheet.append(['id', 'formula']); sheet.append([123, '=1+1'])
        sheet['A2'].number_format = '00000'
        stream = BytesIO(); book.save(stream)
        table = load_table(stream.getvalue(), 'test.xlsx')
        self.assertEqual(table['rows'][0]['values'], ['123', '=1+1'])
        with self.assertRaises(ValueError): load_table(stream.getvalue(), 'test.xlsx', sheet='missing')

class LocalAPITests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.server = HTTPServer(('127.0.0.1', 0), make_handler())
        cls.thread = threading.Thread(target=cls.server.serve_forever, daemon=True); cls.thread.start()
        cls.base = f'http://127.0.0.1:{cls.server.server_port}'

    @classmethod
    def tearDownClass(cls):
        cls.server.shutdown(); cls.server.server_close(); cls.thread.join()

    def request(self, path, body, headers=None):
        req = Request(self.base+path, json.dumps(body).encode(), headers=headers or {'Content-Type':'application/json'})
        return urlopen(req, timeout=5)

    def test_import_preview_export(self):
        with self.request('/api/import', {'data':base64.b64encode(b'id;name\n001; Ana \n').decode(), 'delimiter': ';'}) as response:
            imported = json.load(response)
        request = {'token':imported['token'], 'recipe':{'version':1,'trim_values':True}}
        with self.request('/api/preview', request) as response:
            preview = json.load(response)
        self.assertEqual(preview['rows'][0]['values'], ['001','Ana'])
        with self.request('/api/export', request) as response:
            archive = zipfile.ZipFile(BytesIO(response.read()))
        self.assertEqual(set(archive.namelist()), {'clean.csv','recipe.json','changes.json','quarantine.json','replay.py'})

    def test_cross_origin_blocked(self):
        with self.assertRaises(HTTPError) as error:
            self.request('/api/import', {}, {'Content-Type':'application/json','Origin':'https://evil.example'})
        self.assertEqual(error.exception.code, 403)

    def test_non_object_request_rejected(self):
        with self.assertRaises(HTTPError) as error: self.request('/api/import', [])
        self.assertEqual(error.exception.code, 400)

if __name__ == '__main__': unittest.main()
