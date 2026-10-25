"""Python コアの単体テスト（不具合の再発防止）。PowerShell 版との一致は test_parity.py で確認する。"""
import json
import re
import unittest
import zipfile
from collections import OrderedDict
from unittest import mock

from support import CORE as core, config, context, quiet, record, rewrite_zip, workdir, write_xlsx


def api(**values):
    with quiet():
        return core.normalize_api(record(**values), 0, config())


class QueryPlacementTests(unittest.TestCase):
    def test_only_get_uses_parameters_tab(self):
        # JMeter 5.6.3: DELETE の Parameters は URL とボディの両方に送られ、HEAD / TRACE の Parameters は送られない
        self.assertEqual(api(method='GET', query='q=value')['args'], [('q', 'value', True)])
        for method in ('DELETE', 'OPTIONS', 'HEAD', 'TRACE', 'POST', 'PUT', 'PATCH'):
            with self.subTest(method=method):
                a = api(method=method, query='q=value')
                self.assertEqual(a['path'], '/api?q=value')
                self.assertEqual(a['args'], [])

    def test_body_is_ignored_for_methods_without_body(self):
        for method in ('GET', 'DELETE', 'OPTIONS', 'HEAD', 'TRACE'):
            with self.subTest(method=method):
                a = api(method=method, body_type='raw', body='text')
                self.assertEqual((a['body'], a['kind']), ('', ''))
                self.assertTrue(any(method in w for w in core.WARNINGS))


class EncodingTests(unittest.TestCase):
    def test_raw_pairs_are_encoded_by_jmeter(self):
        self.assertEqual(api(method='GET', query='q=a b 日本')['args'], [('q', 'a b 日本', True)])

    def test_mixed_encoding_is_completed_without_double_encoding(self):
        a = api(method='GET', query='a%26b=\nq=x%20y 日本語')
        self.assertEqual(a['args'], [('a%26b', '', False), ('q', 'x%20y%20%E6%97%A5%E6%9C%AC%E8%AA%9E', False)])

    def test_encoded_name_with_plain_value(self):
        self.assertEqual(api(method='GET', query='filter%5Bname%5D=x')['args'], [('filter%5Bname%5D', 'x', False)])

    def test_url_query_plus_means_space(self):
        self.assertEqual(api(method='POST', url='https://example.test/?q=a+b&r=c%2Bd')['path'], '/?q=a%20b&r=c%2Bd')
        self.assertEqual(api(method='GET', url='https://example.test/?q=a+b')['args'], [('q', 'a%20b', False)])

    def test_plus_in_query_column_is_literal(self):
        self.assertEqual(api(method='GET', query='q=a+b')['args'], [('q', 'a+b', True)])

    def test_encoded_query_keeps_structural_separators(self):
        self.assertEqual(api(method='POST', query='ids=a%20b,c%2Cd')['path'], '/api?ids=a%20b,c%2Cd')

    def test_variables_are_not_encoded(self):
        self.assertEqual(api(method='POST', query='q=${VALUE} x%20')['path'], '/api?q=${VALUE}%20x%20')

    def test_form_pairs_use_the_same_rules(self):
        a = api(body_type='form', body='a%26b=x%26y%25\nplain=1+2')
        self.assertEqual(a['args'], [('a%26b', 'x%26y%25', False), ('plain', '1+2', True)])

    def test_multipart_values_are_decoded(self):
        a = api(body_type='multipart', body='a%26b=%20x%26y%25%E6%97%A5%E6%9C%AC%E8%AA%9E%20\nname=日本 語')
        self.assertEqual(a['args'], [('a&b', ' x&y%日本語 ', False), ('name', '日本 語', False)])
        self.assertEqual(core.pct_decode('%FF%FE'), '%FF%FE', 'UTF-8 として読めない場合は元のまま')


class UrlTests(unittest.TestCase):
    def test_invalid_urls_are_skipped(self):
        for url in ('http:///x', 'ftp://example.test/x', 'https://example.test:99999/x', 'https://exa mple.test/x'):
            with self.subTest(url=url):
                self.assertIsNone(api(url=url))
                self.assertTrue(core.WARNINGS)

    def test_userinfo_is_removed(self):
        a = api(url='https://user:pass@example.test/x')
        self.assertEqual((a['url'], a['host']), ('https://example.test/x', 'example.test'))
        self.assertTrue(any('認証情報' in w for w in core.WARNINGS))

    def test_variable_scheme_host_and_port(self):
        a = api(url='${PROTOCOL}://${HOST}:${PORT}/v')
        self.assertEqual((a['scheme'], a['host'], a['port'], a['path']), ('${PROTOCOL}', '${HOST}', '${PORT}', '/v'))

    def test_ipv6_and_trace(self):
        a = api(url='http://[::1]:8080/x', method='TRACE')
        self.assertEqual((a['host'], a['port']), ('[::1]', '8080'))
        self.assertFalse(core.WARNINGS)

    def test_cookie_header_warning_depends_on_cookie_manager(self):
        with quiet():
            core.normalize_api(record(headers='Cookie: a=b'), 0, config())
            self.assertTrue(any('Cookie' in w for w in core.WARNINGS))
        with quiet():
            core.normalize_api(record(headers='Cookie: a=b'), 0, config(cookie_manager='OFF'))
            self.assertFalse(core.WARNINGS)


class TableTests(unittest.TestCase):
    def test_body_whitespace_is_kept(self):
        rows = [(1, ['url', 'method', 'body_type', 'body', 'pre_script']),
                (2, ['https://example.test', 'PUT', 'raw', '  text\n ', '  log.info("x")  ']),
                (3, ['https://example.test', 'PUT', 'raw', '   ', ''])]
        recs = core.load_api_table(rows, 'test')
        self.assertEqual(recs[0]['body'], '  text\n ')
        self.assertEqual(recs[0]['pre_script'], 'log.info("x")')
        self.assertEqual(recs[1]['body'], '')
        self.assertEqual(api(method='PUT', body_type='raw', body='  text\n ')['body'], '  text\n ')

    def test_unclosed_csv_quote_is_an_error(self):
        with self.assertRaises(core.ToolError) as cm:
            core.parse_csv('url,body\nhttps://example.test,"open\n', ',', 'api.csv')
        self.assertIn('2 行目', str(cm.exception))

    def test_settings_out_of_range_use_defaults(self):
        for values in ({'threads': '-1'}, {'threads': '0'}, {'loops': '-2'}, {'ramp_up': '-1'}, {'duration': '-2'},
                       {'connect_timeout': '-5'}, {'threads': '99999999999'}):
            with self.subTest(values=values):
                cfg = config(**values)
                key = list(values)[0]
                self.assertEqual(cfg[key], core._setting_default(key))
                self.assertTrue(core.WARNINGS)
        cfg = config(random_min='10', random_max='1')
        self.assertEqual((cfg['random_min'], cfg['random_max']), ('1', '99999'))
        self.assertEqual(config(loops='-1', threads='1')['loops'], '-1')
        self.assertEqual(config(csv_data='none')['csv_data'], 'OFF')


class ExcelTests(unittest.TestCase):
    def test_xhhhh_escape_round_trip(self):
        self.assertEqual(core.xl_unescape('_x0041_ _x005F_x0041_ _x000D_'), 'A _x0041_ \r')
        self.assertEqual(core.xl_escape('_x0041_'), '_x005F_x0041_')
        with workdir() as td:
            write_xlsx(td / 'a.xlsx', [('API一覧', [['url', 'body'], ['https://example.test', '_x0041_ 日本語']])])
            self.assertEqual(core.xlsx_read(str(td / 'a.xlsx'))[0][1][1][1], ['https://example.test', '_x0041_ 日本語'])

    def test_relative_sheet_target(self):
        with workdir() as td:
            write_xlsx(td / 'a.xlsx', [('API一覧', [['url'], ['https://example.test']])])
            rewrite_zip(td / 'a.xlsx', td / 'b.xlsx', lambda n, d: d.replace(
                b'worksheets/sheet1.xml', b'./worksheets/../worksheets/sheet1.xml') if n.endswith('workbook.xml.rels') else d)
            self.assertEqual(core.xlsx_read(str(td / 'b.xlsx'))[0][1][1][1], ['https://example.test'])

    def test_unreadable_cells_fail_only_on_used_sheets(self):
        with workdir() as td:
            write_xlsx(td / 'a.xlsx', [('API一覧', [['url'], ['https://example.test']]), ('メモ', [['x']])])
            # 1 枚目の A2 を計算結果の無い数式、2 枚目の A1 をエラー値にする
            with zipfile.ZipFile(str(td / 'a.xlsx')) as z:
                s1 = z.read('xl/worksheets/sheet1.xml').decode('utf-8')
                s2 = z.read('xl/worksheets/sheet2.xml').decode('utf-8')
            s1 = re.sub(r'<c r="A2"([^>]*?) t="(s|inlineStr)"([^>]*)>.*?</c>', r'<c r="A2"\1\3><f>1+1</f></c>', s1, flags=re.S)
            s2 = re.sub(r'<c r="A1"([^>]*?) t="(s|inlineStr)"([^>]*)>.*?</c>', r'<c r="A1"\1\3 t="e"><v>#REF!</v></c>', s2, flags=re.S)
            rewrite_zip(td / 'a.xlsx', td / 'b.xlsx', lambda n, d: s1.encode() if n.endswith('sheet1.xml') else (
                s2.encode() if n.endswith('sheet2.xml') else d))
            core.xlsx_read(str(td / 'b.xlsx'))
            with self.assertRaises(core.ToolError) as cm:
                core.xlsx_check(str(td / 'b.xlsx'), 'API一覧')
            self.assertIn('A2', str(cm.exception))
            with self.assertRaises(core.ToolError):
                core.xlsx_check(str(td / 'b.xlsx'), 'メモ')
            # 使うのは「API一覧」だけなので、「メモ」のエラー値は generate では問題にならない
            rewrite_zip(td / 'a.xlsx', td / 'c.xlsx', lambda n, d: s2.encode() if n.endswith('sheet2.xml') else d)
            with quiet():
                raw, recs = core.load_input(str(td / 'c.xlsx'), None, {})
            self.assertEqual(recs[0]['url'], 'https://example.test')


class OpenApiTests(unittest.TestCase):
    def path_of(self, style, explode, example, name='id'):
        op = {'parameters': [{'name': name, 'in': 'path', 'required': True, 'style': style, 'explode': explode,
                              'example': example}]}
        with quiet():
            return core.build_operation(context(), dict(path='/x/{%s}' % name, item={}, op=op, method='GET', key='get',
                                                         is_any=False), 'test')['url']

    def test_path_styles(self):
        cases = [('simple', False, ['a b', 'c/d'], '/x/a%20b,c%2Fd'), ('simple', True, {'k': 'v', 'n': 1}, '/x/k=v,n=1'),
                 ('simple', False, {'k': 'v'}, '/x/k,v'), ('label', False, ['a', 'b'], '/x/.a,b'),
                 ('label', True, ['a', 'b'], '/x/.a.b'), ('label', True, {'k': 'v', 'n': 1}, '/x/.k=v.n=1'),
                 ('matrix', False, 'blue', '/x/;id=blue'), ('matrix', False, ['a', 'b'], '/x/;id=a,b'),
                 ('matrix', True, ['a', 'b'], '/x/;id=a;id=b'), ('matrix', True, {'k': 'v'}, '/x/;k=v'),
                 ('matrix', False, {'k': 'v', 'n': 'x y'}, '/x/;id=k,v,n,x%20y')]
        for style, explode, example, expected in cases:
            with self.subTest(style=style, explode=explode, example=example):
                self.assertEqual(self.path_of(style, explode, example), expected)

    def test_query_styles_keep_values_distinguishable(self):
        q = core.query_pairs
        self.assertEqual(q('ids', ['a,b', 'c d'], 'form', False), [('ids', 'a%2Cb,c%20d', True)])
        self.assertEqual(q('ids', ['a', 'b'], 'form', True), [('ids', 'a', False), ('ids', 'b', False)])
        self.assertEqual(q('ids', ['x', 'y'], 'spaceDelimited', False), [('ids', 'x%20y', True)])
        self.assertEqual(q('ids', ['p', 'q'], 'pipeDelimited', False), [('ids', 'p%7Cq', True)])
        self.assertEqual(q('f', {'a': 'b c'}, 'deepObject', True), [('f[a]', 'b c', False)])
        self.assertEqual(q('o', core.to_gen({'k': 1, 'v': 'x,y'}), 'form', False), [('o', 'k,1,v,x%2Cy', True)])
        self.assertEqual(core.oa_row_pairs(q('ids', ['a,b', 'c'], 'form', False)), 'ids=a%2Cb,c')
        self.assertEqual(api(method='GET', query='ids=a%2Cb,c')['args'], [('ids', 'a%2Cb,c', False)])

    def test_row_pairs_encode_names_and_values_separately(self):
        self.assertEqual(core.oa_row_pairs([('a&b', 'x')]), 'a%26b=x')
        self.assertEqual(core.oa_row_pairs([('a', ' x&y ')]), 'a=%20x%26y%20')
        self.assertEqual(api(method='GET', query=core.oa_row_pairs([('a&b', '')]))['args'], [('a%26b', '', False)])

    def test_form_encoding_and_content_parameters(self):
        op = {'requestBody': {'content': {'application/x-www-form-urlencoded': {
            'example': {'ids': ['a', 'b'], 'n': 'x y'}, 'encoding': {'ids': {'explode': False}}}}}}
        self.assertEqual(core.build_body(context(), op)['pairs'], [('ids', 'a,b', True), ('n', 'x y', False)])
        p = {'name': 'filter', 'in': 'query', 'content': {'application/json': {'example': {'ids': [1, 2]}}}}
        self.assertEqual(core.param_text(context(), p), '{"ids":[1,2]}')
        p = {'name': 'q', 'in': 'query', 'content': {'text/plain': {'example': 'abc'}}}
        self.assertEqual(core.param_text(context(), p), 'abc')

    def test_required_only_skips_optional_body(self):
        op = {'requestBody': {'content': {'application/json': {'schema': {'type': 'object'}}}}}
        self.assertIsNone(core.build_body(context(required_only=True), op))
        op['requestBody']['required'] = True
        self.assertIsNotNone(core.build_body(context(required_only=True), op))

    def test_accept_skips_wildcards(self):
        op = {'responses': {'200': {'content': {'*/*': {}, 'application/xml': {}}}}}
        self.assertEqual(core.pick_accept(context(), op), 'application/xml')

    def test_operation_and_path_servers(self):
        doc = {'servers': [{'url': 'https://api.example.test/v1'}]}
        ctx = context(doc)
        item = {'servers': [{'url': 'https://files.example.test:8443/up'}]}
        op = {'servers': [{'url': '/v2'}]}
        self.assertEqual(core.op_server_url(ctx, dict(op=op, item=item)), 'https://api.example.test/v2')
        self.assertEqual(core.op_server_url(ctx, dict(op={}, item=item)), 'https://files.example.test:8443/up')
        self.assertEqual(core.op_server_url(ctx, dict(op={}, item={})), '')

    def test_same_seed_generates_the_same_test_data(self):
        # データ生成は JMeter_Tools の openapi2jmx と同じ（同じシードなら同じ値）。v1.0.0 の出力で固定する
        schema = json.loads('''{"type": "object", "required": ["id", "name"], "properties": {
          "id": {"type": "integer", "minimum": 1, "maximum": 999},
          "name": {"type": "string", "minLength": 3, "maxLength": 8},
          "code": {"type": "string", "pattern": "^[A-Z]{2}-[0-9]{3}$"},
          "mail": {"type": "string", "format": "email"},
          "uuid": {"type": "string", "format": "uuid"},
          "at": {"type": "string", "format": "date-time"},
          "price": {"type": "number", "minimum": 0.5, "maximum": 99.5},
          "tags": {"type": "array", "items": {"type": "string", "enum": ["a", "b", "c"]}, "minItems": 2, "maxItems": 3,
                   "uniqueItems": true},
          "flag": {"type": "boolean"},
          "nested": {"type": "object", "properties": {"x": {"type": "integer", "multipleOf": 5, "minimum": 10, "maximum": 50}}}
        }}''', object_pairs_hook=OrderedDict)
        ctx = core.Ctx({}, 12345, {'required_only': False, 'use_examples': False, 'any_method': 'GET'})
        values = [core.to_json(core.gen_value(ctx, schema, 0, []), 0, False) for _ in range(3)]
        self.assertEqual(values, [
            '{"id":316,"name":"AbyX0eYQ","code":"II-555","mail":"ysdq4s1j@example.com",'
            '"uuid":"8236478b-76df-40b3-8ecf-378da1f9955e","at":"2026-07-07T00:36:43Z","price":16.95,"tags":["a","c"],'
            '"flag":true,"nested":{"x":15}}',
            '{"id":532,"name":"tuiLKtPp","code":"LB-733","mail":"58wjhn7p@example.com",'
            '"uuid":"fff176ca-f4cd-4de4-ab0e-585334c85ca4","at":"2025-12-02T06:55:30Z","price":80.14,"tags":["c","b"],'
            '"flag":true,"nested":{"x":25}}',
            '{"id":740,"name":"hXUNCt5n","code":"IX-303","mail":"9ybqb3ro@example.com",'
            '"uuid":"31cc944b-bb76-47fe-ac6a-8b2cf8f5bff7","at":"2025-03-19T20:09:13Z","price":62.33,"tags":["a","c"],'
            '"flag":false,"nested":{"x":35}}'])


class CurlTests(unittest.TestCase):
    def test_multiple_commands(self):
        split = core.curl_commands
        bash = "curl 'https://a.test/x' \\\n  -H 'a: b' ;\ncurl 'https://a.test/y' -d 'k=v' ;\n"
        self.assertEqual(split(bash), [['curl', 'https://a.test/x', '-H', 'a: b'], ['curl', 'https://a.test/y', '-d', 'k=v']])
        cmd = 'curl "https://a.test/x" ^\r\n  -H "a: b" &\r\ncurl "https://a.test/y"\r\n'
        self.assertEqual(split(cmd), [['curl', 'https://a.test/x', '-H', 'a: b'], ['curl', 'https://a.test/y']])
        self.assertEqual(split('curl https://a.test/x\ncurl https://a.test/y'), [['curl', 'https://a.test/x'], ['curl', 'https://a.test/y']])
        self.assertEqual(split('curl https://a.test/x | jq . && curl -I https://a.test/y'),
                         [['curl', 'https://a.test/x'], ['jq', '.'], ['curl', '-I', 'https://a.test/y']])
        # 継続記号の無い折り返し（次の行が curl で始まらない）は 1 つのコマンドとして読む
        self.assertEqual(split('curl https://a.test/x\n  -H "a: b"'), [['curl', 'https://a.test/x', '-H', 'a: b']])

    def test_add_curl_appends_every_command(self):
        with workdir() as td:
            core.write_text(str(td / 'api.csv'), core.to_csv([['No', 'URL', 'メソッド'], ['1', 'https://a.test/0', 'GET']]), bom=True)
            core.write_text(str(td / 'c.txt'), 'curl https://a.test/1 ;\ncurl -X DELETE https://a.test/2\n')
            with quiet():
                self.assertEqual(core.cmd_add_curl({'input': str(td / 'api.csv'), 'curl_file': str(td / 'c.txt')}, []), 0)
            rows = core.parse_csv(core.read_text(str(td / 'api.csv'))[0])
            self.assertEqual([r[:3] for r in rows[1:]], [['1', 'https://a.test/0', 'GET'], ['2', 'https://a.test/1', 'GET'],
                                                         ['3', 'https://a.test/2', 'DELETE']])
            self.assertEqual(len(list(td.glob('api.csv.*.bak'))), 1)


class FileTests(unittest.TestCase):
    def test_reserved_windows_names(self):
        self.assertEqual(core.safe_file_name('CON.jmx', 'x.jmx'), '_CON.jmx')
        self.assertEqual(core.safe_file_name('lpt1', 'x.jmx'), '_lpt1')
        self.assertEqual(core.safe_file_name('console.jmx', 'x.jmx'), 'console.jmx')

    def test_backups_in_the_same_second_are_all_kept(self):
        with workdir() as td, quiet(), mock.patch.object(core, 'timestamp', return_value='20260101_000000'):
            path = str(td / 'api.jmx')
            for text in ('one', 'two', 'three'):
                core.write_jmx(path, text)
            self.assertEqual(sorted(p.name for p in td.iterdir()),
                             ['api.jmx', 'api.jmx.20260101_000000.bak', 'api.jmx.20260101_000000_2.bak'])
            self.assertEqual(sorted(p.read_text() for p in td.iterdir()), ['one', 'three', 'two'])

    def test_failed_write_keeps_the_current_file(self):
        with workdir() as td, quiet():
            path = td / 'api.jmx'
            path.write_text('original')
            with mock.patch.object(core, 'write_file', side_effect=core.ToolError('disk full')):
                with self.assertRaises(core.ToolError):
                    core.write_jmx(str(path), 'replacement')
            self.assertEqual([p.name for p in td.iterdir()], ['api.jmx'])
            self.assertEqual(path.read_text(), 'original')

    def test_jtl_summary_reads_csv(self):
        with workdir() as td, quiet():
            path = td / 'result.jtl'
            path.write_text('timeStamp,label,success,message\n1,"API,a",false,"line1\nline2"\n2,API b,true,ok\n',
                            encoding='utf-8')
            self.assertEqual(core.cmd_jtl_summary(str(path)), 3)
            path.write_text('timeStamp,label,success\n1,a,true\n', encoding='utf-8')
            self.assertEqual(core.cmd_jtl_summary(str(path)), 0)
            path.write_text('<testResults/>\n', encoding='utf-8')
            self.assertEqual(core.cmd_jtl_summary(str(path)), 0)
            self.assertTrue(core.WARNINGS)


if __name__ == '__main__':
    unittest.main()
