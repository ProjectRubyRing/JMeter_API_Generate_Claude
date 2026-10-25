"""JMeter 5.6.3 で実行し、ローカルの HTTP サーバが受け取ったリクエストを確認する（外部の API は呼び出さない）。

環境変数 JMETER_HOME に JMeter 5.6.3 のフォルダを設定すると実行します（未設定・Java が無い場合は skip）。
"""
import csv
import json
import os
import re
import shutil
import subprocess
import threading
import unittest
import zipfile
from email import policy
from email.parser import BytesParser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, urlsplit

from support import CORE as core, POWERSHELL, RUNTIMES, cli, option, workdir, write_csv, write_xlsx

JMETER = Path(os.environ.get('JMETER_HOME', '')) if os.environ.get('JMETER_HOME') else None
JAVA = shutil.which('java')
READY = bool(JAVA and JMETER and (JMETER / 'bin' / 'ApacheJMeter.jar').is_file())


class Receiver(BaseHTTPRequestHandler):
    def log_message(self, *args):
        pass

    def receive(self):
        length = int(self.headers.get('Content-Length') or 0)
        body = self.rfile.read(length) if length else b''
        self.server.requests.append((self.command, self.path, dict(self.headers), body))
        callback = getattr(self.server, 'after_request', None)
        if callback:
            callback(len(self.server.requests))
        status = 500 if urlsplit(self.path).path.endswith('/failure') else 200
        data = json.dumps({'ok': True, 'token': 'extracted-value'}).encode()
        self.send_response(status)
        self.send_header('Content-Type', 'application/json')
        self.send_header('Content-Length', str(len(data)))
        self.end_headers()
        if self.command != 'HEAD':
            self.wfile.write(data)

    do_GET = do_POST = do_PUT = do_PATCH = do_DELETE = do_OPTIONS = do_HEAD = do_TRACE = receive


@unittest.skipUnless(READY, 'JMETER_HOME（JMeter 5.6.3）と Java が必要です')
class JMeterTests(unittest.TestCase):
    def setUp(self):
        self.server = ThreadingHTTPServer(('127.0.0.1', 0), Receiver)
        self.server.requests = []
        self.base = 'http://127.0.0.1:%d' % self.server.server_port
        self.thread = threading.Thread(target=self.server.serve_forever, daemon=True)
        self.thread.start()

    def tearDown(self):
        self.server.shutdown()
        self.server.server_close()
        self.thread.join()

    def run_plan(self, jmx, out, *props):
        out.mkdir(parents=True, exist_ok=True)
        cmd = [JAVA, '-Xms128m', '-Xmx256m', '-jar', str(JMETER / 'bin' / 'ApacheJMeter.jar'), '-n', '-t', str(jmx),
               '-l', str(out / 'result.jtl'), '-j', str(out / 'jmeter.log'), '-JRESULT_DIR=' + str(out),
               '-Jjmeter.save.saveservice.subresults=false'] + list(props)
        result = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, timeout=180)
        log = (out / 'jmeter.log').read_text(encoding='utf-8', errors='replace')
        self.assertEqual(result.returncode, 0, result.stdout.decode('utf-8', 'replace') + log)
        with (out / 'result.jtl').open(encoding='utf-8-sig', newline='') as f:
            return list(csv.DictReader(f)), log

    def generate(self, rt, out, *args):
        rc, log = cli(rt, 'generate', '-o', out, *args, cwd=out.parent)
        self.assertEqual(rc, 0, '%s: %s' % (rt, log))

    # ------------------------------------------------------------------ リクエストの形
    def table_rows(self):
        return [['No', 'URL', 'メソッド', 'クエリパラメータ', 'ヘッダ', 'ボディ形式', 'リクエストボディ', '期待ステータス'],
                ['1', '/query', 'GET', 'a%26b=\nq=x%20y 日本語\nrepeat=1\nrepeat=2', '', '', '', '200'],
                ['2', '/plus?p=a+b&r=c%2Bd', 'GET', '', '', '', '', '200'],
                ['3', '/form?q=a%20b', 'POST', '', '', 'form', 'a%26b=x%26y%25\nplain=1+2', '200'],
                ['4', '/multipart', 'POST', '', '', 'multipart', 'a%26b=%20x%26y%25%E6%97%A5%E6%9C%AC%E8%AA%9E%20', '200'],
                ['5', '/raw', 'PUT', '', 'Content-Type: text/plain', 'raw', '  text\n ', '200'],
                ['6', '/json', 'PATCH', '', '', 'json', '{"value": "日本語"}', '200'],
                ['7', '/head', 'HEAD', 'q=value', '', '', '', '200'],
                ['8', '/delete', 'DELETE', 'q=value', '', '', '', '200'],
                ['9', '/options', 'OPTIONS', 'q=value', '', '', '', '200'],
                ['10', '/trace', 'TRACE', 'q=value', '', '', '', '200']]

    def check_table_requests(self, requests):
        self.assertEqual(len(requests), 10, requests)
        for method, target, headers, body in requests:
            path = urlsplit(target).path
            query = parse_qs(urlsplit(target).query, keep_blank_values=True)
            with self.subTest(method=method, path=path):
                if path == '/query':
                    self.assertEqual(query, {'a&b': [''], 'q': ['x y 日本語'], 'repeat': ['1', '2']})
                elif path == '/plus':
                    self.assertEqual(query, {'p': ['a b'], 'r': ['c+d']})
                elif path == '/form':
                    self.assertEqual(query, {'q': ['a b']})
                    self.assertEqual(parse_qs(body.decode(), keep_blank_values=True), {'a&b': ['x&y%'], 'plain': ['1+2']})
                elif path == '/multipart':
                    msg = BytesParser(policy=policy.default).parsebytes(
                        ('Content-Type: ' + headers['Content-Type'] + '\r\n\r\n').encode() + body)
                    parts = list(msg.iter_parts())
                    self.assertEqual([p.get_param('name', header='content-disposition') for p in parts], ['a&b'])
                    self.assertEqual(parts[0].get_payload(decode=True).decode('utf-8'), ' x&y%日本語 ')
                elif path == '/raw':
                    self.assertEqual(body, b'  text\n ')
                elif path == '/json':
                    self.assertEqual(json.loads(body.decode('utf-8')), {'value': '日本語'})
                else:
                    # HEAD / DELETE / OPTIONS / TRACE: クエリは URL だけに付き、ボディは送らない
                    self.assertEqual(query, {'q': ['value']})
                    self.assertEqual(body, b'')

    def test_table_requests_on_all_runtimes(self):
        with workdir() as td:
            write_xlsx(td / 'api.xlsx', [('API一覧', self.table_rows())])
            write_csv(td / 'api.csv', self.table_rows())
            write_csv(td / 'settings.csv', [['設定キー', '設定値'], ['base_url', self.base], ['csv_data', 'OFF'],
                                            ['excel_data', 'OFF']])
            for rt in RUNTIMES:
                for ext in ('xlsx', 'csv'):
                    with self.subTest(runtime=rt, input=ext):
                        self.server.requests.clear()
                        out = td / (rt + ext)
                        out.mkdir()
                        self.generate(rt, out, '-i', td / ('api.' + ext), '-s', td / 'settings.csv')
                        results, log = self.run_plan(out / 'api_test.jmx', out / 'run')
                        self.assertTrue(all(r['success'] == 'true' for r in results), results)
                        self.check_table_requests(self.server.requests)

    def openapi_spec(self):
        return {
            'openapi': '3.0.3', 'info': {'title': 'wire', 'version': '1'}, 'servers': [{'url': self.base + '/v1'}],
            'paths': {
                '/styles/{label}/{matrix}': {'get': {'parameters': [
                    {'name': 'label', 'in': 'path', 'required': True, 'style': 'label', 'explode': True,
                     'schema': {'type': 'array', 'items': {'type': 'string'}}, 'example': ['a', 'b']},
                    {'name': 'matrix', 'in': 'path', 'required': True, 'style': 'matrix',
                     'schema': {'type': 'object'}, 'example': {'k': 'v w'}}],
                    'responses': {'200': {'description': 'ok'}}}},
                '/query': {'get': {'parameters': [
                    {'name': 'csv', 'in': 'query', 'explode': False, 'schema': {'type': 'array', 'items': {'type': 'string'}},
                     'example': ['a,b', 'c']},
                    {'name': 'a&b', 'in': 'query', 'schema': {'type': 'string'}, 'example': ''},
                    {'name': 'json', 'in': 'query', 'content': {'application/json': {'example': {'k': [1, 2]}}}}],
                    'responses': {'200': {'description': 'ok'}}}},
                '/upload': {'post': {'requestBody': {'required': True, 'content': {'multipart/form-data': {
                    'schema': {'type': 'object'}, 'example': {'title': ' a&b 日本 ', 'memo': '50%'}}}},
                    'responses': {'200': {'description': 'ok'}}}},
                '/items': {'delete': {'parameters': [{'name': 'ids', 'in': 'query', 'explode': False,
                                                      'schema': {'type': 'array', 'items': {'type': 'integer'}},
                                                      'example': [1, 2]}],
                                      'responses': {'200': {'description': 'ok'}}}},
                '/other': {'servers': [{'url': self.base + '/alt'}],
                           'get': {'responses': {'200': {'description': 'ok'}}}},
            }}

    def test_openapi_requests_on_all_runtimes(self):
        with workdir() as td:
            (td / 'api.json').write_text(json.dumps(self.openapi_spec(), ensure_ascii=False), encoding='utf-8')
            write_csv(td / 'settings.csv', [['設定キー', '設定値'], ['csv_data', 'OFF'], ['excel_data', 'OFF']])
            write_csv(td / 'oa.csv', [['設定キー', '設定値'], ['API_HEADERS', 'なし']])
            for rt in RUNTIMES:
                with self.subTest(runtime=rt):
                    self.server.requests.clear()
                    out = td / rt
                    out.mkdir()
                    self.generate(rt, out, '-i', td / 'api.json', '-s', td / 'settings.csv', '-c', td / 'oa.csv',
                                  option(rt, 'use-examples'), option(rt, 'seed'), '1')
                    results, log = self.run_plan(out / 'wire.jmx', out / 'run')
                    self.assertTrue(all(r['success'] == 'true' for r in results), results)
                    got = dict((urlsplit(t).path, (m, t, h, b)) for m, t, h, b in self.server.requests)
                    self.assertEqual(sorted(got), ['/alt/other', '/v1/items', '/v1/query', '/v1/styles/.a.b/;matrix=k,v%20w',
                                                   '/v1/upload'])
                    raw_query = urlsplit(got['/v1/query'][1]).query
                    self.assertIn('csv=a%2Cb,c', raw_query)      # 値の中のカンマと区切りのカンマを区別できる
                    self.assertEqual(parse_qs(raw_query, keep_blank_values=True)['a&b'], [''])
                    self.assertEqual(json.loads(parse_qs(raw_query)['json'][0]), {'k': [1, 2]})
                    method, target, headers, body = got['/v1/items']
                    self.assertEqual((method, parse_qs(urlsplit(target).query), body), ('DELETE', {'ids': ['1,2']}, b''))
                    method, target, headers, body = got['/v1/upload']
                    msg = BytesParser(policy=policy.default).parsebytes(
                        ('Content-Type: ' + headers['Content-Type'] + '\r\n\r\n').encode() + body)
                    values = dict((p.get_param('name', header='content-disposition'), p.get_payload(decode=True).decode('utf-8'))
                                  for p in msg.iter_parts())
                    self.assertEqual(values, {'title': ' a&b 日本 ', 'memo': '50%'})

    # ------------------------------------------------------------------ 実行時の Groovy
    def test_excel_csv_variables_and_token_rotation(self):
        with workdir() as td:
            tokens = td / 'tokens'
            tokens.mkdir()
            token_file = tokens / 'access_token.txt'
            token_file.write_text(' first-token\n', encoding='utf-8-sig')

            def rotate(count):
                if count == 1:
                    token_file.write_text('second-token', encoding='utf-8')
            self.server.after_request = rotate
            write_xlsx(td / 'api.xlsx', [('API一覧', [
                ['URL', 'メソッド', 'クエリパラメータ', 'ヘッダ', '期待ステータス', 'JSON抽出'],
                [self.base + '/vars', 'GET', 'user=${USER_ID}\nitem=${ITEM_CODE}\nseq=${SEQ}', 'X-Previous: ${PREVIOUS}', '200',
                 'PREVIOUS=$.token'],
                [self.base + '/vars', 'GET', 'user=${USER_ID}\nitem=${ITEM_CODE}\nseq=${SEQ}', 'X-Previous: ${PREVIOUS}', '200', '']])])
            # Excel データは ~/ と \ の区切りで指定（JMX のフォルダ基準）
            write_csv(td / 'settings.csv', [['設定キー', '設定値'], ['loops', '2'], ['access_token_dir', str(tokens)],
                                            ['excel_file', '~/data\\testdata.xlsx']])
            out = td / 'out'
            out.mkdir()
            self.generate('python', out, '-i', td / 'api.xlsx', '-s', td / 'settings.csv')
            core.build_testdata_xlsx(str(out / 'data' / 'testdata.xlsx'))
            results, log = self.run_plan(out / 'api_test.jmx', td / 'run')
            self.assertEqual(len(results), 4, results)
            requests = self.server.requests
            queries = [parse_qs(urlsplit(r[1]).query) for r in requests]
            self.assertEqual(queries[0]['user'], [core.SPEC['sample_csv']['rows'][0][0]])
            self.assertEqual(queries[2]['user'], [core.SPEC['sample_csv']['rows'][1][0]])
            self.assertEqual(queries[0]['item'], queries[1]['item'])
            self.assertNotEqual(queries[0]['item'], queries[2]['item'])
            self.assertEqual(requests[0][2]['Authorization'], 'Bearer first-token')
            for request in requests[1:]:
                self.assertEqual(request[2]['Authorization'], 'Bearer second-token')
                self.assertEqual(request[2]['X-Previous'], 'extracted-value')

    def test_excel_load_failure_stops_before_any_request(self):
        with workdir() as td:
            write_csv(td / 'api.csv', [['URL', 'メソッド'], [self.base + '/must-not-run', 'GET']])
            write_xlsx(td / 'header_only.xlsx', [('data', [['ITEM_CODE']])])
            write_xlsx(td / 'ok.xlsx', [('data', [['VALUE'], ['v']])])
            # 計算結果の無い数式（A2）
            with zipfile.ZipFile(str(td / 'ok.xlsx')) as z:
                entries = [(i.filename, z.read(i.filename)) for i in z.infolist()]
            with zipfile.ZipFile(str(td / 'formula.xlsx'), 'w') as z:
                for name, data in entries:
                    if name == 'xl/worksheets/sheet1.xml':
                        data = re.sub(rb'<c r="A2"([^>]*?) t="s"([^>]*)><v>\d+</v></c>', rb'<c r="A2"\1\2><f>1+1</f></c>', data)
                    z.writestr(name, data)
            for name in ('missing.xlsx', 'header_only.xlsx', 'formula.xlsx'):
                with self.subTest(excel=name):
                    self.server.requests.clear()
                    write_csv(td / 'settings.csv', [['設定キー', '設定値'], ['csv_data', 'OFF'], ['excel_file', str(td / name)]])
                    out = td / name.replace('.', '_')
                    out.mkdir()
                    self.generate('python', out, '-i', td / 'api.csv', '-s', td / 'settings.csv')
                    results, log = self.run_plan(out / 'api_test.jmx', out / 'run')
                    self.assertEqual(self.server.requests, [])
                    self.assertTrue(any(r['success'] == 'false' and 'Excel' in r['label'] for r in results), results)
            # ファイルパスを空にすると読み込まずに実行する
            self.server.requests.clear()
            out = td / 'empty'
            out.mkdir()
            write_csv(td / 'settings.csv', [['設定キー', '設定値'], ['csv_data', 'OFF']])
            self.generate('python', out, '-i', td / 'api.csv', '-s', td / 'settings.csv')
            results, log = self.run_plan(out / 'api_test.jmx', out / 'run', '-JEXCEL_FILE=')
            self.assertEqual(len(self.server.requests), 1)
            self.assertTrue(all(r['success'] == 'true' for r in results), results)

    @unittest.skipUnless(os.name == 'nt', 'ファイルのロックは Windows で確認する')
    def test_token_read_failure_keeps_the_previous_token(self):
        import msvcrt
        with workdir() as td:
            tokens = td / 'tokens'
            tokens.mkdir()
            token_file = tokens / 'access_token.txt'
            token_file.write_text('first-token', encoding='utf-8')
            state = {}

            def lock(count):
                if count == 1:
                    token_file.write_text('second-token', encoding='utf-8')
                    state['f'] = open(str(token_file), 'r+b')
                    msvcrt.locking(state['f'].fileno(), msvcrt.LK_NBLCK, 1)
                elif count == 2:
                    msvcrt.locking(state['f'].fileno(), msvcrt.LK_UNLCK, 1)
                    state['f'].close()
            self.server.after_request = lock
            write_csv(td / 'api.csv', [['URL', 'メソッド'], [self.base + '/a', 'GET'], [self.base + '/b', 'GET'],
                                       [self.base + '/c', 'GET']])
            write_csv(td / 'settings.csv', [['設定キー', '設定値'], ['csv_data', 'OFF'], ['excel_data', 'OFF'],
                                            ['access_token_dir', str(tokens)]])
            out = td / 'out'
            out.mkdir()
            self.generate('python', out, '-i', td / 'api.csv', '-s', td / 'settings.csv')
            results, log = self.run_plan(out / 'api_test.jmx', out / 'run')
            self.assertEqual([r[2]['Authorization'] for r in self.server.requests],
                             ['Bearer first-token', 'Bearer first-token', 'Bearer second-token'])
            # jmeter.log の文字コードは Java の既定（日本語の Windows は MS932）のため、例外のクラス名で確認する
            self.assertIn('java.io.IOException', log)

    # ------------------------------------------------------------------ run コマンド
    @unittest.skipUnless((os.name == 'nt' and POWERSHELL) or (os.name != 'nt' and 'bash' in RUNTIMES), 'run コマンドを実行できません')
    def test_run_reports_failure_and_creates_report(self):
        with workdir() as td:
            rt = 'powershell' if os.name == 'nt' else 'bash'
            write_csv(td / 'api.csv', [['API名', 'URL', 'メソッド', '期待ステータス'],
                                       ['failure,quoted', self.base + '/failure', 'GET', '200']])
            out = td / 'out'
            out.mkdir()
            self.generate(rt, out, '-i', td / 'api.csv')
            rc, log = cli(rt, 'run', '-t', out / 'api_test.jmx', '-o', td / 'result', option(rt, 'jmeter-home'), JMETER,
                          option(rt, 'heap'), '256m', option(rt, 'fail-on-error'), cwd=td)
            self.assertEqual(rc, 2, log)
            self.assertEqual(len(list((td / 'result').glob('*/report/index.html'))), 1, log)


if __name__ == '__main__':
    unittest.main()
