"""Python / PowerShell / bash の各実装が、同じ入力から同じファイルを出力することの確認（CLI 経由）。"""
import re
import shutil
import unittest
import xml.etree.ElementTree as ET
import zipfile

from support import (CORE as core, ROOT, RUNTIMES, cli, comparable_files, option, prop, samplers, workdir, write_csv,
                     write_xlsx)

FIXTURES = ROOT / 'tests' / 'fixtures'


class ParityCase(unittest.TestCase):
    def run_all(self, td, make_args, expect_rc=0, prepare=None):
        """各実装で実行し、出力フォルダの内容を返す（実装間で一致することを確認）。"""
        results = {}
        for rt in RUNTIMES:
            out = td / rt
            out.mkdir()
            if prepare:
                prepare(out)
            rc, log = cli(rt, *make_args(rt, out), cwd=out)
            self.assertEqual(rc, expect_rc, '%s: %s' % (rt, log))
            results[rt] = comparable_files(out)
        first = RUNTIMES[0]
        for rt in RUNTIMES[1:]:
            self.assertEqual(sorted(results[rt]), sorted(results[first]), rt)
            for name in results[first]:
                self.assertEqual(results[rt][name], results[first][name], '%s と %s で %s が違います' % (first, rt, name))
        return results[first]


class SampleTests(ParityCase):
    def test_checked_in_samples_are_current(self):
        # output/ のサンプルは input/ のテンプレートから生成したもの（xlsx / csv のどちらからでも同じ）
        checked_in = comparable_files(ROOT / 'output')
        with workdir() as td:
            for src in ('api_definition.xlsx', 'api_definition.csv'):
                folder = td / src.replace('.', '_')
                folder.mkdir()
                files = self.run_all(folder, lambda rt, out: ['generate', '-i', ROOT / 'input' / src, '-o', out])
                for name in ('api_test.jmx', 'api_test_linux_cmd.txt'):
                    # テストプランのコメントには入力ファイル名が入る（xlsx / csv の違い以外は同じ）
                    got = files[name].replace('入力: ' + src, '入力: api_definition.xlsx')
                    self.assertEqual(got, checked_in[name], src + ' → ' + name)

    def test_templates_and_manual_are_current(self):
        with workdir() as td:
            files = self.run_all(td, lambda rt, out: ['init', '-o', out / 'input'])
            checked_in = comparable_files(ROOT / 'input')
            for name in ('api_definition.xlsx', 'api_definition.csv', 'api_definition_settings.csv', 'openapi_config.csv'):
                self.assertEqual(files['input/' + name], checked_in[name], name)
        with workdir() as td:
            files = self.run_all(td, lambda rt, out: ['manual', '-o', out / 'docs'])
            checked_in = comparable_files(ROOT / 'docs')
            for name in ('JMeter_Usage_Guide.md', 'JMeter_Usage_Guide.xlsx'):
                self.assertEqual(files['docs/' + name], checked_in[name], name)


class GenerateTests(ParityCase):
    def test_openapi_features(self):
        spec = FIXTURES / 'openapi_features.json'
        with workdir() as td:
            files = self.run_all(td, lambda rt, out: ['generate', '-i', spec, '-o', out, option(rt, 'seed'), '42',
                                                      option(rt, 'use-examples')])
        self.assertIn('features-api.jmx', files)
        with workdir() as td:
            self.run_all(td, lambda rt, out: ['generate', '-i', spec, '-o', out, option(rt, 'seed'), '7',
                                              option(rt, 'required-only')])

    def test_servers_with_settings_and_connection_overrides(self):
        spec = FIXTURES / 'openapi_features.json'
        with workdir() as td:
            write_csv(td / 'settings.csv', [['設定キー', '設定値'], ['base_url', 'https://override.test/base']])
            write_csv(td / 'oa.csv', [['設定キー', '設定値'], ['HOST', 'configured.test']])
            for mode, extra, hosts in (
                    ('definition', [], {'/v1/query': 'localhost', '/up/upload': 'files.example.com', '/x/xml': 'xml.example.com'}),
                    ('settings', ['-s', td / 'settings.csv'], {'/base/query': 'override.test', '/base/upload': 'override.test'}),
                    ('connection', ['-c', td / 'oa.csv'], {'/v1/query': 'configured.test', '/up/upload': 'configured.test'})):
                (td / mode).mkdir()
                self.run_all(td / mode, lambda rt, out: ['generate', '-i', spec, '-o', out, option(rt, 'seed'), '1'] + extra)
                jmx = td / mode / RUNTIMES[0] / 'features-api.jmx'
                root = ET.parse(str(jmx)).getroot()
                udv = dict((prop(a, 'Argument.name'), prop(a, 'Argument.value'))
                           for a in root.findall('.//Arguments/collectionProp/elementProp'))
                found = {}
                for s in samplers(jmx):
                    var = prop(s, 'HTTPSampler.domain')[2:-1]               # ${HOST_2} → HOST_2
                    host = udv[var].split(',', 1)[1][:-2]                    # ${__P(HOST_2,files.example.com)} → files.example.com
                    found[prop(s, 'HTTPSampler.path').split('?')[0]] = host
                for path, host in hosts.items():
                    with self.subTest(mode=mode, path=path):
                        self.assertEqual(found.get(path), host)

    def test_split_and_reserved_names(self):
        with workdir() as td:
            write_csv(td / 'api.csv', [['No', 'URL', 'メソッド'], ['1', 'https://example.test/a', 'GET'],
                                       ['1', 'https://example.test/b', 'GET']])
            (td / 'out').mkdir()
            files = self.run_all(td / 'out', lambda rt, out: ['generate', '-i', td / 'api.csv', '-o', out, '-n', 'CON.jmx',
                                                              option(rt, 'split')])
            self.assertEqual(sorted(n for n in files if n.endswith('.jmx')), ['_CON_API01.jmx', '_CON_API01_2.jmx'])

    def test_invalid_inputs_fail_without_jmx(self):
        with workdir() as td:
            (td / 'bad.csv').write_text('URL,メソッド,リクエストボディ\nhttps://example.test,POST,"open\n', encoding='utf-8')
            (td / 'bad.json').write_text('{"openapi": "3.0.3", "paths": ', encoding='utf-8')
            write_xlsx(td / 'f.xlsx', [('API一覧', [['URL', 'メソッド'], ['https://example.test/a', 'GET']])])
            with zipfile.ZipFile(str(td / 'f.xlsx')) as z:
                entries = [(i.filename, z.read(i.filename)) for i in z.infolist()]
            with zipfile.ZipFile(str(td / 'formula.xlsx'), 'w') as z:
                for name, data in entries:
                    if name == 'xl/worksheets/sheet1.xml':
                        data = re.sub(rb'<c r="A2"([^>]*?) t="s"([^>]*)><v>\d+</v></c>', rb'<c r="A2"\1\2><f>1+1</f></c>', data)
                    z.writestr(name, data)
            for name in ('bad.csv', 'bad.json', 'formula.xlsx'):
                (td / name.replace('.', '_')).mkdir()
                files = self.run_all(td / name.replace('.', '_'), lambda rt, out: ['generate', '-i', td / name, '-o', out],
                                     expect_rc=1)
                self.assertFalse([n for n in files if n.endswith('.jmx')], name)


class AddCurlTests(ParityCase):
    def test_multiple_commands_into_csv_and_xlsx(self):
        with workdir() as td:
            (td / 'curl.txt').write_text(
                "curl 'https://api.example.test/a?x=1&y=a+b' -H 'accept: application/json' ;\n"
                "curl 'https://api.example.test/b' -H 'content-type: application/json' --data-raw '{\"k\":\"v\"}' ;\n"
                "curl -u user:pass -F 'title=日本 語' -F 'memo=50%' https://api.example.test/c | jq .\n", encoding='utf-8')

            def prepare(out):
                shutil.copy(str(ROOT / 'input' / 'api_definition.csv'), str(out))
                shutil.copy(str(ROOT / 'input' / 'api_definition.xlsx'), str(out))
            for ext in ('csv', 'xlsx'):
                (td / ext).mkdir()
                self.run_all(td / ext, lambda rt, out: ['add-curl', '-i', out / ('api_definition.' + ext),
                                                        option(rt, 'curl-file'), td / 'curl.txt'], prepare=prepare)
                for rt in RUNTIMES:
                    # 元のファイルは 1 回の実行で 1 つだけ退避する（.bak は名前に日時を含むため比較の対象外）
                    self.assertEqual(len(list((td / ext / rt).glob('api_definition.%s.*.bak' % ext))), 1, rt)
            rows = core.parse_csv(core.read_text(str(td / 'csv' / RUNTIMES[0] / 'api_definition.csv'))[0])
            self.assertEqual([r[4] for r in rows[-3:]], ['https://api.example.test/a', 'https://api.example.test/b',
                                                          'https://api.example.test/c'])
            self.assertEqual([r[0] for r in rows[-3:]], ['5', '6', '7'])


class UpdateTests(ParityCase):
    def test_update_keeps_gui_changes_and_applies_definition_changes(self):
        with workdir() as td:
            rows = [['No', 'API名', 'URL', 'メソッド', 'クエリパラメータ'],
                    ['1', 'search', 'https://example.test/search', 'GET', 'q=a'],
                    ['2', 'remove', 'https://example.test/item', 'DELETE', 'id=1']]
            write_csv(td / 'v1.csv', rows)
            write_csv(td / 'v2.csv', rows[:2] + [['2', 'remove', 'https://example.test/item', 'DELETE', 'id=2'],
                                                 ['3', 'add', 'https://example.test/add', 'POST', '']])

            def prepare(out):
                rc, log = cli('python', 'generate', '-i', td / 'v1.csv', '-o', out, cwd=out)
                self.assertEqual(rc, 0, log)
                jmx = out / 'api_test.jmx'
                # JMeter GUI での変更の代わり: API01 のパスを書き換える
                jmx.write_text(jmx.read_text(encoding='utf-8').replace('>/search<', '>/search-gui<'), encoding='utf-8')
            files = self.run_all(td, lambda rt, out: ['generate', '-i', td / 'v2.csv', option(rt, 'update'), out / 'api_test.jmx'],
                                 prepare=prepare)
            jmx = td / RUNTIMES[0] / 'api_test.jmx'
            paths = [prop(s, 'HTTPSampler.path') for s in samplers(jmx)]
            self.assertEqual(paths, ['/search-gui', '/item?id=2', '/add'])
            self.assertIn('api_test.jmx.base.json', files)


if __name__ == '__main__':
    unittest.main()
