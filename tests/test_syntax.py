"""構文・文字コード・2 実装の共通データ（SPEC）の検査。"""
import ast
import json
import os
import re
import subprocess
import unittest

from support import BASH, POWERSHELL, PS1, ROOT, SH, SOURCE


class SyntaxTests(unittest.TestCase):
    def test_python_core_supports_python36(self):
        # RHEL の OS 標準の python3（3.6 以上）で動く構文であること
        ast.parse(SOURCE, feature_version=(3, 6))

    @unittest.skipUnless(BASH, 'bash がありません')
    def test_bash_syntax(self):
        subprocess.run([BASH, '-n', str(SH)], check=True)

    @unittest.skipUnless(POWERSHELL, 'PowerShell がありません')
    def test_powershell_syntax(self):
        script = ("$t = $null; $e = $null; [void][System.Management.Automation.Language.Parser]::ParseFile('%s', [ref]$t, [ref]$e); "
                  "if ($e.Count) { $e | Out-String -Width 200; exit 1 }" % str(PS1).replace("'", "''"))
        result = subprocess.run([POWERSHELL, '-NoProfile', '-Command', script], stdout=subprocess.PIPE, stderr=subprocess.STDOUT)
        self.assertEqual(result.returncode, 0, result.stdout.decode('utf-8', 'replace'))

    @unittest.skipUnless(os.name == 'nt', 'Windows のみ')
    def test_batch_launcher(self):
        result = subprocess.run(['cmd.exe', '/d', '/c', str(ROOT / 'jmeter_api_tool.bat'), 'version'],
                                stdout=subprocess.PIPE, stderr=subprocess.STDOUT, timeout=120)
        self.assertEqual(result.returncode, 0, result.stdout)


class SourceFileTests(unittest.TestCase):
    def test_encodings_and_line_endings(self):
        ps = PS1.read_bytes()
        self.assertTrue(ps.startswith(b'\xef\xbb\xbf'), '.ps1 は BOM 付き UTF-8（Windows PowerShell 5.1 の文字化け防止）')
        self.assertEqual(ps.count(b'\r\n'), ps.count(b'\n'), '.ps1 の改行は CRLF')
        sh = SH.read_bytes()
        self.assertFalse(sh.startswith(b'\xef\xbb\xbf'))
        self.assertNotIn(b'\r', sh, '.sh の改行は LF（RHEL で実行するため）')

    def test_spec_is_identical_in_both_scripts(self):
        sh = re.search(r"^SPEC = json\.loads\(r'''(.*?)'''\)", SH.read_text(encoding='utf-8'), re.S | re.M).group(1)
        ps = re.search(r"^\$script:SpecJson = @'\n(.*?)\n'@", PS1.read_text(encoding='utf-8-sig').replace('\r\n', '\n'),
                       re.S | re.M).group(1)
        self.assertEqual(sh, ps, '設定項目・テンプレート・Groovy・利用方法は両スクリプトで同じ内容にする')
        self.assertEqual(json.dumps(json.loads(sh), ensure_ascii=False, indent=1), sh)

    def test_versions_agree(self):
        spec = json.loads(re.search(r"^SPEC = json\.loads\(r'''(.*?)'''\)", SH.read_text(encoding='utf-8'), re.S | re.M).group(1))
        version = spec['tool']['version']
        self.assertIn('TOOL_VERSION="%s"' % version, SH.read_text(encoding='utf-8'))
        self.assertIn("$script:ToolVersion = '%s'" % version, PS1.read_text(encoding='utf-8-sig'))
        self.assertIn('v%s' % version, SH.read_text(encoding='utf-8').split('\n', 3)[2])


if __name__ == '__main__':
    unittest.main()
