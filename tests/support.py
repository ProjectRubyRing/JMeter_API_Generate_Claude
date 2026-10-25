"""テスト共通: シェルに埋め込んだ Python コアの読み込みと、各実装（Python / PowerShell / bash）の実行。

環境変数:
  TEST_POWERSHELL  PowerShell の実行ファイル（既定: powershell → pwsh の順に探す）
  TEST_RUNTIMES    比較する実装（カンマ区切り。既定: python,powershell,bash のうち使えるもの）
  JMETER_HOME      JMeter 5.6.3 のフォルダ（test_jmeter.py の実通信テストで使用。未設定なら skip）
"""
import ast
import contextlib
import io
import os
import re
import shutil
import subprocess
import sys
import tempfile
import types
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SH = ROOT / 'jmeter_api_tool.sh'
PS1 = ROOT / 'jmeter_api_tool.ps1'
SOURCE = SH.read_text(encoding='utf-8').split("<<'__PYTHON_CORE__'\n", 1)[1].split('\n__PYTHON_CORE__', 1)[0]


def load_core():
    tree = ast.parse(SOURCE)
    # テストの出力先（sys.stdout / sys.stderr）を差し替える行は読み込まない
    tree.body = [node for node in tree.body if not (
        isinstance(node, ast.Assign) and any(
            isinstance(t, ast.Attribute) and isinstance(t.value, ast.Name)
            and t.value.id == 'sys' and t.attr in ('stdout', 'stderr') for t in node.targets))]
    module = types.ModuleType('jmeter_core')
    exec(compile(tree, str(SH) + ' (Python コア)', 'exec'), module.__dict__)
    return module


CORE = load_core()
POWERSHELL = os.environ.get('TEST_POWERSHELL') or shutil.which('powershell') or shutil.which('pwsh')
BASH = shutil.which('bash')
_CORE_FILE = Path(tempfile.gettempdir()) / ('jmeter_api_tool_core_%d.py' % os.getpid())


def runtimes():
    available = ['python'] + (['powershell'] if POWERSHELL else []) + (['bash'] if BASH else [])
    wanted = [x.strip() for x in os.environ.get('TEST_RUNTIMES', '').split(',') if x.strip()]
    return [x for x in available if not wanted or x in wanted]


RUNTIMES = runtimes()


def _decode(data):
    # Windows PowerShell 5.1 はパイプへの出力に OEM コードページ（日本語環境は cp932）を使う
    try:
        return data.decode('utf-8')
    except UnicodeDecodeError:
        return data.decode('oem' if os.name == 'nt' else 'utf-8', errors='replace')


def option(runtime, name):
    """sh / Python と PowerShell で名前が違うオプション（例: option(rt, 'seed') → --seed / -Seed）。"""
    ps = {'seed': '-Seed', 'use-examples': '-UseExamples', 'required-only': '-RequiredOnly', 'split': '-Split',
          'dry-run': '-DryRun', 'update': '-UpdateJmx', 'curl-file': '-CurlFile', 'curl': '-Curl', 'name': '-ApiName',
          'jmeter-home': '-JMeterHome', 'heap': '-Heap', 'fail-on-error': '-FailOnError', 'any-method': '-AnyMethod'}
    sh = {'update': '-u'}
    if runtime == 'powershell':
        return ps[name]
    return sh.get(name, '--' + name)


def cli(runtime, *args, cwd=None, timeout=300):
    """ツールを実行し (終了コード, 出力) を返す。"""
    if runtime == 'powershell':
        command = [POWERSHELL, '-NoLogo', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', str(PS1)]
    elif runtime == 'bash':
        command = [BASH, str(SH)]
    else:
        if not _CORE_FILE.exists():
            _CORE_FILE.write_text(SOURCE, encoding='utf-8')
        command = [sys.executable, '-X', 'utf8', str(_CORE_FILE)]
    result = subprocess.run(command + [str(a) for a in args], cwd=str(cwd or ROOT), stdout=subprocess.PIPE,
                            stderr=subprocess.STDOUT, timeout=timeout)
    return result.returncode, _decode(result.stdout)


def record(**values):
    """normalize_api に渡す API一覧の 1 行。"""
    rec = dict((c['key'], '') for c in CORE.SPEC['api_columns'])
    rec.update(_row=2, no='1', enabled='Y', method='POST', url='https://example.test/api')
    rec.update(values)
    return rec


def config(**values):
    with quiet():
        return CORE.resolve_settings(dict((k, str(v)) for k, v in values.items()))


def context(doc=None, **values):
    opts = dict(required_only=False, use_examples=True, any_method='GET')
    opts.update(values)
    ctx = CORE.Ctx(doc or {}, 12345, opts)
    ctx.oa_conf = None
    ctx.doc_server = CORE.oa_server_url(doc or {}, None)
    return ctx


@contextlib.contextmanager
def quiet():
    """Python コアの [INFO] / [WARN] の出力を捨てる（WARNINGS はテストごとに空にする）。"""
    del CORE.WARNINGS[:]
    with contextlib.redirect_stdout(io.StringIO()):
        yield


@contextlib.contextmanager
def workdir():
    path = Path(tempfile.mkdtemp(prefix='jmeter_api_tool_test_'))
    try:
        yield path
    finally:
        shutil.rmtree(str(path), ignore_errors=True)


def prop(node, name):
    return node.findtext("./*[@name='%s']" % name, '')


def samplers(path):
    return list(ET.parse(str(path)).getroot().iter('HTTPSamplerProxy'))


def arguments(sampler):
    return [(prop(a, 'Argument.name'), prop(a, 'Argument.value'), prop(a, 'HTTPArgument.always_encode'))
            for a in sampler.findall(".//collectionProp[@name='Arguments.arguments']/elementProp")]


def normalized(text):
    """生成日時・乱数シード（未指定時）など、実行ごとに変わる部分をそろえる。"""
    text = text.replace('\r\n', '\n')
    text = re.sub(r'\d{4}-\d{2}-\d{2}([ T]\d{2}:\d{2}:\d{2}Z?)?', '<date>', text)
    return text


def comparable_files(folder):
    """フォルダ内のファイル → 比較用の内容（xlsx は各 XML、.bak は除く）。"""
    out = {}
    for path in sorted(Path(folder).rglob('*')):
        if not path.is_file() or path.suffix == '.bak':
            continue
        key = path.relative_to(folder).as_posix()
        if path.suffix == '.xlsx':
            with zipfile.ZipFile(str(path)) as z:
                out[key] = [(n, normalized(z.read(n).decode('utf-8', 'replace'))) for n in sorted(z.namelist())]
        else:
            out[key] = normalized(path.read_bytes().decode('utf-8', 'replace'))
    return out


def write_csv(path, rows, bom=True):
    CORE.write_text(str(path), CORE.to_csv(rows), bom=bom)


def write_xlsx(path, sheets):
    """[(シート名, 行の一覧)] から xlsx を作る（ツールの xlsx 出力を使用）。"""
    objs = []
    for name, rows in sheets:
        sheet = CORE.Sheet(name)
        for r, cells in enumerate(rows, 1):
            for c, text in enumerate(cells, 1):
                sheet.put(r, c, text, 'input')
        objs.append(sheet)
    CORE.save_book(objs, str(path), 'Test')


def rewrite_zip(src, dst, change):
    """xlsx（ZIP）の各エントリを change(名前, データ) → データ で書き換えてコピーする。"""
    with zipfile.ZipFile(str(src)) as z:
        entries = [(i, z.read(i.filename)) for i in z.infolist()]
    with zipfile.ZipFile(str(dst), 'w', zipfile.ZIP_DEFLATED) as z:
        for info, data in entries:
            z.writestr(info.filename, change(info.filename, data))
