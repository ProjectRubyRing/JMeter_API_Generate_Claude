# JMeter API テスト JMX 生成ツール（jmeter_api_tool）

REST API の定義（URL・HTTP メソッド・クエリパラメータ・ヘッダ・リクエストボディ）を **Excel または CSV** に記入するか、**OpenAPI 定義（JSON）** を指定すると、
**Apache JMeter 5.6.3** 用の JMX ファイルを自動生成し、CLI モードで実行して JTL・HTML レポート・ログを出力するツールです。

- RHEL 9.8 用: `jmeter_api_tool.sh`（bash + OS 標準の python3。追加インストール不要）
- Windows 11 用: `jmeter_api_tool.ps1` ＋ 起動用 `jmeter_api_tool.bat`（Windows PowerShell 5.1 / PowerShell 7。追加モジュール・Excel 不要）
- 両スクリプトの処理内容は同一で、同じ入力から同じ JMX・Excel・Markdown を出力します。

詳しい使い方は **`docs/JMeter_Usage_Guide.xlsx`**（Meiryo UI・モノトーン）または **`docs/JMeter_Usage_Guide.md`** を参照してください。

## v1.1.0 の主な変更

同じ基準（[ProjectRubyRing/JMeter_API_Generate](https://github.com/ProjectRubyRing/JMeter_API_Generate)）から作られた別の実装（Argon / Grok / Codex 版）と比べ、
JMeter 5.6.3 での実通信で確認したうえで、良い点を取り入れました。

- **クエリの送り方**: 「Parameters」タブに置くのは GET だけにし、それ以外のメソッドは URL に付けます（JMeter 5.6.3 は DELETE の Parameters を URL とボディの両方に送り、HEAD / TRACE の Parameters は送らないため）。TRACE をメソッドとして扱い、ボディは送りません。
- **URL エンコード**: 名前か値に `%XX` を含む組はエンコード済みとして残りの文字だけをエンコードします（`x%20y 日本語` のような混在で JMeter がリクエストを送れなかった問題と、`a%26b` の名前の二重エンコードを修正）。URL のクエリの `+` は空白として、multipart の値は `%XX` を戻して送ります。リクエストボディの前後の空白・改行はそのまま送ります。
- **入力の確認**: 不正な URL（ホスト名なし・http / https 以外・ポート範囲外）は警告してスキップし、`ユーザー:パスワード@` は警告して取り除きます。`${PROTOCOL}://${HOST}/…` の形の URL に対応しました。範囲外の数値設定（スレッド数 0 など）は警告して既定値を使います。CSV の閉じていない引用符や、Excel の計算結果の無い数式・エラー値のセルはエラーにします（使うシートだけ）。Excel の `_xHHHH_` を正しく読み書きし、相対パスのシートも読めます。ヘッダの Cookie がクッキー管理に置き換えられる場合は警告します。
- **OpenAPI**: パスパラメータの style（simple / label / matrix）、explode しない配列・spaceDelimited / pipeDelimited のクエリ（値の中の区切り文字と区別）、フォームの encoding、content で定義したパラメータ、operation / path 単位の servers、`--required-only` での任意のリクエストボディの省略、ワイルドカードを除いた Accept に対応しました。データの生成規則は変えていないため、同じシードなら v1.0.0 と同じテストデータになります。
- **実行時（Groovy）**: Excel データ読込に失敗した場合は失敗として記録してテストを停止します（値の無い変数で API を呼ばない）。アクセストークンの読み込みに失敗したときは前回のトークンを使います。`~/`・`\` のパスに対応し、トークンのキャッシュを開始・終了時に消去します。
- **ファイル**: JMX・API一覧は書き終えてから置き換え（失敗しても元のファイルが残る）、同じ秒の退避ファイルには連番を付けます。`CON.jmx` などの Windows の予約名を避けます。HTML レポートが無い場合の集計は JTL を CSV として読みます。
- **add-curl**: 「Copy all as cURL」などで並べた複数の curl コマンド（改行・`;`・`&` 区切り）を、コマンドごとに 1 行ずつ追加します。
- **テスト**: `tests/` に、Python / PowerShell / bash の出力の一致、不具合の再発防止、JMeter 5.6.3 での実通信を確認するテストを追加しました（下記「テスト」）。

## フォルダ構成

| パス | 内容 |
|---|---|
| `jmeter_api_tool.sh` | RHEL 9.8 用スクリプト |
| `jmeter_api_tool.ps1` / `jmeter_api_tool.bat` | Windows 11 用スクリプト / 起動ファイル（実行ポリシーを回避して .ps1 を起動） |
| `input/api_definition.xlsx` | 入力テンプレート（Excel）：「API一覧」「設定」「記入方法」シート |
| `input/api_definition.csv` / `input/api_definition_settings.csv` | 入力テンプレート（CSV）：API 一覧 / 設定 |
| `input/openapi_config.csv` | OpenAPI 用設定ファイルのテンプレート：接続先（PROTOCOL / HOST / PORT）・HTTP ヘッダ・アクセストークンのフォルダ・OpenAPI 定義の変数（PUBLIC_BASE_URL / NAME_PREFIX） |
| `output/api_test.jmx` | テンプレートの記入例から生成した JMX（サンプル） |
| `output/api_test_linux_cmd.txt` | Linux サーバで JMeter の CLI を実行するコマンドと各パラメータの説明（サンプル。generate で JMX と一緒に出力） |
| `output/data/testdata.csv` / `testdata.xlsx` | CSV / Excel データ読込のサンプルデータ |
| `docs/JMeter_Usage_Guide.xlsx` / `.md` | 利用方法（manual コマンドで再出力できます） |
| `docs/JMeter_Scenario_Build_Guide.xlsx` | JMeter シナリオ組み立てガイド（生成した JMX に、レスポンス値の引継ぎ・CSV データ投入・ランダム ID・特定 API のループを JMeter GUI で組み込む手順） |
| `tests/` | テスト（Python の unittest。下記「テスト」） |

## クイックスタート

### RHEL 9.8

```bash
chmod +x jmeter_api_tool.sh
./jmeter_api_tool.sh init                                   # 入力テンプレートを作成（input/）
./jmeter_api_tool.sh generate -i input/api_definition.xlsx  # JMX を生成（output/）
./jmeter_api_tool.sh generate -i input/api_definition.xlsx -n order_test.jmx   # 出力する JMX のファイル名を指定
./jmeter_api_tool.sh generate -i openapi.json --seed 12345  # OpenAPI 定義（JSON）から JMX を生成
./jmeter_api_tool.sh generate -i openapi.json -c input/openapi_config.csv   # 接続先・ヘッダ・トークンを OpenAPI 用設定ファイルで指定
./jmeter_api_tool.sh add-curl -i input/api_definition.xlsx --curl-file curl.txt  # curl コマンドを解析して API一覧に追加（1 コマンド = 1 行）
./jmeter_api_tool.sh generate -i input/api_definition.xlsx -u output/api_test.jmx --dry-run   # 既存の JMX の更新箇所を確認（--dry-run を外すと更新）
./jmeter_api_tool.sh linux-cmd -t output/api_test.jmx       # Linux サーバでの実行コマンドを出力（generate でも出力）
./jmeter_api_tool.sh run -t output/api_test.jmx             # CLI 実行（result/JMX名_日時/）
./jmeter_api_tool.sh run -t output/api_test.jmx --threads 10 --loops 100   # 大量実行
./jmeter_api_tool.sh manual                                 # 利用方法を出力（docs/）
```

### Windows 11

```powershell
.\jmeter_api_tool.bat init
.\jmeter_api_tool.bat generate -i input\api_definition.xlsx
.\jmeter_api_tool.bat generate -i input\api_definition.xlsx -JmxName order_test.jmx
.\jmeter_api_tool.bat generate -i openapi.json -Seed 12345
.\jmeter_api_tool.bat generate -i openapi.json -c input\openapi_config.csv
.\jmeter_api_tool.bat add-curl -i input\api_definition.xlsx -CurlFile curl.txt
.\jmeter_api_tool.bat generate -i input\api_definition.xlsx -u output\api_test.jmx -DryRun
.\jmeter_api_tool.bat linux-cmd -t output\api_test.jmx
.\jmeter_api_tool.bat run -t output\api_test.jmx
.\jmeter_api_tool.bat run -t output\api_test.jmx -Threads 10 -Loops 100
.\jmeter_api_tool.bat manual
```

`run` コマンドには Java（17 推奨）と JMeter 5.6.3 が必要です（`JMETER_HOME` を設定するか `--jmeter-home` / `-JMeterHome` で指定）。
`init` / `generate` / `add-curl` / `linux-cmd` / `manual` は JMeter・Java が無くても動作します。

## OpenAPI 定義（JSON）からの生成

`-i` に OpenAPI 3.0.x の JSON（拡張子 `.json`）を指定すると、Excel / CSV と同じ「API一覧」の 1 行 = 1 API に変換してから JMX を生成します。
定義の読み込みとテストデータの生成は JMeter_Tools の `openapi2jmx`（`OpenApi2Jmx.ps1`）と同じ処理で、同じシードなら同じテストデータになります。

| API一覧の列 | OpenAPI から設定する値 |
|---|---|
| No / 有効 / API名 | 定義の順番 / Y / `メソッド パス [operationId]`（ANY は `ANY(GET) パス`） |
| URL | パス（パスパラメータはデータ型に合ったランダム値。style: simple / label / matrix に対応）。接続先は設定 `base_url`（operation / path に servers があればその接続先） |
| クエリパラメータ・ヘッダ・ボディ形式・リクエストボディ | パラメータ・requestBody の定義から生成（style / explode・content、JSON / フォーム（encoding）/ multipart / XML / テキスト） |
| 期待ステータス | 2xx の応答コード（`2XX` だけの場合は空欄） |
| 備考 | 概要・説明・operationId・認可（security）・API Gateway 統合など |

- 設定は Excel / CSV と同じく `-s` または「入力ファイル名_settings.csv」から読み込みます。設定ファイルで空欄の `base_url`（`servers[0].url`。無い場合は `http://localhost:8080`）、`test_plan_name`、`output_file`（`<タイトル>.jmx`）は OpenAPI 定義の値を使います（`output_file` は `-n` / `-JmxName` を指定するとその値が優先）。
- `servers[0].url` のプレースホルダ `${public_base_url}` とタイトル（`info.title`）の `${name_prefix}` は、OpenAPI 用設定ファイルの `PUBLIC_BASE_URL` / `NAME_PREFIX` の値に置き換えます。未設定（既定）の場合は取り除くため、パスの先頭（コンテキストルート）やテストプラン名・JMX のファイル名に `${...}` は残りません（例: `${name_prefix}-orders-api` → `orders-api`）。
- 認証ヘッダ（Authorization・API-Key）はランダム値では通らないため OpenAPI 定義からは生成しません。Authorization は設定 `common_headers`、または下記の OpenAPI 用設定ファイル（`HEADERS` / `TOKEN_DIR`）で指定してください。API-Key は下記の `API_HEADERS` で各 API のヘッダに初期値 `XXXXXXX` を設定するため、`API_HEADERS` の値を実際のキーに書き換えてください（`HEADERS` / `common_headers` に書いた API-Key より、実行時は各 API 側の値が優先されます）。
- 各 API のヘッダ（API ごとの HTTP Header Manager）には、OpenAPI 定義から求めたヘッダに加えて、OpenAPI 用設定ファイルの `API_HEADERS`（既定: `Accept: application/json`・`Transaction-ID: XXXXXXXXX`・`Transaction-Class: 1`・`User-ID: xxxxxxxx`・`API-Key: XXXXXXX`）を初期値として設定します（`-c` を指定しない場合は既定値）。
- OpenAPI 用のオプション（sh / PowerShell）: `-c` / `-OpenApiConfig`（OpenAPI 用設定ファイル。下記）、`--seed` / `-Seed`（乱数シード）、`--any-method` / `-AnyMethod`、`--required-only` / `-RequiredOnly`、`--use-examples` / `-UseExamples`、`--placeholder-value` / `-PlaceholderValue`（意味は JMeter_Tools と同じ）。
- GET / DELETE / HEAD / OPTIONS / TRACE のリクエストボディは Excel / CSV と同じ規則で送信しません（警告を表示）。YAML・Swagger 2.0・外部ファイルへの `$ref` は対象外です。
- operation / path に servers がある API は、その servers[0].url に送ります（相対 URL はドキュメントの接続先のホストを使用）。設定ファイル（`-s`）で `base_url` を指定した場合は、すべての API をその接続先に送ります。
- `--required-only` / `-RequiredOnly` は、任意（required でない）のリクエストボディも生成しません。
- 利用方法（`docs/JMeter_Usage_Guide.xlsx` / `.md`）の「6.OpenAPI定義から生成」にも、手順・変換ルール・OpenAPI 用設定ファイルをまとめています。

### OpenAPI 用設定ファイル（`-c` / `-OpenApiConfig`）

OpenAPI 定義から生成する場合だけ使う設定ファイルです。`init` で作成される `input/openapi_config.csv`（設定 CSV と同じ列: 設定キー / 設定項目 / 設定値 / 既定値 / 説明）の「設定値」を記入し、`generate -i openapi.json -c input/openapi_config.csv` のように指定します。
設定値が空欄の項目は使われず（`API_HEADERS` は既定値を使用）、テンプレートのままなら JMX は指定しない場合と同じです。記入した項目は、設定ファイル（`-s`）・OpenAPI 定義の値より優先します。

| 設定キー | 内容 |
|---|---|
| `PROTOCOL` / `HOST` / `PORT` | 共通変数（User Defined Variables）PROTOCOL / HOST / PORT の既定値。記入した項目だけ `base_url`（`servers[0].url`）と operation / path の servers の接続先を置き換え、パス（例: `/v1`）はそのまま使います。PORT が空欄の場合はポート番号を付けないため（未設定。http は 80 / https は 443 で接続）、ホスト名だけでも指定できます（3 項目とも空欄なら `base_url` のまま）。実行時の `-JPROTOCOL=` `-JHOST=` `-JPORT=` も従来どおり使えます |
| `HEADERS` | 全 API に付ける HTTP ヘッダ（「名前: 値」を改行区切り。例: `X-Api-Key: キーの値`）。共通ヘッダ（HTTP Header Manager）に追加し、設定 `common_headers` と同じ名前のヘッダはこちらを使います |
| `API_HEADERS` | 各 API のヘッダ（API ごとの HTTP Header Manager）に入れる初期値（「名前: 値」を改行区切り）。既定値は `Accept: application/json` / `Transaction-ID: XXXXXXXXX` / `Transaction-Class: 1` / `User-ID: xxxxxxxx` / `API-Key: XXXXXXX` で、空欄の場合も既定値を使います（付けない場合は `なし`）。OpenAPI 定義から求めた同じ名前のヘッダ（応答の形式から求めた `Accept` など）はこちらの値を使います。生成後は JMeter で API ごとに変更でき、`HEADERS`（共通ヘッダ）と同じ名前のヘッダは実行時に API 側が優先されます |
| `TOKEN_DIR` | アクセストークンのフォルダ。指定すると共通ヘッダに `Authorization: Bearer ${ACCESS_TOKEN}` を追加し、JMeter の実行中に API を呼び出すたびに、このフォルダ直下の `*.token` のうち **更新日時が最も新しいファイル** の内容（前後の空白・改行・BOM を除く）を `${ACCESS_TOKEN}` に設定します（`HEADERS` に Authorization を書いた場合はその値を使用） |
| `PUBLIC_BASE_URL` | OpenAPI 定義の `servers[0].url` のプレースホルダ `${public_base_url}` に入れる値。URL（例: `https://api.example.com/v1`）は接続先とパスの先頭に、パス（例: `/v1`）はパスの先頭（コンテキストルート）に使います（`PROTOCOL` / `HOST` / `PORT` を記入した場合は、その項目だけさらに置き換え）。空欄（既定）は未設定とし、`${public_base_url}` を取り除いてパスに付けません |
| `NAME_PREFIX` | OpenAPI 定義のタイトル（`info.title`）のプレースホルダ `${name_prefix}` に入れる値（例: `dev` → テストプラン名 `dev-orders-api 1.0.0（OpenAPI 3.0.1）`、JMX `dev-orders-api.jmx`）。空欄（既定）は未設定とし、`${name_prefix}` を隣の区切り文字（`-` `_` `.` `/` 空白）ごと取り除きます |

- `TOKEN_DIR` の書き方（Windows / Linux 共通）: 区切りは `\` と `/` のどちらでも可。相対パスは JMX のフォルダ基準（例: `tokens`、`..\tokens`）。Windows と Linux のフォルダを `;` で区切って書くと、実行する環境に存在する方を使うため、同じ JMX を両方で使えます（例: `C:\work\tokens;/home/user/tokens`）。実行時に `-JACCESS_TOKEN_DIR=フォルダ` で変更することもできます。
- トークンは JMX には書き込まず、JMeter の実行時に読み込みます。新しい `.token` ファイルを置く・ファイルを書き換えると、実行中でも次の API 呼び出しから切り替わります（フォルダの確認は 1 秒に 1 回）。拡張子が `.token` 以外のファイルは使いません。
- 生成時には、そのとき最新の `.token` ファイル（またはフォルダ・ファイルが無いことの警告）を表示します。
- `PUBLIC_BASE_URL` / `NAME_PREFIX` は OpenAPI 定義（`servers[0].url`・タイトル）に入れる値のため、設定ファイル（`-s`）で `base_url` / `test_plan_name` / `output_file` を指定した場合は、その設定の値を使います。

## curl コマンドからの追加・JMX の差分更新・Linux 用の実行コマンド

- **add-curl**: curl コマンド（ブラウザの「Copy as cURL」など）を解析して、API一覧（Excel は「API一覧」シートの空白行、CSV は末尾）に追加します。「Copy all as cURL」のように複数のコマンドを改行・`;`・`&` で区切って並べると、コマンドごとに 1 行ずつ追加します。メソッド・URL・クエリパラメータ・ヘッダ・ボディ（JSON / フォーム / multipart / テキスト）・Basic 認証（`-u`）・クッキー（`-b`）などを各列に変換し、JMeter で表現できないオプションは警告を表示して無視します。bash・コマンドプロンプト（`^`）・PowerShell（`` ` ``）形式の複数行のコマンドに対応し、元のファイルは `.日時.bak` に退避します。
- **generate -u / -UpdateJmx**: 新しい JMX を作らずに、既存の JMX を入力ファイル（Excel / CSV / OpenAPI）との差分箇所だけ更新します。前回生成した内容（`JMX名.jmx.base.json`）・現在の JMX・今回の定義の 3 つを比べるため、JMeter GUI で追加した要素や変更した値は残ります。更新箇所（[追加] / [変更] / [削除] / [競合] / [維持]）は書き換える前に表示し、`--dry-run` / `-DryRun` では表示だけ行います。更新前の JMX は `.日時.bak` に退避します。
- **JMX名_linux_cmd.txt**: generate で JMX と一緒に、Linux サーバで JMeter の CLI を実行するコマンド（そのまま貼り付けて実行できる形）と、持ち込むファイル・JMeter のオプション・JMX のプロパティ（`-J名前=値`）ごとの説明と既定値を出力します。既存の JMX からは `linux-cmd` コマンドで出力できます。
- 詳しくは利用方法の「3.JMX生成」（add-curl）・「8.JMXの更新」・「9.実行方法」を参照してください。

## 生成される JMX の主な内容

- 既定は 1 スレッド × 1 回。`-JTHREADS` `-JLOOPS` `-JRAMP_UP`（`-JDURATION`）で大量件数の負荷試験に切り替え
- 接続先は `${__P(HOST,...)}` 形式の共通変数（`-JHOST=` で環境切替）
- API ごと: ヘッダ、定数タイマ（待機）、前処理・後処理（JSR223 / Groovy）、正規表現抽出、JSON 抽出、ステータス・本文・応答時間の検証
- 共通: 連番（Counter）、乱数、CSV データ読込、Excel データ読込（Apache POI 不要）、クッキー管理、共通前処理・後処理、スループット制御（既定は無効）
- アクセストークン: 設定 `access_token_dir` を指定すると、そのフォルダ配下のファイル（`access_token_file`、既定 `access_token.txt`）からトークンを読み込み、共通ヘッダに `Authorization: Bearer ${ACCESS_TOKEN}` を付けます（実行中のファイル更新に追従。`-JACCESS_TOKEN_DIR` / `-JACCESS_TOKEN_FILE` で変更可）。`access_token_file` を `*.token` のように `*` を含めて書くと最も新しいファイルを使い、フォルダは `\` / `/` どちらの区切りでも、`;` 区切りで複数（存在する最初のフォルダを使用）でも書けます
- 結果出力: JTL（CSV）、エラー詳細（XML）、結果ログ（api_result.log）、jmeter.log、HTML レポート（run コマンド）
- Excel データ読込は、ファイル・データ行が無いなどで読み込めない場合に失敗として記録し、テストを停止します（Excel を使わない場合は設定 `excel_data` を OFF、実行時だけなら `-JEXCEL_FILE=`）
- 要素名は `[必須]` / `[任意]` で始まり、`[任意]` の要素は JMeter GUI で削除しても他の要素に影響しません。
  生成時に設定で `OFF`（生成しない）/ `DISABLE`（無効状態で生成）も選べます。

## テスト

`tests/` に Python 標準の unittest で書いたテストがあります（テストの実行に使う Python は開発時だけ必要です。利用時に Windows へ Python を入れる必要はありません）。

```bash
cd tests
python -X utf8 -m unittest -v                                             # すべて
JMETER_HOME=/opt/apache-jmeter-5.6.3 python -X utf8 -m unittest -v test_jmeter   # JMeter 5.6.3 での実通信（127.0.0.1 のサーバ。外部の API は呼びません）
```

| ファイル | 内容 |
|---|---|
| `test_syntax.py` | Python 3.6 の構文、`bash -n`、PowerShell の構文解析、.bat の起動、文字コード・改行コード、両スクリプトの共通データ（SPEC）の一致、版数の一致 |
| `test_regressions.py` | Python コアの単体テスト（クエリの置き場所・エンコード・URL・設定・CSV / Excel・OpenAPI の style / servers・add-curl・退避ファイル・JTL の集計、同じシードで同じテストデータになること） |
| `test_parity.py` | Python / PowerShell / bash で init・generate・manual・add-curl・generate -u を実行し、出力が一致することと、`input/`・`output/`・`docs/` のファイルが最新であること |
| `test_jmeter.py` | 生成した JMX を JMeter 5.6.3 で実行し、ローカルの HTTP サーバが受け取ったリクエスト（URL・ボディ・multipart・ヘッダ）、Excel / CSV データ、トークンの切り替え、Excel 読込失敗時の停止、run コマンドの終了コードと HTML レポートを確認（`JMETER_HOME` が無い場合は skip） |

- 使う実装は環境変数 `TEST_RUNTIMES`（例: `python,powershell`）、PowerShell は `TEST_POWERSHELL`（例: `pwsh` のパス）で変更できます。

## 動作確認済みの環境

| 環境 | 確認内容 |
|---|---|
| RHEL 9.8（Red Hat UBI 9.8 コンテナ、bash 5.1.8、python 3.9.25、OpenJDK 17.0.20、ロケール未設定） | init / generate / manual / run（JMeter 5.6.3 で実行、HTML レポート生成まで） |
| Windows 11（Windows PowerShell 5.1、Java 17.0.11） | init / generate / manual / run（成功・失敗・`-FailOnError`） |
| PowerShell 7（公式コンテナ） | init / generate / manual |
| Windows 11（Windows PowerShell 5.1 / Git Bash + python 3.12） | add-curl（bash・Chrome・コマンドプロンプト・PowerShell 形式など 15 種類の curl、CSV の文字コード 5 種類、Excel で保存した xlsx）/ generate -u（Excel・CSV・OpenAPI、JMeter で保存し直した JMX、CRLF の JMX）/ linux-cmd |
| Ubuntu 22.04（WSL、bash 5.1.16、python 3.10.12） | add-curl / generate -u / linux-cmd（Windows 版と出力が一致） |
| JMeter 5.6.3（`bin/jmeter`、Java 17.0.11） | `JMX名_linux_cmd.txt` の 1 行のコマンドをそのまま実行（add-curl・generate -u で更新した JMX。エラー 0 件・HTML レポート生成） |
| Microsoft Excel（Office 16） | 生成した xlsx の読込（修復なし）、全セルの折り返し高さ（はみ出しなし）、Excel で編集・保存したテンプレートの再読込 |
| Windows 11（Windows PowerShell 5.1 / PowerShell 7.6 / Git Bash + python 3.12、Java 17.0.11、JMeter 5.6.3） | v1.1.0: `tests/` のすべてのテスト（Python・PowerShell・bash の出力の一致、JMeter での実通信） |
| RHEL 9.6 コンテナ（bash 5.1.8、python 3.9.25、OpenJDK 21.0.7、JMeter 5.6.3） | v1.1.0: `tests/` のテスト（Python・bash の出力の一致、JMeter での実通信、run コマンド。PowerShell 関連は対象外） |

- RHEL 版と Windows 版の出力（JMX・xlsx の各 XML・md・csv）が、生成日時以外は完全に一致することを確認しています。
  add-curl で追記した CSV・xlsx のシート、generate -u で更新した JMX・基準ファイル、`JMX名_linux_cmd.txt` も同様です。
- v1.1.0 の変更は、Windows 上の Python・Windows PowerShell 5.1・PowerShell 7・Git Bash と RHEL 9.6 コンテナの bash・python 3.9 で出力が一致することと、JMeter 5.6.3 での実通信をテスト（`tests/`）で確認しています。RHEL 9.8 の実機・JMeter GUI・Excel アプリでの確認は v1.0.0 時点のものです。
- 生成した JMX は JMeter 5.6.3 の CLI / GUI で読み込みエラーがないこと、記入例の 4 API がエラー 0 件で実行できることを確認しています
  （記入例は公開サービス postman-echo.com を呼び出すため、インターネット接続が必要です）。

## 注意事項

- RHEL へは改行コード LF のまま転送してください（CRLF になった場合: `sed -i 's/\r$//' jmeter_api_tool.sh`）。
- `.ps1` は UTF-8（BOM 付き）で保存しています。編集する場合も BOM 付き UTF-8 のまま保存してください（Windows PowerShell 5.1 の文字化け防止）。
- generate で JMX と一緒に作られる `JMX名.jmx.base.json` は、JMX の更新（`-u` / `-UpdateJmx`）で JMeter 側の変更を残すために使います。JMX と同じフォルダに置いたままにしてください（無い場合は、JMX と定義を直接比べて定義の内容に更新します）。
- 両スクリプトには同じ定義データ（設定項目・テンプレート・Groovy スクリプト・利用方法の本文）が埋め込まれています。改修する場合は両方を同じ内容に変更してください。
