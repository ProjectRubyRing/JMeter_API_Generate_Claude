#!/usr/bin/env bash
# =============================================================================
#  jmeter_api_tool.sh  v1.0.0   （RHEL 9.8 用 / bash + python3 標準ライブラリ）
#
#  REST API の定義（Excel / CSV / OpenAPI）から Apache JMeter 5.6.3 用の JMX を生成し、
#  CLI モードで実行して JTL・HTML レポート・ログを出力するツールです。
#
#  使い方（詳細は ./jmeter_api_tool.sh help）:
#    ./jmeter_api_tool.sh init      入力テンプレート（Excel / CSV）を作成
#    ./jmeter_api_tool.sh generate  入力ファイルから JMX を生成
#    ./jmeter_api_tool.sh manual    利用方法（Excel / Markdown）を出力
#    ./jmeter_api_tool.sh run       JMeter を CLI モードで実行
#
#  必要なもの : bash、python3（3.6 以上。RHEL 9 は OS 標準の 3.9 で動作。追加モジュール不要）
#  run のみ   : Java 8 以上（17 推奨）、Apache JMeter 5.6.3
#  ※ 処理内容は Windows 版 jmeter_api_tool.ps1 と同一です。
# =============================================================================
if [ -z "${BASH_VERSION:-}" ]; then
  exec bash "$0" "$@"
fi
set -u
set -o pipefail

TOOL_VERSION="1.0.0"
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PYTHON_BIN=""

err()  { printf '[ERROR] %s\n' "$*" >&2; }
warn() { printf '[WARN] %s\n' "$*"; }
info() { printf '[INFO] %s\n' "$*"; }
die()  { err "$*"; exit 1; }

usage() {
  cat <<'EOF'
jmeter_api_tool.sh - REST API の JMeter テスト（JMX）生成・実行ツール

使い方:
  ./jmeter_api_tool.sh init     [-o フォルダ] [-f]
      入力テンプレート（api_definition.xlsx / .csv / _settings.csv・openapi_config.csv）を作成します（既定: input）

  ./jmeter_api_tool.sh generate -i 入力ファイル [-s 設定CSV] [-o フォルダ] [-n ファイル名] [--split] [-f] [OpenAPI用オプション]
      Excel（.xlsx）・CSV・OpenAPI 定義（.json）の API 定義から JMX を生成します（既定: output）
        -i FILE    入力ファイル（.xlsx / .csv / .json）
        -s FILE    設定 CSV（省略時は「入力ファイル名_settings.csv」→ 既定値）
        -o DIR     出力先フォルダ
        -n NAME    出力する JMX のファイル名（設定 output_file より優先。例: -n order_test.jmx）
        --split    API ごとに別々の JMX を出力
        -f         サンプルデータ（data/testdata.*）を上書き
        -u JMX     既存の JMX（このツールで生成したもの）を、入力ファイルとの差分箇所だけ更新（新しい JMX は作りません。
                   更新箇所を表示し、更新前の JMX は「.日時.bak」を付けて退避。-o・-n・--split は使いません）
        --dry-run  -u と併用: 更新箇所を表示するだけで、ファイルは書き換えない
      OpenAPI 定義（.json）用のオプション:
        -c FILE                OpenAPI 用設定ファイル（接続先 PROTOCOL / HOST / PORT・HTTP ヘッダ・アクセストークンの
                               フォルダ TOKEN_DIR・定義の変数 PUBLIC_BASE_URL / NAME_PREFIX。
                               テンプレート: init で作成される input/openapi_config.csv）
        --seed N               テストデータの乱数シード（同じ値なら同じデータ。既定: 毎回ランダム）
        --any-method M         x-amazon-apigateway-any-method（ANY）で送るメソッド（既定: GET）
        --required-only        必須（required）のパラメータ・プロパティだけを生成
        --use-examples         example / default があればランダム値より優先
        --placeholder-value N  値の位置に引用符なしで書かれた ${...} に仮に入れる数値（既定: 3000）

  ./jmeter_api_tool.sh add-curl -i API一覧 (--curl "curl ..." | --curl-file ファイル | -- curl ...) [--name API名] [--dry-run]
      curl コマンドを解析して、API一覧（Excel は「API一覧」シートの空白行、CSV は末尾）に 1 行追加します
      （元のファイルは「ファイル名.日時.bak」に退避）
        -i FILE            追加先の API一覧（.xlsx / .csv）
        --curl "..."       curl コマンド（文字列）
        --curl-file FILE   curl コマンドを書いたテキストファイル（ブラウザの「Copy as cURL」の内容を保存したもの）
        -- curl ...        -- の後ろに curl コマンドをそのまま書く
        --name NAME        API名（省略時: メソッド + パス）
        --dry-run          追加する行を表示するだけで、ファイルは書き換えない

  ./jmeter_api_tool.sh linux-cmd -t JMX [-o フォルダ]
      Linux サーバで JMeter の CLI を実行するコマンドと、各パラメータの説明を「JMX名_linux_cmd.txt」に出力します
      （generate でも JMX と同じフォルダに出力されます。既定: JMX と同じフォルダ）

  ./jmeter_api_tool.sh manual   [-o フォルダ]
      利用方法（JMeter_Usage_Guide.xlsx / .md）を出力します（既定: docs）

  ./jmeter_api_tool.sh run -t JMX [オプション] [-- JMeterの追加引数...]
      JMeter を CLI モードで実行し、result/JMX名_日時/ に JTL・ログ・HTML レポートを出力します
        --threads N        スレッド数（-JTHREADS）
        --rampup SEC       ランプアップ秒数（-JRAMP_UP）
        --loops N          ループ回数（-JLOOPS、-1 = 無限）
        --duration SEC     実行時間（-JDURATION、duration 指定で生成した JMX のみ）
        -J 名前=値         任意のプロパティ（複数指定可。例: -J HOST=stg.example.com）
        -o DIR             結果の出力先（既定: result）
        --heap SIZE        JVM ヒープ（例: 2g。既定: 1g）
        --jmeter-home DIR  JMeter のフォルダ（省略時: JMETER_HOME → PATH → /opt 等を探索）
        --fail-on-error    エラーが 1 件でもあれば終了コード 2

  ./jmeter_api_tool.sh help | version

例:
  ./jmeter_api_tool.sh init
  ./jmeter_api_tool.sh generate -i input/api_definition.xlsx
  ./jmeter_api_tool.sh generate -i openapi.json --seed 12345
  ./jmeter_api_tool.sh generate -i openapi.json -c input/openapi_config.csv
  ./jmeter_api_tool.sh generate -i input/api_definition.xlsx -u output/api_test.jmx --dry-run
  ./jmeter_api_tool.sh add-curl -i input/api_definition.xlsx --curl-file curl.txt
  ./jmeter_api_tool.sh linux-cmd -t output/api_test.jmx
  ./jmeter_api_tool.sh run -t output/api_test.jmx --threads 10 --loops 100
EOF
}

find_python() {
  local cand
  for cand in "${PYTHON:-}" python3 /usr/bin/python3 /usr/libexec/platform-python python3.12 python3.11 python3.9 python; do
    [ -n "$cand" ] || continue
    command -v "$cand" >/dev/null 2>&1 || continue
    if "$cand" -c 'import sys; sys.exit(0 if sys.version_info >= (3, 6) else 1)' >/dev/null 2>&1; then
      PYTHON_BIN="$(command -v "$cand")"
      return 0
    fi
  done
  return 1
}

need_value() {
  [ $# -ge 2 ] || die "$1 には値が必要です"
}

find_jmeter() {
  local explicit="$1" cand resolved
  local -a cands=()
  if [ -n "$explicit" ]; then
    if [ -f "$explicit/bin/jmeter" ] && [ -f "$explicit/bin/ApacheJMeter.jar" ]; then
      (cd "$explicit" && pwd)
      return 0
    fi
    err "指定されたフォルダに JMeter がありません: $explicit"
    return 1
  fi
  [ -n "${JMETER_HOME:-}" ] && cands+=("$JMETER_HOME")
  if command -v jmeter >/dev/null 2>&1; then
    resolved="$(readlink -f "$(command -v jmeter)" 2>/dev/null || command -v jmeter)"
    cands+=("$(dirname "$(dirname "$resolved")")")
  fi
  cands+=("/opt/apache-jmeter-5.6.3" "/opt/jmeter" "/usr/local/apache-jmeter-5.6.3" "/usr/local/jmeter"
          "$HOME/apache-jmeter-5.6.3" "$SCRIPT_DIR/apache-jmeter-5.6.3")
  for cand in "${cands[@]}"; do
    if [ -f "$cand/bin/jmeter" ] && [ -f "$cand/bin/ApacheJMeter.jar" ]; then
      (cd "$cand" && pwd)
      return 0
    fi
  done
  return 1
}

check_java() {
  local java_cmd="java" ver major
  if [ -n "${JAVA_HOME:-}" ] && [ -x "${JAVA_HOME}/bin/java" ]; then
    java_cmd="${JAVA_HOME}/bin/java"
  fi
  command -v "$java_cmd" >/dev/null 2>&1 || die "Java が見つかりません。sudo dnf install -y java-17-openjdk でインストールしてください。"
  ver="$("$java_cmd" -version 2>&1 | awk -F'"' '/version/ {print $2; exit}')"
  major="${ver%%.*}"
  if [ "$major" = "1" ]; then
    major="$(printf '%s' "$ver" | cut -d. -f2)"
  fi
  if [[ "$major" =~ ^[0-9]+$ ]]; then
    [ "$major" -ge 8 ] || die "Java 8 以上が必要です（現在: $ver）。Java 17 を推奨します。"
    info "Java: $ver"
  else
    warn "Java のバージョンを判別できませんでした（$("$java_cmd" -version 2>&1 | head -n 1)）"
  fi
}

run_core() {
  if [ -z "$PYTHON_BIN" ]; then
    find_python || die "python3（3.6 以上）が見つかりません。sudo dnf install -y python3 でインストールしてください。"
  fi
  "$PYTHON_BIN" -I - "$@" <<'__PYTHON_CORE__'
# -*- coding: utf-8 -*-
# =============================================================================
#  jmeter_api_tool - Python コア（標準ライブラリのみ / Python 3.6 以上）
#  jmeter_api_tool.sh に埋め込まれ、python3 で実行されます。
#  ※ 処理内容は jmeter_api_tool.ps1（Windows 版）と同一です。
# =============================================================================
import base64
import datetime
import fnmatch
import hashlib
import io
import json
import os
import re
import sys
import unicodedata
import zipfile
import xml.etree.ElementTree as ET
from collections import OrderedDict
from decimal import Decimal, ROUND_CEILING, ROUND_FLOOR

SPEC = json.loads(r'''{
 "tool": {
  "name": "jmeter_api_tool",
  "version": "1.0.0",
  "jmeter": "5.6.3"
 },
 "settings": [
  {
   "category": "基本",
   "key": "test_plan_name",
   "label": "テストプラン名",
   "default": "APIテスト",
   "type": "text",
   "desc": "JMeter 上に表示されるテストプラン名"
  },
  {
   "category": "基本",
   "key": "output_file",
   "label": "出力JMXファイル名",
   "default": "api_test.jmx",
   "type": "text",
   "desc": "生成する JMX のファイル名（英数字・_・-・. 以外の文字は _ に置き換えます）"
  },
  {
   "category": "基本",
   "key": "base_url",
   "label": "ベースURL",
   "default": "",
   "type": "text",
   "desc": "API一覧の URL が「/」で始まる場合に先頭に付ける URL（例: https://api.example.com/v1）"
  },
  {
   "category": "基本",
   "key": "content_encoding",
   "label": "文字コード",
   "default": "UTF-8",
   "type": "text",
   "desc": "リクエストの文字コード"
  },
  {
   "category": "負荷",
   "key": "threads",
   "label": "スレッド数",
   "default": "1",
   "type": "int",
   "desc": "同時に実行する仮想ユーザー数。-JTHREADS"
  },
  {
   "category": "負荷",
   "key": "ramp_up",
   "label": "ランプアップ(秒)",
   "default": "1",
   "type": "int",
   "desc": "全スレッドが起動し終わるまでの秒数。-JRAMP_UP"
  },
  {
   "category": "負荷",
   "key": "loops",
   "label": "ループ回数",
   "default": "1",
   "type": "int",
   "desc": "各スレッドが API 一覧を繰り返す回数。-1 = 無限（実行時間と併用）。-JLOOPS"
  },
  {
   "category": "負荷",
   "key": "duration",
   "label": "実行時間(秒)",
   "default": "",
   "type": "int",
   "desc": "指定するとスケジューラを有効にし、この秒数で終了（ループ回数は -1 を推奨）。-JDURATION"
  },
  {
   "category": "負荷",
   "key": "on_sample_error",
   "label": "エラー時の動作",
   "default": "continue",
   "type": "enum",
   "desc": "continue = 続行 / startnextloop = 次のループへ / stopthread = スレッド停止 / stoptest = テスト停止 / stoptestnow = 即時停止",
   "choices": [
    "continue",
    "startnextloop",
    "stopthread",
    "stoptest",
    "stoptestnow"
   ]
  },
  {
   "category": "負荷",
   "key": "throughput_timer",
   "label": "スループット制御",
   "default": "DISABLE",
   "type": "onoff",
   "desc": "1 分あたりのリクエスト数を一定に保つ Constant Throughput Timer",
   "choices": [
    "ON",
    "OFF",
    "DISABLE"
   ]
  },
  {
   "category": "負荷",
   "key": "throughput_per_min",
   "label": "目標スループット(件/分)",
   "default": "600",
   "type": "int",
   "desc": "スループット制御の目標値（メイン処理全体の合計）。-JTHROUGHPUT"
  },
  {
   "category": "HTTP",
   "key": "http_defaults",
   "label": "HTTPリクエスト初期値",
   "default": "ON",
   "type": "onoff",
   "desc": "文字コード・タイムアウトの共通設定（HTTP Request Defaults）",
   "choices": [
    "ON",
    "OFF",
    "DISABLE"
   ]
  },
  {
   "category": "HTTP",
   "key": "connect_timeout",
   "label": "接続タイムアウト(ms)",
   "default": "10000",
   "type": "int",
   "desc": "-JCONNECT_TIMEOUT"
  },
  {
   "category": "HTTP",
   "key": "response_timeout",
   "label": "応答タイムアウト(ms)",
   "default": "30000",
   "type": "int",
   "desc": "-JRESPONSE_TIMEOUT"
  },
  {
   "category": "HTTP",
   "key": "common_headers",
   "label": "共通ヘッダ",
   "default": "",
   "type": "lines",
   "desc": "全 API に付けるヘッダ。「名前: 値」を改行区切り",
   "template": "Accept: application/json\nUser-Agent: JMeter-API-Test/1.0\nX-Client-Id: ${CLIENT_ID}"
  },
  {
   "category": "HTTP",
   "key": "access_token_dir",
   "label": "アクセストークンのフォルダ",
   "default": "",
   "type": "text",
   "desc": "指定すると、このフォルダ配下のファイルからアクセストークンを読み込み、共通ヘッダに Authorization: Bearer ${ACCESS_TOKEN} を付けます（共通ヘッダに Authorization がある場合はその値を使用）。空欄 = 使わない。JMX からの相対パスまたは絶対パス（区切りは \\ と / のどちらでも可。; で区切って複数書くと、存在する最初のフォルダを使用）。-JACCESS_TOKEN_DIR"
  },
  {
   "category": "HTTP",
   "key": "access_token_file",
   "label": "アクセストークンのファイル名",
   "default": "access_token.txt",
   "type": "text",
   "desc": "アクセストークンを書いたファイルの名前（フォルダからの相対パス）。*.token のように * を含めると、一致するファイルのうち更新日時が最も新しいファイルを使います。前後の空白・改行は除き、実行中にファイルが更新されると読み直します。-JACCESS_TOKEN_FILE"
  },
  {
   "category": "HTTP",
   "key": "follow_redirects",
   "label": "リダイレクト追従",
   "default": "true",
   "type": "bool",
   "desc": "true = 3xx 応答の Location を自動で追従",
   "choices": [
    "true",
    "false"
   ]
  },
  {
   "category": "HTTP",
   "key": "keep_alive",
   "label": "KeepAlive",
   "default": "true",
   "type": "bool",
   "desc": "true = 接続を再利用する",
   "choices": [
    "true",
    "false"
   ]
  },
  {
   "category": "HTTP",
   "key": "default_wait_ms",
   "label": "既定の待機時間(ms)",
   "default": "0",
   "type": "int",
   "desc": "API一覧の「待機時間」が空欄の API に使う値（0 = 待機なし）"
  },
  {
   "category": "クッキー",
   "key": "cookie_manager",
   "label": "クッキー管理",
   "default": "ON",
   "type": "onoff",
   "desc": "HTTP Cookie Manager（Set-Cookie を保存して以降のリクエストで送信）",
   "choices": [
    "ON",
    "OFF",
    "DISABLE"
   ]
  },
  {
   "category": "クッキー",
   "key": "cookie_clear_each_iteration",
   "label": "反復ごとにクリア",
   "default": "true",
   "type": "bool",
   "desc": "true = ループのたびにクッキーを消去（毎回新しいセッション）",
   "choices": [
    "true",
    "false"
   ]
  },
  {
   "category": "クッキー",
   "key": "cookies",
   "label": "事前設定クッキー",
   "default": "",
   "type": "lines",
   "desc": "最初から送信するクッキー。「名前=値」を改行区切り",
   "template": "LANG=ja"
  },
  {
   "category": "クッキー",
   "key": "cookie_domain",
   "label": "クッキーのドメイン",
   "default": "",
   "type": "text",
   "desc": "事前設定クッキーのドメイン。空欄 = 1 つ目の接続先（${HOST}）"
  },
  {
   "category": "変数",
   "key": "user_variables",
   "label": "共通変数",
   "default": "",
   "type": "lines",
   "desc": "「名前=値」を改行区切り。${名前} で参照でき、実行時に -J名前=値 で変更可",
   "template": "ENV=staging\nCLIENT_ID=jmeter-client"
  },
  {
   "category": "変数",
   "key": "counter",
   "label": "連番",
   "default": "ON",
   "type": "onoff",
   "desc": "反復ごとに連番を設定する Counter",
   "choices": [
    "ON",
    "OFF",
    "DISABLE"
   ]
  },
  {
   "category": "変数",
   "key": "counter_name",
   "label": "連番の変数名",
   "default": "SEQ",
   "type": "text",
   "desc": "連番を入れる変数名"
  },
  {
   "category": "変数",
   "key": "counter_start",
   "label": "連番の開始値",
   "default": "1",
   "type": "int",
   "desc": ""
  },
  {
   "category": "変数",
   "key": "counter_increment",
   "label": "連番の増分",
   "default": "1",
   "type": "int",
   "desc": ""
  },
  {
   "category": "変数",
   "key": "counter_format",
   "label": "連番の書式",
   "default": "000000",
   "type": "text",
   "desc": "0 埋めの桁数など（Java の DecimalFormat 形式）。空欄 = 書式なし"
  },
  {
   "category": "変数",
   "key": "counter_per_thread",
   "label": "スレッドごとの連番",
   "default": "false",
   "type": "bool",
   "desc": "true = スレッドごとに開始値から採番 / false = 全スレッドで一意",
   "choices": [
    "true",
    "false"
   ]
  },
  {
   "category": "変数",
   "key": "random_variable",
   "label": "乱数",
   "default": "ON",
   "type": "onoff",
   "desc": "反復ごとに乱数を設定する Random Variable",
   "choices": [
    "ON",
    "OFF",
    "DISABLE"
   ]
  },
  {
   "category": "変数",
   "key": "random_name",
   "label": "乱数の変数名",
   "default": "RANDOM_NO",
   "type": "text",
   "desc": ""
  },
  {
   "category": "変数",
   "key": "random_min",
   "label": "乱数の最小値",
   "default": "1",
   "type": "int",
   "desc": ""
  },
  {
   "category": "変数",
   "key": "random_max",
   "label": "乱数の最大値",
   "default": "99999",
   "type": "int",
   "desc": ""
  },
  {
   "category": "CSVデータ",
   "key": "csv_data",
   "label": "CSVデータ読込",
   "default": "ON",
   "type": "onoff",
   "desc": "CSV Data Set Config（1 反復につき 1 行を読み込み、列を変数に設定）",
   "choices": [
    "ON",
    "OFF",
    "DISABLE"
   ]
  },
  {
   "category": "CSVデータ",
   "key": "csv_file",
   "label": "CSVファイル",
   "default": "data/testdata.csv",
   "type": "text",
   "desc": "JMX からの相対パスまたは絶対パス。-JCSV_FILE"
  },
  {
   "category": "CSVデータ",
   "key": "csv_variables",
   "label": "CSVの変数名",
   "default": "",
   "type": "text",
   "desc": "カンマ区切り。空欄 = 生成時に CSV ファイルの 1 行目を読み取って設定"
  },
  {
   "category": "CSVデータ",
   "key": "csv_encoding",
   "label": "CSVの文字コード",
   "default": "UTF-8",
   "type": "text",
   "desc": ""
  },
  {
   "category": "CSVデータ",
   "key": "csv_delimiter",
   "label": "CSVの区切り文字",
   "default": ",",
   "type": "text",
   "desc": ""
  },
  {
   "category": "CSVデータ",
   "key": "csv_recycle",
   "label": "終端で先頭に戻る",
   "default": "true",
   "type": "bool",
   "desc": "true = 最終行の次は先頭行に戻る",
   "choices": [
    "true",
    "false"
   ]
  },
  {
   "category": "CSVデータ",
   "key": "csv_stop_thread",
   "label": "終端でスレッド停止",
   "default": "false",
   "type": "bool",
   "desc": "true = 最終行まで読んだらスレッドを停止（csv_recycle=false と併用）",
   "choices": [
    "true",
    "false"
   ]
  },
  {
   "category": "CSVデータ",
   "key": "csv_share_mode",
   "label": "CSVの共有範囲",
   "default": "all",
   "type": "enum",
   "desc": "all = 全スレッドで共有 / group = スレッドグループ単位 / thread = スレッドごとに先頭から",
   "choices": [
    "all",
    "group",
    "thread"
   ]
  },
  {
   "category": "Excelデータ",
   "key": "excel_data",
   "label": "Excelデータ読込",
   "default": "ON",
   "type": "onoff",
   "desc": "Excel（.xlsx）の読み込み。Apache POI などの追加ライブラリは不要",
   "choices": [
    "ON",
    "OFF",
    "DISABLE"
   ]
  },
  {
   "category": "Excelデータ",
   "key": "excel_file",
   "label": "Excelファイル",
   "default": "data/testdata.xlsx",
   "type": "text",
   "desc": "JMX からの相対パスまたは絶対パス。-JEXCEL_FILE"
  },
  {
   "category": "Excelデータ",
   "key": "excel_sheet",
   "label": "シート名",
   "default": "",
   "type": "text",
   "desc": "空欄 = 先頭のシート。-JEXCEL_SHEET"
  },
  {
   "category": "前後処理",
   "key": "setup_teardown",
   "label": "開始・終了処理",
   "default": "ON",
   "type": "onoff",
   "desc": "テスト開始・終了時のログ出力と後片付け（setUp / tearDown Thread Group）",
   "choices": [
    "ON",
    "OFF",
    "DISABLE"
   ]
  },
  {
   "category": "前後処理",
   "key": "common_preprocessor",
   "label": "共通前処理",
   "default": "ON",
   "type": "onoff",
   "desc": "全 API の直前に実行する JSR223 PreProcessor（REQUEST_ID・TIMESTAMP 等を生成）",
   "choices": [
    "ON",
    "OFF",
    "DISABLE"
   ]
  },
  {
   "category": "前後処理",
   "key": "common_postprocessor",
   "label": "共通後処理",
   "default": "ON",
   "type": "onoff",
   "desc": "全 API の直後に実行する JSR223 PostProcessor（LAST_STATUS 等を設定）",
   "choices": [
    "ON",
    "OFF",
    "DISABLE"
   ]
  },
  {
   "category": "結果出力",
   "key": "result_dir",
   "label": "結果出力フォルダ",
   "default": "~/result",
   "type": "text",
   "desc": "JTL・結果ログの出力先。~/ = JMX のフォルダ。-JRESULT_DIR（run コマンドは自動設定）"
  },
  {
   "category": "結果出力",
   "key": "view_results_tree",
   "label": "結果ツリー",
   "default": "ON",
   "type": "onoff",
   "desc": "View Results Tree（GUI での確認用）",
   "choices": [
    "ON",
    "OFF",
    "DISABLE"
   ]
  },
  {
   "category": "結果出力",
   "key": "summary_report",
   "label": "集計レポート",
   "default": "ON",
   "type": "onoff",
   "desc": "Summary Report（GUI での確認用）",
   "choices": [
    "ON",
    "OFF",
    "DISABLE"
   ]
  },
  {
   "category": "結果出力",
   "key": "aggregate_report",
   "label": "統計レポート",
   "default": "ON",
   "type": "onoff",
   "desc": "Aggregate Report（GUI での確認用。パーセンタイル）",
   "choices": [
    "ON",
    "OFF",
    "DISABLE"
   ]
  },
  {
   "category": "結果出力",
   "key": "jtl_writer",
   "label": "JTL出力(全件)",
   "default": "ON",
   "type": "onoff",
   "desc": "全リクエストの結果を result_all.jtl（CSV）に出力",
   "choices": [
    "ON",
    "OFF",
    "DISABLE"
   ]
  },
  {
   "category": "結果出力",
   "key": "error_writer",
   "label": "エラー詳細出力",
   "default": "ON",
   "type": "onoff",
   "desc": "失敗したリクエストのみ error_detail.xml（リクエスト・レスポンス全文）に出力",
   "choices": [
    "ON",
    "OFF",
    "DISABLE"
   ]
  },
  {
   "category": "結果出力",
   "key": "result_log",
   "label": "結果ログ出力",
   "default": "ON",
   "type": "onoff",
   "desc": "1 リクエスト 1 行の結果ログ api_result.log を出力（JSR223 Listener）",
   "choices": [
    "ON",
    "OFF",
    "DISABLE"
   ]
  },
  {
   "category": "結果出力",
   "key": "log_response",
   "label": "結果ログの本文出力",
   "default": "error",
   "type": "enum",
   "desc": "none = 出力しない / error = 失敗時のみ / all = 全件。-JLOG_RESPONSE",
   "choices": [
    "none",
    "error",
    "all"
   ]
  },
  {
   "category": "結果出力",
   "key": "log_max_chars",
   "label": "本文の最大文字数",
   "default": "1000",
   "type": "int",
   "desc": "結果ログに出力する本文の最大文字数。-JLOG_MAX_CHARS"
  },
  {
   "category": "結果出力",
   "key": "debug_sampler",
   "label": "変数確認(デバッグ)",
   "default": "DISABLE",
   "type": "onoff",
   "desc": "Debug Sampler（全変数を結果ツリーに表示）",
   "choices": [
    "ON",
    "OFF",
    "DISABLE"
   ]
  }
 ],
 "openapi_config": [
  {
   "category": "接続先",
   "key": "PROTOCOL",
   "label": "プロトコル",
   "default": "",
   "desc": "共通変数 PROTOCOL の既定値（http / https）。空欄 = OpenAPI 定義（servers[0].url）または設定 base_url の値。実行時は -JPROTOCOL=値 で変更できます"
  },
  {
   "category": "接続先",
   "key": "HOST",
   "label": "ホスト名",
   "default": "",
   "desc": "共通変数 HOST の既定値（例: api.example.com）。空欄 = OpenAPI 定義（servers[0].url）または設定 base_url の値。実行時は -JHOST=値 で変更できます"
  },
  {
   "category": "接続先",
   "key": "PORT",
   "label": "ポート番号",
   "default": "",
   "desc": "共通変数 PORT の既定値（1～65535）。空欄 = 未設定（ポート番号なし。http は 80 / https は 443 で接続）。PROTOCOL と HOST も空欄の場合は OpenAPI 定義（servers[0].url）または設定 base_url の値。実行時は -JPORT=値 で変更できます"
  },
  {
   "category": "HTTPヘッダ",
   "key": "HEADERS",
   "label": "HTTPヘッダ",
   "default": "",
   "desc": "全 API に付ける HTTP ヘッダ（共通ヘッダの HTTP Header Manager に追加）。「名前: 値」を改行区切り（例: X-Api-Key: キーの値）。設定 common_headers と同じ名前のヘッダはこちらを使います"
  },
  {
   "category": "HTTPヘッダ",
   "key": "API_HEADERS",
   "label": "各APIのHTTPヘッダ（初期値）",
   "default": "Accept: application/json\nTransaction-ID: XXXXXXXXX\nTransaction-Class: 1\nUser-ID: xxxxxxxx\nAPI-Key: XXXXXXX",
   "desc": "OpenAPI 定義から生成する各 API のヘッダ（API ごとの HTTP Header Manager）の初期値（生成後に JMeter で API ごとに変更できます）。「名前: 値」を改行区切り。OpenAPI 定義から求めた同じ名前のヘッダ（Accept など）はこちらの値を使います。空欄 = 既定値。付けない場合は「なし」と記入"
  },
  {
   "category": "アクセストークン",
   "key": "TOKEN_DIR",
   "label": "アクセストークンのフォルダ",
   "default": "",
   "desc": "指定すると、共通ヘッダに Authorization: Bearer ${ACCESS_TOKEN} を追加し、JMeter の実行中に API を呼び出すたびに、このフォルダ直下の拡張子 .token のファイルのうち更新日時が最も新しいファイルの内容を ${ACCESS_TOKEN} に設定します（HEADERS に Authorization を書いた場合はその値を使用）。区切りは \\ と / のどちらでも可、相対パスは JMX のフォルダ基準。Windows と Linux のフォルダを ; で区切って書くと、存在する方を使います（例: C:\\work\\tokens;/home/user/tokens）。実行時は -JACCESS_TOKEN_DIR=フォルダ で変更できます"
  },
  {
   "category": "OpenAPI定義の変数",
   "key": "PUBLIC_BASE_URL",
   "label": "ベースURL（${public_base_url}）",
   "default": "",
   "desc": "OpenAPI 定義（servers[0].url）のプレースホルダ ${public_base_url} に入れる値。URL（例: https://api.example.com/v1）は接続先とパスの先頭に、パス（例: /v1）はパスの先頭（コンテキストルート）に使います。空欄 = 未設定（${public_base_url} を取り除き、パスに付けません）"
  },
  {
   "category": "OpenAPI定義の変数",
   "key": "NAME_PREFIX",
   "label": "名前の接頭辞（${name_prefix}）",
   "default": "",
   "desc": "OpenAPI 定義（info.title）のプレースホルダ ${name_prefix} に入れる値（例: dev）。テストプラン名と出力する JMX のファイル名に使われます。空欄 = 未設定（${name_prefix} を隣の区切り文字（- _ . / 空白）ごと取り除きます。例: ${name_prefix}-api → api）"
  }
 ],
 "api_columns": [
  {
   "key": "no",
   "header": "No",
   "aliases": [
    "番号",
    "no."
   ],
   "required": false,
   "width": 6,
   "hint": "#説明",
   "desc": "実行順の番号（ラベル APIxx に使用）。空欄は自動採番。# で始まる行は読み飛ばす",
   "example": "1"
  },
  {
   "key": "enabled",
   "header": "有効",
   "aliases": [
    "enabled",
    "有効/無効",
    "有効フラグ"
   ],
   "required": false,
   "width": 6,
   "hint": "Y / N",
   "desc": "Y = 有効、N = 無効化した状態で生成。空欄 = Y",
   "example": "Y"
  },
  {
   "key": "name",
   "header": "API名",
   "aliases": [
    "name",
    "名前",
    "api"
   ],
   "required": false,
   "width": 22,
   "hint": "任意",
   "desc": "HTTP リクエストの名前（レポートのラベル）。空欄はメソッドとパス",
   "example": "ログイン"
  },
  {
   "key": "method",
   "header": "メソッド",
   "aliases": [
    "method",
    "httpメソッド",
    "httpmethod"
   ],
   "required": true,
   "width": 10,
   "hint": "GET / POST",
   "desc": "HTTP メソッド（GET / POST。PUT / PATCH / DELETE も可）",
   "example": "POST"
  },
  {
   "key": "url",
   "header": "URL",
   "aliases": [
    "url",
    "エンドポイント",
    "endpoint"
   ],
   "required": true,
   "width": 42,
   "hint": "必須（https://～）",
   "desc": "完全な URL、またはベースURL 設定時は「/」で始まるパス。${変数} 可。? 以降のクエリも可",
   "example": "https://api.example.com/v1/users"
  },
  {
   "key": "query",
   "header": "クエリパラメータ",
   "aliases": [
    "query",
    "params",
    "クエリ",
    "パラメータ"
   ],
   "required": false,
   "width": 28,
   "hint": "名前=値（改行 / & 区切り）",
   "desc": "URL のクエリパラメータ",
   "example": "page=1\nsize=20"
  },
  {
   "key": "headers",
   "header": "ヘッダ",
   "aliases": [
    "headers",
    "header",
    "リクエストヘッダ",
    "ヘッダ情報"
   ],
   "required": false,
   "width": 34,
   "hint": "名前: 値（改行区切り）",
   "desc": "この API だけに付けるヘッダ（全 API 共通のヘッダは設定 common_headers）",
   "example": "Authorization: Bearer ${AUTH_TOKEN}"
  },
  {
   "key": "body_type",
   "header": "ボディ形式",
   "aliases": [
    "body_type",
    "bodytype",
    "形式"
   ],
   "required": false,
   "width": 11,
   "hint": "auto / json / form / raw",
   "desc": "リクエストボディの形式。空欄 = auto（自動判定）",
   "example": "json"
  },
  {
   "key": "body",
   "header": "リクエストボディ",
   "aliases": [
    "body",
    "requestbody",
    "ボディ",
    "postパラメータ"
   ],
   "required": false,
   "width": 40,
   "hint": "JSON 等 / 名前=値",
   "desc": "POST 等で送る本文。form 形式は「名前=値」を改行 / & 区切り",
   "example": "{\"userId\":\"${USER_ID}\"}"
  },
  {
   "key": "wait_ms",
   "header": "待機時間(ms)",
   "aliases": [
    "wait",
    "wait_ms",
    "待機時間",
    "待機"
   ],
   "required": false,
   "width": 11,
   "hint": "数値（ms）",
   "desc": "API 呼び出し前に待つ時間（Constant Timer）。空欄 = 設定 default_wait_ms、0 = 待機なし",
   "example": "1000"
  },
  {
   "key": "expected_status",
   "header": "期待ステータス",
   "aliases": [
    "status",
    "expected_status",
    "ステータス"
   ],
   "required": false,
   "width": 13,
   "hint": "200 / 200,201",
   "desc": "期待する HTTP ステータス（カンマ区切りで複数可）。4xx / 5xx を書くと異常系テストとして扱う",
   "example": "200"
  },
  {
   "key": "expected_text",
   "header": "応答検証文字列",
   "aliases": [
    "expected_text",
    "assert",
    "検証文字列"
   ],
   "required": false,
   "width": 22,
   "hint": "文字列（改行区切り）",
   "desc": "応答本文に含まれるべき文字列。複数書くとすべて含む場合に成功",
   "example": "\"result\":\"OK\""
  },
  {
   "key": "max_time_ms",
   "header": "応答時間上限(ms)",
   "aliases": [
    "max_time",
    "max_time_ms",
    "応答時間上限"
   ],
   "required": false,
   "width": 13,
   "hint": "数値（ms）",
   "desc": "この時間を超えたら失敗（Duration Assertion）",
   "example": "3000"
  },
  {
   "key": "regex_extract",
   "header": "正規表現抽出",
   "aliases": [
    "regex",
    "regex_extract",
    "正規表現"
   ],
   "required": false,
   "width": 34,
   "hint": "変数名=正規表現（改行区切り）",
   "desc": "応答本文から ( ) 内を抽出して変数に設定。見つからない場合は NOT_FOUND",
   "example": "TOKEN=\"token\"\\s*:\\s*\"([^\"]+)\""
  },
  {
   "key": "json_extract",
   "header": "JSON抽出",
   "aliases": [
    "json",
    "json_extract",
    "jsonpath"
   ],
   "required": false,
   "width": 28,
   "hint": "変数名=JSONPath（改行区切り）",
   "desc": "応答 JSON から値を抽出して変数に設定。見つからない場合は NOT_FOUND",
   "example": "USER_ID=$.data.id"
  },
  {
   "key": "pre_script",
   "header": "前処理スクリプト",
   "aliases": [
    "pre",
    "pre_script",
    "前処理"
   ],
   "required": false,
   "width": 36,
   "hint": "Groovy",
   "desc": "API 呼び出しの直前に実行する Groovy（JSR223 PreProcessor）",
   "example": "vars.put('X', '1')"
  },
  {
   "key": "post_script",
   "header": "後処理スクリプト",
   "aliases": [
    "post",
    "post_script",
    "後処理"
   ],
   "required": false,
   "width": 36,
   "hint": "Groovy",
   "desc": "API 呼び出しの直後に実行する Groovy（JSR223 PostProcessor）。抽出の後に実行",
   "example": "log.info(prev.getResponseCode())"
  },
  {
   "key": "comment",
   "header": "備考",
   "aliases": [
    "comment",
    "コメント",
    "メモ"
   ],
   "required": false,
   "width": 24,
   "hint": "任意",
   "desc": "JMeter のコメント欄に設定",
   "example": ""
  }
 ],
 "sample_apis": [
  {
   "no": "1",
   "enabled": "Y",
   "name": "ログイン(POST JSON)",
   "method": "POST",
   "url": "https://postman-echo.com/post",
   "query": "",
   "headers": "Content-Type: application/json\nX-Correlation-ID: ${REQUEST_ID}\nX-Signature: ${SIGNATURE}",
   "body_type": "json",
   "body": "{\n  \"userId\": \"${USER_ID}\",\n  \"password\": \"${PASSWORD}\",\n  \"requestNo\": \"${SEQ}\",\n  \"timestamp\": \"${TIMESTAMP}\"\n}",
   "wait_ms": "500",
   "expected_status": "200",
   "expected_text": "${USER_ID}",
   "max_time_ms": "5000",
   "regex_extract": "ECHO_URL=\"url\"\\s*:\\s*\"([^\"]+)\"",
   "json_extract": "AUTH_TOKEN=$.json.userId\nLOGIN_SEQ=$.json.requestNo",
   "pre_script": "// 署名の生成例: USER_ID + TIMESTAMP の SHA-256（16進）を変数 SIGNATURE に設定\nString src = (vars.get('USER_ID') ?: '') + (vars.get('TIMESTAMP') ?: '')\ndef digest = java.security.MessageDigest.getInstance('SHA-256').digest(src.getBytes('UTF-8'))\nvars.put('SIGNATURE', digest.encodeHex().toString())",
   "post_script": "// 応答 JSON の解析例: 値を取り出して次の API で使う変数に設定\ntry {\n    def json = new groovy.json.JsonSlurper().parseText(prev.getResponseDataAsString())\n    vars.put('LOGIN_USER', String.valueOf(json?.json?.userId))\n} catch (Exception e) {\n    log.warn('ログイン応答の解析に失敗しました: ' + e.getMessage())\n}",
   "comment": "JSON を POST する例（postman-echo.com は送信内容をそのまま返す）"
  },
  {
   "no": "2",
   "enabled": "Y",
   "name": "商品検索(GET)",
   "method": "GET",
   "url": "https://postman-echo.com/get",
   "query": "keyword=${KEYWORD}\nitemCode=${ITEM_CODE}\nrand=${RANDOM_NO}\npage=1",
   "headers": "Authorization: Bearer ${AUTH_TOKEN}",
   "body_type": "",
   "body": "",
   "wait_ms": "1000",
   "expected_status": "200",
   "expected_text": "${ITEM_CODE}",
   "max_time_ms": "5000",
   "regex_extract": "",
   "json_extract": "SEARCH_KEYWORD=$.args.keyword",
   "pre_script": "",
   "post_script": "",
   "comment": "GET とクエリパラメータの例（CSV・Excel・乱数の変数を使用）"
  },
  {
   "no": "3",
   "enabled": "Y",
   "name": "注文登録(POST form)",
   "method": "POST",
   "url": "https://postman-echo.com/post?source=jmeter",
   "query": "channel=web",
   "headers": "Authorization: Bearer ${AUTH_TOKEN}",
   "body_type": "form",
   "body": "itemCode=${ITEM_CODE}\nitemName=${ITEM_NAME}\nquantity=${QUANTITY}\norderNo=ORD${SEQ}",
   "wait_ms": "1000",
   "expected_status": "200,201",
   "expected_text": "ORD${SEQ}",
   "max_time_ms": "5000",
   "regex_extract": "ORDER_NO=\"orderNo\"\\s*:\\s*\"([^\"]+)\"",
   "json_extract": "",
   "pre_script": "",
   "post_script": "// 抽出した値の利用例（jmeter.log に出力）\nlog.info('注文番号=' + vars.get('ORDER_NO') + ' / ステータス=' + prev.getResponseCode())",
   "comment": "form 形式の POST の例（POST のクエリパラメータは URL に付与）"
  },
  {
   "no": "4",
   "enabled": "Y",
   "name": "クッキー確認(GET)",
   "method": "GET",
   "url": "https://postman-echo.com/cookies/set",
   "query": "session_id=S${SEQ}",
   "headers": "",
   "body_type": "",
   "body": "",
   "wait_ms": "500",
   "expected_status": "200",
   "expected_text": "session_id",
   "max_time_ms": "5000",
   "regex_extract": "",
   "json_extract": "SESSION_COOKIE=$.cookies.session_id",
   "pre_script": "",
   "post_script": "",
   "comment": "Set-Cookie とリダイレクト追従の例（クッキー管理を使用）"
  }
 ],
 "sample_csv": {
  "header": [
   "USER_ID",
   "PASSWORD",
   "KEYWORD"
  ],
  "rows": [
   [
    "user001",
    "Passw0rd!01",
    "テスト"
   ],
   [
    "user002",
    "Passw0rd!02",
    "JMeter"
   ],
   [
    "user003",
    "Passw0rd!03",
    "負荷試験"
   ],
   [
    "user004",
    "Passw0rd!04",
    "API"
   ],
   [
    "user005",
    "Passw0rd!05",
    "サンプル"
   ]
  ]
 },
 "sample_excel": {
  "sheet": "data",
  "header": [
   "ITEM_CODE",
   "ITEM_NAME",
   "QUANTITY"
  ],
  "rows": [
   [
    "A001",
    "りんご",
    "1"
   ],
   [
    "A002",
    "みかん",
    "2"
   ],
   [
    "B001",
    "ぶどう",
    "3"
   ],
   [
    "B002",
    "もも",
    "1"
   ],
   [
    "C001",
    "メロン",
    "5"
   ]
  ]
 },
 "names": {
  "udv": "[必須] 共通変数 (User Defined Variables)",
  "http_defaults": "[任意] HTTPリクエスト初期値 (HTTP Request Defaults)",
  "common_headers": "[任意] 共通ヘッダ (HTTP Header Manager)",
  "cookie": "[任意] クッキー管理 (HTTP Cookie Manager)",
  "setup_tg": "[任意] 事前処理 (setUp Thread Group)",
  "setup_start": "[任意] テスト開始処理 (JSR223 Sampler)",
  "excel_loader": "[任意] Excelデータ読込 (JSR223 Sampler)",
  "main_tg": "[必須] メイン処理 (Thread Group)",
  "counter": "[任意] 連番 @@NAME@@ (Counter)",
  "random": "[任意] 乱数 @@NAME@@ (Random Variable)",
  "csv": "[任意] CSVデータ読込 (CSV Data Set Config)",
  "excel_row": "[任意] Excelデータ行取得 (JSR223 PreProcessor)",
  "access_token": "[任意] アクセストークン読込 (JSR223 PreProcessor)",
  "common_pre": "[任意] 共通前処理 (JSR223 PreProcessor)",
  "common_post": "[任意] 共通後処理 (JSR223 PostProcessor)",
  "throughput": "[任意] スループット制御 (Constant Throughput Timer)",
  "api_headers": "[任意] @@ID@@ ヘッダ (HTTP Header Manager)",
  "api_timer": "[任意] @@ID@@ 待機 @@MS@@ms (Constant Timer)",
  "api_pre": "[任意] @@ID@@ 前処理 (JSR223 PreProcessor)",
  "api_regex": "[任意] @@ID@@ 正規表現抽出 @@VAR@@ (Regular Expression Extractor)",
  "api_json": "[任意] @@ID@@ JSON抽出 @@VAR@@ (JSON Extractor)",
  "api_post": "[任意] @@ID@@ 後処理 (JSR223 PostProcessor)",
  "api_status": "[任意] @@ID@@ ステータス検証 @@CODES@@ (Response Assertion)",
  "api_text": "[任意] @@ID@@ 応答本文検証 (Response Assertion)",
  "api_duration": "[任意] @@ID@@ 応答時間検証 @@MS@@ms (Duration Assertion)",
  "debug": "[任意] 変数確認 (Debug Sampler)",
  "teardown_tg": "[任意] 事後処理 (tearDown Thread Group)",
  "teardown_end": "[任意] テスト終了処理 (JSR223 Sampler)",
  "view_tree": "[任意] 結果ツリー (View Results Tree)",
  "summary": "[任意] 集計レポート (Summary Report)",
  "aggregate": "[任意] 統計レポート (Aggregate Report)",
  "jtl": "[任意] JTL出力 全件 (Simple Data Writer)",
  "error_xml": "[任意] エラー詳細出力 XML (Simple Data Writer)",
  "result_log": "[任意] 結果ログ出力 (JSR223 Listener)"
 },
 "comments": {
  "test_plan": "jmeter_api_tool v1.0.0 で生成（@@DATE@@ / 入力: @@INPUT@@）\n名前が [必須] の要素は削除しないでください。[任意] の要素は不要であれば削除できます（各要素のコメント参照）。\n実行時の変更例: -JTHREADS=10 -JLOOPS=100 -JHOST=example.com",
  "udv": "[必須] 接続先と共通変数。値は ${__P(名前,既定値)} 形式のため、実行時に -J名前=値 で変更できます。\n各 API（HTTP リクエスト）が ${PROTOCOL} ${HOST} ${PORT} を参照しているため削除しないでください。",
  "udv_protocol": "プロトコル（http / https）。-J@@VAR@@=値 で変更",
  "udv_host": "ホスト名。-J@@VAR@@=値 で変更",
  "udv_port": "ポート番号（空欄 = http は 80 / https は 443）。-J@@VAR@@=値 で変更",
  "udv_user": "共通変数。-J@@VAR@@=値 で変更",
  "http_defaults": "[任意] 全 HTTP リクエストの初期値（文字コード・タイムアウト）。\nタイムアウトは -JCONNECT_TIMEOUT / -JRESPONSE_TIMEOUT（ミリ秒）で変更できます。削除するとタイムアウトは無制限になります。",
  "common_headers": "[任意] 全 API に付けるヘッダ。API 固有のヘッダは各 API の配下にあります（同じ名前の場合は API 側が優先）。",
  "cookie": "[任意] クッキーの自動管理（Set-Cookie を保存し、以降のリクエストで送信）。反復ごとにクリア = @@CLEAR@@。\nクッキーを使わない API だけの場合は削除できます。",
  "setup_tg": "[任意] テスト開始前に 1 回だけ実行されます。配下の処理が不要であれば、このスレッドグループごと削除できます。",
  "setup_start": "[任意] 実行条件（パラメータ欄）を jmeter.log に出力し、結果ログ用の集計値をリセットします。",
  "excel_loader": "[任意] Excel（.xlsx）を読み込みます（Apache POI 不要）。1 行目 = 変数名、2 行目以降 = データ。\nパラメータ: ファイルパス|シート名（-JEXCEL_FILE / -JEXCEL_SHEET で変更可。相対パスは JMX のフォルダ基準）。\n削除する場合は、メイン処理の「Excelデータ行取得」も削除してください。",
  "main_tg": "[必須] API を実行するスレッドグループ。\nスレッド数 = -JTHREADS / ランプアップ秒 = -JRAMP_UP / ループ回数 = -JLOOPS（-1 = 無限）@@DURATION@@",
  "main_tg_duration": " / 実行時間（秒）= -JDURATION",
  "counter": "[任意] 反復（ループ）ごとに連番を ${@@NAME@@} に設定します（@@SCOPE@@）。",
  "counter_scope_global": "全スレッドで一意",
  "counter_scope_thread": "スレッドごとに採番",
  "random": "[任意] 反復ごとに @@MIN@@ ～ @@MAX@@ の乱数を ${@@NAME@@} に設定します。",
  "csv": "[任意] CSV ファイルを 1 反復につき 1 行読み込み、列を変数に設定します（変数: @@VARS@@）。\nファイルは -JCSV_FILE で変更できます（相対パスは JMX のフォルダ基準）。ファイルが無いとスレッドが停止するため、使わない場合は削除してください。",
  "excel_row": "[任意] 「Excelデータ読込」で読み込んだデータから、反復ごとに次の 1 行を取り出して変数に設定します。",
  "access_token": "[任意] 全 API の呼び出し直前に、指定したフォルダ配下のファイルからアクセストークンを読み込み ${ACCESS_TOKEN} に設定します（ファイルが更新されると読み直します）。共通ヘッダの Authorization がこの変数を参照します。\nパラメータ: フォルダ|ファイル名（-JACCESS_TOKEN_DIR / -JACCESS_TOKEN_FILE で変更可。相対パスは JMX のフォルダ基準）。\nファイル名が *.token のように * を含む場合は、一致するファイルのうち更新日時が最も新しいファイルを使います。フォルダは ; で区切って複数書けます（存在する最初のフォルダを使用。例: C:\\tokens;/opt/tokens）。",
  "common_pre": "[任意] 全 API の呼び出し直前に毎回実行（Groovy）。${REQUEST_ID} ${TIMESTAMP} ${EPOCH_MS} を設定します。",
  "common_post": "[任意] 全 API の呼び出し直後に毎回実行（Groovy）。${LAST_STATUS} ${LAST_ELAPSED_MS} を設定します。",
  "throughput": "[任意] スループット制御。メイン処理全体で 1 分あたり -JTHROUGHPUT 件（既定 @@VALUE@@）になるよう待機を調整します。\n無効の場合は、使うときに右クリック → Enable で有効にしてください。",
  "sampler": "@@ID@@: @@METHOD@@ @@URL@@",
  "api_headers": "[任意] この API だけに付けるヘッダ。",
  "api_timer": "[任意] この API の呼び出し前に @@MS@@ ミリ秒待機します。",
  "api_pre": "[任意] この API の呼び出し直前に実行（Groovy）。",
  "api_post": "[任意] この API の呼び出し直後に実行（Groovy）。正規表現・JSON 抽出の後に実行されます。",
  "api_regex": "[任意] 応答本文から正規表現で抽出し ${@@VAR@@} に設定します（見つからない場合は NOT_FOUND）。",
  "api_json": "[任意] 応答 JSON から @@PATH@@ を抽出し ${@@VAR@@} に設定します（見つからない場合は NOT_FOUND）。",
  "api_status": "[任意] HTTP ステータスが @@CODES@@ であることを検証します。",
  "api_status_ignore": "（4xx / 5xx を期待するため「Ignore Status」を有効にしています）",
  "api_text": "[任意] 応答本文に指定した文字列がすべて含まれることを検証します。",
  "api_duration": "[任意] 応答時間が @@MS@@ ミリ秒以内であることを検証します。",
  "debug": "[任意] デバッグ用。全変数を結果ツリーに表示します。無効の場合は、使うときに右クリック → Enable で有効にしてください。",
  "teardown_tg": "[任意] 全スレッドの終了後に 1 回だけ実行されます。",
  "teardown_end": "[任意] 合計・成功・失敗件数を jmeter.log と結果ログに出力し、結果ログを閉じます。",
  "view_tree": "[任意] GUI でリクエスト / レスポンスを確認するためのリスナー。CLI 実行時は不要です（削除・無効化可）。",
  "summary": "[任意] GUI で集計結果を確認するためのリスナー。",
  "aggregate": "[任意] GUI で統計（パーセンタイル）を確認するためのリスナー。",
  "jtl": "[任意] 全リクエストの結果を JTL（CSV）に出力します。出力先: -JRESULT_DIR（既定 @@DIR@@、~/ は JMX のフォルダ）。\nrun コマンドや -l 指定時は同じ内容が result.jtl にも出力されるため、大量実行時は削除・無効化してかまいません。",
  "error_xml": "[任意] 失敗したリクエストだけを、リクエスト・レスポンス全文付きで XML に出力します（結果ツリーで読み込み可能）。",
  "result_log": "[任意] 1 リクエスト 1 行の結果ログ（api_result.log）を出力します。パラメータ: 出力フォルダ|本文出力(none/error/all)|最大文字数。\n大量実行時はファイル書き込みが負荷になるため、削除・無効化を推奨します。"
 },
 "scripts": {
  "excel_loader": "// ============================================================\n// [任意] Excelデータ読込（setUp Thread Group / JSR223 Sampler）\n//  テスト開始前に1回だけ実行され、Excel(.xlsx)の内容をメモリに読み込みます。\n//  ・Apache POI などの追加ライブラリは不要です（xlsx を ZIP + XML として直接解析）。\n//  ・1行目 = 列名（そのまま変数名になります）、2行目以降 = データ。空行は無視します。\n//  ・読み込んだデータは props 'EXCEL_DATA' に保存され、メイン処理の\n//    「Excelデータ行取得」が反復ごとに1行ずつ変数に設定します。\n//  ・パラメータ欄: ファイルパス|シート名（シート名が空の場合は先頭シート）\n//    相対パスは JMX ファイルのフォルダが基準です。-JEXCEL_FILE / -JEXCEL_SHEET で変更できます。\n//  ・数値・日付はセルに保存されている値（表示書式の適用前）になります。\n//    表示どおりの値を使いたい列は、Excel 側でセルの書式を「文字列」にしてください。\n// ============================================================\nimport java.util.concurrent.atomic.AtomicLong\nimport java.util.zip.ZipFile\nimport javax.xml.parsers.DocumentBuilderFactory\nimport org.apache.jmeter.services.FileServer\n\nSampleResult.setIgnore()   // この処理はレポート・JTL に含めない\n\nString[] params = (Parameters ?: '').split('\\\\|', -1)\nString path = params.length > 0 ? params[0].trim() : ''\nString sheetName = params.length > 1 ? params[1].trim() : ''\nprops.remove('EXCEL_DATA')\nprops.put('EXCEL_ROW_COUNTER', new AtomicLong(0))\nif (path.isEmpty()) {\n    log.warn('[Excel] ファイルパスが指定されていません')\n    return\n}\nFile file = new File(path)\nif (!file.isAbsolute()) {\n    file = new File(FileServer.getFileServer().getBaseDir(), path)\n}\nif (!file.isFile()) {\n    log.warn('[Excel] ファイルが見つかりません: ' + file.getAbsolutePath())\n    return\n}\n\ndef factory = DocumentBuilderFactory.newInstance()\nfactory.setNamespaceAware(true)\nfactory.setFeature('http://apache.org/xml/features/disallow-doctype-decl', true)\ndef builder = factory.newDocumentBuilder()\n\n// 子要素をローカル名で取得する\ndef children = { node, String name ->\n    def list = []\n    if (node == null) {\n        return list\n    }\n    def nodes = node.getChildNodes()\n    for (int i = 0; i < nodes.getLength(); i++) {\n        def n = nodes.item(i)\n        if (n.getNodeType() == org.w3c.dom.Node.ELEMENT_NODE && n.getLocalName() == name) {\n            list.add(n)\n        }\n    }\n    return list\n}\n// <t> のテキストを連結する（ふりがな rPh は除外）\ndef textOf\ntextOf = { node ->\n    StringBuilder sb = new StringBuilder()\n    def nodes = node.getChildNodes()\n    for (int i = 0; i < nodes.getLength(); i++) {\n        def n = nodes.item(i)\n        if (n.getNodeType() != org.w3c.dom.Node.ELEMENT_NODE || n.getLocalName() == 'rPh') {\n            continue\n        }\n        sb.append(n.getLocalName() == 't' ? n.getTextContent() : textOf(n))\n    }\n    return sb.toString()\n}\n// セル参照 \"AB12\" → 列番号（0 始まり）\ndef columnIndex = { String ref ->\n    int n = 0\n    for (char ch : ref.toCharArray()) {\n        if (!Character.isLetter(ch)) {\n            break\n        }\n        n = n * 26 + (Character.toUpperCase(ch) - ('A' as char) + 1)\n    }\n    return n - 1\n}\n\ndef zip = new ZipFile(file)\ntry {\n    def parse = { String entryName ->\n        def entry = zip.getEntry(entryName)\n        if (entry == null) {\n            return null\n        }\n        def input = zip.getInputStream(entry)\n        try {\n            return builder.parse(input).getDocumentElement()\n        } finally {\n            input.close()\n        }\n    }\n    // 共有文字列\n    List<String> shared = []\n    children(parse('xl/sharedStrings.xml'), 'si').each { shared.add(textOf(it)) }\n    // シートの特定（workbook.xml → workbook.xml.rels）\n    Map<String, String> rels = [:]\n    children(parse('xl/_rels/workbook.xml.rels'), 'Relationship').each {\n        rels.put(it.getAttribute('Id'), it.getAttribute('Target'))\n    }\n    String target = null\n    String foundName = null\n    def sheetsNode = children(parse('xl/workbook.xml'), 'sheets')\n    for (def sheet : children(sheetsNode ? sheetsNode[0] : null, 'sheet')) {\n        if (sheetName.isEmpty() || sheet.getAttribute('name') == sheetName) {\n            def attrs = sheet.getAttributes()\n            for (int i = 0; i < attrs.getLength(); i++) {\n                def a = attrs.item(i)\n                if (a.getLocalName() == 'id' && a.getNamespaceURI() != null) {\n                    target = rels.get(a.getNodeValue())\n                }\n            }\n            foundName = sheet.getAttribute('name')\n            break\n        }\n    }\n    if (target == null) {\n        log.warn('[Excel] シートが見つかりません: ' + (sheetName ?: '(先頭シート)') + ' / ' + file.getName())\n        return\n    }\n    def sheetRoot = parse(target.startsWith('/') ? target.substring(1) : 'xl/' + target)\n    def sheetData = children(sheetRoot, 'sheetData')\n    // セルの読み取り\n    List<List<String>> table = []\n    for (def row : children(sheetData ? sheetData[0] : null, 'row')) {\n        Map<Integer, String> cells = [:]\n        int next = 0\n        for (def c : children(row, 'c')) {\n            String ref = c.getAttribute('r')\n            int col = ref ? columnIndex(ref) : next\n            next = col + 1\n            String type = c.getAttribute('t')\n            String value = ''\n            if (type == 'inlineStr') {\n                def inline = children(c, 'is')\n                value = inline ? textOf(inline[0]) : ''\n            } else {\n                def v = children(c, 'v')\n                String raw = v ? v[0].getTextContent() : ''\n                if (type == 's') {\n                    value = raw ? shared[Integer.parseInt(raw.trim())] : ''\n                } else if (type == 'b') {\n                    value = raw.trim() == '1' ? 'TRUE' : 'FALSE'\n                } else {\n                    value = raw\n                }\n            }\n            cells.put(col, value)\n        }\n        int size = cells.isEmpty() ? 0 : cells.keySet().max() + 1\n        List<String> values = []\n        for (int i = 0; i < size; i++) {\n            values.add(cells.containsKey(i) ? cells.get(i) : '')\n        }\n        if (values.any { it.trim() }) {\n            table.add(values)\n        }\n    }\n    if (table.isEmpty()) {\n        log.warn('[Excel] データがありません: ' + file.getName() + ' / ' + foundName)\n        return\n    }\n    List<String> headers = table[0].collect { it.trim() }\n    List<Map<String, String>> rows = []\n    for (int r = 1; r < table.size(); r++) {\n        Map<String, String> m = new LinkedHashMap<>()\n        for (int i = 0; i < headers.size(); i++) {\n            if (headers[i]) {\n                m.put(headers[i], i < table[r].size() ? table[r][i] : '')\n            }\n        }\n        rows.add(m)\n    }\n    props.put('EXCEL_DATA', Collections.unmodifiableList(rows))\n    log.info('[Excel] 読込完了: ' + file.getAbsolutePath() + ' / シート=' + foundName\n            + ' / 列=' + headers.findAll { it } + ' / データ=' + rows.size() + '行')\n} catch (Exception e) {\n    log.error('[Excel] 読込エラー: ' + file.getAbsolutePath(), e)\n} finally {\n    zip.close()\n}\n",
  "excel_row": "// ============================================================\n// [任意] Excelデータ行取得（JSR223 PreProcessor）\n//  「Excelデータ読込」（setUp）で読み込んだデータから、反復（ループ）ごとに\n//  次の1行を取り出し、1行目の列名を変数名として設定します。\n//    例: 列名 ITEM_CODE → 変数 ITEM_CODE、使用中の行番号 → 変数 EXCEL_ROW_NO\n//  ・全スレッドで共有し、上から順に払い出します。最終行の次は先頭に戻ります。\n//  ・同じ反復内の2つ目以降の API では同じ行の値を使います。\n// ============================================================\ndef rows = props.get('EXCEL_DATA')\nif (rows == null || rows.isEmpty()) {\n    return\n}\nString iteration = String.valueOf(vars.getIteration())\nif (iteration == vars.get('EXCEL_ROW_ITERATION')) {\n    return\n}\nvars.put('EXCEL_ROW_ITERATION', iteration)\ndef counter = props.get('EXCEL_ROW_COUNTER')\nlong n = counter != null ? counter.getAndIncrement() : 0L\nint index = (int) (n % rows.size())\nrows.get(index).each { k, v -> vars.put(k, v) }\nvars.put('EXCEL_ROW_NO', String.valueOf(index + 1))\n",
  "access_token": "// ============================================================\n// [任意] アクセストークン読込（JSR223 PreProcessor）\n//  メイン処理内のすべての API の呼び出し直前に毎回実行されます。\n//  指定したフォルダ配下のファイルからアクセストークンを読み込み、変数 ACCESS_TOKEN に設定します。\n//  共通ヘッダ（HTTP Header Manager）の Authorization（Bearer + 変数 ACCESS_TOKEN）で送信されます。\n//  ・パラメータ欄: フォルダ|ファイル名\n//    相対パスは JMX ファイルのフォルダ基準です。-JACCESS_TOKEN_DIR / -JACCESS_TOKEN_FILE で変更できます。\n//    区切りは \\ と / のどちらでもかまいません（Windows / Linux 共通）。フォルダを ; で区切って複数書くと、\n//    最初に存在するフォルダを使います（例: C:\\tokens;/opt/tokens）。\n//  ・ファイル名に * ? を含む場合（例: *.token）は、一致するファイルのうち更新日時が最も新しいファイルを使います\n//    （フォルダの確認は 1 秒に 1 回。新しいファイルを置くと自動で切り替わります）。\n//  ・ファイルの内容の前後の空白・改行（BOM）は除きます。\n//  ・ファイルの更新日時・サイズが変わったときだけ読み直します（実行中にトークンを差し替えできます）。\n// ============================================================\nimport java.nio.file.FileSystems\nimport org.apache.jmeter.services.FileServer\n\nString[] params = (Parameters ?: '').split('\\\\|', -1)\nString dir = params.length > 0 ? params[0].trim() : ''\nString name = params.length > 1 ? params[1].trim().replace('\\\\', '/') : ''\n// パスの解決: \\ は / として扱い、相対パス（~/ で始まるものを含む）は JMX のフォルダ基準\ndef resolve = { String path ->\n    String p = path.trim().replace('\\\\', '/')\n    File f = new File(p.startsWith('~/') ? p.substring(2) : p)\n    return f.isAbsolute() ? f : new File(FileServer.getFileServer().getBaseDir(), f.getPath())\n}\n// フォルダを ; で区切って複数書いた場合は、最初に存在するフォルダを使う\nList<File> dirs = dir.split(';').toList().findAll { it.trim() }.collect { resolve(it) }\nFile folder = dirs.find { it.isDirectory() } ?: (dirs ? dirs[0] : null)\nFile file = folder == null ? resolve(name) : new File(folder, name)\n// ファイル名に * ? がある場合: 一致するファイルのうち更新日時が最も新しいもの（同時刻は名前が後のもの）。\n// フォルダの確認は全スレッドで共有し、1 秒に 1 回だけ行う\nif (name.contains('*') || name.contains('?')) {\n    def latest = props.get('ACCESS_TOKEN_LATEST')\n    long now = System.currentTimeMillis()\n    if (latest == null || latest[0] != file.getPath() || now - latest[1] >= 1000) {\n        def matcher = FileSystems.getDefault().getPathMatcher('glob:' + file.getName())\n        File found = null\n        for (File f : (file.getParentFile()?.listFiles() ?: new File[0])) {\n            if (f.isFile() && matcher.matches(f.toPath().getFileName()) && (found == null || f.lastModified() > found.lastModified()\n                    || (f.lastModified() == found.lastModified() && f.getName() > found.getName()))) {\n                found = f\n            }\n        }\n        latest = [file.getPath(), now, found ?: file]\n        props.put('ACCESS_TOKEN_LATEST', latest)\n    }\n    file = latest[2]\n}\n// 読み込んだトークンは全スレッドで共有し、ファイルが変わったときだけ読み直す\nString stamp = file.getAbsolutePath() + '|' + file.lastModified() + '|' + file.length()\ndef cache = props.get('ACCESS_TOKEN_CACHE')\nif (cache == null || cache[0] != stamp) {\n    String token = ''\n    if (file.isFile()) {\n        token = file.getText('UTF-8').replace('﻿', '').trim()\n        if (token.isEmpty()) {\n            log.warn('[AccessToken] ファイルが空です: ' + file.getAbsolutePath())\n        } else {\n            log.info('[AccessToken] 読み込みました: ' + file.getAbsolutePath())\n        }\n    } else {\n        log.warn('[AccessToken] ファイルが見つかりません: ' + file.getAbsolutePath())\n    }\n    cache = [stamp, token]\n    props.put('ACCESS_TOKEN_CACHE', cache)\n}\nvars.put('ACCESS_TOKEN', cache[1])",
  "common_pre": "// ============================================================\n// [任意] 共通前処理（JSR223 PreProcessor）\n//  メイン処理内のすべての API の呼び出し直前に毎回実行されます。\n//  API 呼び出し前にパラメータを生成・加工する処理をここに書きます。\n//  ※ JSR223 内では JMeter の変数参照（ドル記号＋波括弧）は使わず、\n//     vars.get('変数名') / vars.put('変数名', 値) を使います。\n// ============================================================\nimport java.time.ZonedDateTime\nimport java.time.format.DateTimeFormatter\n\n// リクエストごとの一意な ID とタイムスタンプ → 変数 REQUEST_ID / TIMESTAMP / EPOCH_MS\nvars.put('REQUEST_ID', UUID.randomUUID().toString())\nvars.put('TIMESTAMP', ZonedDateTime.now().format(DateTimeFormatter.ISO_OFFSET_DATE_TIME))\nvars.put('EPOCH_MS', String.valueOf(System.currentTimeMillis()))\n\n// 加工の例（必要に応じてコメントを外して使用）\n// String userId = vars.get('USER_ID') ?: ''\n// vars.put('USER_ID_UPPER', userId.toUpperCase())\n// vars.put('TODAY', new Date().format('yyyyMMdd'))\n",
  "common_post": "// ============================================================\n// [任意] 共通後処理（JSR223 PostProcessor）\n//  メイン処理内のすべての API の呼び出し直後に毎回実行されます。\n//  応答からパラメータを取得し、次の API で使う変数に設定する処理をここに書きます。\n//  prev = 直前の API の実行結果（SampleResult）\n//  ※ 検証（Assertion）より前に実行されるため、prev.isSuccessful() は HTTP レベルの成否です。\n// ============================================================\nvars.put('LAST_STATUS', prev.getResponseCode())\nvars.put('LAST_ELAPSED_MS', String.valueOf(prev.getTime()))\n\n// 応答 JSON から値を取り出す例（必要に応じてコメントを外して使用）\n// if ((prev.getContentType() ?: '').contains('json')) {\n//     def json = new groovy.json.JsonSlurper().parseText(prev.getResponseDataAsString())\n//     vars.put('RESULT_CODE', String.valueOf(json?.resultCode))\n// }\n// 応答ヘッダから値を取り出す例\n// def m = prev.getResponseHeaders() =~ /(?im)^X-Request-Id:\\s*(.+)$/\n// if (m.find()) {\n//     vars.put('RESPONSE_REQUEST_ID', m.group(1).trim())\n// }\n",
  "setup_start": "// ============================================================\n// [任意] テスト開始処理（setUp Thread Group / JSR223 Sampler）\n//  テスト開始時に1回だけ実行されます。\n//  ・実行条件（パラメータ欄）を jmeter.log に出力します。\n//  ・前回の実行（GUI で連続実行した場合）の集計値・結果ログをリセットします。\n//  削除しても他の要素の動作には影響しません。\n// ============================================================\nSampleResult.setIgnore()   // この処理はレポート・JTL に含めない\n\ndef oldWriter = props.remove('API_RESULT_LOG_WRITER')\nif (oldWriter != null) {\n    try {\n        oldWriter.close()\n    } catch (Exception ignored) {\n    }\n}\n['API_RESULT_TOTAL', 'API_RESULT_OK', 'API_RESULT_NG'].each { props.remove(it) }\nprops.put('API_TEST_START_MS', String.valueOf(System.currentTimeMillis()))\nlog.info('==== テスト開始: ' + (Parameters ?: ''))\n",
  "teardown_end": "// ============================================================\n// [任意] テスト終了処理（tearDown Thread Group / JSR223 Sampler）\n//  すべてのスレッドの終了後に1回だけ実行されます。\n//  ・合計 / 成功 / 失敗件数と経過時間を jmeter.log と結果ログに出力します。\n//    （件数は「結果ログ出力」が有効な場合のみ集計されます）\n//  ・結果ログファイルを閉じ、読み込んだ Excel データを解放します。\n// ============================================================\nSampleResult.setIgnore()   // この処理はレポート・JTL に含めない\n\ndef total = props.get('API_RESULT_TOTAL')\nString counts = total == null ? '件数=集計なし（結果ログ出力が無効）'\n        : '合計=' + total.get() + ' 成功=' + (props.get('API_RESULT_OK')?.get() ?: 0) + ' 失敗=' + (props.get('API_RESULT_NG')?.get() ?: 0)\nString start = props.get('API_TEST_START_MS')\nString elapsed = start ? ' 経過秒=' + (System.currentTimeMillis() - Long.parseLong(start)).intdiv(1000) : ''\nString summary = '==== テスト終了: ' + counts + elapsed\nlog.info(summary)\n\ndef writer = props.remove('API_RESULT_LOG_WRITER')\nif (writer != null) {\n    synchronized (writer) {\n        writer.println('# ' + summary)\n        writer.close()\n    }\n}\nprops.remove('EXCEL_DATA')\nprops.remove('EXCEL_ROW_COUNTER')\n",
  "result_listener": "// ============================================================\n// [任意] 結果ログ出力（JSR223 Listener）\n//  すべての API の実行結果を 1リクエスト1行（タブ区切り）で api_result.log に出力します。\n//  パラメータ欄: 出力フォルダ|本文出力|最大文字数\n//   ・出力フォルダ: -JRESULT_DIR で変更（~/ は JMX のフォルダ。run コマンドは自動設定）\n//   ・本文出力  : none = 出力しない / error = 失敗時のみ / all = 全件（-JLOG_RESPONSE）\n//   ・最大文字数: リクエスト・レスポンス本文の最大文字数（-JLOG_MAX_CHARS）\n//  ※ 大量実行時はファイル書き込みが負荷になるため、無効化または none を推奨します。\n// ============================================================\nimport java.text.SimpleDateFormat\nimport java.util.concurrent.atomic.AtomicLong\nimport java.util.function.Function\nimport org.apache.jmeter.services.FileServer\n\ndef res = sampleResult\nif (res == null) {\n    return\n}\nString[] p = (Parameters ?: '').split('\\\\|', -1)\nString dirText = (p.length > 0 && p[0].trim()) ? p[0].trim() : '~/result'\nString mode = (p.length > 1 && p[1].trim()) ? p[1].trim().toLowerCase() : 'error'\nint maxChars = (p.length > 2 && p[2].trim().isInteger()) ? p[2].trim().toInteger() : 1000\n\n// 件数の集計（テスト終了処理で出力）\ndef count = { String key -> props.computeIfAbsent(key, { k -> new AtomicLong(0) } as Function).incrementAndGet() }\ncount('API_RESULT_TOTAL')\ncount(res.isSuccessful() ? 'API_RESULT_OK' : 'API_RESULT_NG')\n\n// 出力ファイル（最初の1件で作成し、全スレッドで共有）\ndef writer = props.computeIfAbsent('API_RESULT_LOG_WRITER', { k ->\n    File dir = dirText.startsWith('~/') ? new File(FileServer.getFileServer().getBaseDir(), dirText.substring(2)) : new File(dirText)\n    dir.mkdirs()\n    File file = new File(dir, 'api_result.log')\n    boolean isNew = !file.exists() || file.length() == 0\n    def w = new PrintWriter(new OutputStreamWriter(new FileOutputStream(file, true), 'UTF-8'), true)\n    if (isNew) {\n        w.println(['日時', '結果', 'ステータス', '応答時間ms', '受信バイト', 'ラベル', 'リクエスト', 'スレッド', 'メッセージ'].join('\\t'))\n    }\n    log.info('[結果ログ] 出力先: ' + file.getAbsolutePath())\n    return w\n} as Function)\n\n// 改行・タブを除去し、最大文字数で切り詰める\ndef oneLine = { String s, int limit ->\n    if (s == null) {\n        return ''\n    }\n    String t = s.replace('\\r', '').replace('\\n', '\\\\n').replace('\\t', ' ')\n    return (limit > 0 && t.length() > limit) ? t.substring(0, limit) + '...(全' + t.length() + '文字)' : t\n}\nString request = ''\nif (res instanceof org.apache.jmeter.protocol.http.sampler.HTTPSampleResult) {\n    request = (res.getHTTPMethod() ?: '') + ' ' + (res.getUrlAsString() ?: '')\n}\nString message = res.getResponseMessage() ?: ''\ndef failures = res.getAssertionResults().findAll { it.isFailure() || it.isError() }\nif (failures) {\n    message += ' / ' + failures.collect { it.getName() + ': ' + it.getFailureMessage() }.join(' / ')\n}\nString time = new SimpleDateFormat('yyyy-MM-dd HH:mm:ss.SSS').format(new Date(res.getStartTime()))\nString line = [time, res.isSuccessful() ? 'OK' : 'NG', res.getResponseCode(), res.getTime(), res.getBytesAsLong(),\n               res.getSampleLabel(), request.trim(), res.getThreadName(), oneLine(message, 1000)].join('\\t')\nsynchronized (writer) {\n    writer.println(line)\n    if (mode == 'all' || (mode == 'error' && !res.isSuccessful())) {\n        writer.println('\\t[REQUEST] ' + oneLine(res.getSamplerData(), maxChars))\n        writer.println('\\t[RESPONSE] ' + oneLine(res.getResponseDataAsString(), maxChars))\n    }\n}\nif (!res.isSuccessful()) {\n    log.warn('[API NG] ' + res.getSampleLabel() + ' status=' + res.getResponseCode() + ' ' + oneLine(message, 500))\n}\n"
 },
 "linux_cmd": {
  "suffix": "_linux_cmd.txt",
  "order": [
   "THREADS",
   "RAMP_UP",
   "LOOPS",
   "DURATION",
   "THROUGHPUT",
   "PROTOCOL",
   "HOST",
   "PORT",
   "CONNECT_TIMEOUT",
   "RESPONSE_TIMEOUT",
   "CSV_FILE",
   "EXCEL_FILE",
   "EXCEL_SHEET",
   "ACCESS_TOKEN_DIR",
   "ACCESS_TOKEN_FILE",
   "RESULT_DIR",
   "LOG_RESPONSE",
   "LOG_MAX_CHARS"
  ],
  "props": {
   "THREADS": "スレッド数（同時に実行する仮想ユーザー数）",
   "RAMP_UP": "ランプアップ秒数（全スレッドが起動し終わるまでの秒数）",
   "LOOPS": "ループ回数（各スレッドが API を繰り返す回数。-1 = 無限。DURATION と併用）",
   "DURATION": "実行時間（秒）。この秒数で終了します（LOOPS は -1 を推奨）",
   "THROUGHPUT": "スループット制御の目標（1 分あたりの件数。メイン処理全体の合計。スループット制御を有効にした場合のみ）",
   "PROTOCOL": "接続先のプロトコル（http / https）",
   "HOST": "接続先のホスト名",
   "PORT": "接続先のポート番号（空欄 = http は 80 / https は 443）",
   "CONNECT_TIMEOUT": "接続タイムアウト（ミリ秒）",
   "RESPONSE_TIMEOUT": "応答タイムアウト（ミリ秒）",
   "CSV_FILE": "CSV データファイル（相対パスは JMX のフォルダ基準）",
   "EXCEL_FILE": "Excel データファイル（相対パスは JMX のフォルダ基準）",
   "EXCEL_SHEET": "Excel データのシート名（空欄 = 先頭のシート）",
   "ACCESS_TOKEN_DIR": "アクセストークンのフォルダ（相対パスは JMX のフォルダ基準。; で区切ると存在する最初のフォルダを使用）",
   "ACCESS_TOKEN_FILE": "アクセストークンのファイル名（*.token のように * を含むと、更新日時が最も新しいファイル）",
   "RESULT_DIR": "結果ログ（api_result.log）・JTL 出力（result_all.jtl）・エラー詳細（error_detail.xml）の出力先。上のコマンドでは \"$RUN_DIR\" を指定しています",
   "LOG_RESPONSE": "結果ログへの本文の出力（none = 出力しない / error = 失敗時のみ / all = 全件）",
   "LOG_MAX_CHARS": "結果ログに出力する本文の最大文字数"
  },
  "nth_host": "（@@N@@ 番目の接続先）",
  "user_desc": "共通変数（設定 user_variables）。JMX の中の ${名前} で参照されます",
  "default_note": "。既定値: @@VALUE@@",
  "empty": "（空欄）",
  "no_props": "  （この JMX には、実行時に -J で変更できるプロパティがありません）",
  "file_jmx": "テスト計画（JMX）",
  "file_rel": "@@DESC@@（@@KEY@@）",
  "file_abs": "@@DESC@@（@@KEY@@。絶対パスのため、サーバ上の同じ場所に配置）",
  "file_desc": {
   "CSV_FILE": "CSV データ",
   "EXCEL_FILE": "Excel データ",
   "ACCESS_TOKEN_DIR": "アクセストークン"
  },
  "template": "################################################################################\n#  JMeter CLI の実行コマンド（Linux サーバ用）\n#    JMX          : @@JMX@@\n#    テストプラン : @@PLAN@@\n#    作成         : jmeter_api_tool @@VERSION@@（@@DATE@@）\n################################################################################\nこのファイルは、JMX を Linux サーバに持ち込み、JMeter の CLI（非 GUI）モードで実行するためのコマンドと、\n各パラメータの説明です。-J で指定している値は JMX の既定値です。必要に応じて書き換えてから実行してください。\n\n================================================================================\n1. 前提\n================================================================================\n- Java 17（8 以上）と Apache JMeter 5.6.3 がインストールされていること（例: /opt/apache-jmeter-5.6.3）。\n  確認: java -version / jmeter -v\n- jmeter にパスが通っていない場合は、下のコマンドの jmeter を /opt/apache-jmeter-5.6.3/bin/jmeter のように\n  フルパスで書いてください。\n\n================================================================================\n2. サーバに持ち込むファイル（JMX と同じフォルダ構成で配置します）\n================================================================================\n  JMX を置くフォルダ（例: /opt/jmeter_test）\n@@FILES@@\n\n  転送の例: scp -r <JMX のフォルダ>/ ユーザー@サーバ:/opt/jmeter_test/\n  ※ 相対パスは JMX のフォルダが基準です。ファイルが無いと CSV データ読込などの要素がエラーになるため、\n     使わない要素は JMX から削除するか、ファイルを配置してください。\n\n================================================================================\n3. 実行コマンド（そのまま貼り付けて実行できます）\n================================================================================\ncd /opt/jmeter_test\nRUN_DIR=\"result/@@STEM@@_$(date +%Y%m%d_%H%M%S)\"\nmkdir -p \"$RUN_DIR\"\nexport HEAP=\"-Xms1g -Xmx1g -XX:MaxMetaspaceSize=256m\"\n@@CMD_ML@@\n\n1 行で書く場合:\n@@CMD_1L@@\n\n長時間の実行で端末を閉じても続ける場合は、コマンドの前に nohup、後ろに > \"$RUN_DIR/console.log\" 2>&1 & を付けます。\n  例: nohup jmeter -n -t ...（上と同じオプション） > \"$RUN_DIR/console.log\" 2>&1 &\n\n================================================================================\n4. JMeter のオプション\n================================================================================\n  -n                          CLI（非 GUI）モードで実行します。\n  -t JMX                      実行するテスト計画（JMX）。\n  -l ファイル                 全リクエストの結果（JTL。CSV 形式）の出力先。\n  -j ファイル                 JMeter の実行ログ（jmeter.log）の出力先。\n  -e                          テスト終了後に HTML レポートを作成します。\n  -o フォルダ                 HTML レポートの出力先（存在しないか空のフォルダ）。\n  -J名前=値                   JMeter プロパティ。JMX の ${__P(名前,既定値)} の値を実行時に変更します（下記 5.）。\n  -Jjmeter.save.saveservice.subresults=false\n                              リダイレクトなどのサブ結果を JTL に出力しません（件数が API の実行回数と一致します）。\n  -f                          （任意）既存の結果ファイル・レポートフォルダを削除してから実行します。\n  -H ホスト -P ポート         （任意）プロキシサーバー経由で接続します。\n  -L レベル                   （任意）ログレベル（例: -LDEBUG）。\n  環境変数 HEAP               JVM のヒープサイズ。スレッド数が多い場合は増やします（例: -Xms4g -Xmx4g）。\n\n================================================================================\n5. プロパティ（-J名前=値）の説明（この JMX で使用しているもの）\n================================================================================\n@@PROPS@@\n\n================================================================================\n6. 実行後の確認\n================================================================================\n  ls \"$RUN_DIR\"                                                   # result.jtl / jmeter.log / report/ / api_result.log など\n  awk -F',' 'NR>1 && $8==\"false\"' \"$RUN_DIR/result.jtl\" | wc -l   # 失敗件数（ラベルにカンマが無い場合）\n  grep -E ' (ERROR|WARN) ' \"$RUN_DIR/jmeter.log\" | tail -n 20     # エラー・警告の確認\n  HTML レポート: \"$RUN_DIR/report/index.html\" をブラウザで開きます\n  （手元に取得する例: scp -r ユーザー@サーバ:/opt/jmeter_test/result ./）\n\n================================================================================\n7. 実行中の停止\n================================================================================\n  /opt/apache-jmeter-5.6.3/bin/shutdown.sh   # 実行中のリクエストの完了を待って停止\n  /opt/apache-jmeter-5.6.3/bin/stoptest.sh   # すぐに停止\n\n================================================================================\n8. 補足\n================================================================================\n- RESULT_DIR を相対パスで指定した場合は、jmeter を実行したフォルダ（カレントディレクトリ）が基準になります。\n  上のコマンドでは、結果をすべて \"$RUN_DIR\" にまとめています。\n- このツール（jmeter_api_tool.sh）もサーバに持ち込んだ場合は、次のコマンドでも同じように実行できます\n  （結果フォルダの作成・集計結果の表示まで行います）。\n    ./jmeter_api_tool.sh run -t @@JMX@@ --threads 10 --loops 100\n"
 },
 "update": {
  "base_suffix": ".base.json",
  "base_note": "jmeter_api_tool の generate -u（JMX の更新）で、前回生成した内容と比べるための基準ファイルです。JMX と同じフォルダに置いてください（削除すると、JMeter で変更した箇所も定義と違えば定義の内容に更新されます）。",
  "protect": [
   "setup_start",
   "teardown_end",
   "excel_loader",
   "excel_row",
   "access_token",
   "common_pre",
   "common_post",
   "result_log"
  ],
  "title": "==== JMX の更新内容（@@PATH@@） ====",
  "mode_3way": "比較: 前回生成した内容（@@BASE@@）と今回の定義の差分だけを反映します。JMeter で変更・追加・削除した箇所はそのまま残します。",
  "mode_2way": "比較: 基準ファイルが無いため JMX と今回の定義を比べ、違う箇所を定義の内容に更新します（JMeter で追加した要素と、共通スクリプトの変更は残します）。",
  "no_base": "基準ファイル（@@BASE@@）が無いため、JMeter で変更した箇所も、定義と違えば定義の内容に更新します。--dry-run（-DryRun）で更新箇所を確認してから実行してください",
  "plan_label": "テストプラン",
  "children": "（配下の要素 @@N@@ 件を含む）",
  "none": "（なし）",
  "order_only": "（値の並び順だけが変わります）",
  "max_lines": 30,
  "more_lines": "…（ほか @@N@@ 行）",
  "save_config": "（結果の保存項目）",
  "notes": {
   "競合": "（JMeter 側でも変更されていましたが、定義の値で更新します）",
   "維持": "（JMeter 側で変更されたスクリプトのため更新せず、JMeter 側の内容を残します）"
  },
  "summary": "追加 @@ADD@@ 件 / 変更 @@MOD@@ 件 / 削除 @@DEL@@ 件 / 競合 @@CONF@@ 件 / 維持 @@KEEP@@ 件",
  "count_keys": {
   "追加": "ADD",
   "変更": "MOD",
   "削除": "DEL",
   "競合": "CONF",
   "維持": "KEEP"
  },
  "summary_kept": "（定義に変更が無く JMeter 側の変更を残した箇所: @@KEPT@@、JMeter で削除済みのため追加しなかった要素: @@DELETED@@）",
  "dry_run": "（--dry-run のため、JMX は更新していません）",
  "tags": {
   "HTTPSamplerProxy": "HTTP リクエスト",
   "HeaderManager": "HTTP ヘッダ",
   "ConstantTimer": "待機（Constant Timer）",
   "JSR223PreProcessor": "JSR223 PreProcessor",
   "JSR223PostProcessor": "JSR223 PostProcessor",
   "JSR223Sampler": "JSR223 Sampler",
   "JSR223Listener": "JSR223 Listener",
   "RegexExtractor": "正規表現抽出",
   "JSONPostProcessor": "JSON 抽出",
   "ResponseAssertion": "応答の検証",
   "DurationAssertion": "応答時間の検証",
   "Arguments": "共通変数",
   "ConfigTestElement": "HTTP リクエスト初期値",
   "CookieManager": "クッキー管理",
   "SetupThreadGroup": "setUp Thread Group",
   "ThreadGroup": "Thread Group",
   "PostThreadGroup": "tearDown Thread Group",
   "CounterConfig": "連番",
   "RandomVariableConfig": "乱数",
   "CSVDataSet": "CSV データ読込",
   "ConstantThroughputTimer": "スループット制御",
   "DebugSampler": "Debug Sampler",
   "ResultCollector": "リスナー"
  },
  "labels": {
   "@testname": "名前",
   "@enabled": "有効（true）/ 無効（false）",
   "@guiclass": "GUI クラス",
   "@testclass": "クラス",
   "TestPlan.comments": "コメント",
   "HTTPSampler.domain": "サーバー名",
   "HTTPSampler.port": "ポート",
   "HTTPSampler.protocol": "プロトコル",
   "HTTPSampler.path": "パス",
   "HTTPSampler.method": "メソッド",
   "HTTPSampler.contentEncoding": "文字コード",
   "HTTPsampler.Arguments": "パラメータ / リクエストボディ",
   "HTTPSampler.postBodyRaw": "ボディを直接送信",
   "HTTPSampler.DO_MULTIPART_POST": "multipart/form-data で送信",
   "HTTPSampler.follow_redirects": "リダイレクトに従う",
   "HTTPSampler.use_keepalive": "KeepAlive",
   "HTTPSampler.connect_timeout": "接続タイムアウト",
   "HTTPSampler.response_timeout": "応答タイムアウト",
   "HeaderManager.headers": "ヘッダ",
   "ConstantTimer.delay": "待機時間（ミリ秒）",
   "script": "スクリプト",
   "parameters": "パラメータ",
   "Asserion.test_strings": "検証する値",
   "Assertion.assume_success": "ステータスを無視（Ignore Status）",
   "Assertion.test_type": "検証方法",
   "DurationAssertion.duration": "応答時間の上限（ミリ秒）",
   "RegexExtractor.refname": "変数名",
   "RegexExtractor.regex": "正規表現",
   "RegexExtractor.template": "テンプレート",
   "JSONPostProcessor.referenceNames": "変数名",
   "JSONPostProcessor.jsonPathExprs": "JSONPath",
   "ThreadGroup.num_threads": "スレッド数",
   "ThreadGroup.ramp_time": "ランプアップ秒数",
   "ThreadGroup.main_controller": "ループ回数",
   "ThreadGroup.scheduler": "スケジューラ",
   "ThreadGroup.duration": "実行時間（秒）",
   "ThreadGroup.on_sample_error": "エラー時の動作",
   "Arguments.arguments": "変数",
   "CookieManager.cookies": "クッキー",
   "CookieManager.clearEachIteration": "反復ごとにクリア",
   "filename": "ファイル",
   "variableNames": "変数名",
   "delimiter": "区切り文字",
   "fileEncoding": "文字コード",
   "ignoreFirstLine": "1 行目を読み飛ばす",
   "recycle": "終端で先頭に戻る",
   "stopThread": "終端でスレッド停止",
   "shareMode": "共有範囲",
   "CounterConfig.start": "開始値",
   "CounterConfig.incr": "増分",
   "CounterConfig.name": "変数名",
   "CounterConfig.format": "書式",
   "CounterConfig.per_user": "スレッドごと",
   "maximumValue": "最大値",
   "minimumValue": "最小値",
   "variableName": "変数名",
   "throughput": "目標スループット（件/分）",
   "ResultCollector.error_logging": "失敗のみ出力"
  }
 },
 "template": {
  "api_sheet": "API一覧",
  "settings_sheet": "設定",
  "guide_sheet": "記入方法",
  "api_title": "API定義一覧",
  "api_note": "1行 = 1API（上から順に実行）。複数の値はセル内改行（Alt + Enter）で区切ります。4行目（No が # で始まる行）は記入説明のため読み込まれません。詳しくは「記入方法」シートを参照してください。",
  "settings_title": "全体設定",
  "settings_note": "「設定値」列を変更してください（空欄は既定値）。ON = 生成して有効 / OFF = 生成しない / DISABLE = 生成するが無効化。説明に -J名前 とある項目は、実行時にも -J名前=値 で変更できます。",
  "settings_headers": [
   "設定キー",
   "設定項目",
   "設定値",
   "既定値",
   "説明"
  ],
  "settings_widths": [
   28,
   24,
   40,
   16,
   74
  ],
  "empty_default": "（空欄）",
  "blank_rows": 20,
  "validation_rows": 300
 },
 "styles": {
  "fonts": [
   {
    "sz": 10,
    "color": "FF000000"
   },
   {
    "sz": 10,
    "b": true,
    "color": "FF000000"
   },
   {
    "sz": 16,
    "b": true,
    "color": "FF1A1A1A"
   },
   {
    "sz": 11,
    "b": true,
    "color": "FFFFFFFF"
   },
   {
    "sz": 10,
    "b": true,
    "color": "FFFFFFFF"
   },
   {
    "sz": 9,
    "color": "FF595959"
   },
   {
    "sz": 10,
    "color": "FF404040"
   },
   {
    "sz": 20,
    "b": true,
    "color": "FF1A1A1A"
   },
   {
    "sz": 10,
    "u": true,
    "color": "FF262626"
   },
   {
    "sz": 11,
    "b": true,
    "color": "FF262626"
   }
  ],
  "fills": [
   "",
   "",
   "FF262626",
   "FF404040",
   "FF595959",
   "FFD9D9D9",
   "FFF2F2F2",
   "FFF8F8F8",
   "FFEDEDED",
   "FFFFFFFF"
  ],
  "borders": [
   {},
   {
    "left": [
     "thin",
     "FFA6A6A6"
    ],
    "right": [
     "thin",
     "FFA6A6A6"
    ],
    "top": [
     "thin",
     "FFA6A6A6"
    ],
    "bottom": [
     "thin",
     "FFA6A6A6"
    ]
   },
   {
    "left": [
     "thin",
     "FF7F7F7F"
    ],
    "right": [
     "thin",
     "FF7F7F7F"
    ],
    "top": [
     "thin",
     "FF7F7F7F"
    ],
    "bottom": [
     "thin",
     "FF7F7F7F"
    ]
   },
   {
    "bottom": [
     "medium",
     "FF1A1A1A"
    ]
   },
   {
    "bottom": [
     "thin",
     "FF7F7F7F"
    ]
   },
   {
    "left": [
     "thin",
     "FFBFBFBF"
    ],
    "right": [
     "thin",
     "FFBFBFBF"
    ],
    "top": [
     "thin",
     "FFBFBFBF"
    ]
   },
   {
    "left": [
     "thin",
     "FFBFBFBF"
    ],
    "right": [
     "thin",
     "FFBFBFBF"
    ]
   },
   {
    "left": [
     "thin",
     "FFBFBFBF"
    ],
    "right": [
     "thin",
     "FFBFBFBF"
    ],
    "bottom": [
     "thin",
     "FFBFBFBF"
    ]
   },
   {
    "left": [
     "thin",
     "FFBFBFBF"
    ],
    "right": [
     "thin",
     "FFBFBFBF"
    ],
    "top": [
     "thin",
     "FFBFBFBF"
    ],
    "bottom": [
     "thin",
     "FFBFBFBF"
    ]
   },
   {
    "left": [
     "medium",
     "FF595959"
    ]
   },
   {
    "left": [
     "thin",
     "FF595959"
    ],
    "right": [
     "thin",
     "FF595959"
    ],
    "top": [
     "thin",
     "FF595959"
    ],
    "bottom": [
     "thin",
     "FF595959"
    ]
   }
  ],
  "xfs": [
   [
    "default",
    0,
    0,
    0,
    "",
    "",
    false,
    0,
    0
   ],
   [
    "title",
    2,
    0,
    3,
    "",
    "center",
    false,
    0,
    0
   ],
   [
    "cover_title",
    7,
    0,
    0,
    "",
    "center",
    false,
    0,
    0
   ],
   [
    "cover_sub",
    6,
    0,
    0,
    "",
    "center",
    false,
    0,
    0
   ],
   [
    "h3",
    3,
    3,
    0,
    "",
    "center",
    false,
    1,
    0
   ],
   [
    "h4",
    9,
    0,
    4,
    "",
    "center",
    false,
    0,
    0
   ],
   [
    "p",
    0,
    0,
    0,
    "",
    "top",
    true,
    0,
    0
   ],
   [
    "li",
    0,
    0,
    0,
    "",
    "top",
    true,
    1,
    0
   ],
   [
    "quote",
    6,
    6,
    9,
    "",
    "top",
    true,
    1,
    0
   ],
   [
    "code_first",
    0,
    6,
    5,
    "",
    "top",
    true,
    1,
    0
   ],
   [
    "code_mid",
    0,
    6,
    6,
    "",
    "top",
    true,
    1,
    0
   ],
   [
    "code_last",
    0,
    6,
    7,
    "",
    "top",
    true,
    1,
    0
   ],
   [
    "code_single",
    0,
    6,
    8,
    "",
    "top",
    true,
    1,
    0
   ],
   [
    "th",
    4,
    4,
    2,
    "",
    "center",
    true,
    0,
    0
   ],
   [
    "td",
    0,
    0,
    1,
    "",
    "top",
    true,
    0,
    0
   ],
   [
    "td_alt",
    0,
    7,
    1,
    "",
    "top",
    true,
    0,
    0
   ],
   [
    "link",
    8,
    0,
    0,
    "",
    "center",
    false,
    0,
    0
   ],
   [
    "api_hdr",
    4,
    3,
    2,
    "center",
    "center",
    true,
    0,
    0
   ],
   [
    "api_hdr_req",
    4,
    2,
    2,
    "center",
    "center",
    true,
    0,
    0
   ],
   [
    "hint",
    5,
    8,
    1,
    "",
    "top",
    true,
    0,
    49
   ],
   [
    "cell",
    0,
    0,
    1,
    "",
    "top",
    true,
    0,
    49
   ],
   [
    "cat",
    1,
    5,
    1,
    "",
    "center",
    false,
    0,
    0
   ],
   [
    "key",
    0,
    6,
    1,
    "",
    "top",
    true,
    0,
    0
   ],
   [
    "input",
    0,
    9,
    10,
    "",
    "top",
    true,
    0,
    49
   ],
   [
    "note",
    5,
    0,
    0,
    "",
    "center",
    true,
    0,
    0
   ],
   [
    "data_hdr",
    4,
    4,
    2,
    "center",
    "center",
    false,
    0,
    0
   ],
   [
    "data_cell",
    0,
    0,
    1,
    "",
    "center",
    false,
    0,
    49
   ],
   [
    "td_center",
    0,
    0,
    1,
    "center",
    "top",
    true,
    0,
    0
   ],
   [
    "toc_no",
    0,
    0,
    1,
    "center",
    "center",
    false,
    0,
    0
   ],
   [
    "toc_link",
    8,
    0,
    1,
    "",
    "center",
    false,
    0,
    0
   ],
   [
    "toc_td",
    0,
    0,
    1,
    "",
    "center",
    true,
    0,
    0
   ]
  ]
 },
 "xlsx": {
  "font_name": "Meiryo UI",
  "line_h10": 144,
  "pad_h10": 36,
  "max_h10": 4090,
  "units": {
   "upper": 22,
   "digit": 20,
   "lower": 17,
   "punct": 15,
   "latin": 20,
   "hankaku": 17,
   "kana": 27,
   "wide": 33
  },
  "min_total_width": 110,
  "manual_file": "JMeter_Usage_Guide",
  "cover_sheet": "表紙",
  "toc_title": "目次",
  "toc_headers": [
   "No",
   "シート",
   "内容"
  ],
  "cover_widths": [
   6,
   26,
   82
  ]
 },
 "manual_md": "# JMeter APIテスト 利用方法\n\n> 対象: Apache JMeter 5.6.3 ／ RHEL 9.8 ／ Windows 11 ／ ツール: jmeter_api_tool v1.0.0（@@DATE@@ 出力）\n\n本書は、REST API の定義（URL・HTTPメソッド・クエリパラメータ・ヘッダ・リクエストボディ）を Excel または CSV に記入するか OpenAPI 定義（JSON）を指定して、JMeter のテスト計画ファイル（JMX）を自動生成して実行し、結果を確認するまでの手順をまとめたものです。\n\n## 1.概要\n<!-- summary: ツールでできること・ファイル構成・作業の流れ -->\n\n### このツールでできること\n- Excel / CSV に記入した API 定義、または OpenAPI 3.0.x の定義ファイル（JSON）から、JMeter 5.6.3 用の JMX ファイルを自動生成します（OpenAPI は「6.OpenAPI定義から生成」参照）。\n- curl コマンド（ブラウザの開発者ツールの「Copy as cURL」など）を解析して、API一覧（Excel / CSV）に 1 行追加できます（add-curl コマンド。「3.JMX生成」参照）。\n- JMeter GUI で調整した JMX を、API 定義（Excel / CSV / OpenAPI）の変更に合わせて、違う箇所だけ更新できます。更新する箇所は書き換える前に一覧で表示します（generate -u。「8.JMXの更新」参照）。\n- JMX を Linux サーバに持ち込んで JMeter の CLI で実行するためのコマンドと、各パラメータの説明をテキストファイル（JMX名_linux_cmd.txt）に出力します（「9.実行方法」参照）。\n- 生成される JMX には、API テストでよく使う要素（待機タイマ、連番・乱数、CSV / Excel データ読込、クッキー、前処理・後処理、正規表現 / JSON 抽出、応答の検証、結果出力）があらかじめ組み込まれています。\n- 既定は「1スレッド × 1回」の実行です。実行時のプロパティ指定（例: -JTHREADS=10 -JLOOPS=100）だけで、大量件数の負荷試験に切り替えられます。\n- 実行結果は JTL（CSV）、HTML レポート、jmeter.log、結果ログ（api_result.log）、エラー詳細（error_detail.xml）に出力されます。\n- 生成される要素の名前には [必須] / [任意] を付けています。[任意] の要素は JMeter GUI で右クリック →「Remove」で削除するだけで、他の要素に影響を与えずに取り除けます。\n\n### ファイル構成\n<!-- cols: 36,74 -->\n| パス | 内容 |\n|---|---|\n| jmeter_api_tool.sh | RHEL 9.8 用スクリプト（bash + python3 標準ライブラリのみ使用） |\n| jmeter_api_tool.ps1 | Windows 11 用スクリプト（Windows PowerShell 5.1 / PowerShell 7 で動作） |\n| jmeter_api_tool.bat | Windows 用の起動ファイル（実行ポリシーを一時的に回避して .ps1 を起動） |\n| input/api_definition.xlsx | 入力テンプレート（Excel）：「API一覧」「設定」「記入方法」シート |\n| input/api_definition.csv | 入力テンプレート（CSV）：API 一覧 |\n| input/api_definition_settings.csv | 入力テンプレート（CSV）：設定 |\n| input/openapi_config.csv | OpenAPI 定義（.json）から生成する場合の設定（接続先・HTTP ヘッダ・アクセストークン・定義の変数 ${public_base_url} / ${name_prefix} の値）。generate の -c / -OpenApiConfig で指定 |\n| output/api_test.jmx | 生成された JMX（サンプル） |\n| output/api_test_linux_cmd.txt | Linux サーバで JMeter の CLI を実行するコマンドと各パラメータの説明（サンプル。generate で JMX と一緒に出力） |\n| output/data/testdata.csv | CSV データ読込のサンプルデータ |\n| output/data/testdata.xlsx | Excel データ読込のサンプルデータ |\n| docs/JMeter_Usage_Guide.xlsx / .md | 本書（manual コマンドで出力） |\n| result/JMX名_日時/ | run コマンドの実行結果（JTL・ログ・HTML レポート） |\n\n### 作業の流れ\n1. テンプレートを作成する（init コマンド）\n2. API 定義と設定を記入する（Excel または CSV。curl コマンドから 1 行追加することもできます（add-curl）。OpenAPI 定義から生成する場合は「6.OpenAPI定義から生成」）\n3. JMX を生成する（generate コマンド）\n4. JMeter GUI で内容を確認し、不要な [任意] 要素を削除・調整する（API 定義を変更した場合は generate -u で JMX を差分だけ更新する）\n5. CLI モードで実行する（run コマンド。Linux サーバでは JMX と一緒に出力される「JMX名_linux_cmd.txt」のコマンド）\n6. 結果（HTML レポート・JTL・ログ）を確認する\n\n> ファイル名・フォルダ名はすべて英数字にしています。Windows と Linux の間でファイルを転送したときの文字化けを防ぐためです。\n\n## 2.環境構築\n<!-- summary: Java・JMeter 5.6.3 のインストールと動作確認（RHEL 9.8 / Windows 11） -->\n\n### 必要なソフトウェア\n<!-- cols: 18,30,30,40 -->\n| ソフトウェア | RHEL 9.8 | Windows 11 | 備考 |\n|---|---|---|---|\n| Java | java-17-openjdk（dnf） | JDK 17（Eclipse Temurin 等） | JMeter 5.6.3 は Java 8 以上で動作。17 を推奨 |\n| Apache JMeter | 5.6.3（.tgz） | 5.6.3（.zip） | https://archive.apache.org/dist/jmeter/binaries/ |\n| スクリプト実行環境 | bash・python3（OS 標準） | Windows PowerShell 5.1（OS 標準） | 追加のモジュール・ライブラリは不要 |\n| Excel | 不要 | 不要（記入・閲覧には推奨） | xlsx の読み書きはスクリプト自身が行います |\n\n### RHEL 9.8 へのインストール\n```bash\n# Java 17（GUI を使わないサーバーなら java-17-openjdk-headless でも可）\nsudo dnf install -y java-17-openjdk\njava -version\n\n# JMeter 5.6.3（ダウンロード・チェックサム確認・展開）\ncd /opt\nsudo curl -LO https://archive.apache.org/dist/jmeter/binaries/apache-jmeter-5.6.3.tgz\nsudo curl -LO https://archive.apache.org/dist/jmeter/binaries/apache-jmeter-5.6.3.tgz.sha512\nsha512sum -c apache-jmeter-5.6.3.tgz.sha512\nsudo tar xzf apache-jmeter-5.6.3.tgz\n\n# 環境変数（~/.bashrc に追記）\necho 'export JMETER_HOME=/opt/apache-jmeter-5.6.3' >> ~/.bashrc\necho 'export PATH=$JMETER_HOME/bin:$PATH' >> ~/.bashrc\nsource ~/.bashrc\njmeter -v\n```\n- python3 は RHEL 9 のパッケージ管理（dnf）が使用しているため、通常は導入済みです（確認: `python3 --version`）。無い場合は `sudo dnf install -y python3` で導入します。\n- ツールのファイルを Windows から転送して改行コードが CRLF になった場合は、`sed -i 's/\\r$//' jmeter_api_tool.sh` で LF に戻してください。\n- GUI を使う場合に日本語が □ で表示されるときは、日本語フォントを導入します（例: `sudo dnf install -y google-noto-sans-cjk-ttc-fonts`）。\n\n### Windows 11 へのインストール\n1. JDK 17 をインストールします（例: Eclipse Temurin 17 の .msi。インストール時に「JAVA_HOME 変数の設定」と「PATH への追加」を有効にします）。\n2. https://archive.apache.org/dist/jmeter/binaries/apache-jmeter-5.6.3.zip をダウンロードし、ファイルのプロパティで「許可する（ブロック解除）」にチェックしてから、C:\\apache-jmeter-5.6.3 などスペースや日本語を含まないフォルダに展開します。\n3. 環境変数 JMETER_HOME に展開先のフォルダを設定します（下記のコマンド、または「システムの詳細設定」→「環境変数」）。\n4. PowerShell を開き直して動作を確認します。\n\n```powershell\n[Environment]::SetEnvironmentVariable('JMETER_HOME', 'C:\\apache-jmeter-5.6.3', 'User')\n# PowerShell を開き直してから確認\njava -version\n& \"$env:JMETER_HOME\\bin\\jmeter.bat\" -v\n```\n\n### ツールの配置と動作確認\nツール一式を任意のフォルダ（例: RHEL は ~/JMeter_API_Generate、Windows は C:\\work\\JMeter_API_Generate）に配置し、ヘルプが表示されることを確認します。\n\n```bash\n# RHEL 9.8\ncd ~/JMeter_API_Generate\nchmod +x jmeter_api_tool.sh\n./jmeter_api_tool.sh help\n```\n\n```powershell\n# Windows 11（PowerShell またはコマンドプロンプト）\ncd C:\\work\\JMeter_API_Generate\n.\\jmeter_api_tool.bat help\n```\n- jmeter_api_tool.bat は PowerShell の実行ポリシー（既定では .ps1 の実行が禁止）を、その実行に限って回避して起動します。PowerShell から直接実行する場合は `powershell -ExecutionPolicy Bypass -File .\\jmeter_api_tool.ps1 help` としてください。\n- インターネットから取得した ZIP を展開した場合は、事前に `Get-ChildItem -Recurse | Unblock-File` でブロックを解除しておくと確実です。\n\n## 3.JMX生成\n<!-- summary: テンプレート作成から JMX 生成・GUI での確認までの手順 -->\n\n### 手順1：テンプレートを作成する\n```bash\n./jmeter_api_tool.sh init          # RHEL 9.8\n```\n```powershell\n.\\jmeter_api_tool.bat init         # Windows 11\n```\ninput フォルダに次の 4 ファイルが作成されます（既存ファイルは上書きしません。上書きする場合は -f / -Force を付けます）。\n- api_definition.xlsx（Excel 形式のテンプレート）\n- api_definition.csv（CSV 形式：API 一覧）\n- api_definition_settings.csv（CSV 形式：設定）\n- openapi_config.csv（CSV 形式：OpenAPI 定義から生成する場合の接続先・HTTP ヘッダ・アクセストークン・定義の変数 ${public_base_url} / ${name_prefix} の値。generate の -c / -OpenApiConfig で指定）\n\n### 手順2：API 定義を記入する\n- Excel 形式：「API一覧」シートに 1行 = 1API で記入し、「設定」シートの「設定値」列で全体設定を変更します。\n- CSV 形式：api_definition.csv に API を、api_definition_settings.csv に設定を記入します。\n- テンプレートには記入例として、postman-echo.com（送信した内容をそのまま返す公開サービス）を呼び出す 4 件の API が入っています。自分の API に書き換えてください。\n- 入力項目の詳細は「4.入力項目_設定」「5.入力項目_API」シートを参照してください。\n\n> CSV を Excel で保存する場合は「CSV UTF-8（コンマ区切り）」を選んでください。Shift_JIS の「CSV（コンマ区切り）」で保存したファイルも自動判別して読み込めます。\n\n### curl コマンドから API を追加する（add-curl）\nブラウザの開発者ツール（ネットワーク → 右クリック →「Copy as cURL」）や API の仕様書にある curl コマンドを解析して、API一覧に 1 行追加します。Excel は「API一覧」シートの最後の API の次の行（空白行）に、CSV はファイルの末尾に追加します。元のファイルは「ファイル名.日時.bak」に退避してから書き換えます。\n\n```bash\n# RHEL 9.8\n./jmeter_api_tool.sh add-curl -i input/api_definition.xlsx --curl-file curl.txt\n./jmeter_api_tool.sh add-curl -i input/api_definition.csv --name 注文登録 -- curl -X POST https://api.example.com/v1/orders -H 'Content-Type: application/json' -d '{\"item\":\"A001\"}'\n```\n```powershell\n# Windows 11\n.\\jmeter_api_tool.bat add-curl -i input\\api_definition.xlsx -CurlFile curl.txt\n.\\jmeter_api_tool.bat add-curl -i input\\api_definition.csv -CurlFile curl.txt -ApiName 注文登録 -DryRun\n```\n\n<!-- cols: 24,28,66 -->\n| sh のオプション | PowerShell のオプション | 説明 |\n|---|---|---|\n| -i ファイル | -InputFile（-i） | 追加先の API一覧（.xlsx / .csv）。必須 |\n| --curl-file ファイル | -CurlFile ファイル | curl コマンドを書いたテキストファイル（コピーした内容をそのまま保存したもの。行末の \\ ・ ^ ・ ` で改行した複数行のコマンドも可） |\n| --curl \"curl …\" | -Curl \"curl …\" | curl コマンドを文字列で指定（\" を含むコマンドは --curl-file / -CurlFile で指定してください） |\n| -- curl … | （なし） | -- の後ろに curl コマンドをそのまま書く（sh のみ） |\n| --name API名 | -ApiName API名 | API名（省略時は「メソッド パス」。例: POST /v1/orders） |\n| --dry-run | -DryRun | 追加する行を表示するだけで、ファイルは書き換えない |\n\n<!-- cols: 40,78 -->\n| curl のオプション | API一覧に設定する値 |\n|---|---|\n| URL | URL（スキーム・ホスト・ポート・パス）とクエリパラメータ（? 以降を 1 行 1 つ。%XX は元の文字に戻す）。スキームが無い場合は http |\n| -X / --request | メソッド（省略時は -d・-F があれば POST、-G は GET、-I は HEAD、-T は PUT、それ以外は GET） |\n| -H / --header | ヘッダ（1 行 1 つ）。Content-Length・Transfer-Encoding は JMeter が自動で付けるため追加しない |\n| -d / --data / --data-raw / --data-binary / --data-urlencode / --json | リクエストボディ。名前=値 の形式は form、JSON は json、それ以外は raw（-G を付けた場合はクエリパラメータ。@ファイル はファイルの内容） |\n| -F / --form / --form-string | multipart のボディ（名前=値）。ファイルの送信（名前=@ファイル）は追加しないため、JMeter の「Files Upload」タブで設定する |\n| -u / --user、URL の ユーザー:パスワード@ | ヘッダ Authorization: Basic …（パスワードが無い場合は作らない） |\n| --oauth2-bearer | ヘッダ Authorization: Bearer … |\n| -b / --cookie、-A / --user-agent、-e / --referer | ヘッダ Cookie / User-Agent / Referer（-b ファイル名 は読み込まない） |\n| -s・-L・-k・--compressed など | JMeter の HTTP リクエストに関係しないため無視する（-x プロキシ・--cert など JMeter で別に設定するものは、設定方法を警告で表示して無視） |\n\n- 1 回に追加するのは 1 行です（&& や ; でつないだ 2 つ目以降のコマンドと、--next 以降は読みません）。\n- No は既存の No の最大値 + 1、有効は Y、備考は「curl コマンドから追加」です。Excel は書式（罫線・折り返し）を上の行から引き継ぎ、行の高さを内容に合わせます。\n- CSV は元の文字コード（UTF-8 / UTF-8（BOM 付き）/ Shift_JIS / UTF-16）と改行コードのまま追記します。Shift_JIS で表せない文字を含む場合はエラーになるため、CSV を UTF-8 で保存し直してください。\n- 追加した後は generate で JMX を生成します。生成済みの JMX がある場合は generate -u で、追加した API だけを JMX に追加できます（「8.JMXの更新」参照）。\n\n### 手順3：JMX を生成する\n```bash\n# RHEL 9.8\n./jmeter_api_tool.sh generate -i input/api_definition.xlsx\n./jmeter_api_tool.sh generate -i input/api_definition.csv -s input/api_definition_settings.csv\n```\n```powershell\n# Windows 11\n.\\jmeter_api_tool.bat generate -i input\\api_definition.xlsx\n.\\jmeter_api_tool.bat generate -i input\\api_definition.csv -s input\\api_definition_settings.csv\n```\n\n<!-- cols: 22,30,66 -->\n| sh のオプション | PowerShell のオプション | 説明 |\n|---|---|---|\n| -i ファイル | -InputFile（-i） | 入力ファイル（.xlsx / .csv、OpenAPI 定義は .json）。必須 |\n| -s ファイル | -SettingsFile（-s） | 設定 CSV。省略時は「入力ファイル名_settings.csv」を探し、無ければ既定値を使用（Excel 入力時は「設定」シートを使用） |\n| -o フォルダ | -OutDir（-o） | 出力先フォルダ（既定: output） |\n| -n ファイル名 | -JmxName（-n） | 出力する JMX のファイル名（例: order_test.jmx）。設定「出力JMXファイル名」（output_file）より優先します |\n| --split | -Split | API ごとに別々の JMX を出力する |\n| -f | -Force（-f） | サンプルデータ（data/testdata.csv / .xlsx）が既にあっても上書きする |\n| -u JMX | -UpdateJmx（-u） | 新しい JMX を作らずに、既存の JMX を入力ファイルとの差分だけ更新する（「8.JMXの更新」参照） |\n| --dry-run | -DryRun | -u と併用: 更新箇所を表示するだけで、ファイルは書き換えない |\n| -c ファイル など | -OpenApiConfig（-c） など | OpenAPI 定義（.json）用のオプション（「6.OpenAPI定義から生成」参照） |\n\n### 生成されるファイル\n<!-- cols: 34,84 -->\n| ファイル | 説明 |\n|---|---|\n| output/出力JMXファイル名 | 生成された JMX。同名のファイルがある場合は「.日時.bak」を付けて退避してから作成します |\n| output/JMX名_linux_cmd.txt | Linux サーバで JMeter の CLI を実行するコマンドと各パラメータの説明（「9.実行方法」参照。JMX名は .jmx を除いた名前）。--split の場合は JMX ごとに出力します |\n| output/出力JMXファイル名.base.json | JMX の更新（generate -u）で使う基準ファイル（生成した内容の控え）。JMX と同じフォルダに置いたままにしてください（--split の場合は作成しません） |\n| output/data/testdata.csv | CSV データ読込用のサンプル（ファイルが無い場合のみ作成） |\n| output/data/testdata.xlsx | Excel データ読込用のサンプル（ファイルが無い場合のみ作成） |\n\n### 手順4：JMeter GUI で確認・調整する\n- RHEL（デスクトップ環境あり）：`jmeter -t output/api_test.jmx`\n- Windows：`%JMETER_HOME%\\bin\\jmeter.bat` をダブルクリックし、「File」→「Open」で JMX を選択\n- 不要な [任意] 要素は右クリック →「Remove」で削除、一時的に外す場合は右クリック →「Disable」で無効化できます。\n- GUI モードは動作確認・デバッグ用です。負荷試験は必ず CLI モード（run コマンド）で実行してください。\n\n## 4.入力項目_設定\n<!-- summary: 「設定」シート（設定 CSV）の全項目と既定値 -->\n「設定」シート（CSV の場合は api_definition_settings.csv）の項目です。設定値が空欄の項目は既定値が使われます。\n\n- ON / OFF / DISABLE の項目は、ON = 生成して有効、OFF = 生成しない、DISABLE = 生成するが無効化（GUI で右クリック →「Enable」ですぐ使える）を表します。\n- 説明欄の「-J名前」は、実行時に -J名前=値（run コマンドでは -J 名前=値 / -Prop 名前=値）で値を変更できることを表します。\n\n<!-- @SETTINGS_TABLE@ -->\n\n## 5.入力項目_API\n<!-- summary: 「API一覧」シート（API CSV）の列と記入ルール -->\n「API一覧」シート（CSV の場合は api_definition.csv）の列です。1行 = 1API で、上から順に実行されます。\n\n<!-- @API_COLUMNS_TABLE@ -->\n\n### 記入のルール\n- 複数の値を書く列（クエリパラメータ、ヘッダ、正規表現抽出など）は、セル内改行（Excel では Alt + Enter）で区切ります。クエリパラメータと form 形式のボディは & 区切りでも記入できます。\n- 先頭の列（No）が # で始まる行はコメントとして読み飛ばします（テンプレートの 4 行目の記入説明など）。\n- URL・パラメータ・ヘッダ・ボディには ${変数名} で JMeter の変数・関数を埋め込めます（例: `${USER_ID}`、`${SEQ}`、`${__UUID()}`）。\n- 「有効」を N にした API は、生成はされますが無効化された状態になります。\n- GET / DELETE / OPTIONS のクエリパラメータは HTTP リクエストの「Parameters」タブに設定されます。それ以外（POST / PUT / PATCH / HEAD など）の場合は URL のパスに付与されます（値は UTF-8 で URL エンコード。${...} の部分はそのまま）。\n- 既に URL エンコード済みの値（%E3%83%86 など）は、二重にエンコードしないよう自動で判定します。\n- GET / DELETE / HEAD / OPTIONS ではリクエストボディは使用できません（指定した場合は警告を表示して無視します）。\n\n### ボディ形式（auto の場合の自動判定）\n<!-- cols: 26,16,76 -->\n| 本文の内容 | 判定 | 送信方法 |\n|---|---|---|\n| { または [ で始まる | json | 本文をそのまま送信。Content-Type が未指定なら application/json を付与 |\n| < で始まる | raw（XML） | 本文をそのまま送信。Content-Type が未指定なら application/xml を付与 |\n| 名前=値 の形式（改行 / & 区切り） | form | パラメータとして送信（application/x-www-form-urlencoded） |\n| 上記以外 | raw | 本文をそのまま送信（Content-Type はヘッダ列で指定） |\n\n## 6.OpenAPI定義から生成\n<!-- summary: OpenAPI 定義（JSON）から JMX を生成する手順・変換ルール・OpenAPI 用設定ファイル -->\ngenerate コマンドの -i に OpenAPI 3.0.x の定義ファイル（JSON、拡張子 .json）を指定すると、定義の各操作（パスとメソッドの組）を「API一覧」と同じ 1行 = 1API に変換してから JMX を生成します。API は定義に書かれた順番に実行されます。\n\n### 生成手順\n```bash\n# RHEL 9.8\n./jmeter_api_tool.sh generate -i openapi.json\n./jmeter_api_tool.sh generate -i openapi.json -c input/openapi_config.csv --seed 12345\n```\n```powershell\n# Windows 11\n.\\jmeter_api_tool.bat generate -i openapi.json\n.\\jmeter_api_tool.bat generate -i openapi.json -c input\\openapi_config.csv -Seed 12345\n```\n\n<!-- cols: 24,28,66 -->\n| sh のオプション | PowerShell のオプション | 説明 |\n|---|---|---|\n| -c ファイル | -OpenApiConfig（-c） | OpenAPI 用設定ファイル（下記）。省略時は各項目の既定値を使用 |\n| --seed 数 | -Seed 数 | テストデータの乱数シード（0 以上の整数・12 桁以内）。同じ値なら同じテストデータを再生成します。省略時は毎回ランダム（使ったシードは生成時に表示） |\n| --any-method メソッド | -AnyMethod メソッド | x-amazon-apigateway-any-method（ANY）の API で送るメソッド（GET / POST / PUT / PATCH / DELETE / HEAD / OPTIONS。既定: GET） |\n| --required-only | -RequiredOnly | 必須（required）のパラメータ・プロパティだけを生成 |\n| --use-examples | -UseExamples | example / default があればランダム値より優先して使用 |\n| --placeholder-value 数値 | -PlaceholderValue 数値 | 値の位置に引用符なしで書かれた ${...}（JSON として読めない部分）に仮に入れる数値（既定: 3000） |\n\n- -s（設定ファイル）・-o・-n・--split・-f の使い方は Excel / CSV と同じです。-s を省略した場合は「入力ファイル名_settings.csv」があれば読み込み、無ければ既定値を使います。\n- 対象は OpenAPI 3.0.x の JSON です（3.1 など 3.0.x 以外は警告を表示し、3.0.x と共通の項目だけを解釈します）。YAML・Swagger 2.0・外部ファイルへの $ref には対応していません（YAML は JSON に変換してから指定してください）。\n\n### API一覧への変換ルール\n<!-- cols: 24,94 -->\n| API一覧の列 | OpenAPI 定義から設定する値 |\n|---|---|\n| No / 有効 / API名 | 定義の順番 / Y / 「メソッド パス [operationId]」（ANY は「ANY(GET) パス」） |\n| メソッド / URL | 操作のメソッドとパス。パスパラメータはデータ型に合ったランダム値（URL エンコード済み）。接続先とパスの先頭は設定 base_url |\n| クエリパラメータ | in: query のパラメータ（style / explode に従って展開） |\n| ヘッダ | Content-Type（リクエストボディの形式）、Accept（成功応答の形式）、in: header / in: cookie のパラメータ、OpenAPI 用設定 API_HEADERS の値（同じ名前のヘッダは API_HEADERS の値を使用） |\n| ボディ形式 / リクエストボディ | requestBody の定義から生成（JSON → json、application/x-www-form-urlencoded → form、multipart/form-data → multipart、XML・テキストなど → raw） |\n| 期待ステータス | 2xx の応答コード（カンマ区切り。2XX だけの場合は空欄） |\n| 備考 | 概要・説明・operationId・認可（security）・API Gateway 統合（x-amazon-apigateway-integration）・乱数シードなど |\n\n- パラメータ・リクエストボディの値は、データ型・format・pattern・minimum / maximum・enum などの定義からランダムに生成します。生成の規則は JMeter_Tools の openapi2jmx と同じで、同じシードなら同じ値になります。\n- 待機時間は設定 default_wait_ms を使います。正規表現抽出・JSON抽出・前処理・後処理・応答検証文字列・応答時間上限は生成しないため、必要に応じて JMeter GUI で追加してください。\n- 認証ヘッダ（Authorization・API-Key）はランダム値では通らないため、OpenAPI 定義のパラメータからは生成しません。Authorization は設定 common_headers、または OpenAPI 用設定ファイルの HEADERS / TOKEN_DIR で、API-Key は API_HEADERS で指定します（API_HEADERS の既定値の API-Key: XXXXXXX は実際のキーに書き換えてください）。指定が無い場合は生成時に警告を表示します。\n- GET / DELETE / HEAD / OPTIONS のリクエストボディは、Excel / CSV と同じ規則で送信しません（警告を表示します）。\n- 設定ファイルで空欄の base_url（servers[0].url。無い場合は http://localhost:8080）・test_plan_name（「タイトル APIのバージョン（OpenAPI のバージョン）」）・output_file（「タイトル.jmx」）は OpenAPI 定義の値を使います（output_file は -n / -JmxName が優先）。\n- 共通ヘッダ・クッキー・連番・CSV / Excel データ読込・結果出力などの要素は、Excel / CSV から生成する場合と同じく設定（「4.入力項目_設定」）に従って生成されます。CSV / Excel データ読込は既定で ON のため、サンプルデータ（data/testdata.csv / .xlsx）も作成されます。使わない場合は設定 csv_data / excel_data を OFF にしてください。\n\n### OpenAPI 用設定ファイル（openapi_config.csv）\nOpenAPI 定義から生成する場合だけ使う設定ファイルです。init で作成される input/openapi_config.csv（列は設定 CSV と同じ）の「設定値」を記入し、generate の -c / -OpenApiConfig で指定します。設定値が空欄の項目は使われず（API_HEADERS は既定値を使用）、テンプレートのままなら -c を指定しない場合と同じ JMX になります。記入した項目は、設定ファイル（-s）・OpenAPI 定義の値より優先します。\n\n<!-- @OPENAPI_CONFIG_TABLE@ -->\n\n- PROTOCOL / HOST / PORT は、記入した項目だけ base_url の接続先を置き換え、パス（例: /v1）はそのまま使います。PORT が空欄の場合はポート番号を付けません（http は 80 / https は 443 で接続）。実行時の -JPROTOCOL= / -JHOST= / -JPORT= も使えます。\n- API_HEADERS の値は各 API のヘッダ（API ごとの HTTP Header Manager）の初期値です。生成後は JMeter で API ごとに変更でき、HEADERS（共通ヘッダ）と同じ名前のヘッダは実行時に API 側が優先されます。\n- TOKEN_DIR を指定すると、共通ヘッダに Authorization: Bearer ${ACCESS_TOKEN} を追加し、JMeter の実行中に API を呼び出すたびに、フォルダ直下の *.token のうち更新日時が最も新しいファイルの内容（前後の空白・改行・BOM を除く）を ${ACCESS_TOKEN} に設定します。トークンは JMX に書き込まず、実行時に読み込みます。\n- 新しい .token ファイルを置く・ファイルを書き換えると、実行中でも次の API 呼び出しから切り替わります（フォルダの確認は 1 秒に 1 回）。拡張子が .token 以外のファイルは使いません。生成時には、そのとき最新の .token ファイル（またはフォルダ・ファイルが無いことの警告）を表示します。\n- TOKEN_DIR の区切りは \\ と / のどちらでも可、相対パスは JMX のフォルダ基準です。Windows と Linux のフォルダを ; で区切って書くと、実行する環境に存在する方を使うため、同じ JMX を両方で使えます（例: C:\\work\\tokens;/home/user/tokens）。実行時に -JACCESS_TOKEN_DIR=フォルダ で変更することもできます。\n- PUBLIC_BASE_URL / NAME_PREFIX は OpenAPI 定義（servers[0].url・info.title）に入れる値です。未設定（空欄）の場合は ${public_base_url} / ${name_prefix} を取り除くため、パスの先頭やテストプラン名・JMX のファイル名に ${...} は残りません（例: ${name_prefix}-orders-api → orders-api）。設定ファイル（-s）で base_url / test_plan_name / output_file を指定した場合は、その設定の値を使います。\n\n## 7.JMX構成\n<!-- summary: 生成される JMX の要素一覧と、削除してよいかの判断基準 -->\n生成される JMX のツリー構成です（記入例を生成した場合）。名前が [必須] の要素は削除しないでください。[任意] の要素は、他の要素に影響を与えずに削除できます。\n\n```text\nテストプラン\n├─ [必須] 共通変数 (User Defined Variables)\n├─ [任意] HTTPリクエスト初期値 (HTTP Request Defaults)\n├─ [任意] 共通ヘッダ (HTTP Header Manager)\n├─ [任意] クッキー管理 (HTTP Cookie Manager)\n├─ [任意] 事前処理 (setUp Thread Group)\n│   ├─ [任意] テスト開始処理 (JSR223 Sampler)\n│   └─ [任意] Excelデータ読込 (JSR223 Sampler)\n├─ [必須] メイン処理 (Thread Group)\n│   ├─ [任意] 連番 SEQ (Counter)\n│   ├─ [任意] 乱数 RANDOM_NO (Random Variable)\n│   ├─ [任意] CSVデータ読込 (CSV Data Set Config)\n│   ├─ [任意] Excelデータ行取得 (JSR223 PreProcessor)\n│   ├─ [任意] 共通前処理 (JSR223 PreProcessor)\n│   ├─ [任意] 共通後処理 (JSR223 PostProcessor)\n│   ├─ [任意] スループット制御 (Constant Throughput Timer) ※既定は無効\n│   ├─ API01_ログイン(POST JSON)  ← HTTP リクエスト（1API = 1つ）\n│   │   ├─ [任意] API01 ヘッダ (HTTP Header Manager)\n│   │   ├─ [任意] API01 待機 500ms (Constant Timer)\n│   │   ├─ [任意] API01 前処理 (JSR223 PreProcessor)\n│   │   ├─ [任意] API01 正規表現抽出 ECHO_URL (Regular Expression Extractor)\n│   │   ├─ [任意] API01 JSON抽出 AUTH_TOKEN (JSON Extractor)\n│   │   ├─ [任意] API01 後処理 (JSR223 PostProcessor)\n│   │   ├─ [任意] API01 ステータス検証 200 (Response Assertion)\n│   │   ├─ [任意] API01 応答本文検証 (Response Assertion)\n│   │   └─ [任意] API01 応答時間検証 5000ms (Duration Assertion)\n│   ├─ API02_商品検索(GET) / API03_注文登録(POST form) / API04_クッキー確認(GET)\n│   └─ [任意] 変数確認 (Debug Sampler) ※既定は無効\n├─ [任意] 事後処理 (tearDown Thread Group)\n│   └─ [任意] テスト終了処理 (JSR223 Sampler)\n├─ [任意] 結果ツリー (View Results Tree)\n├─ [任意] 集計レポート (Summary Report)\n├─ [任意] 統計レポート (Aggregate Report)\n├─ [任意] JTL出力 全件 (Simple Data Writer)\n├─ [任意] エラー詳細出力 XML (Simple Data Writer)\n└─ [任意] 結果ログ出力 (JSR223 Listener)\n```\n\n### 要素の役割と削除時の影響\n<!-- cols: 30,28,52,44 -->\n| 要素 | JMeter の要素 | 役割 | 削除したときの影響 |\n|---|---|---|---|\n| [必須] 共通変数 | User Defined Variables | 接続先（PROTOCOL / HOST / PORT）と共通変数を定義。値は `${__P(名前,既定値)}` 形式で、実行時に -J で変更可能 | 削除不可（各 API が ${HOST} 等を参照） |\n| [任意] HTTPリクエスト初期値 | HTTP Request Defaults | 文字コードと接続 / 応答タイムアウト（CONNECT_TIMEOUT / RESPONSE_TIMEOUT） | タイムアウトが無制限になる |\n| [任意] 共通ヘッダ | HTTP Header Manager | 全 API に付与するヘッダ（設定 common_headers） | 共通ヘッダが送信されない |\n| [任意] クッキー管理 | HTTP Cookie Manager | 応答の Set-Cookie を保存し、以降のリクエストで送信。事前設定クッキーも送信 | クッキーが送信されない（セッションを使う API は失敗） |\n| [任意] 事前処理 | setUp Thread Group | テスト開始前に 1 回だけ実行する処理の入れ物 | 配下の処理（Excel 読込など）も削除される |\n| [任意] テスト開始処理 | JSR223 Sampler | 実行条件をログに出力し、集計値をリセット | 開始ログが出ない（動作には影響なし） |\n| [任意] Excelデータ読込 | JSR223 Sampler | Excel（.xlsx）をメモリに読み込む（Apache POI 不要） | Excel の値が変数に設定されない |\n| [必須] メイン処理 | Thread Group | API を実行する。THREADS / RAMP_UP / LOOPS（/ DURATION）で制御 | 削除不可 |\n| [任意] 連番 | Counter | 反復ごとに連番を ${SEQ} に設定（全スレッドで一意） | ${SEQ} が未定義になる |\n| [任意] 乱数 | Random Variable | 反復ごとに乱数を ${RANDOM_NO} に設定 | ${RANDOM_NO} が未定義になる |\n| [任意] CSVデータ読込 | CSV Data Set Config | CSV を 1 反復につき 1 行読み込み、列を変数に設定 | CSV の変数が未定義になる |\n| [任意] Excelデータ行取得 | JSR223 PreProcessor | Excel データから反復ごとに 1 行を取り出して変数に設定 | Excel の変数が未定義になる |\n| [任意] 共通前処理 | JSR223 PreProcessor | 全 API の直前に実行。${REQUEST_ID} ${TIMESTAMP} ${EPOCH_MS} を生成 | 左記の変数が未定義になる |\n| [任意] 共通後処理 | JSR223 PostProcessor | 全 API の直後に実行。${LAST_STATUS} ${LAST_ELAPSED_MS} を設定 | 左記の変数が未定義になる |\n| [任意] スループット制御 | Constant Throughput Timer | 1 分あたりの目標リクエスト数（THROUGHPUT）に調整。既定は無効 | なし |\n| APIxx_名前 | HTTP Request | API 定義 1 行分のリクエスト | その API が実行されない（配下の要素も削除される） |\n| [任意] APIxx ヘッダ | HTTP Header Manager | その API だけに付けるヘッダ | そのヘッダが送信されない |\n| [任意] APIxx 待機 | Constant Timer | API 呼び出し前の待機時間 | 待機しなくなる |\n| [任意] APIxx 前処理 / 後処理 | JSR223 PreProcessor / PostProcessor | API ごとの前処理・後処理（Groovy） | その処理が行われない |\n| [任意] APIxx 正規表現抽出 / JSON抽出 | Regular Expression Extractor / JSON Extractor | 応答から値を取り出して変数に設定 | 後続の API でその変数が未定義になる |\n| [任意] APIxx ステータス検証 / 応答本文検証 / 応答時間検証 | Response Assertion / Duration Assertion | 応答の検証。失敗はエラーとして集計 | 検証されなくなる |\n| [任意] 変数確認 | Debug Sampler | 変数の一覧を結果ツリーに表示（デバッグ用）。既定は無効 | なし |\n| [任意] 事後処理 / テスト終了処理 | tearDown Thread Group / JSR223 Sampler | 合計・成功・失敗件数をログに出力し、結果ログを閉じる | 終了時のサマリが出ない |\n| [任意] 結果ツリー / 集計レポート / 統計レポート | View Results Tree / Summary Report / Aggregate Report | GUI での結果確認用（ファイル出力なし） | GUI で結果が見られない（CLI 実行には影響なし） |\n| [任意] JTL出力 全件 | Simple Data Writer | 全リクエストの結果を result_all.jtl（CSV）に出力 | JTL が出力されない（run コマンドは別途 result.jtl を出力） |\n| [任意] エラー詳細出力 XML | Simple Data Writer | 失敗したリクエストのみ、リクエスト・レスポンス全文を error_detail.xml に出力 | エラー詳細が出力されない |\n| [任意] 結果ログ出力 | JSR223 Listener | 1 リクエスト 1 行の結果ログ api_result.log を出力 | 結果ログが出力されない |\n\n### 記入例で使っている変数の流れ\n<!-- cols: 30,34,54 -->\n| 変数 | 設定する要素 | 使っている場所（記入例） |\n|---|---|---|\n| PROTOCOL / HOST / PORT | [必須] 共通変数 | 全 API の接続先 |\n| SEQ | [任意] 連番 | API01 のボディ、API03 の orderNo、API04 のクエリ |\n| RANDOM_NO | [任意] 乱数 | API02 のクエリ rand |\n| USER_ID / PASSWORD / KEYWORD | [任意] CSVデータ読込 | API01 のボディ、API02 のクエリ |\n| ITEM_CODE / ITEM_NAME / QUANTITY | [任意] Excelデータ行取得 | API02 のクエリ、API03 のボディ |\n| REQUEST_ID / TIMESTAMP | [任意] 共通前処理 | API01 のヘッダ・ボディ |\n| SIGNATURE | API01 前処理 | API01 のヘッダ X-Signature |\n| AUTH_TOKEN | API01 JSON抽出 | API02 / API03 の Authorization ヘッダ |\n| ORDER_NO | API03 正規表現抽出 | API03 後処理のログ出力 |\n| CLIENT_ID | 共通変数（設定 user_variables） | 共通ヘッダ X-Client-Id |\n\n### 要素を削除・無効化するときのポイント\n- 一時的に外す場合は、削除より「Disable」（右クリック → Disable、または Ctrl + T）が安全です。\n- 変数を設定している要素を削除すると、その変数を使っている箇所では ${変数名} という文字列がそのまま送信されます。上の「変数の流れ」を確認してから削除してください。\n- CSV データ読込は、ファイルが見つからないとスレッドがエラーで停止します。CSV を使わない場合は削除してください。\n- 生成時に設定を OFF にすると、その要素は最初から生成されません（「4.入力項目_設定」参照）。\n\n## 8.JMXの更新\n<!-- summary: JMeter で調整した JMX を、API 定義の変更に合わせて違う箇所だけ更新する（generate -u） -->\ngenerate に -u（PowerShell は -UpdateJmx）で生成済みの JMX を指定すると、新しい JMX は作らずに、入力ファイル（Excel / CSV / OpenAPI 定義）から生成した内容と比べて、違う箇所だけを既存の JMX に反映します。JMeter GUI で追加した要素や変更した値を残したまま、API の追加・変更・削除を取り込めます。更新する箇所は、書き換える前に一覧で表示します。\n\n### 更新手順\n```bash\n# RHEL 9.8\n# 1. 更新箇所を確認する（ファイルは書き換えない）\n./jmeter_api_tool.sh generate -i input/api_definition.xlsx -u output/api_test.jmx --dry-run\n# 2. 更新する\n./jmeter_api_tool.sh generate -i input/api_definition.xlsx -u output/api_test.jmx\n# OpenAPI 定義の場合\n./jmeter_api_tool.sh generate -i openapi.json -c input/openapi_config.csv -u output/OrdersApi.jmx\n```\n```powershell\n# Windows 11\n.\\jmeter_api_tool.bat generate -i input\\api_definition.xlsx -u output\\api_test.jmx -DryRun\n.\\jmeter_api_tool.bat generate -i input\\api_definition.xlsx -u output\\api_test.jmx\n.\\jmeter_api_tool.bat generate -i openapi.json -c input\\openapi_config.csv -u output\\OrdersApi.jmx\n```\n\n<!-- cols: 22,30,66 -->\n| sh のオプション | PowerShell のオプション | 説明 |\n|---|---|---|\n| -u JMX | -UpdateJmx（-u） | 更新する JMX（このツールで生成したもの）。同じ場所・同じファイル名のまま更新します |\n| --dry-run | -DryRun | 更新箇所を表示するだけで、ファイルは書き換えない |\n\n- -i・-s・-c・--seed など入力に関するオプションは、JMX を生成するときと同じです。-o・-n は使いません（指定しても無視します）。--split とは同時に指定できません。\n- 更新前の JMX は「JMX名.jmx.日時.bak」に退避します。更新した後は、基準ファイル（JMX名.jmx.base.json）と Linux サーバでの実行コマンド（JMX名_linux_cmd.txt）も作り直します。\n- 違う箇所が無い場合は、JMX を書き換えません。\n\n### 更新箇所の表示\n```text\n==== JMX の更新内容（output/api_test.jmx） ====\n比較: 前回生成した内容（output/api_test.jmx.base.json）と今回の定義の差分だけを反映します。（以下略）\n[変更] [必須] メイン処理 (Thread Group) / API02_商品検索(GET)\n       パラメータ / リクエストボディ:\n         + sort=new\n[競合] [必須] メイン処理 (Thread Group) / API01_ログイン(POST JSON)\n       パス: /post-gui → /post-v2\n       （JMeter 側でも変更されていましたが、定義の値で更新します）\n[削除] [必須] メイン処理 (Thread Group) / API04_クッキー確認(GET)\n       HTTP リクエスト（配下の要素 5 件を含む）\n[追加] [必須] メイン処理 (Thread Group) / API05_注文取消(DELETE)\n       HTTP リクエスト（配下の要素 2 件を含む）\n----\n追加 1 件 / 変更 1 件 / 削除 1 件 / 競合 1 件 / 維持 0 件（定義に変更が無く JMeter 側の変更を残した箇所: 2、JMeter で削除済みのため追加しなかった要素: 0）\n```\n\n<!-- cols: 14,104 -->\n| 表示 | 内容 |\n|---|---|\n| [追加] | 定義に追加された API・要素を JMX に追加します（生成した順番で直前にある要素の後ろ） |\n| [変更] | 定義が変わった項目（パス・パラメータ・ヘッダ・ボディ・検証・設定値など）を更新します。複数行の値は、- の行を削除し + の行を追加します |\n| [削除] | 定義から無くなった API・要素を、配下の要素ごと JMX から削除します |\n| [競合] | JMeter 側と定義の両方で変更されていた項目です。定義の値で更新します（JMeter 側の値は退避した .bak で確認できます） |\n| [維持] | ツールが作る共通のスクリプト（テスト開始処理・共通前処理など）を JMeter 側で変更していたため、定義の内容にせず JMeter 側の内容を残した項目です |\n\n### 更新のルール\n- 前回生成した内容（基準ファイル JMX名.jmx.base.json）・現在の JMX・今回の定義から生成した内容の 3 つを比べ、前回から定義が変わった箇所だけを JMX に反映します。定義が変わっていない箇所は、JMeter で変更していてもそのまま残します。\n- JMeter で追加した要素（このツールが付けない名前の要素）は、更新・削除の対象外です。JMeter で削除した要素は、定義が変わらない限り追加し直しません。\n- 要素は名前で対応付けます。HTTP リクエストは名前（APIxx_API名）が同じもの、無ければ番号（コメント欄の「APIxx:」または名前の先頭の APIxx_）、API名の順で探します（OpenAPI 定義の場合は、API名（メソッド パス [operationId]）を番号より先に使います）。[必須] / [任意] で始まる要素の名前を JMeter で変更すると、このツールの要素として扱われなくなり、更新されません。\n- JMeter で並べ替えた要素の順番は、そのまま残します（順番の違いは更新しません）。\n- JMeter で保存し直した JMX（書式・属性の順番が変わったもの）も更新できます。改行コード（LF / CRLF）は元の JMX に合わせます。\n- 基準ファイルが無い場合（ファイルを削除した場合や、基準ファイルを作らない以前の版で生成した JMX）は、現在の JMX と今回の定義を直接比べ、違う箇所を定義の内容に更新します。JMeter で変更した値も定義の値に戻り、JMeter で削除した要素も追加し直すため（JMeter で追加した要素と、共通のスクリプトの変更は残します）、先に --dry-run / -DryRun で確認してください。1 回更新すると基準ファイルが作られ、次からは JMeter の変更を残す比較になります。\n- OpenAPI 定義の場合は、同じ API名の API の番号（APIxx）を引き継ぎ、追加された API には既存の番号の続きを付けます。基準ファイルがあれば、操作の定義（パラメータ・リクエストボディ・参照先のスキーマ）が変わっていない API は前回と同じテストデータを使います。乱数シードは前回の値（基準ファイル、無ければ JMX のコメントに書かれた値）を使います（--seed / -Seed で変更できます）。\n\n## 9.実行方法\n<!-- summary: GUI / CLI での実行方法、run コマンド、実行時に変更できるプロパティ -->\n\n### 実行モードの使い分け\n<!-- cols: 16,46,56 -->\n| モード | 用途 | コマンド例 |\n|---|---|---|\n| GUI モード | シナリオの作成・動作確認・デバッグ（少量の実行） | `jmeter -t output/api_test.jmx` |\n| CLI モード | 負荷試験・大量件数の実行・定期実行 | `jmeter -n -t output/api_test.jmx -l result.jtl -e -o report` |\n\n### run コマンド（推奨）\nrun コマンドは CLI モードで JMeter を実行し、結果フォルダ（result/JMX名_日時/）に JTL・ログ・HTML レポートをまとめて出力し、最後に集計結果を表示します。\n\n```bash\n# RHEL 9.8\n# 既定（1スレッド × 1回）\n./jmeter_api_tool.sh run -t output/api_test.jmx\n# 10 スレッド × 100 回（1 API あたり 1,000 件）、ランプアップ 10 秒\n./jmeter_api_tool.sh run -t output/api_test.jmx --threads 10 --loops 100 --rampup 10\n# 接続先を切り替え、ヒープを 2GB にして実行\n./jmeter_api_tool.sh run -t output/api_test.jmx --threads 50 --loops 200 -J HOST=stg.example.com --heap 2g\n```\n\n```powershell\n# Windows 11\n.\\jmeter_api_tool.bat run -t output\\api_test.jmx\n.\\jmeter_api_tool.bat run -t output\\api_test.jmx -Threads 10 -Loops 100 -RampUp 10\n.\\jmeter_api_tool.bat run -t output\\api_test.jmx -Threads 50 -Loops 200 -Prop HOST=stg.example.com -Heap 2g\n```\n\n<!-- cols: 24,30,64 -->\n| sh のオプション | PowerShell のオプション | 説明 |\n|---|---|---|\n| -t ファイル | -Jmx（-t） | 実行する JMX。必須 |\n| --threads 数 | -Threads | スレッド数（THREADS） |\n| --rampup 秒 | -RampUp | ランプアップ秒数（RAMP_UP） |\n| --loops 数 | -Loops | ループ回数（LOOPS）。-1 で無限 |\n| --duration 秒 | -Duration | 実行時間（DURATION）。設定 duration を指定して生成した JMX のみ有効 |\n| -J 名前=値 | -Prop 名前=値,名前=値 | 任意のプロパティ。sh は複数回指定可、PowerShell はカンマ区切り |\n| -o フォルダ | -OutDir（-o） | 結果の出力先（既定: result） |\n| --heap サイズ | -Heap | JVM のヒープサイズ（例: 2g）。既定 1g |\n| --jmeter-home フォルダ | -JMeterHome | JMeter のフォルダ。省略時は環境変数 JMETER_HOME → PATH → 標準的な場所の順に探索 |\n| --fail-on-error | -FailOnError | 1 件でもエラーがあれば終了コード 2 を返す（CI 用） |\n| -- 以降の引数 | -ExtraArgs \"引数\" | JMeter にそのまま渡す追加オプション（例: プロキシ -H proxy.example.com -P 8080） |\n\n### 実行時に変更できるプロパティ\n<!-- cols: 34,32,52 -->\n| プロパティ | 既定値 | 説明 |\n|---|---|---|\n| THREADS | 設定 threads | 同時に実行するスレッド数（仮想ユーザー数） |\n| RAMP_UP | 設定 ramp_up | 全スレッドが起動するまでの秒数 |\n| LOOPS | 設定 loops | 各スレッドの繰り返し回数（-1 = 無限） |\n| DURATION | 設定 duration | 実行時間（秒）。duration を指定して生成した JMX のみ |\n| THROUGHPUT | 設定 throughput_per_min | スループット制御を有効にした場合の目標（件 / 分） |\n| PROTOCOL / HOST / PORT | URL から自動設定 | 接続先。2 つ目以降のホストは HOST_2、HOST_3 … |\n| CONNECT_TIMEOUT / RESPONSE_TIMEOUT | 10000 / 30000 | 接続 / 応答タイムアウト（ミリ秒） |\n| CSV_FILE | data/testdata.csv | CSV データファイル（相対パスは JMX のフォルダ基準） |\n| EXCEL_FILE / EXCEL_SHEET | data/testdata.xlsx / 先頭シート | Excel データファイルとシート名 |\n| RESULT_DIR | ~/result | JTL・結果ログの出力先（~/ は JMX のフォルダ。run コマンドは自動設定） |\n| LOG_RESPONSE | error | 結果ログへの本文出力（none / error / all） |\n| LOG_MAX_CHARS | 1000 | 結果ログに出力する本文の最大文字数 |\n| 共通変数の名前（例: ENV） | 設定 user_variables | 設定で定義した共通変数 |\n\n### Linux サーバで実行するコマンド（JMX名_linux_cmd.txt）\ngenerate で JMX を生成すると、同じフォルダに「JMX名_linux_cmd.txt」（例: api_test_linux_cmd.txt）を出力します。JMX を Linux サーバに持ち込み、このツールを使わずに JMeter の CLI で実行するためのテキストファイルで、次の内容を記載しています。\n- 前提（Java・JMeter のインストール）と、サーバに持ち込むファイル（JMX・CSV / Excel のデータファイル・アクセストークンのフォルダ）\n- そのまま貼り付けて実行できるコマンド（結果フォルダの作成と -n / -t / -l / -j / -e -o / -J の指定。複数行と 1 行の 2 通り）\n- JMeter のオプションと、JMX で使っているプロパティ（-J名前=値）ごとの説明・既定値\n- 実行後の確認方法と、実行中の停止方法\n\nJMeter で編集した JMX などから作り直す場合は、linux-cmd コマンドを使います。\n```bash\n./jmeter_api_tool.sh linux-cmd -t output/api_test.jmx       # RHEL 9.8\n```\n```powershell\n.\\jmeter_api_tool.bat linux-cmd -t output\\api_test.jmx      # Windows 11\n```\n\n<!-- cols: 24,30,64 -->\n| sh のオプション | PowerShell のオプション | 説明 |\n|---|---|---|\n| -t ファイル | -Jmx（-t） | 対象の JMX。必須 |\n| -o フォルダ | -OutDir（-o） | 出力先フォルダ（既定: JMX と同じフォルダ） |\n\n- コマンドの -J の値は、JMX に書かれた既定値（${__P(名前,既定値)} の既定値）です。接続先・スレッド数など、必要な値に書き換えてから実行してください。\n- JMX の ${__P(名前,既定値)} を読み取って作るため、JMeter で追加・変更したプロパティも反映されます。\n\n### JMeter を直接実行する場合\n```bash\n# RHEL 9.8（-f は既存の結果ファイル・レポートフォルダを削除してから実行）\njmeter -n -t output/api_test.jmx -l result/result.jtl -j result/jmeter.log \\\n  -e -o result/report -f -JTHREADS=10 -JLOOPS=100 -JRESULT_DIR=result\n```\n```powershell\n# Windows 11\n& \"$env:JMETER_HOME\\bin\\jmeter.bat\" -n -t output\\api_test.jmx -l result\\result.jtl -j result\\jmeter.log -e -o result\\report -f -JTHREADS=10 -JLOOPS=100 -JRESULT_DIR=result\n```\n- RESULT_DIR を相対パスで指定した場合は、JMeter を起動したフォルダ（カレントディレクトリ）が基準になります。\n- HTML レポートの出力先（-o）は、存在しないか空のフォルダである必要があります（-f を付けると自動で削除されます）。\n\n### GUI モードで実行する場合\n- メニューの「Run」→「Start」（Ctrl + R）で実行し、「結果ツリー」「集計レポート」で結果を確認します。\n- GUI で実行した場合も、JTL・結果ログは RESULT_DIR（既定: JMX と同じフォルダの result）に出力されます。\n- GUI での大量実行は避けてください。画面の描画で JMeter 自体が遅くなり、正しく測定できません。\n\n## 10.大量実行\n<!-- summary: 件数の考え方、負荷試験の推奨設定、JVM・OS の調整 -->\n\n### リクエスト件数の考え方\n総リクエスト件数 = THREADS（スレッド数）× LOOPS（ループ回数）× 有効な API の数\n\n<!-- cols: 22,50,46 -->\n| 目的 | 指定例 | 件数（API 4 本の場合） |\n|---|---|---|\n| 疎通確認（既定） | THREADS=1, LOOPS=1 | 4 件 |\n| 少量の連続実行 | THREADS=1, LOOPS=100 | 400 件 |\n| 並列実行 | THREADS=10, LOOPS=100 | 4,000 件 |\n| 時間指定の負荷試験 | 設定 duration=600 で生成し、THREADS=50, LOOPS=-1 | 600 秒間継続 |\n| CSV の行数だけ実行 | LOOPS=-1、設定 csv_recycle=false・csv_stop_thread=true | CSV 行数 × API 数 |\n\n### 負荷試験の推奨設定\n- CLI モード（run コマンド、または jmeter -n）で実行します。\n- [任意] 結果ログ出力は 1 件ごとにファイルへ書き込むため、毎秒数百件を超える場合は無効化するか、LOG_RESPONSE=none にします。\n- [任意] JTL出力 全件 は run コマンドの result.jtl と同じ内容を別ファイルに書き込むため、大量実行時は無効化してかまいません。\n- [任意] 結果ツリー / 集計レポート / 統計レポートは CLI では負荷になりにくいですが、不要なら削除または無効化します。\n- 待機時間（Constant Timer）で実際の操作間隔を再現し、スループット制御で目標の件数 / 分に合わせます。\n- ランプアップは「スレッド数 ÷ 1 秒あたりに増やしたいスレッド数」を目安にし、急激に負荷をかけないようにします。\n- Excel データは全行をメモリに読み込むため、数万行を超えるデータは CSV を使ってください。\n\n### JVM のヒープサイズ\n<!-- cols: 22,30,66 -->\n| 規模の目安 | ヒープ | 指定方法 |\n|---|---|---|\n| 〜100 スレッド | 1g（既定） | 指定不要 |\n| 〜500 スレッド | 2g 〜 4g | run コマンドの --heap 4g / -Heap 4g |\n| 500 スレッド超 | 4g 以上、または分散実行 | 同上。JMeter を直接実行する場合は環境変数 HEAP（例: -Xms4g -Xmx4g -XX:MaxMetaspaceSize=256m） |\n\n### OS の調整（RHEL 9.8）\n```bash\n# ファイルディスクリプタの上限（スレッド数が多い場合）\nulimit -n                 # 現在値の確認\nulimit -n 65535           # 一時的に変更（恒久的には /etc/security/limits.conf）\n# エフェメラルポートの範囲\nsysctl net.ipv4.ip_local_port_range\n```\n\n### OS の調整（Windows 11）\n- 動的ポートの範囲は `netsh int ipv4 show dynamicport tcp` で確認できます。\n- 長時間の試験では、スリープ・省電力の設定を無効にしてから実行します。\n- ウイルス対策ソフトのリアルタイムスキャンが結果ファイルの書き込みを遅くする場合があります（除外設定は社内のルールに従ってください）。\n\n### 分散実行（参考）\n1 台で必要な負荷をかけられない場合は、JMeter の分散実行（コントローラ 1 台 + 負荷生成サーバー複数台）を利用します。各サーバーで `jmeter-server` を起動し、`jmeter -n -t output/api_test.jmx -R サーバー1,サーバー2` で実行します。CSV / Excel のデータファイルは各サーバーの同じパスに配置が必要です。詳細は JMeter 公式マニュアルの「Remote Testing」を参照してください。\n\n## 11.結果確認\n<!-- summary: 出力ファイルの場所と見方（HTML レポート・JTL・ログ） -->\n\n### run コマンドの出力先\n```text\nresult/api_test_20261011_103000/\n├─ result.jtl          全リクエストの結果（CSV 形式）\n├─ jmeter.log          JMeter の実行ログ\n├─ report/index.html   HTML ダッシュボードレポート\n├─ api_result.log      結果ログ（[任意] 結果ログ出力）\n├─ result_all.jtl      JTL（[任意] JTL出力 全件）\n└─ error_detail.xml    失敗したリクエストの詳細（[任意] エラー詳細出力。失敗がある場合のみ）\n```\n\n### HTML レポート（report/index.html）\n<!-- cols: 28,90 -->\n| 画面 | 確認するポイント |\n|---|---|\n| Dashboard | APDEX（応答の満足度）、成功 / 失敗の割合、Statistics（API ごとの件数・エラー率・平均・90/95/99 パーセンタイル・スループット）、Errors（エラーの種類） |\n| Charts → Over Time | 応答時間・スループットの時間推移（性能劣化やエラー急増のタイミング） |\n| Charts → Throughput | 1 秒あたりのリクエスト数、応答コード別の件数 |\n| Charts → Response Times | 応答時間の分布、パーセンタイル |\n\nJTL から HTML レポートを作り直す場合は `jmeter -g result.jtl -o report2` を実行します。\n\n### JTL（result.jtl）の主な列\n<!-- cols: 22,96 -->\n| 列 | 内容 |\n|---|---|\n| timeStamp | 開始時刻（エポックミリ秒） |\n| elapsed | 応答時間（ミリ秒） |\n| label | API 名（HTTP リクエストの名前） |\n| responseCode / responseMessage | HTTP ステータスとメッセージ |\n| threadName | スレッド名 |\n| success | 成否（true / false） |\n| failureMessage | 検証（Assertion）エラーの内容 |\n| bytes / sentBytes | 受信 / 送信バイト数 |\n| grpThreads / allThreads | 実行中のスレッド数 |\n| URL | リクエストの URL |\n| Latency / Connect / IdleTime | 最初の応答までの時間 / 接続時間 / 待機時間（ミリ秒） |\n\n### 結果ログ（api_result.log）\nタブ区切りのテキストで、1 リクエスト 1 行です。列は「日時・結果（OK / NG）・ステータス・応答時間ms・受信バイト・ラベル・リクエスト（メソッドと URL）・スレッド・メッセージ」です。失敗時（LOG_RESPONSE=all の場合は全件）は次の行に [REQUEST] / [RESPONSE] として送信内容と応答本文（LOG_MAX_CHARS 文字まで）を出力します。最終行には合計・成功・失敗件数が出力されます。\n\n### エラー詳細（error_detail.xml）\n失敗したリクエストだけを、リクエストヘッダ・本文とレスポンスヘッダ・本文付きで出力します。JMeter GUI の「結果ツリー（View Results Tree）」の「Filename」欄でこのファイルを開くと、内容を画面で確認できます。\n\n### jmeter.log\n- 「==== テスト開始」「==== テスト終了」の行に、実行条件と合計・成功・失敗件数が出力されます。\n- 失敗したリクエストは「[API NG]」で始まる警告として出力されます。\n- スクリプトのエラーは ERROR / WARN として出力されます。うまく動かない場合はまずこのファイルを確認してください。\n\n## 12.変数と関数\n<!-- summary: 生成される変数、よく使う関数、JSR223（Groovy）・抽出の書き方 -->\n\n### 生成される変数\n<!-- cols: 34,26,58 -->\n| 変数 | 設定する要素 | 内容 |\n|---|---|---|\n| `${PROTOCOL}` `${HOST}` `${PORT}` | 共通変数 | 接続先 |\n| `${SEQ}` | 連番 | 000001 から始まる連番（全スレッドで一意） |\n| `${RANDOM_NO}` | 乱数 | 1 〜 99999 の乱数 |\n| `${REQUEST_ID}` | 共通前処理 | リクエストごとの UUID |\n| `${TIMESTAMP}` / `${EPOCH_MS}` | 共通前処理 | 現在時刻（ISO 8601 形式 / エポックミリ秒） |\n| `${LAST_STATUS}` / `${LAST_ELAPSED_MS}` | 共通後処理 | 直前の API のステータスと応答時間 |\n| `${EXCEL_ROW_NO}` | Excelデータ行取得 | 使用中の Excel の行番号 |\n| CSV / Excel の列名 | CSVデータ読込 / Excelデータ行取得 | 1 行目の列名がそのまま変数名になる |\n\n### よく使う関数\n<!-- cols: 18,46,54 -->\n| 関数 | 記入例 | 結果 |\n|---|---|---|\n| `__P` | `${__P(THREADS,1)}` | プロパティの値（未指定なら既定値） |\n| `__UUID` | `${__UUID()}` | ランダムな UUID |\n| `__time` | `${__time(yyyyMMddHHmmss,)}` | 現在時刻（書式指定） |\n| `__Random` | `${__Random(1,100,)}` | 1 〜 100 の乱数 |\n| `__RandomString` | `${__RandomString(8,abcdef0123456789,)}` | 指定した文字からなるランダムな文字列 |\n| `__counter` | `${__counter(FALSE,)}` | 全スレッド共通の連番（TRUE にするとスレッドごと） |\n| `__threadNum` | `${__threadNum}` | スレッド番号 |\n| `__urlencode` | `${__urlencode(${KEYWORD})}` | URL エンコード |\n| `__digest` | `${__digest(SHA-256,${USER_ID},,,)}` | ハッシュ値 |\n| `__timeShift` | `${__timeShift(yyyy-MM-dd,,P1D,,)}` | 日付の計算（例: 翌日） |\n| `__groovy` | `${__groovy(vars.get('SEQ').toInteger() * 2,)}` | Groovy の式の結果 |\n| `__V` | `${__V(USER_${SEQ})}` | 変数名を組み立てて参照 |\n\n### JSR223（Groovy）の書き方\n```groovy\n// 変数の取得・設定（値は文字列）\nString userId = vars.get('USER_ID')\nvars.put('USER_ID_UPPER', userId.toUpperCase())\n\n// プロパティ（全スレッド共通）の取得・設定\nString host = props.getProperty('HOST', 'localhost')\nprops.put('SHARED_TOKEN', vars.get('AUTH_TOKEN'))\n\n// 直前の応答（後処理・リスナーで使用）\nString code = prev.getResponseCode()\ndef json = new groovy.json.JsonSlurper().parseText(prev.getResponseDataAsString())\nvars.put('ORDER_ID', String.valueOf(json.data.id))\n\n// ログ出力（jmeter.log）\nlog.info('order id = ' + vars.get('ORDER_ID'))\n```\n- JSR223 のスクリプト内では `${変数名}` を使わず、`vars.get('変数名')` を使います（${} を使うとスクリプトのキャッシュが効かず、値によっては構文エラーになります）。\n- 言語は groovy、「Cache compiled script if available」はオンにします（生成した JMX は設定済み）。\n- vars はスレッドごと、props は全スレッド共通です。スレッドをまたいで値を共有する場合は props を使います。\n\n### 正規表現抽出の書き方\n<!-- cols: 40,78 -->\n| 抽出したいもの | 正規表現の例 |\n|---|---|\n| JSON の \"token\":\"abc\" の abc | `\"token\"\\s*:\\s*\"([^\"]+)\"` |\n| 数値の項目 \"id\": 123 の 123 | `\"id\"\\s*:\\s*(\\d+)` |\n| HTML の hidden 項目の値 | `name=\"csrf\" value=\"([^\"]+)\"` |\n| 応答ヘッダの値（GUI で対象を Response Headers に変更） | `Location: (.+)` |\n- ( ) で囲んだ部分が変数に入ります（テンプレート `$1$`）。見つからない場合は NOT_FOUND が入ります。\n- 一致したものをすべて取得するには、GUI で「Match No.」を -1 にします。`変数名_1`、`変数名_2` … と、件数を表す `変数名_matchNr` が作られます。\n\n### JSON 抽出（JSONPath）の書き方\n<!-- cols: 40,78 -->\n| 抽出したいもの | JSONPath |\n|---|---|\n| ルートの id | `$.id` |\n| 入れ子の値 | `$.data.user.name` |\n| 配列の先頭要素 | `$.items[0].id` |\n| 条件に一致する要素 | `$.items[?(@.status == 'active')].id` |\n| すべての id（Match No. を -1 にする） | `$..id` |\n\n## 13.トラブル対応\n<!-- summary: よくあるエラーと対処方法 -->\n<!-- cols: 36,36,48 -->\n| 症状 | 主な原因 | 対処 |\n|---|---|---|\n| `python3: command not found`（RHEL） | python3 が未導入 | `sudo dnf install -y python3` |\n| `$'\\r': command not found`（RHEL） | 改行コードが CRLF になっている | `sed -i 's/\\r$//' jmeter_api_tool.sh` |\n| このシステムではスクリプトの実行が無効になっている（Windows） | PowerShell の実行ポリシー | jmeter_api_tool.bat から実行する |\n| 日本語が文字化けする（コンソール） | 端末の文字コード | RHEL は `export LANG=ja_JP.UTF-8` など UTF-8 の端末を使用 |\n| JMeter が見つからない | JMETER_HOME が未設定 | 環境変数 JMETER_HOME を設定するか、--jmeter-home / -JMeterHome で指定 |\n| Java が見つからない・バージョンが古い | Java が未導入、PATH が未設定 | Java 17 をインストールし、JAVA_HOME と PATH を設定 |\n| `${USER_ID}` などがそのまま送信される | 変数を設定する要素が削除・無効、または変数名の誤り | 結果ツリーでリクエストを確認し、変数名と設定元の要素を見直す |\n| CSV の値が `<EOF>` になる | CSV の最終行に達した（csv_recycle=false） | csv_recycle=true にするか、ループ回数を見直す |\n| `File … must exist and be readable` | CSV ファイルが見つからない | CSV_FILE を JMX からの相対パス、または絶対パスで指定 |\n| 応答の日本語が文字化けする | 応答に文字コードの指定がない | JMeter 5.6.3 は既定で UTF-8。異なる場合は `-Jsampleresult.default.encoding=文字コード` |\n| Connection refused / Connect timed out | 接続先・ポートの誤り、ファイアウォール、プロキシ | URL を確認。プロキシ経由の場合は -H プロキシ -P ポート を追加（run は -- / -ExtraArgs で指定） |\n| SSLHandshakeException | TLS のバージョン・証明書の問題 | jmeter.log で詳細を確認。クライアント証明書が必要な場合は system.properties で keystore を設定 |\n| `java.lang.OutOfMemoryError` | ヒープ不足 | --heap / -Heap で増やす。CLI で実行し、結果ツリーを無効化 |\n| `BindException: Address already in use` | クライアント側のポート枯渇 | KeepAlive を有効（既定）にする、スレッド数を減らす、OS のポート範囲を広げる |\n| `Too many open files`（RHEL） | ファイルディスクリプタの上限 | `ulimit -n` で上限を引き上げる |\n| HTML レポートが作成されない | JTL が空、出力先フォルダが空でない | jmeter.log を確認。`jmeter -g result.jtl -o 新しいフォルダ` で再作成 |\n| Excel のデータが読み込まれない | ファイルパス・シート名の誤り、.xls 形式 | jmeter.log の [Excel] の行を確認。.xlsx 形式で保存する（.xls は非対応） |\n| `JSON の構文エラーです（YAML 形式は未対応です）`（generate） | OpenAPI 定義が YAML、または JSON の書式誤り | YAML は JSON に変換してから指定する。JSON はエラーメッセージの行・列を確認して修正 |\n| `Swagger 2.0 形式には対応していません`（generate） | Swagger 2.0（OpenAPI 2.0）の定義 | OpenAPI 3.0.x 形式に変換してから指定する |\n| `curl コマンドとして読めません`（add-curl） | 先頭が curl ではない、またはコマンドの貼り付け誤り | curl で始まるコマンドを指定する（ブラウザの「Copy as cURL」の内容をそのまま保存したファイルを --curl-file / -CurlFile で指定すると確実） |\n| `CSV の文字コード（Shift_JIS(CP932)）で表せない文字が含まれています`（add-curl） | Shift_JIS の CSV に、Shift_JIS に無い文字を追加しようとした | CSV を UTF-8 で保存し直してから実行する |\n| `jmeter_api_tool で生成した JMX ではありません`（generate -u） | このツールで生成していない JMX、または「[必須] メイン処理」の名前を変更した | このツールで生成した JMX を指定する（新しく作る場合は -u / -UpdateJmx を付けない） |\n| JMeter で変更した値が更新で元に戻った（generate -u） | 基準ファイル（JMX名.jmx.base.json）が無かった、または [競合] の項目だった | 表示された [競合] / [変更] を確認し、退避した .bak から値を戻す。基準ファイルは JMX と同じフォルダに残しておく |\n\n## 14.コマンド一覧\n<!-- summary: ツールと JMeter のコマンドラインオプション -->\n\n### ツールのコマンド\n<!-- cols: 12,46,46,30 -->\n| コマンド | RHEL 9.8（sh） | Windows 11（bat / ps1） | 説明 |\n|---|---|---|---|\n| init | `./jmeter_api_tool.sh init [-o フォルダ] [-f]` | `jmeter_api_tool.bat init [-OutDir フォルダ] [-Force]` | 入力テンプレートを作成（既定: input） |\n| generate | `./jmeter_api_tool.sh generate -i ファイル [-s 設定CSV] [-o フォルダ] [-n ファイル名] [--split] [-f] [-c OpenAPI用設定] [--seed N] [-u JMX [--dry-run]]` | `jmeter_api_tool.bat generate -i ファイル [-s 設定CSV] [-OutDir フォルダ] [-JmxName ファイル名] [-Split] [-Force] [-c OpenAPI用設定] [-Seed N] [-UpdateJmx JMX [-DryRun]]` | JMX を生成（既定: output）。OpenAPI 用のオプションは「6.OpenAPI定義から生成」、-u（既存の JMX の差分更新）は「8.JMXの更新」参照 |\n| add-curl | `./jmeter_api_tool.sh add-curl -i API一覧 (--curl-file ファイル / --curl \"curl …\" / -- curl …) [--name API名] [--dry-run]` | `jmeter_api_tool.bat add-curl -i API一覧 (-CurlFile ファイル / -Curl \"curl …\") [-ApiName API名] [-DryRun]` | curl コマンドを解析して API一覧に 1 行追加（「3.JMX生成」参照） |\n| linux-cmd | `./jmeter_api_tool.sh linux-cmd -t JMX [-o フォルダ]` | `jmeter_api_tool.bat linux-cmd -t JMX [-OutDir フォルダ]` | Linux サーバで実行するコマンドと説明を出力（「9.実行方法」参照） |\n| manual | `./jmeter_api_tool.sh manual [-o フォルダ]` | `jmeter_api_tool.bat manual [-OutDir フォルダ]` | 本書（xlsx / md）を出力（既定: docs） |\n| run | `./jmeter_api_tool.sh run -t JMX [オプション]` | `jmeter_api_tool.bat run -t JMX [オプション]` | CLI モードで実行（既定: result） |\n| help | `./jmeter_api_tool.sh help` | `jmeter_api_tool.bat help` | ヘルプを表示 |\n\n### JMeter の主なオプション\n<!-- cols: 24,94 -->\n| オプション | 説明 |\n|---|---|\n| -n | CLI（非 GUI）モードで実行 |\n| -t ファイル | 実行する JMX |\n| -l ファイル | 結果（JTL）の出力先 |\n| -j ファイル | ログ（jmeter.log）の出力先 |\n| -e | テスト終了後に HTML レポートを作成 |\n| -o フォルダ | HTML レポートの出力先（存在しないか空のフォルダ） |\n| -f | 既存の結果ファイル・レポートフォルダを削除してから実行 |\n| -g ファイル | 既存の JTL から HTML レポートだけを作成（-o と併用） |\n| -J名前=値 | JMeter プロパティを設定（${__P(名前,既定値)} で参照） |\n| -G名前=値 | 分散実行時に全サーバーへプロパティを設定 |\n| -D名前=値 | Java のシステムプロパティを設定 |\n| -q ファイル | 追加のプロパティファイルを読み込む |\n| -H ホスト / -P ポート | プロキシサーバーのホスト / ポート |\n| -R ホスト1,ホスト2 | 分散実行（指定したサーバーで実行） |\n| -L レベル | ログレベル（例: -LDEBUG） |\n| -v | バージョンを表示 |\n",
 "template_guide_md": "## 記入方法（API定義テンプレート）\n\n### 基本的な使い方\n1. 「API一覧」シートに、呼び出す API を 1行 = 1API で記入します。上から順番に実行されます。\n2. 「設定」シートの「設定値」列で、スレッド数・ループ回数などの全体設定を変更します（空欄は既定値）。\n3. 保存後、次のコマンドで JMX を生成します。\n\n```text\nRHEL 9.8  : ./jmeter_api_tool.sh generate -i input/api_definition.xlsx\nWindows 11: jmeter_api_tool.bat generate -i input\\api_definition.xlsx\n```\n\n### 記入のルール\n- 「API一覧」の 4 行目（No 列が # で始まる行）は記入説明のため読み込まれません。API は 5 行目以降に記入します。行の追加・削除は自由です。\n- No 列が # で始まる行はコメントとして読み飛ばします。一時的に外したい API は No の先頭に # を付けるか、「有効」を N にします。\n- 複数の値を書く列（クエリパラメータ、ヘッダ、正規表現抽出など）は、セル内改行（Alt + Enter）で区切ります。\n- URL・パラメータ・ヘッダ・ボディには `${変数名}` で変数を埋め込めます。\n- シート名（API一覧・設定）と見出し行の列名は変更しないでください。列の並び順は入れ替えても読み込めます。\n- 入力セルは書式を「文字列」にしてあるため、0 始まりの値やカンマを含む値もそのまま保存されます。\n\n### API一覧の列\n<!-- @API_COLUMNS_TABLE@ -->\n\n### 記入例で使っている変数\n<!-- cols: 30,34,56 -->\n| 変数 | 設定する要素 | 内容 |\n|---|---|---|\n| `${USER_ID}` `${PASSWORD}` `${KEYWORD}` | CSVデータ読込（data/testdata.csv） | CSV の 1 行目の列名がそのまま変数名 |\n| `${ITEM_CODE}` `${ITEM_NAME}` `${QUANTITY}` | Excelデータ行取得（data/testdata.xlsx） | Excel の 1 行目の列名がそのまま変数名 |\n| `${SEQ}` / `${RANDOM_NO}` | 連番 / 乱数 | 反復ごとの連番（000001〜）/ 乱数 |\n| `${REQUEST_ID}` `${TIMESTAMP}` | 共通前処理 | リクエストごとの UUID / 現在時刻 |\n| `${AUTH_TOKEN}` | API01 の JSON抽出 | 前の API の応答から取り出した値 |\n| `${CLIENT_ID}` | 設定 user_variables | 共通変数 |\n\n### 生成されるもの\n- 生成された JMX の要素名は [必須] / [任意] で始まります。[任意] の要素は不要なら JMeter GUI で削除できます。\n- 詳しい使い方は利用方法（docs/JMeter_Usage_Guide.xlsx / .md）を参照してください。\n"
}''')
TOOL = SPEC['tool']
VERSION = TOOL['version']
XL = SPEC['xlsx']
TPL = SPEC['template']
NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
NS_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'


# ----------------------------------------------------------------------------- console
def _reopen(stream):
    try:
        return io.TextIOWrapper(stream.buffer, encoding=(stream.encoding or 'utf-8'),
                                errors='replace', line_buffering=True)
    except Exception:
        return stream


sys.stdout = _reopen(sys.stdout)
sys.stderr = _reopen(sys.stderr)
WARNINGS = []


class ToolError(Exception):
    pass


def info(msg):
    print('[INFO] ' + msg)


def warn(msg):
    WARNINGS.append(msg)
    print('[WARN] ' + msg)


# ----------------------------------------------------------------------------- text utilities
_BAD_XML = re.compile('[\x00-\x08\x0b\x0c\x0e-\x1f\ufffe\uffff]')


def norm_nl(s):
    return (s or '').replace('\r\n', '\n').replace('\r', '\n')


def clean(s):
    return _BAD_XML.sub('', norm_nl(s))


def xesc(s):
    return (clean(s).replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
            .replace('"', '&quot;').replace("'", '&apos;'))


def fill(template, values):
    s = template
    for k in sorted(values):
        s = s.replace('@@' + k + '@@', values[k])
    return s


def normalize_key(s):
    s = unicodedata.normalize('NFKC', s or '')
    return re.sub(r'\s+', '', s).lower()


def java_hash(s):
    h = 0
    b = s.encode('utf-16-le')
    for i in range(0, len(b), 2):
        h = (31 * h + (b[i] | (b[i + 1] << 8))) % 4294967296
    return h - 4294967296 if h >= 2147483648 else h


def char_units(cp):
    u = XL['units']
    if cp < 0x80:
        if 0x41 <= cp <= 0x5A:
            return u['upper']
        if 0x30 <= cp <= 0x39:
            return u['digit']
        if 0x61 <= cp <= 0x7A:
            return u['lower']
        return u['punct']
    if cp < 0x1100:
        return u['latin']
    if 0xFF61 <= cp <= 0xFF9F:
        return u['hankaku']
    if 0x3040 <= cp <= 0x30FF:
        return u['kana']
    return u['wide']


def text_units(s):
    return sum(char_units(ord(ch)) for ch in s)


def wrap_lines(text, width):
    usable = max(XL['units']['digit'], width * 20 - 30) * 100
    n = 0
    for line in norm_nl(text).split('\n'):
        u = text_units(line) * 105
        n += max(1, (u + usable - 1) // usable)
    return n


def row_h10(lines, pad=None):
    h = lines * XL['line_h10'] + (XL['pad_h10'] if pad is None else pad)
    return min(h, XL['max_h10'])


def today():
    return datetime.date.today().strftime('%Y-%m-%d')


def timestamp():
    return datetime.datetime.now().strftime('%Y%m%d_%H%M%S')


# ----------------------------------------------------------------------------- files / csv
def read_text(path):
    with open(path, 'rb') as f:
        data = f.read()
    if data[:3] == b'\xef\xbb\xbf':
        return norm_nl(data[3:].decode('utf-8', 'replace')), 'UTF-8(BOM)'
    if data[:2] == b'\xff\xfe':
        return norm_nl(data[2:].decode('utf-16-le', 'replace')), 'UTF-16LE'
    if data[:2] == b'\xfe\xff':
        return norm_nl(data[2:].decode('utf-16-be', 'replace')), 'UTF-16BE'
    try:
        return norm_nl(data.decode('utf-8')), 'UTF-8'
    except UnicodeDecodeError:
        pass
    try:
        return norm_nl(data.decode('cp932')), 'Shift_JIS(CP932)'
    except UnicodeDecodeError:
        raise ToolError('文字コードを判別できません（UTF-8 または Shift_JIS で保存してください）: ' + path)


def write_file(path, data):
    try:
        d = os.path.dirname(path)
        if d and not os.path.isdir(d):
            os.makedirs(d)
        with open(path, 'wb') as f:
            f.write(data)
    except (IOError, OSError) as e:
        raise ToolError('ファイルに書き込めません（権限・空き容量・他のアプリで開いていないか確認してください）: %s (%s)' % (path, e))


def write_text(path, text, bom=False):
    data = text.encode('utf-8')
    write_file(path, (b'\xef\xbb\xbf' + data) if bom else data)


def parse_csv(text, delim=','):
    rows, row, field = [], [], []
    inq = False
    started = False
    i, n = 0, len(text)
    while i < n:
        c = text[i]
        if inq:
            if c == '"':
                if i + 1 < n and text[i + 1] == '"':
                    field.append('"')
                    i += 2
                    continue
                inq = False
            else:
                field.append(c)
            i += 1
            continue
        if c == '"' and not field:
            inq = True
            started = True
        elif c == delim:
            row.append(''.join(field))
            field = []
            started = True
        elif c == '\n':
            row.append(''.join(field))
            rows.append(row)
            row, field = [], []
            started = False
        else:
            field.append(c)
            started = True
        i += 1
    if started or field or row:
        row.append(''.join(field))
        rows.append(row)
    return rows


def csv_field(s):
    s = s or ''
    if (',' in s or '"' in s or '\n' in s or '\r' in s) or s != s.strip(' '):
        return '"' + s.replace('"', '""') + '"'
    return s


def to_csv(rows):
    return ''.join(','.join(csv_field(f) for f in r) + '\r\n' for r in rows)


# ----------------------------------------------------------------------------- xlsx reader
def _ln(tag):
    return tag.rsplit('}', 1)[-1] if '}' in tag else tag


def _xl_text(el):
    parts = []

    def walk(e):
        for c in e:
            ln = _ln(c.tag)
            if ln == 'rPh':
                continue
            if ln == 't':
                parts.append(c.text or '')
            else:
                walk(c)
    walk(el)
    return ''.join(parts)


def _col_index(ref):
    n = 0
    for ch in ref:
        if 'A' <= ch <= 'Z':
            n = n * 26 + (ord(ch) - 64)
        elif 'a' <= ch <= 'z':
            n = n * 26 + (ord(ch) - 96)
        else:
            break
    return n - 1


def _xl_rows(root, shared):
    rows = []
    for sd in root:
        if _ln(sd.tag) != 'sheetData':
            continue
        last = 0
        for row in sd:
            if _ln(row.tag) != 'row':
                continue
            r = row.get('r') or ''
            rnum = int(r) if re.match(r'^[0-9]+$', r) else last + 1
            last = rnum
            cells = {}
            nextc = 0
            for c in row:
                if _ln(c.tag) != 'c':
                    continue
                ref = c.get('r') or ''
                ci = _col_index(ref) if ref else nextc
                if ci < 0:
                    ci = nextc
                nextc = ci + 1
                t = c.get('t') or ''
                val = ''
                if t == 'inlineStr':
                    for x in c:
                        if _ln(x.tag) == 'is':
                            val = _xl_text(x)
                else:
                    v = ''
                    for x in c:
                        if _ln(x.tag) == 'v':
                            v = x.text or ''
                    if t == 's':
                        try:
                            val = shared[int(v.strip())]
                        except (ValueError, IndexError):
                            val = ''
                    elif t == 'b':
                        val = 'TRUE' if v.strip() == '1' else 'FALSE'
                    else:
                        val = v
                cells[ci] = val
            width = (max(cells) + 1) if cells else 0
            rows.append((rnum, [cells.get(i, '') for i in range(width)]))
    return rows


def xlsx_read(path):
    try:
        z = zipfile.ZipFile(path)
    except Exception:
        raise ToolError('Excel ファイルを開けません（.xlsx 形式で保存してください）: ' + path)
    with z:
        names = set(z.namelist())

        def load(name):
            if name not in names:
                return None
            return ET.fromstring(z.read(name))
        shared = []
        sst = load('xl/sharedStrings.xml')
        if sst is not None:
            for si in sst:
                if _ln(si.tag) == 'si':
                    shared.append(_xl_text(si))
        wb = load('xl/workbook.xml')
        rels = load('xl/_rels/workbook.xml.rels')
        if wb is None or rels is None:
            raise ToolError('Excel ファイルの形式が正しくありません: ' + path)
        relmap = {}
        for r in rels:
            if _ln(r.tag) == 'Relationship':
                relmap[r.get('Id')] = r.get('Target') or ''
        sheets = []
        for el in wb.iter():
            if _ln(el.tag) != 'sheet':
                continue
            rid = ''
            for k, v in el.attrib.items():
                if k.startswith('{') and _ln(k) == 'id':
                    rid = v
            target = relmap.get(rid, '')
            if not target:
                continue
            part = target[1:] if target.startswith('/') else 'xl/' + target
            root = load(part)
            sheets.append((el.get('name') or '', _xl_rows(root, shared) if root is not None else []))
        return sheets


# ----------------------------------------------------------------------------- xlsx writer
XF_NAMES = [x[0] for x in SPEC['styles']['xfs']]
XF_INDEX = dict((name, i) for i, name in enumerate(XF_NAMES))
XF_FONT = dict((x[0], x[1]) for x in SPEC['styles']['xfs'])


def col_letter(c):
    s = ''
    while c > 0:
        c, rem = divmod(c - 1, 26)
        s = chr(65 + rem) + s
    return s


def cref(r, c):
    return col_letter(c) + str(r)


def fmt_h10(h):
    return str(h // 10) if h % 10 == 0 else '%d.%d' % (h // 10, h % 10)


def plain_of(v):
    if v is None:
        return ''
    if isinstance(v, str):
        return v
    return ''.join(t for t, b in v)


class Sheet(object):
    def __init__(self, name):
        self.name = name
        self.cells = {}
        self.h10 = {}
        self.widths = {}
        self.merges = []
        self.validations = []
        self.links = []
        self.freeze_rows = 0
        self.freeze_cols = 0
        self.selected = False
        self.paper = 9
        self.fit_width = 1

    def put(self, r, c, value, style):
        self.cells[(r, c)] = (value, style)

    def block(self, r, c1, c2, value, style):
        self.put(r, c1, value, style)
        for c in range(c1 + 1, c2 + 1):
            self.put(r, c, None, style)
        if c2 > c1:
            self.merges.append((r, c1, r, c2))


def _t_xml(s, force_preserve=False):
    s = clean(s)
    sp = ' xml:space="preserve"' if (force_preserve or (s and (s[0] in ' \t\n' or s[-1] in ' \t\n'))) else ''
    return '<t%s>%s</t>' % (sp, xesc(s))


def _rpr_xml(font_idx, bold):
    f = SPEC['styles']['fonts'][font_idx]
    b = bold or f.get('b', False)
    return '<rPr>%s<sz val="%s"/><color rgb="%s"/><rFont val="%s"/><family val="3"/><charset val="128"/></rPr>' % (
        '<b/>' if b else '', f['sz'], f['color'], XL['font_name'])


def _pane_xml(fr, fc):
    if not fr and not fc:
        return ''
    tl = cref(fr + 1, fc + 1)
    if fr and fc:
        return ('<pane xSplit="%d" ySplit="%d" topLeftCell="%s" activePane="bottomRight" state="frozen"/>'
                '<selection pane="topRight"/><selection pane="bottomLeft"/>'
                '<selection pane="bottomRight" activeCell="%s" sqref="%s"/>') % (fc, fr, tl, tl, tl)
    if fr:
        return ('<pane ySplit="%d" topLeftCell="%s" activePane="bottomLeft" state="frozen"/>'
                '<selection pane="bottomLeft" activeCell="%s" sqref="%s"/>') % (fr, tl, tl, tl)
    return ('<pane xSplit="%d" topLeftCell="%s" activePane="topRight" state="frozen"/>'
            '<selection pane="topRight" activeCell="%s" sqref="%s"/>') % (fc, tl, tl, tl)


def _sheet_xml(sh, sidx):
    by_row = {}
    for (r, c) in sh.cells:
        by_row.setdefault(r, []).append(c)
    rows = sorted(set(list(by_row.keys()) + list(sh.h10.keys())))
    maxr = max([r for (r, c) in sh.cells] or [1])
    maxc = max([c for (r, c) in sh.cells] or [1])
    o = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n',
         '<worksheet xmlns="%s" xmlns:r="%s">' % (NS_MAIN, NS_REL),
         '<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>',
         '<dimension ref="A1:%s"/>' % cref(maxr, maxc),
         '<sheetViews><sheetView showGridLines="0"%s workbookViewId="0">%s</sheetView></sheetViews>' % (
             ' tabSelected="1"' if sh.selected else '', _pane_xml(sh.freeze_rows, sh.freeze_cols)),
         '<sheetFormatPr defaultRowHeight="14.4"/>']
    if sh.widths:
        o.append('<cols>')
        for c in sorted(sh.widths):
            o.append('<col min="%d" max="%d" width="%s" customWidth="1"/>' % (c, c, sh.widths[c]))
        o.append('</cols>')
    o.append('<sheetData>')
    for r in rows:
        attrs = ' r="%d"' % r
        if r in sh.h10:
            attrs += ' ht="%s" customHeight="1"' % fmt_h10(sh.h10[r])
        cols = sorted(by_row.get(r, []))
        if not cols:
            o.append('<row%s/>' % attrs)
            continue
        o.append('<row%s>' % attrs)
        for c in cols:
            value, style = sh.cells[(r, c)]
            s = XF_INDEX[style]
            if value is None or plain_of(value) == '':
                o.append('<c r="%s" s="%d"/>' % (cref(r, c), s))
            else:
                o.append('<c r="%s" s="%d" t="s"><v>%d</v></c>' % (cref(r, c), s, sidx(value, XF_FONT[style])))
        o.append('</row>')
    o.append('</sheetData>')
    if sh.merges:
        o.append('<mergeCells count="%d">' % len(sh.merges))
        for (r1, c1, r2, c2) in sh.merges:
            o.append('<mergeCell ref="%s:%s"/>' % (cref(r1, c1), cref(r2, c2)))
        o.append('</mergeCells>')
    if sh.validations:
        o.append('<dataValidations count="%d">' % len(sh.validations))
        for sqref, items in sh.validations:
            o.append('<dataValidation type="list" errorStyle="warning" allowBlank="1" showErrorMessage="1" sqref="%s">'
                     '<formula1>%s</formula1></dataValidation>' % (sqref, xesc('"' + items + '"')))
        o.append('</dataValidations>')
    if sh.links:
        o.append('<hyperlinks>')
        for refx, target in sh.links:
            o.append('<hyperlink ref="%s" location="%s" display="%s"/>' % (
                refx, xesc("'" + target.replace("'", "''") + "'!A1"), xesc(target)))
        o.append('</hyperlinks>')
    o.append('<pageMargins left="0.4" right="0.4" top="0.5" bottom="0.5" header="0.3" footer="0.3"/>')
    o.append('<pageSetup paperSize="%d" orientation="landscape" fitToWidth="%d" fitToHeight="0"/>' % (sh.paper, sh.fit_width))
    o.append('<headerFooter><oddFooter>&amp;C&amp;P / &amp;N</oddFooter></headerFooter>')
    o.append('</worksheet>')
    return ''.join(o)


def _styles_xml():
    st = SPEC['styles']
    fn = XL['font_name']
    o = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<styleSheet xmlns="%s">' % NS_MAIN]
    o.append('<fonts count="%d">' % len(st['fonts']))
    for f in st['fonts']:
        o.append('<font>%s%s%s<sz val="%s"/><color rgb="%s"/><name val="%s"/><family val="3"/><charset val="128"/></font>' % (
            '<b/>' if f.get('b', False) else '', '<i/>' if f.get('i', False) else '',
            '<u/>' if f.get('u', False) else '', f['sz'], f['color'], fn))
    o.append('</fonts>')
    o.append('<fills count="%d">' % len(st['fills']))
    for i, color in enumerate(st['fills']):
        if i == 0:
            o.append('<fill><patternFill patternType="none"/></fill>')
        elif i == 1:
            o.append('<fill><patternFill patternType="gray125"/></fill>')
        else:
            o.append('<fill><patternFill patternType="solid"><fgColor rgb="%s"/><bgColor indexed="64"/></patternFill></fill>' % color)
    o.append('</fills>')
    o.append('<borders count="%d">' % len(st['borders']))
    for b in st['borders']:
        parts = []
        for side in ('left', 'right', 'top', 'bottom'):
            if side in b:
                parts.append('<%s style="%s"><color rgb="%s"/></%s>' % (side, b[side][0], b[side][1], side))
            else:
                parts.append('<%s/>' % side)
        o.append('<border>' + ''.join(parts) + '<diagonal/></border>')
    o.append('</borders>')
    o.append('<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>')
    o.append('<cellXfs count="%d">' % len(st['xfs']))
    for name, font, fillx, border, h, v, wrap, indent, numfmt in st['xfs']:
        attrs = 'numFmtId="%d" fontId="%d" fillId="%d" borderId="%d" xfId="0"' % (numfmt, font, fillx, border)
        if numfmt:
            attrs += ' applyNumberFormat="1"'
        attrs += ' applyFont="1" applyFill="1" applyBorder="1"'
        a = []
        if h:
            a.append('horizontal="%s"' % h)
        if v:
            a.append('vertical="%s"' % v)
        if wrap:
            a.append('wrapText="1"')
        if indent:
            a.append('indent="%d"' % indent)
        if a:
            o.append('<xf %s applyAlignment="1"><alignment %s/></xf>' % (attrs, ' '.join(a)))
        else:
            o.append('<xf %s/>' % attrs)
    o.append('</cellXfs>')
    o.append('<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>')
    o.append('<dxfs count="0"/><tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>')
    o.append('</styleSheet>')
    return ''.join(o)


def save_book(sheets, path, title):
    sst, sst_map = [], {}
    refs = [0]

    def sidx(value, font):
        refs[0] += 1
        if isinstance(value, str):
            key = 'S\x01' + value
        else:
            key = 'R\x01' + str(font) + '\x01' + '\x02'.join(('1' if b else '0') + t for t, b in value)
        i = sst_map.get(key)
        if i is None:
            i = len(sst)
            sst_map[key] = i
            sst.append((value, font))
        return i
    sheet_parts = [_sheet_xml(sh, sidx) for sh in sheets]
    o = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<sst xmlns="%s" count="%d" uniqueCount="%d">' % (
        NS_MAIN, refs[0], len(sst))]
    for value, font in sst:
        if isinstance(value, str):
            o.append('<si>' + _t_xml(value) + '</si>')
        else:
            o.append('<si>' + ''.join('<r>' + _rpr_xml(font, b) + _t_xml(t, True) + '</r>' for t, b in value) + '</si>')
    o.append('</sst>')
    sst_xml = ''.join(o)
    active = 0
    for i, sh in enumerate(sheets):
        if sh.selected:
            active = i
    n = len(sheets)
    ct = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n',
          '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
          '<Default Extension="xml" ContentType="application/xml"/>',
          '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>']
    for i in range(n):
        ct.append('<Override PartName="/xl/worksheets/sheet%d.xml" '
                  'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' % (i + 1))
    ct.append('<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>')
    ct.append('<Override PartName="/xl/sharedStrings.xml" '
              'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>')
    ct.append('<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>')
    ct.append('<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>')
    ct.append('</Types>')
    rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
            '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
            '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>'
            '</Relationships>')
    now = datetime.datetime.now(datetime.timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ')
    core = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
            'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" '
            'xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
            '<dc:title>%s</dc:title><dc:creator>%s</dc:creator>'
            '<dcterms:created xsi:type="dcterms:W3CDTF">%s</dcterms:created>'
            '<dcterms:modified xsi:type="dcterms:W3CDTF">%s</dcterms:modified>'
            '</cp:coreProperties>') % (xesc(title), xesc(TOOL['name'] + ' ' + VERSION), now, now)
    app = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
           '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" '
           'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
           '<Application>Microsoft Excel</Application></Properties>')
    wb = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n',
          '<workbook xmlns="%s" xmlns:r="%s">' % (NS_MAIN, NS_REL),
          '<bookViews><workbookView xWindow="0" yWindow="0" windowWidth="28800" windowHeight="15000" activeTab="%d"/></bookViews>' % active,
          '<sheets>']
    for i, sh in enumerate(sheets):
        wb.append('<sheet name="%s" sheetId="%d" r:id="rId%d"/>' % (xesc(sh.name), i + 1, i + 1))
    wb.append('</sheets></workbook>')
    wbrels = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n',
              '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">']
    for i in range(n):
        wbrels.append('<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
                      'Target="worksheets/sheet%d.xml"/>' % (i + 1, i + 1))
    wbrels.append('<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" '
                  'Target="styles.xml"/>' % (n + 1))
    wbrels.append('<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" '
                  'Target="sharedStrings.xml"/>' % (n + 2))
    wbrels.append('</Relationships>')
    parts = [('[Content_Types].xml', ''.join(ct)), ('_rels/.rels', rels), ('docProps/app.xml', app),
             ('docProps/core.xml', core), ('xl/workbook.xml', ''.join(wb)),
             ('xl/_rels/workbook.xml.rels', ''.join(wbrels)), ('xl/styles.xml', _styles_xml()),
             ('xl/sharedStrings.xml', sst_xml)]
    for i, x in enumerate(sheet_parts):
        parts.append(('xl/worksheets/sheet%d.xml' % (i + 1), x))
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as z:
        for name, text in parts:
            z.writestr(name, text.encode('utf-8'))
    write_file(path, buf.getvalue())


# ----------------------------------------------------------------------------- markdown -> sheets
_LIST_RE = re.compile(r'^(\s*)([-*]|[0-9]+\.)\s+(.*)$')


def md_split_row(line):
    s = line.strip()
    if s.startswith('|'):
        s = s[1:]
    if s.endswith('|') and not s.endswith('\\|'):
        s = s[:-1]
    cells, cur = [], []
    i = 0
    while i < len(s):
        if s[i] == '\\' and i + 1 < len(s) and s[i + 1] == '|':
            cur.append('\\|')
            i += 2
            continue
        if s[i] == '|':
            cells.append(''.join(cur).strip())
            cur = []
        else:
            cur.append(s[i])
        i += 1
    cells.append(''.join(cur).strip())
    return cells


def md_inline(text):
    text = re.sub(r'\[([^\]]+)\]\(([^)]+)\)', r'\1 (\2)', text)
    runs = []
    pos = 0
    for m in re.finditer(r'`([^`]*)`|\*\*(.+?)\*\*', text):
        if m.start() > pos:
            runs.append([text[pos:m.start()].replace('<br>', '\n').replace('\\|', '|'), False])
        if m.group(1) is not None:
            runs.append([m.group(1).replace('\\|', '|'), False])
        else:
            runs.append([m.group(2).replace('<br>', '\n').replace('\\|', '|'), True])
        pos = m.end()
    if pos < len(text):
        runs.append([text[pos:].replace('<br>', '\n').replace('\\|', '|'), False])
    merged = []
    for t, b in runs:
        if t == '':
            continue
        if merged and merged[-1][1] == b:
            merged[-1][0] += t
        else:
            merged.append([t, b])
    if not any(b for t, b in merged):
        return ''.join(t for t, b in merged)
    return merged


def md_parse(md):
    lines = norm_nl(md).split('\n')
    doc = {'title': '', 'cover': [], 'sections': []}
    blocks = doc['cover']
    cols = None
    i, n = 0, len(lines)
    while i < n:
        line = lines[i]
        s = line.strip()
        if s.startswith('```'):
            code = []
            i += 1
            while i < n and not lines[i].strip().startswith('```'):
                code.append(lines[i])
                i += 1
            i += 1
            blocks.append({'t': 'code', 'lines': code})
            continue
        if s.startswith('<!--'):
            text = s
            while '-->' not in text and i + 1 < n:
                i += 1
                text += '\n' + lines[i]
            i += 1
            m = re.match(r'^<!--\s*cols:\s*([0-9,\s]+?)\s*-->$', text)
            if m:
                cols = [int(x) for x in re.split(r'[,\s]+', m.group(1).strip()) if x]
            m = re.match(r'^<!--\s*summary:\s*([^\n]*?)\s*-->$', text)
            if m and doc['sections']:
                doc['sections'][-1]['summary'] = m.group(1)
            continue
        if s == '':
            i += 1
            continue
        if line.startswith('# '):
            doc['title'] = line[2:].strip()
            i += 1
            continue
        if line.startswith('## '):
            sec = {'title': line[3:].strip(), 'summary': '', 'blocks': []}
            doc['sections'].append(sec)
            blocks = sec['blocks']
            i += 1
            continue
        if line.startswith('### '):
            blocks.append({'t': 'h3', 'text': line[4:].strip()})
            i += 1
            continue
        if line.startswith('#### '):
            blocks.append({'t': 'h4', 'text': line[5:].strip()})
            i += 1
            continue
        if s.startswith('|'):
            rows = []
            while i < n and lines[i].strip().startswith('|'):
                rows.append(md_split_row(lines[i]))
                i += 1
            sep = len(rows) > 1 and all(re.match(r'^:?-+:?$', c) for c in rows[1] if c != '')
            blocks.append({'t': 'table', 'header': rows[0], 'rows': rows[2:] if sep else rows[1:], 'cols': cols})
            cols = None
            continue
        if _LIST_RE.match(line):
            items = []
            while i < n:
                m = _LIST_RE.match(lines[i])
                if not m:
                    break
                items.append({'level': len(m.group(1)) // 2, 'mark': m.group(2), 'text': m.group(3).strip()})
                i += 1
            blocks.append({'t': 'list', 'items': items})
            continue
        if s.startswith('>'):
            q = []
            while i < n and lines[i].strip().startswith('>'):
                q.append(re.sub(r'^\s*>\s?', '', lines[i]))
                i += 1
            blocks.append({'t': 'quote', 'text': '\n'.join(q).strip()})
            continue
        if re.match(r'^(-{3,}|\*{3,})$', s):
            i += 1
            continue
        para = []
        while i < n:
            s2 = lines[i].strip()
            if (s2 == '' or s2.startswith('```') or s2.startswith('<!--') or s2.startswith('|')
                    or s2.startswith('>') or s2.startswith('#') or _LIST_RE.match(lines[i])):
                break
            para.append(s2)
            i += 1
        blocks.append({'t': 'p', 'text': '\n'.join(para)})
    return doc


def _max_units(text):
    return max([text_units(x) for x in norm_nl(text).split('\n')] or [0])


def section_widths(blocks):
    tables = [b for b in blocks if b['t'] == 'table']
    if not tables:
        return [XL['min_total_width']]
    n = max(len(t['header']) for t in tables)
    w = [0] * n
    for t in tables:
        k = len(t['header'])
        for j in range(k):
            if j == k - 1 and k < n:
                continue
            if t['cols'] and j < len(t['cols']):
                cw = t['cols'][j]
            else:
                units = 0
                for row in [t['header']] + t['rows']:
                    cell = row[j] if j < len(row) else ''
                    units = max(units, _max_units(plain_of(md_inline(cell))))
                cw = min(60, max(8, (units + 19) // 20 + 2))
            w[j] = max(w[j], cw)
    for j in range(n):
        if w[j] == 0:
            w[j] = 10
    total = sum(w)
    if total < XL['min_total_width']:
        w[n - 1] += XL['min_total_width'] - total
    return w


def _prefix(prefix, v):
    if isinstance(v, str):
        return prefix + v
    if v and not v[0][1]:
        return [[prefix + v[0][0], False]] + [list(x) for x in v[1:]]
    return [[prefix, False]] + [list(x) for x in v]


def render_table(sh, b, r, first, last, widths):
    header = b['header']
    ncol = len(header)

    def span(j):
        c1 = first + j
        c2 = last if j == ncol - 1 else c1
        return c1, c2, sum(widths[c - first] for c in range(c1, c2 + 1))
    lines = 1
    for j, h in enumerate(header):
        c1, c2, w = span(j)
        v = md_inline(h)
        sh.block(r, c1, c2, v, 'th')
        lines = max(lines, wrap_lines(plain_of(v), w))
    sh.h10[r] = row_h10(lines)
    r += 1
    for k, row in enumerate(b['rows']):
        style = 'td' if k % 2 == 0 else 'td_alt'
        lines = 1
        for j in range(ncol):
            c1, c2, w = span(j)
            v = md_inline(row[j] if j < len(row) else '')
            sh.block(r, c1, c2, v, style)
            lines = max(lines, wrap_lines(plain_of(v), w))
        sh.h10[r] = row_h10(lines)
        r += 1
    return r


def render_blocks(sh, blocks, r, first, last, widths):
    total = sum(widths)
    for k, b in enumerate(blocks):
        t = b['t']
        if t == 'h3':
            if k > 0:
                sh.h10[r] = 80
                r += 1
            sh.block(r, first, last, plain_of(md_inline(b['text'])), 'h3')
            sh.h10[r] = 240
            r += 1
            sh.h10[r] = 40
            r += 1
            continue
        if t == 'h4':
            sh.block(r, first, last, plain_of(md_inline(b['text'])), 'h4')
            sh.h10[r] = 220
            r += 1
            continue
        if t == 'p':
            v = md_inline(b['text'])
            sh.block(r, first, last, v, 'p')
            sh.h10[r] = row_h10(wrap_lines(plain_of(v), total))
            r += 1
        elif t == 'quote':
            v = md_inline(b['text'])
            sh.block(r, first, last, v, 'quote')
            sh.h10[r] = row_h10(wrap_lines(plain_of(v), total - 2))
            r += 1
        elif t == 'list':
            for it in b['items']:
                mark = '・' if it['mark'] in ('-', '*') else it['mark'] + ' '
                v = _prefix('　' * it['level'] + mark, md_inline(it['text']))
                sh.block(r, first, last, v, 'li')
                sh.h10[r] = row_h10(wrap_lines(plain_of(v), total - 2))
                r += 1
        elif t == 'code':
            code = b['lines'] or ['']
            for j, line in enumerate(code):
                if len(code) == 1:
                    style = 'code_single'
                elif j == 0:
                    style = 'code_first'
                elif j == len(code) - 1:
                    style = 'code_last'
                else:
                    style = 'code_mid'
                text = line if line != '' else ' '
                sh.block(r, first, last, text, style)
                sh.h10[r] = row_h10(wrap_lines(text, total - 2), 18)
                r += 1
        elif t == 'table':
            r = render_table(sh, b, r, first, last, widths)
        sh.h10[r] = 60
        r += 1
    return r


def render_section(name, title, blocks):
    sh = Sheet(name)
    widths = section_widths(blocks)
    sh.widths[1] = 2
    for j, w in enumerate(widths):
        sh.widths[j + 2] = w
    first, last = 2, 1 + len(widths)
    sh.block(1, first, last, title, 'title')
    sh.h10[1] = 300
    sh.h10[2] = 60
    render_blocks(sh, blocks, 3, first, last, widths)
    return sh


def sheet_name_of(title, used):
    name = re.sub(r'[\\/?*\[\]:]', '', title).strip()[:31] or 'Sheet'
    base, k = name, 2
    while name.lower() in used:
        suffix = '(%d)' % k
        name = base[:31 - len(suffix)] + suffix
        k += 1
    used.add(name.lower())
    return name


def render_cover(doc, names):
    sh = Sheet(XL['cover_sheet'])
    sh.selected = True
    widths = XL['cover_widths']
    sh.widths[1] = 2
    for j, w in enumerate(widths):
        sh.widths[j + 2] = w
    first, last = 2, 1 + len(widths)
    sh.h10[1] = 120
    sh.block(2, first, last, doc['title'], 'cover_title')
    sh.h10[2] = 420
    sh.h10[3] = 120
    r = render_blocks(sh, doc['cover'], 4, first, last, widths)
    sh.h10[r] = 80
    r += 1
    sh.block(r, first, last, XL['toc_title'], 'h3')
    sh.h10[r] = 240
    r += 1
    sh.h10[r] = 40
    r += 1
    for j, h in enumerate(XL['toc_headers']):
        sh.put(r, first + j, h, 'th')
    sh.h10[r] = row_h10(1)
    r += 1
    for k, sec in enumerate(doc['sections']):
        sh.put(r, first, str(k + 1), 'toc_no')
        sh.put(r, first + 1, names[k], 'toc_link')
        sh.links.append((cref(r, first + 1), names[k]))
        sh.put(r, first + 2, sec['summary'], 'toc_td')
        sh.h10[r] = row_h10(wrap_lines(sec['summary'], widths[2]))
        r += 1
    return sh


def md_cell(s):
    return (s or '').replace('|', '\\|').replace('\n', '<br>')


def md_code(s):
    return '<br>'.join('`' + x.replace('|', '\\|') + '`' for x in norm_nl(s).split('\n') if x != '')


def settings_table_md(items=None, cols='12,28,22,16,62'):
    out = ['<!-- cols: %s -->' % cols, '| 区分 | 設定キー | 設定項目 | 既定値 | 説明 |', '|---|---|---|---|---|']
    for s in items or SPEC['settings']:
        d = s['default'] if s['default'] != '' else TPL['empty_default']
        desc = s['desc']
        if 'template' in s:
            desc += '（テンプレートの記入例: ' + s['template'].replace('\n', ' / ') + '）'
        out.append('| %s | %s | %s | %s | %s |' % (md_cell(s['category']), md_code(s['key']), md_cell(s['label']),
                                                   md_cell(d), md_cell(desc)))
    return '\n'.join(out)


def api_columns_table_md():
    out = ['<!-- cols: 18,6,62,36 -->', '| 列名 | 必須 | 説明 | 記入例 |', '|---|---|---|---|']
    for c in SPEC['api_columns']:
        out.append('| %s | %s | %s | %s |' % (md_cell(c['header']), '○' if c['required'] else '', md_cell(c['desc']),
                                             md_code(c['example'])))
    return '\n'.join(out)


def expand_md(md):
    md = md.replace('1.0.0', VERSION).replace('@@DATE@@', today())
    md = md.replace('<!-- @SETTINGS_TABLE@ -->', settings_table_md())
    md = md.replace('<!-- @OPENAPI_CONFIG_TABLE@ -->', settings_table_md(SPEC['openapi_config'], '14,28,24,26,60'))
    return md.replace('<!-- @API_COLUMNS_TABLE@ -->', api_columns_table_md())


# ----------------------------------------------------------------------------- templates / data / manual
def template_api_rows():
    cols = SPEC['api_columns']
    rows = [[c['header'] for c in cols], [c['hint'] for c in cols]]
    for api in SPEC['sample_apis']:
        rows.append([api.get(c['key'], '') for c in cols])
    return rows


def template_settings_rows(items=None):
    rows = [list(TPL['settings_headers'])]
    cat = None
    for s in items or SPEC['settings']:
        if s['category'] != cat:
            cat = s['category']
            rows.append(['#■ ' + cat, '', '', '', ''])
        rows.append([s['key'], s['label'], s.get('template', s['default']),
                     s['default'] if s['default'] != '' else TPL['empty_default'], s['desc']])
    return rows


def build_template_xlsx(path):
    cols = SPEC['api_columns']
    n = len(cols)
    sh = Sheet(TPL['api_sheet'])
    sh.selected = True
    for j, c in enumerate(cols):
        sh.widths[j + 1] = c['width']
    total = sum(c['width'] for c in cols)
    sh.block(1, 1, n, TPL['api_title'], 'title')
    sh.h10[1] = 300
    sh.block(2, 1, n, TPL['api_note'], 'note')
    sh.h10[2] = row_h10(wrap_lines(TPL['api_note'], total))
    lines = 1
    for j, c in enumerate(cols):
        sh.put(3, j + 1, c['header'], 'api_hdr_req' if c['required'] else 'api_hdr')
        lines = max(lines, wrap_lines(c['header'], c['width']))
    sh.h10[3] = row_h10(lines)
    lines = 1
    for j, c in enumerate(cols):
        sh.put(4, j + 1, c['hint'], 'hint')
        lines = max(lines, wrap_lines(c['hint'], c['width']))
    sh.h10[4] = row_h10(lines)
    r = 5
    for api in SPEC['sample_apis']:
        lines = 1
        for j, c in enumerate(cols):
            v = api.get(c['key'], '')
            sh.put(r, j + 1, v, 'cell')
            lines = max(lines, wrap_lines(v, c['width']))
        sh.h10[r] = row_h10(lines)
        r += 1
    for k in range(TPL['blank_rows']):
        for j in range(n):
            sh.put(r, j + 1, None, 'cell')
        sh.h10[r] = row_h10(1)
        r += 1
    sh.freeze_rows, sh.freeze_cols = 4, 3
    sh.paper, sh.fit_width = 8, 2
    keys = [c['key'] for c in cols]
    last = TPL['validation_rows']
    for key, items in (('enabled', 'Y,N'), ('method', 'GET,POST,PUT,PATCH,DELETE'), ('body_type', 'auto,json,form,raw')):
        cl = col_letter(keys.index(key) + 1)
        sh.validations.append(('%s5:%s%d' % (cl, cl, last), items))
    # 設定シート
    st = Sheet(TPL['settings_sheet'])
    widths = TPL['settings_widths']
    for j, w in enumerate(widths):
        st.widths[j + 1] = w
    st.block(1, 1, 5, TPL['settings_title'], 'title')
    st.h10[1] = 300
    st.block(2, 1, 5, TPL['settings_note'], 'note')
    st.h10[2] = row_h10(wrap_lines(TPL['settings_note'], sum(widths)))
    for j, h in enumerate(TPL['settings_headers']):
        st.put(3, j + 1, h, 'th')
    st.h10[3] = row_h10(1)
    r = 4
    cat = None
    for s in SPEC['settings']:
        if s['category'] != cat:
            cat = s['category']
            st.block(r, 1, 5, '■ ' + cat, 'cat')
            st.h10[r] = 220
            r += 1
        d = s['default'] if s['default'] != '' else TPL['empty_default']
        cells = [(s['key'], 'key'), (s['label'], 'key'), (s.get('template', s['default']), 'input'), (d, 'key'), (s['desc'], 'td')]
        lines = 1
        for j, (v, style) in enumerate(cells):
            st.put(r, j + 1, v, style)
            lines = max(lines, wrap_lines(v, widths[j]))
        st.h10[r] = row_h10(lines)
        if 'choices' in s:
            st.validations.append(('C%d' % r, ','.join(s['choices'])))
        r += 1
    st.freeze_rows = 3
    # 記入方法シート
    gdoc = md_parse(expand_md(SPEC['template_guide_md']))
    sec = gdoc['sections'][0]
    guide = render_section(TPL['guide_sheet'], sec['title'], sec['blocks'])
    save_book([sh, st, guide], path, TPL['api_title'])


def build_testdata_csv(path):
    d = SPEC['sample_csv']
    write_text(path, to_csv([d['header']] + d['rows']), bom=True)


def build_testdata_xlsx(path):
    d = SPEC['sample_excel']
    sh = Sheet(d['sheet'])
    sh.selected = True
    for j, h in enumerate(d['header']):
        sh.widths[j + 1] = 18
        sh.put(1, j + 1, h, 'data_hdr')
    sh.h10[1] = 220
    for i, row in enumerate(d['rows']):
        for j, v in enumerate(row):
            sh.put(i + 2, j + 1, v, 'data_cell')
        sh.h10[i + 2] = 200
    sh.freeze_rows = 1
    save_book([sh], path, 'testdata')


def build_manual(out_dir):
    md = expand_md(SPEC['manual_md'])
    base = os.path.join(out_dir, XL['manual_file'])
    write_text(base + '.md', md)
    doc = md_parse(md)
    used = set([XL['cover_sheet'].lower()])
    names = [sheet_name_of(sec['title'], used) for sec in doc['sections']]
    sheets = [render_cover(doc, names)]
    for k, sec in enumerate(doc['sections']):
        sheets.append(render_section(names[k], sec['title'], sec['blocks']))
    save_book(sheets, base + '.xlsx', doc['title'])
    return [base + '.xlsx', base + '.md']


# ----------------------------------------------------------------------------- OpenAPI loading
# =============================================================================
#  OpenAPI 3.0.x 定義（JSON）の読み込み
#  JMeter_Tools の linux/openapi2jmx.sh（Python 本体）から、定義の読み込み・テストデータ生成の
#  処理をそのまま移植したものです（JMX 出力の部分は除く）。移植時の変更点は次のとおりです。
#    - GenError を ToolError の派生クラスにした（[ERROR] として表示するため）
#    - pct_encode を oa_pct_encode に改名した（このツールの pct_encode と名前が重なるため）
#    - build_operation の「6. 組み立て」を、API一覧（Excel / CSV）の 1 行と同じ項目を返すように変更した
#    - ヘッダ API-Key の警告文（このツールには共通の API-Key が無いため）。OpenAPI 用設定 API_HEADERS に API-Key（既定値以外）がある場合は警告しない
#    - 認可（security）の警告は、設定ファイル（-s）を読んだ後に load_input で判定する（設定 common_headers の Authorization を確認するため）
#  末尾の oa_row_pairs / oa_server_url / load_openapi がこのツール用に追加した処理です。
#  Windows 版 jmeter_api_tool.ps1 の同じ処理と同一の乱数列・同一の結果になります。
# =============================================================================
class GenError(ToolError):
    """利用者に伝えるべき入力エラー"""


# ----------------------------------------------------------------- 乱数（両実装共通）
MOD31 = 2147483647


class Rng(object):
    """Park-Miller 最小標準乱数（乗数 48271）。PowerShell 版と完全に同じ数列を返す。"""

    def __init__(self, seed):
        self.state = (seed % (MOD31 - 1)) + 1

    def next31(self):
        self.state = (self.state * 48271) % MOD31
        return self.state

    def between(self, lo, hi):
        if hi <= lo:
            return lo
        span = hi - lo + 1
        if span <= MOD31 - 1:
            return lo + (self.next31() - 1) % span
        r = (self.next31() - 1) * (MOD31 - 1) + (self.next31() - 1)
        return lo + r % span


# ----------------------------------------------------------------- 値の表現
class JNum(object):
    """JSON に数値として書き出す値（表記を文字列で厳密に保持する）"""
    __slots__ = ('text',)

    def __init__(self, text):
        self.text = text


SKIP = object()          # 生成しない（循環参照・深さ超過など）
UNITS_LIMIT = 10 ** 18   # 数値の生成範囲の上限（最小単位に換算した整数で。64bit 整数に収めるため）
MAX_DECIMALS = 12        # 数値の小数桁の上限
MAX_DEPTH = 12           # スキーマの入れ子の上限

ALNUM = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789'
LOWER_ALNUM = 'abcdefghijklmnopqrstuvwxyz0123456789'
HEXD = '0123456789abcdef'
EPOCH_START = 1735689600     # 2025-01-01T00:00:00Z
EPOCH_END = 1798761599       # 2026-12-31T23:59:59Z


def is_obj(x):
    return isinstance(x, dict)


def is_arr(x):
    return isinstance(x, list)


def is_str(x):
    return isinstance(x, str)


def is_bool(x):
    return isinstance(x, bool)


def is_int(x):
    return isinstance(x, int) and not isinstance(x, bool)


def is_dec(x):
    return isinstance(x, Decimal)


def is_num(x):
    return is_int(x) or is_dec(x)


def get(o, k):
    return o.get(k) if isinstance(o, dict) else None


def has(o, k):
    return isinstance(o, dict) and k in o


def to_dec(x):
    if is_int(x):
        return Decimal(x)
    if is_dec(x):
        return x
    return None


def dec_plain(d):
    return format(d, 'f')


def clamp_to(d, lim):
    if d > lim:
        return lim
    if d < -lim:
        return -lim
    return d


def int_or_none(x):
    if is_int(x):
        return x if x >= 0 else None
    if is_dec(x) and x == x.to_integral_value() and x >= 0:
        return int(x)
    return None


def dceil(d):
    return int(d.to_integral_value(rounding=ROUND_CEILING))


def dfloor(d):
    return int(d.to_integral_value(rounding=ROUND_FLOOR))


def to_gen(x):
    """OpenAPI 定義上の値（example/enum など）を生成値の表現へ変換する"""
    if x is None or is_bool(x):
        return x
    if is_int(x):
        return JNum(str(x))
    if is_dec(x):
        return JNum(dec_plain(x))
    if is_str(x):
        return x
    if is_arr(x):
        return [to_gen(i) for i in x]
    if is_obj(x):
        o = OrderedDict()
        for k, v in x.items():
            o[k] = to_gen(v)
        return o
    return str(x)


# ----------------------------------------------------------------- 文字列の整形
def clean_text(s):
    """XML 1.0 で使えない文字（制御文字・孤立サロゲート等）を取り除く"""
    out = []
    for ch in s:
        c = ord(ch)
        if c == 9 or c == 10 or c == 13 or 0x20 <= c <= 0xD7FF or 0xE000 <= c <= 0xFFFD or c >= 0x10000:
            out.append(ch)
    return ''.join(out)


def json_str(s):
    s = clean_text(s)
    out = ['"']
    for ch in s:
        if ch == '"':
            out.append('\\"')
        elif ch == '\\':
            out.append('\\\\')
        elif ch == '\n':
            out.append('\\n')
        elif ch == '\r':
            out.append('\\r')
        elif ch == '\t':
            out.append('\\t')
        else:
            out.append(ch)
    out.append('"')
    return ''.join(out)


def to_json(v, indent, pretty):
    if v is None:
        return 'null'
    if v is True:
        return 'true'
    if v is False:
        return 'false'
    if isinstance(v, JNum):
        return v.text
    if is_str(v):
        return json_str(v)
    if is_obj(v):
        if len(v) == 0:
            return '{}'
        if not pretty:
            return '{' + ','.join(json_str(k) + ':' + to_json(x, 0, False) for k, x in v.items()) + '}'
        pad = '  ' * (indent + 1)
        items = [pad + json_str(k) + ': ' + to_json(x, indent + 1, True) for k, x in v.items()]
        return '{\n' + ',\n'.join(items) + '\n' + '  ' * indent + '}'
    if is_arr(v):
        if len(v) == 0:
            return '[]'
        if not pretty:
            return '[' + ','.join(to_json(x, 0, False) for x in v) + ']'
        pad = '  ' * (indent + 1)
        items = [pad + to_json(x, indent + 1, True) for x in v]
        return '[\n' + ',\n'.join(items) + '\n' + '  ' * indent + ']'
    return json_str(str(v))


def scalar_str(v):
    if v is None:
        return ''
    if v is True:
        return 'true'
    if v is False:
        return 'false'
    if isinstance(v, JNum):
        return v.text
    if is_str(v):
        return v
    return to_json(v, 0, False)


UNRESERVED = set(b'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-._~')


def oa_pct_encode(s, keep_slash):
    out = []
    for b in clean_text(s).encode('utf-8'):
        if b in UNRESERVED or (keep_slash and b == 0x2F):
            out.append(chr(b))
        else:
            out.append('%%%02X' % b)
    return ''.join(out)


def header_safe(s):
    return clean_text(s).replace('\r', ' ').replace('\n', ' ').replace('\t', ' ')


def safe_name(s):
    s = re.sub(r'[^A-Za-z0-9._-]+', '_', s)
    s = re.sub(r'_+', '_', s).strip('_.')
    return s if s else 'openapi'


def cp_len(s):
    return len(s)


# ----------------------------------------------------------------- 生成コンテキスト
class Ctx(object):
    def __init__(self, doc, seed, opts):
        self.doc = doc
        self.seed = seed
        self.rng = Rng(seed)
        self.required_only = opts['required_only']
        self.use_examples = opts['use_examples']
        self.any_method = opts['any_method']
        self.warnings = []

    def warn(self, msg):
        if msg not in self.warnings:
            self.warnings.append(msg)


def unescape_token(tok):
    if '%' in tok:
        try:
            from urllib.parse import unquote
            tok = unquote(tok, encoding='utf-8', errors='strict')
        except Exception:
            pass
    return tok.replace('~1', '/').replace('~0', '~')


def resolve_ref(ctx, ref):
    if not is_str(ref) or not ref.startswith('#'):
        return None
    if ref == '#' or ref == '#/':
        return ctx.doc
    if not ref.startswith('#/'):
        return None
    cur = ctx.doc
    for tok in ref[2:].split('/'):
        key = unescape_token(tok)
        if is_obj(cur) and key in cur:
            cur = cur[key]
        elif is_arr(cur) and re.match(r'^(0|[1-9][0-9]*)$', key) and int(key) < len(cur):
            cur = cur[int(key)]
        else:
            return None
    return cur


def ref_text(ref):
    return ref if is_str(ref) else '(不正な $ref)'


def deref(ctx, obj, what):
    """$ref を辿って実体を返す（Parameter / RequestBody / Response / Example / PathItem 用）"""
    seen = []
    while is_obj(obj) and '$ref' in obj:
        rt = ref_text(obj['$ref'])
        if rt in seen or len(seen) > 32:
            ctx.warn('%s の $ref が循環しています: %s' % (what, rt))
            return None
        seen.append(rt)
        target = resolve_ref(ctx, obj['$ref'])
        if target is None:
            ctx.warn('%s の $ref を解決できません（外部ファイル参照は未対応）: %s' % (what, rt))
            return None
        obj = target
    return obj


# ----------------------------------------------------------------- スキーマ → テストデータ
TYPES = ('object', 'array', 'integer', 'number', 'boolean', 'string', 'null')


def schema_type(schema):
    t = get(schema, 'type')
    if is_arr(t):
        names = [x for x in t if is_str(x)]
        non_null = [x for x in names if x != 'null']
        if non_null:
            t = non_null[0]
        elif names:
            t = 'null'
        else:
            t = None
    if is_str(t) and t in TYPES:
        return t
    if has(schema, 'properties') or has(schema, 'additionalProperties') or has(schema, 'minProperties') \
            or has(schema, 'maxProperties'):
        return 'object'
    if has(schema, 'items') or has(schema, 'minItems') or has(schema, 'maxItems'):
        return 'array'
    fmt = get(schema, 'format')
    if fmt == 'int32' or fmt == 'int64':
        return 'integer'
    if fmt == 'float' or fmt == 'double':
        return 'number'
    if has(schema, 'minimum') or has(schema, 'maximum') or has(schema, 'multipleOf') \
            or has(schema, 'exclusiveMinimum') or has(schema, 'exclusiveMaximum'):
        return 'number'
    return 'string'


def merge_schemas(a, b):
    r = OrderedDict(a)
    for k, v in b.items():
        if k == 'properties' and is_obj(v):
            props = OrderedDict(r['properties']) if is_obj(get(r, 'properties')) else OrderedDict()
            for pk, pv in v.items():
                if pk not in props:
                    props[pk] = pv
            r['properties'] = props
        elif k == 'required' and is_arr(v):
            req = list(r['required']) if is_arr(get(r, 'required')) else []
            for x in v:
                if x not in req:
                    req.append(x)
            r['required'] = req
        elif k not in r:
            r[k] = v
    return r


def follow_schema_ref(ctx, s, stack):
    """スキーマの $ref を辿る。戻り値: (実体 or None, 新しいスタック, 状態 'ok'|'cycle'|'unresolved')"""
    while is_obj(s) and '$ref' in s:
        rt = ref_text(s['$ref'])
        if rt in stack:
            return None, stack, 'cycle'
        t = resolve_ref(ctx, s['$ref'])
        if t is None:
            ctx.warn('スキーマの $ref を解決できません（外部ファイル参照は未対応）: ' + rt)
            return None, stack, 'unresolved'
        stack = stack + [rt]
        s = t
    return s, stack, 'ok'


def merge_all_of(ctx, schema, stack):
    result = OrderedDict((k, v) for k, v in schema.items() if k != 'allOf')
    for sub in schema['allOf']:
        s, s_stack, state = follow_schema_ref(ctx, sub, stack)
        if state != 'ok' or not is_obj(s):
            continue
        if is_arr(get(s, 'allOf')):
            s = merge_all_of(ctx, s, s_stack)
        result = merge_schemas(result, s)
    return result


def gen_value(ctx, schema, depth, stack):
    if schema is None or schema is True:
        schema = OrderedDict()
    if schema is False or not is_obj(schema):
        return SKIP
    if '$ref' in schema:
        s, new_stack, state = follow_schema_ref(ctx, schema, stack)
        if state == 'cycle':
            return SKIP
        if state == 'unresolved':
            return 'UNRESOLVED_REF'
        return gen_value(ctx, s, depth, new_stack)
    if depth > MAX_DEPTH:
        return SKIP
    if is_arr(get(schema, 'allOf')):
        schema = merge_all_of(ctx, schema, stack)
    for key in ('oneOf', 'anyOf'):
        alts = get(schema, key)
        if is_arr(alts) and len(alts) > 0:
            base = OrderedDict((k, v) for k, v in schema.items() if k != 'oneOf' and k != 'anyOf')
            alt, alt_stack, state = follow_schema_ref(ctx, alts[0], stack)
            if state == 'cycle':
                return SKIP
            if state == 'unresolved':
                return 'UNRESOLVED_REF'
            if not is_obj(alt):
                alt = OrderedDict()
            return gen_value(ctx, merge_schemas(base, alt), depth, alt_stack)
    if ctx.use_examples:
        if 'example' in schema:
            return to_gen(schema['example'])
        exs = get(schema, 'examples')
        if is_arr(exs) and len(exs) > 0:
            return to_gen(exs[0])
        if 'default' in schema:
            return to_gen(schema['default'])
    if 'const' in schema:
        return to_gen(schema['const'])
    enum = get(schema, 'enum')
    if is_arr(enum) and len(enum) > 0:
        cands = [e for e in enum if e is not None]
        if not cands:
            return None
        return to_gen(cands[ctx.rng.between(0, len(cands) - 1)])
    t = schema_type(schema)
    if t == 'object':
        return gen_object(ctx, schema, depth, stack)
    if t == 'array':
        return gen_array(ctx, schema, depth, stack)
    if t == 'integer':
        return gen_integer(ctx, schema)
    if t == 'number':
        return gen_number(ctx, schema)
    if t == 'boolean':
        return ctx.rng.between(0, 1) == 1
    if t == 'null':
        return None
    return gen_string(ctx, schema)


def is_read_only(ctx, pschema):
    s = pschema
    guard = 0
    while is_obj(s) and '$ref' in s and guard < 32:
        s = resolve_ref(ctx, s['$ref'])
        guard += 1
    return is_obj(s) and get(s, 'readOnly') is True


def gen_object(ctx, schema, depth, stack):
    result = OrderedDict()
    props = get(schema, 'properties')
    req = get(schema, 'required')
    required = [x for x in req if is_str(x)] if is_arr(req) else []
    if is_obj(props):
        for name, pschema in props.items():
            if ctx.required_only and name not in required:
                continue
            if is_read_only(ctx, pschema):
                continue
            v = gen_value(ctx, pschema, depth + 1, stack)
            if v is SKIP:
                continue
            result[name] = v
    addl = get(schema, 'additionalProperties')
    if len(result) == 0 and not (is_obj(props) and len(props) > 0) and is_obj(addl) and len(addl) > 0 \
            and not ctx.required_only:
        v = gen_value(ctx, addl, depth + 1, stack)
        if v is not SKIP:
            result['key1'] = v
    return result


def gen_array(ctx, schema, depth, stack):
    items = get(schema, 'items')
    if not is_obj(items):
        items = OrderedDict()
    mn = int_or_none(get(schema, 'minItems'))
    mx = int_or_none(get(schema, 'maxItems'))
    n = max(mn if mn is not None else 0, 1)
    if mx is not None and n > mx:
        n = mx
    if n > 50:
        n = 50
    unique = get(schema, 'uniqueItems') is True
    result = []
    seen = []
    for _ in range(n):
        v = gen_value(ctx, items, depth + 1, stack)
        if v is SKIP:
            break
        if unique:
            key = to_json(v, 0, False)
            tries = 0
            while key in seen and tries < 10:
                v = gen_value(ctx, items, depth + 1, stack)
                if v is SKIP:
                    break
                key = to_json(v, 0, False)
                tries += 1
            if v is SKIP:
                break
            seen.append(key)
        result.append(v)
    return result


def num_bounds(schema, scale):
    """minimum/maximum 等を scale（10 の d 乗）単位の整数範囲に変換する"""
    mn = to_dec(get(schema, 'minimum'))
    mx = to_dec(get(schema, 'maximum'))
    exmin = get(schema, 'exclusiveMinimum')
    exmax = get(schema, 'exclusiveMaximum')
    ex_lo = exmin is True
    ex_hi = exmax is True
    if is_num(exmin):
        e = to_dec(exmin)
        if mn is None or e >= mn:
            mn = e
            ex_lo = True
    if is_num(exmax):
        e = to_dec(exmax)
        if mx is None or e <= mx:
            mx = e
            ex_hi = True
    vlim = Decimal(UNITS_LIMIT // scale)
    lo = None
    hi = None
    if mn is not None:
        v = clamp_to(mn, vlim) * scale
        lo = dceil(v)
        if ex_lo and Decimal(lo) == v:
            lo += 1
    if mx is not None:
        v = clamp_to(mx, vlim) * scale
        hi = dfloor(v)
        if ex_hi and Decimal(hi) == v:
            hi -= 1
    one = scale
    if lo is None and hi is None:
        lo = one
        hi = 100 * one
    elif hi is None:
        hi = lo + 99 * one
    elif lo is None:
        lo = one if hi >= one else hi - 99 * one
    if lo > hi:
        hi = lo
    return lo, hi


def ceil_div(a, b):
    return -((-a) // b)


def gen_integer(ctx, schema):
    lo, hi = num_bounds(schema, 1)
    if get(schema, 'format') == 'int32':
        lo = max(lo, -2147483648)
        hi = min(hi, 2147483647)
        if lo > hi:
            hi = lo
    m = to_dec(get(schema, 'multipleOf'))
    if m is not None and m > 0 and m == m.to_integral_value():
        mi = int(m)
        kmin = ceil_div(lo, mi)
        kmax = hi // mi
        if kmin <= kmax:
            return JNum(str(ctx.rng.between(kmin, kmax) * mi))
        return JNum(str(lo))
    return JNum(str(ctx.rng.between(lo, hi)))


def decimals_of(d):
    s = dec_plain(d)
    if '.' in s:
        return len(s.split('.', 1)[1].rstrip('0'))
    return 0


def fmt_units(units, d):
    neg = units < 0
    a = -units if neg else units
    if d == 0:
        s = str(a)
    else:
        p = 10 ** d
        s = str(a // p) + '.' + str(a % p).rjust(d, '0')
    return '-' + s if neg else s


def gen_number(ctx, schema):
    m = to_dec(get(schema, 'multipleOf'))
    if m is not None and m <= 0:
        m = None
    d = decimals_of(m) if m is not None else 2
    if d > MAX_DECIMALS:
        d = MAX_DECIMALS
    scale = 10 ** d
    lo, hi = num_bounds(schema, scale)
    units = None
    if m is not None:
        mu = m * scale
        if mu > 0 and mu == mu.to_integral_value():
            mi = int(mu)
            kmin = ceil_div(lo, mi)
            kmax = hi // mi
            units = ctx.rng.between(kmin, kmax) * mi if kmin <= kmax else lo
    if units is None:
        units = ctx.rng.between(lo, hi)
    return JNum(fmt_units(units, d))


def rand_str(ctx, pool, n):
    out = []
    for _ in range(n):
        out.append(pool[ctx.rng.between(0, len(pool) - 1)])
    return ''.join(out)


def choose_len(ctx, min_l, max_l):
    lo = max(min_l if min_l is not None else 0, 8)
    if max_l is not None and lo > max_l:
        lo = max_l
    if max_l is None:
        hi = max(lo, 16)
    else:
        hi = min(max_l, max(lo, 16))
    if hi < lo:
        hi = lo
    return ctx.rng.between(lo, hi)


def epoch_to_utc(secs):
    return datetime.datetime(1970, 1, 1) + datetime.timedelta(seconds=secs)


def gen_string(ctx, schema):
    fmt = get(schema, 'format')
    fmt = fmt if is_str(fmt) else ''
    min_l = int_or_none(get(schema, 'minLength'))
    max_l = int_or_none(get(schema, 'maxLength'))
    pattern = get(schema, 'pattern')
    if is_str(pattern) and pattern != '':
        s = gen_from_pattern(ctx, pattern, min_l, max_l)
        if s is not None:
            return s
    if fmt == 'date-time':
        return epoch_to_utc(ctx.rng.between(EPOCH_START, EPOCH_END)).strftime('%Y-%m-%dT%H:%M:%SZ')
    if fmt == 'date':
        days = ctx.rng.between(EPOCH_START // 86400, EPOCH_END // 86400)
        return epoch_to_utc(days * 86400).strftime('%Y-%m-%d')
    if fmt == 'time':
        secs = ctx.rng.between(0, 86399)
        return '%02d:%02d:%02d' % (secs // 3600, (secs % 3600) // 60, secs % 60)
    if fmt == 'email':
        return rand_str(ctx, LOWER_ALNUM, 8) + '@example.com'
    if fmt == 'uuid':
        chars = []
        for i in range(32):
            if i == 12:
                chars.append('4')
            elif i == 16:
                chars.append(HEXD[8 + ctx.rng.between(0, 3)])
            else:
                chars.append(HEXD[ctx.rng.between(0, 15)])
        u = ''.join(chars)
        return u[0:8] + '-' + u[8:12] + '-' + u[12:16] + '-' + u[16:20] + '-' + u[20:32]
    if fmt in ('uri', 'url', 'uri-reference', 'iri', 'iri-reference'):
        return 'https://example.com/' + rand_str(ctx, LOWER_ALNUM, 8)
    if fmt in ('hostname', 'idn-hostname'):
        return rand_str(ctx, LOWER_ALNUM, 8) + '.example.com'
    if fmt == 'ipv4':
        return '192.0.2.' + str(ctx.rng.between(1, 254))
    if fmt == 'ipv6':
        return '2001:db8::' + ('%x' % ctx.rng.between(1, 65535))
    if fmt == 'byte':
        raw = bytearray()
        for _ in range(12):
            raw.append(ctx.rng.between(0, 255))
        return base64.b64encode(bytes(raw)).decode('ascii')
    if fmt == 'password':
        return rand_str(ctx, ALNUM, 12)
    if fmt == 'binary':
        return rand_str(ctx, ALNUM, 16)
    return rand_str(ctx, ALNUM, choose_len(ctx, min_l, max_l))


# ----------------------------------------------------------------- pattern（正規表現）から文字列を生成
class PatternError(Exception):
    pass


def _codes(a, b):
    return list(range(a, b + 1))


DIGIT_SET = _codes(0x30, 0x39)
UPPER_SET = _codes(0x41, 0x5A)
LOWER_SET = _codes(0x61, 0x7A)
ALNUM_SET = DIGIT_SET + UPPER_SET + LOWER_SET
WORD_SET = sorted(ALNUM_SET + [0x5F])
NONDIGIT_SET = UPPER_SET + LOWER_SET
NONWORD_SET = [0x2D, 0x2E]
SPACE_SET = [0x20]
NEG_BASE = sorted(ALNUM_SET + [0x2D, 0x2E, 0x5F])
PRINTABLE_SET = _codes(0x21, 0x7E)


def valid_code(c):
    return c == 9 or c == 10 or c == 13 or 0x20 <= c <= 0xD7FF or 0xE000 <= c <= 0xFFFD or 0x10000 <= c <= 0x10FFFF


class PatternParser(object):
    def __init__(self, pattern):
        self.p = pattern
        self.i = 0
        self.n = len(pattern)

    def peek(self, k=0):
        j = self.i + k
        return self.p[j] if j < self.n else None

    def take(self):
        ch = self.p[self.i]
        self.i += 1
        return ch

    def parse(self):
        node = self.alt()
        if self.i != self.n:
            raise PatternError('unexpected')
        return node

    def alt(self):
        options = [self.seq()]
        while self.peek() == '|':
            self.take()
            options.append(self.seq())
        if len(options) == 1:
            return options[0]
        return ('alt', options)

    def seq(self):
        items = []
        while True:
            ch = self.peek()
            if ch is None or ch == '|' or ch == ')':
                break
            items.append(self.quant(self.atom()))
        return ('seq', items)

    def atom(self):
        ch = self.take()
        if ch == '(':
            if self.peek() == '?':
                self.take()
                c2 = self.peek()
                if c2 == ':':
                    self.take()
                elif c2 == 'P' and self.peek(1) == '<':
                    self.skip_group_name(2)
                elif c2 == '<' and self.peek(1) is not None and self.peek(1) not in ('=', '!'):
                    self.skip_group_name(1)
                else:
                    raise PatternError('unsupported group')
            node = self.alt()
            if self.peek() != ')':
                raise PatternError('unbalanced')
            self.take()
            return node
        if ch == '[':
            return self.char_class()
        if ch == '.':
            return ('set', ALNUM_SET)
        if ch == '^' or ch == '$':
            return ('seq', [])
        if ch == '\\':
            return self.escape(False)
        if ch in ('*', '+', '?', ')'):
            raise PatternError('quantifier without target')
        return ('lit', ord(ch))

    def skip_group_name(self, k):
        for _ in range(k):
            self.take()
        while True:
            c = self.peek()
            if c is None:
                raise PatternError('bad group name')
            self.take()
            if c == '>':
                return

    def hex_digits(self, count):
        s = self.p[self.i:self.i + count]
        if len(s) != count or not re.match(r'^[0-9A-Fa-f]+$', s):
            raise PatternError('bad hex escape')
        self.i += count
        return int(s, 16)

    def escape(self, in_class):
        if self.i >= self.n:
            raise PatternError('trailing backslash')
        ch = self.take()
        if ch == 'd':
            return ('set', DIGIT_SET)
        if ch == 'D':
            return ('set', NONDIGIT_SET)
        if ch == 'w':
            return ('set', WORD_SET)
        if ch == 'W':
            return ('set', NONWORD_SET)
        if ch == 's':
            return ('set', SPACE_SET)
        if ch == 'S':
            return ('set', ALNUM_SET)
        if ch == 't':
            return ('lit', 9)
        if ch == 'n':
            return ('lit', 10)
        if ch == 'r':
            return ('lit', 13)
        if ch == 'u':
            c = self.hex_digits(4)
            if not valid_code(c):
                raise PatternError('bad code point')
            return ('lit', c)
        if ch == 'x':
            c = self.hex_digits(2)
            if not valid_code(c):
                raise PatternError('bad code point')
            return ('lit', c)
        if ch in ('b', 'B', 'A', 'z', 'Z', 'G'):
            if in_class:
                raise PatternError('unsupported escape in class')
            return ('seq', [])
        if ch in '0123456789' or ch in ('p', 'P', 'k', 'c', 'f', 'v', 'e', 'a', 'Q', 'E'):
            raise PatternError('unsupported escape')
        return ('lit', ord(ch))

    def class_item(self):
        ch = self.take()
        if ch == '\\':
            return self.escape(True)
        if ch == '[' and self.peek() == ':':
            raise PatternError('posix class')
        return ('lit', ord(ch))

    def char_class(self):
        negate = False
        if self.peek() == '^':
            self.take()
            negate = True
        codes = set()
        first = True
        while True:
            if self.i >= self.n:
                raise PatternError('unterminated class')
            if self.peek() == ']' and not first:
                self.take()
                break
            first = False
            item = self.class_item()
            if item[0] == 'lit' and self.peek() == '-' and self.peek(1) is not None and self.peek(1) != ']':
                self.take()
                item2 = self.class_item()
                if item2[0] != 'lit':
                    raise PatternError('bad range')
                a = item[1]
                b = item2[1]
                if a > b or b - a > 70000:
                    raise PatternError('bad range')
                for c in range(a, b + 1):
                    codes.add(c)
            elif item[0] == 'lit':
                codes.add(item[1])
            elif item[0] == 'set':
                for c in item[1]:
                    codes.add(c)
            else:
                raise PatternError('bad class item')
        if negate:
            pool = [c for c in NEG_BASE if c not in codes]
            if not pool:
                pool = [c for c in PRINTABLE_SET if c not in codes]
        else:
            pool = sorted(codes)
        pool = [c for c in pool if valid_code(c)]
        if not pool:
            raise PatternError('empty class')
        return ('set', pool)

    def quant(self, node):
        while True:
            ch = self.peek()
            if ch == '*':
                self.take()
                mn, mx = 0, None
            elif ch == '+':
                self.take()
                mn, mx = 1, None
            elif ch == '?':
                self.take()
                mn, mx = 0, 1
            elif ch == '{':
                m = re.match(r'\{([0-9]+)(,([0-9]*))?\}', self.p[self.i:])
                if m is None:
                    return node
                self.i += len(m.group(0))
                mn = int(m.group(1))
                if m.group(2) is None:
                    mx = mn
                elif m.group(3) == '':
                    mx = None
                else:
                    mx = int(m.group(3))
                if (mx is not None and mx < mn) or mn > 1000:
                    raise PatternError('bad quantifier')
            else:
                return node
            if self.peek() == '?' or self.peek() == '+':
                self.take()
            node = ('rep', node, mn, mx)


def gen_node(ctx, node, out):
    t = node[0]
    if t == 'lit':
        out.append(chr(node[1]))
    elif t == 'set':
        pool = node[1]
        out.append(chr(pool[ctx.rng.between(0, len(pool) - 1)]))
    elif t == 'seq':
        for x in node[1]:
            gen_node(ctx, x, out)
    elif t == 'alt':
        opts = node[1]
        gen_node(ctx, opts[ctx.rng.between(0, len(opts) - 1)], out)
    elif t == 'rep':
        mn = node[2]
        mx = node[3]
        if mx is None:
            mx = mn + 5
        cap = max(mn, 16)
        if mx > cap:
            mx = cap
        for _ in range(ctx.rng.between(mn, mx)):
            gen_node(ctx, node[1], out)


def gen_from_pattern(ctx, pattern, min_l, max_l):
    try:
        ast = PatternParser(pattern).parse()
    except PatternError:
        ctx.warn('pattern を解釈できないためランダム英数字で代用しました: ' + pattern)
        return None
    try:
        rx = re.compile(pattern)
    except Exception:
        rx = None
    last = None
    for _ in range(10):
        out = []
        gen_node(ctx, ast, out)
        s = ''.join(out)
        last = s
        if rx is not None and rx.search(s) is None:
            continue
        if min_l is not None and cp_len(s) < min_l:
            continue
        if max_l is not None and cp_len(s) > max_l:
            continue
        return s
    if last is not None and (rx is None or rx.search(last) is not None):
        return last
    ctx.warn('pattern に一致する値を生成できないためランダム英数字で代用しました: ' + pattern)
    return None


# ----------------------------------------------------------------- パラメータ
METHOD_KEYS = ('get', 'put', 'post', 'delete', 'options', 'head', 'patch', 'trace')
ANY_KEY = 'x-amazon-apigateway-any-method'
SKIP_HEADERS = ('accept', 'content-type', 'authorization')
MANAGED_HEADERS = ('content-length', 'host', 'transfer-encoding', 'connection')


def collect_params(ctx, item, op):
    table = OrderedDict()
    for src in (get(item, 'parameters'), get(op, 'parameters')):
        if not is_arr(src):
            continue
        for p in src:
            rp = deref(ctx, p, 'parameter')
            if not is_obj(rp):
                continue
            name = get(rp, 'name')
            loc = get(rp, 'in')
            if not is_str(name) or not is_str(loc) or name == '':
                continue
            table[loc + '\n' + name] = rp
    return list(table.values())


def first_value(d):
    for k in d:
        return d[k]
    return None


def param_value(ctx, p):
    if ctx.use_examples:
        if 'example' in p:
            return to_gen(p['example'])
        exs = get(p, 'examples')
        if is_obj(exs) and len(exs) > 0:
            first = deref(ctx, first_value(exs), 'example')
            if is_obj(first) and 'value' in first:
                return to_gen(first['value'])
    schema = get(p, 'schema')
    if schema is None:
        content = get(p, 'content')
        if is_obj(content) and len(content) > 0:
            schema = get(first_value(content), 'schema')
    if schema is None:
        schema = OrderedDict([('type', 'string')])
    v = gen_value(ctx, schema, 0, [])
    if v is SKIP:
        v = ''
    return v


def simple_style(v, explode):
    if is_arr(v):
        return ','.join(scalar_str(x) for x in v)
    if is_obj(v):
        if explode:
            return ','.join(k + '=' + scalar_str(x) for k, x in v.items())
        return ','.join(k + ',' + scalar_str(x) for k, x in v.items())
    return scalar_str(v)


def query_pairs(name, v, style, explode):
    if is_arr(v):
        if style == 'form' and explode:
            return [(name, scalar_str(x)) for x in v]
        if style == 'spaceDelimited':
            sep = ' '
        elif style == 'pipeDelimited':
            sep = '|'
        else:
            sep = ','
        return [(name, sep.join(scalar_str(x) for x in v))]
    if is_obj(v):
        if style == 'deepObject':
            return [(name + '[' + k + ']', scalar_str(x)) for k, x in v.items()]
        if explode:
            return [(k, scalar_str(x)) for k, x in v.items()]
        return [(name, ','.join(k + ',' + scalar_str(x) for k, x in v.items()))]
    return [(name, scalar_str(v))]


# ----------------------------------------------------------------- リクエストボディ
def media_base(mt):
    return mt.split(';', 1)[0].strip().lower()


def choose_media(keys):
    # JSON 系 → フォーム → マルチパート の順に優先し、どれも無ければ定義の先頭のメディアタイプを使う
    preds = (
        lambda b: b == 'application/json',
        lambda b: b.endswith('+json'),
        lambda b: 'json' in b,
        lambda b: b == 'application/x-www-form-urlencoded',
        lambda b: b == 'multipart/form-data',
    )
    for pred in preds:
        for k in keys:
            if pred(media_base(k)):
                return k
    return keys[0]


def media_kind(mt):
    b = media_base(mt)
    if b == 'application/json' or b.endswith('+json') or 'json' in b or b == '*/*' or b == 'application/*':
        return 'json'
    if b == 'application/x-www-form-urlencoded':
        return 'form'
    if b.startswith('multipart/'):
        return 'multipart'
    if b.endswith('/xml') or b.endswith('+xml'):
        return 'xml'
    if b.startswith('text/'):
        return 'text'
    return 'binary'


def content_type_for(mt, kind):
    if '*' in media_base(mt):
        if kind == 'json':
            return 'application/json'
        if kind == 'xml':
            return 'application/xml'
        if kind == 'text':
            return 'text/plain'
        return 'application/octet-stream'
    return mt.strip()


def xml_name(n):
    n = re.sub(r'[^A-Za-z0-9_.-]', '_', n)
    if n == '' or not re.match(r'^[A-Za-z_]', n):
        n = '_' + n
    return n


def xml_text(s):
    return clean_text(s).replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')


def to_xml(v, name, indent):
    pad = '  ' * indent
    if is_obj(v):
        if len(v) == 0:
            return pad + '<' + name + '/>'
        lines = [pad + '<' + name + '>']
        for k, x in v.items():
            kn = xml_name(k)
            if is_arr(x):
                for it in x:
                    lines.append(to_xml(it, kn, indent + 1))
            else:
                lines.append(to_xml(x, kn, indent + 1))
        lines.append(pad + '</' + name + '>')
        return '\n'.join(lines)
    if is_arr(v):
        if len(v) == 0:
            return pad + '<' + name + '/>'
        return '\n'.join(to_xml(it, name, indent) for it in v)
    return pad + '<' + name + '>' + xml_text(scalar_str(v)) + '</' + name + '>'


def xml_root_name(ctx, schema):
    s = schema
    name = None
    guard = 0
    while is_obj(s) and '$ref' in s and guard < 32:
        ref = s['$ref']
        if is_str(ref):
            name = ref.split('/')[-1]
        s = resolve_ref(ctx, ref)
        guard += 1
    x = get(s, 'xml')
    if is_obj(x) and is_str(get(x, 'name')) and x['name'] != '':
        name = x['name']
    if not name:
        name = 'root'
    return xml_name(name)


def build_body(ctx, op):
    if not has(op, 'requestBody'):
        return None
    rb = deref(ctx, op['requestBody'], 'requestBody')
    content = get(rb, 'content')
    if not is_obj(content) or len(content) == 0:
        return None
    mt = choose_media(list(content.keys()))
    media = content[mt]
    schema = get(media, 'schema')
    kind = media_kind(mt)
    have = False
    value = None
    if ctx.use_examples and is_obj(media):
        if 'example' in media:
            value = to_gen(media['example'])
            have = True
        else:
            exs = get(media, 'examples')
            if is_obj(exs) and len(exs) > 0:
                first = deref(ctx, first_value(exs), 'example')
                if is_obj(first) and 'value' in first:
                    value = to_gen(first['value'])
                    have = True
    if not have:
        if schema is None:
            value = OrderedDict() if kind in ('json', 'form', 'multipart', 'xml') else rand_str(ctx, ALNUM, 16)
        else:
            value = gen_value(ctx, schema, 0, [])
            if value is SKIP:
                value = OrderedDict()
    ctype = content_type_for(mt, kind)
    if kind == 'json':
        return {'kind': kind, 'ctype': ctype, 'text': to_json(value, 0, True), 'pairs': []}
    if kind == 'form' or kind == 'multipart':
        pairs = []
        if is_obj(value):
            for k, x in value.items():
                if is_arr(x):
                    for it in x:
                        pairs.append((k, scalar_str(it)))
                else:
                    pairs.append((k, scalar_str(x)))
        else:
            pairs.append(('value', scalar_str(value)))
        return {'kind': kind, 'ctype': ctype, 'text': '', 'pairs': pairs}
    if kind == 'xml':
        text = '<?xml version="1.0" encoding="UTF-8"?>\n' + to_xml(value, xml_root_name(ctx, schema), 0)
        return {'kind': kind, 'ctype': ctype, 'text': text, 'pairs': []}
    text = value if is_str(value) else to_json(value, 0, True)
    return {'kind': kind, 'ctype': ctype, 'text': text, 'pairs': []}


# ----------------------------------------------------------------- レスポンス（Accept・期待ステータス）
def is_2xx(k):
    return re.match(r'^2[0-9][0-9]$', k) is not None or k == '2XX'


def pick_accept(ctx, op):
    responses = get(op, 'responses')
    if not is_obj(responses):
        return None
    for k, resp in responses.items():
        if not is_2xx(str(k).strip().upper()):
            continue
        r = deref(ctx, resp, 'response')
        content = get(r, 'content')
        if is_obj(content) and len(content) > 0:
            mt = list(content.keys())[0]
            if '*' in mt:
                return None
            return mt.strip()
    return None


def success_codes(op):
    responses = get(op, 'responses')
    codes = []
    has_range = False
    if is_obj(responses):
        for k in responses.keys():
            ks = str(k).strip().upper()
            if re.match(r'^2[0-9][0-9]$', ks):
                if ks not in codes:
                    codes.append(ks)
            elif ks == '2XX':
                has_range = True
    return codes, has_range


# ----------------------------------------------------------------- セキュリティ
def effective_security(ctx, op):
    if has(op, 'security'):
        return get(op, 'security')
    return get(ctx.doc, 'security')


def security_text(ctx, op):
    sec = effective_security(ctx, op)
    if not is_arr(sec):
        return None
    if len(sec) == 0:
        return 'なし（認証不要）'
    alts = []
    for req in sec:
        if not is_obj(req):
            continue
        if len(req) == 0:
            alts.append('認証なしも可')
            continue
        parts = []
        for name, scopes in req.items():
            sc = [x for x in scopes if is_str(x)] if is_arr(scopes) else []
            parts.append(name + ('[' + ', '.join(sc) + ']' if sc else ''))
        alts.append(' かつ '.join(parts))
    return ' または '.join(alts) if alts else None


def scheme_uses_authorization(scheme):
    if not is_obj(scheme):
        return False
    t = get(scheme, 'type')
    if t == 'oauth2' or t == 'openIdConnect' or t == 'http':
        return True
    if t == 'apiKey':
        return get(scheme, 'in') == 'header' and is_str(get(scheme, 'name')) \
            and scheme['name'].lower() == 'authorization'
    return False


def needs_auth_header(ctx, ops):
    schemes = get(get(ctx.doc, 'components'), 'securitySchemes')
    for info in ops:
        sec = effective_security(ctx, info['op'])
        if not is_arr(sec):
            continue
        for req in sec:
            if not is_obj(req):
                continue
            for name in req.keys():
                sch = deref(ctx, get(schemes, name), 'securityScheme')
                if scheme_uses_authorization(sch):
                    return True
    return False


# ----------------------------------------------------------------- オペレーション
def collect_operations(ctx):
    ops = []
    paths = get(ctx.doc, 'paths')
    for path, item in paths.items():
        if path.startswith('x-'):
            continue
        item = deref(ctx, item, 'pathItem')
        if not is_obj(item):
            continue
        for key, op in item.items():
            k = key.lower()
            if k in METHOD_KEYS:
                method = k.upper()
                is_any = False
            elif k == ANY_KEY:
                method = ctx.any_method
                is_any = True
            else:
                continue
            if not is_obj(op):
                continue
            ops.append({'path': path, 'item': item, 'op': op, 'key': key, 'method': method, 'is_any': is_any})
    return ops


PLACEHOLDER_RE = re.compile(r'\{([^{}]+)\}')


def build_operation(ctx, info, label):
    path = info['path']
    item = info['item']
    op = info['op']
    method = info['method']
    params = collect_params(ctx, item, op)

    # 1. パスパラメータ
    path_vals = OrderedDict()
    for p in params:
        if get(p, 'in') == 'path':
            path_vals[p['name']] = (param_value(ctx, p), get(p, 'explode') is True)
    for m in PLACEHOLDER_RE.finditer(path):
        ph = m.group(1)
        base = ph[:-1] if ph.endswith('+') else ph
        if base not in path_vals:
            ctx.warn('%s: パス変数 {%s} のパラメータ定義が無いため英数字 8 文字を仮設定しました' % (label, base))
            path_vals[base] = (rand_str(ctx, ALNUM, 8), False)
    parts = []
    pos = 0
    for m in PLACEHOLDER_RE.finditer(path):
        parts.append(path[pos:m.start()])
        ph = m.group(1)
        greedy = ph.endswith('+')
        base = ph[:-1] if greedy else ph
        v, explode = path_vals[base]
        parts.append(oa_pct_encode(simple_style(v, explode), greedy))
        pos = m.end()
    parts.append(path[pos:])
    real_path = ''.join(parts)
    if not real_path.startswith('/'):
        real_path = '/' + real_path

    # 2. クエリパラメータ
    qpairs = []
    for p in params:
        if get(p, 'in') != 'query':
            continue
        if ctx.required_only and get(p, 'required') is not True:
            continue
        v = param_value(ctx, p)
        style = get(p, 'style')
        style = style if is_str(style) else 'form'
        explode = get(p, 'explode')
        explode = explode if is_bool(explode) else (style == 'form')
        qpairs.extend(query_pairs(p['name'], v, style, explode))

    # 3. ヘッダパラメータ
    hparams = []
    for p in params:
        if get(p, 'in') != 'header':
            continue
        name = p['name']
        lname = name.lower()
        if lname in SKIP_HEADERS:
            continue
        if lname == 'api-key':
            if not ctx.api_key_set:
                ctx.warn('%s: ヘッダ %s は API キーの実値が必要なため生成しません（OpenAPI 用設定ファイル（-c）の API_HEADERS に「%s: キーの値」を指定してください）'
                         % (label, name, name))
            continue
        if lname in MANAGED_HEADERS:
            ctx.warn('%s: ヘッダ %s は HTTP クライアントが自動設定するため生成しません' % (label, name))
            continue
        if ctx.required_only and get(p, 'required') is not True:
            continue
        v = param_value(ctx, p)
        hparams.append((name, header_safe(simple_style(v, get(p, 'explode') is True))))

    # 4. クッキーパラメータ
    cookies = []
    for p in params:
        if get(p, 'in') != 'cookie':
            continue
        if ctx.required_only and get(p, 'required') is not True:
            continue
        v = param_value(ctx, p)
        cookies.append(p['name'] + '=' + oa_pct_encode(simple_style(v, False), False))

    # 5. リクエストボディ
    body = build_body(ctx, op)

    # 6. 組み立て（API一覧（Excel / CSV）の 1 行と同じ項目を返す。クエリを Parameters に入れるか URL に
    #    付けるかは、Excel / CSV の行と同じく normalize_api がメソッドで決める）
    body_type = ''
    body_text = ''
    headers = []
    if body is not None:
        if body['kind'] in ('form', 'multipart'):
            body_type = body['kind']
            body_text = oa_row_pairs(body['pairs'])
            if body['kind'] == 'form':
                headers.append(('Content-Type', body['ctype']))
        else:
            body_type = 'json' if body['kind'] == 'json' else 'raw'
            body_text = body['text']
            headers.append(('Content-Type', body['ctype']))
    accept = pick_accept(ctx, op)
    if accept is not None:
        headers.append(('Accept', accept))
    headers.extend(hparams)
    if cookies:
        headers.append(('Cookie', '; '.join(cookies)))

    codes, has_range = success_codes(op)
    has_data = len(path_vals) > 0 or len(qpairs) > 0 or len(hparams) > 0 or len(cookies) > 0 or body is not None
    comment = operation_comment(ctx, info, body, has_data)
    if has_range and not codes:
        comment += '\n※ 成功レスポンスが 2XX（範囲指定）のため期待ステータスは空欄です（JMeter の既定の判定: 2xx / 3xx = 成功）。'
    return {'method': method, 'url': real_path, 'query': oa_row_pairs(qpairs),
            'headers': '\n'.join(k + ': ' + v for k, v in headers), 'body_type': body_type, 'body': body_text,
            'expected_status': ','.join(codes), 'comment': comment}


def operation_comment(ctx, info, body, has_data):
    op = info['op']
    lines = []
    summary = get(op, 'summary')
    if is_str(summary) and summary.strip():
        lines.append('概要: ' + summary.strip())
    desc = get(op, 'description')
    if is_str(desc) and desc.strip():
        lines.append('説明: ' + desc.strip())
    lines.append('OpenAPI 定義: ' + info['key'] + ' ' + info['path'])
    op_id = get(op, 'operationId')
    if is_str(op_id) and op_id:
        lines.append('operationId: ' + op_id)
    if info['is_any']:
        lines.append('ANY メソッド（x-amazon-apigateway-any-method）のため ' + info['method']
                     + ' で代表して送信します（生成時の --any-method / -AnyMethod で変更可）。')
    sec = security_text(ctx, op)
    if sec is not None:
        lines.append('認可（security）: ' + sec)
    integ = get(op, 'x-amazon-apigateway-integration')
    if is_obj(integ):
        items = []
        for k in ('type', 'httpMethod', 'uri', 'connectionType'):
            if is_str(get(integ, k)):
                items.append(k + '=' + integ[k])
        if items:
            lines.append('API Gateway 統合: ' + ', '.join(items))
    if get(op, 'deprecated') is True:
        lines.append('※ deprecated（非推奨）に指定された API です。')
    if body is not None and body['kind'] == 'multipart':
        lines.append('※ multipart/form-data はファイル項目もテキスト値で送信します。実ファイルを送る場合は「ファイルアップロード」タブで設定してください。')
    if has_data:
        lines.append('テストデータ: OpenAPI のデータ型定義からランダム生成（乱数シード ' + str(ctx.seed) + '）')
    return '\n'.join(lines)


# JSON 文字列（エスケープ込み）と ${...} を先頭から順に拾う。文字列は読み飛ばし、${...} だけを置き換える。
# 直後が : の ${...} はキーの位置なので置き換えない（PowerShell 5.1 の JSON 解析は引用符なしのキーを
# 受け付けてしまい、実装間で結果が変わるため。置き換えなければどの実装でも構文エラーになる）
PLACEHOLDER_SCAN = re.compile(r'"[^"\\]*(?:\\.[^"\\]*)*"|\$\{[^{}"\r\n]+\}(?![ \t\r\n]*:)', re.S)
PLACEHOLDER_LINES_SHOWN = 10


def fill_placeholders(text, value):
    """JSON 文字列の外に引用符なしで書かれた ${...}（後で置換するプレースホルダ）を数値 value に置き換える。
    文字列の中の ${...}（"${stageVariables.x}" など）は正しい JSON なので変更しない。
    戻り値: (置換後の文字列, [(プレースホルダ, [行番号, ...]), ...]（初出順）)"""
    if '${' not in text:
        return text, []
    found = OrderedDict()
    out = []
    pos = 0
    line = 1
    for m in PLACEHOLDER_SCAN.finditer(text):
        s = m.start()
        if text[s] != '$':
            continue    # 文字列（読み飛ばすだけ）
        line += text.count('\n', pos, s)
        out.append(text[pos:s])
        out.append(value)
        found.setdefault(m.group(0), []).append(line)
        pos = m.end()
    out.append(text[pos:])
    return ''.join(out), list(found.items())


def placeholder_warning(name, lines, value):
    shown = ', '.join(str(n) for n in lines[:PLACEHOLDER_LINES_SHOWN])
    more = ' ほか' if len(lines) > PLACEHOLDER_LINES_SHOWN else ''
    return ('引用符なしのプレースホルダ %s は JSON として読めないため、数値 %s に置き換えて読み込みました（%d 箇所・行 %s%s）'
            % (name, value, len(lines), shown, more))


def load_json(path, placeholder_value):
    with open(path, 'rb') as f:
        raw = f.read()
    try:
        if raw[:3] == b'\xef\xbb\xbf':
            text = raw[3:].decode('utf-8')
        elif raw[:2] == b'\xff\xfe':
            text = raw[2:].decode('utf-16-le')
        elif raw[:2] == b'\xfe\xff':
            text = raw[2:].decode('utf-16-be')
        else:
            text = raw.decode('utf-8')
    except UnicodeDecodeError:
        raise GenError('入力ファイルの文字コードを UTF-8 として読めません: ' + path)

    def bad_constant(name):
        raise ValueError('JSON では使えない値です: ' + name)

    text, placeholders = fill_placeholders(text, placeholder_value)
    try:
        doc = json.loads(text, object_pairs_hook=OrderedDict, parse_float=Decimal, parse_constant=bad_constant)
    except ValueError as e:
        raise GenError('JSON の構文エラーです（YAML 形式は未対応です）: ' + str(e))
    return doc, placeholders


_OA_ROW_UNSAFE = re.compile(r'[&%$\x00-\x1f\x7f]')


def oa_row_pairs(pairs):
    # (名前, 値) の並びを API一覧の「名前=値」（改行区切り）の形にする。
    # 区切り文字（& 改行 =）・%・$・前後の空白を含むものは、名前・値とも URL エンコードして書く
    out = []
    for k, v in pairs:
        if k == k.strip() and v == v.strip() and '=' not in k and not _OA_ROW_UNSAFE.search(k + v):
            out.append(k + '=' + v)
            continue
        ek = oa_pct_encode(k, False)
        ev = oa_pct_encode(v, False)
        if '%' not in ev:
            # 値に %XX が無いと JMeter がエンコード済みの名前をもう一度エンコードするため、値もすべて %XX にする
            ev = ''.join('%%%02X' % b for b in bytearray(ev.encode('utf-8')))
        out.append(ek + '=' + ev)
    return '\n'.join(out)


OA_VARS = (('public_base_url', 'PUBLIC_BASE_URL'), ('name_prefix', 'NAME_PREFIX'))


def oa_fill_vars(text, conf, where):
    # OpenAPI 定義のプレースホルダ ${public_base_url} / ${name_prefix} を OpenAPI 用設定（-c）の値に置き換える。
    # 未設定（空欄）の場合は隣の区切り文字（- _ . / 空白）ごと取り除き、パスの先頭やテストプラン名に ${...} を残さない
    for name, key in OA_VARS:
        ph = '${' + name + '}'
        if ph not in text:
            continue
        value = (conf or {}).get(key, '')
        if value:
            text = text.replace(ph, value)
            info('OpenAPI 定義 %s の %s: %s（OpenAPI 用設定 %s の値）' % (where, ph, value, key))
        else:
            text = re.sub(r'^\$\{%s\}[-_. /]*|[-_. /]*\$\{%s\}' % (name, name), '', text)
            info('OpenAPI 定義 %s の %s: 未設定のため取り除きます（OpenAPI 用設定ファイル（-c）の %s で設定できます）' % (where, ph, key))
    return text


def oa_server_url(doc, conf):
    # 接続先の既定値: servers[0].url（${public_base_url} は OpenAPI 用設定の値、{変数} は既定値に置換）。http(s) の URL でなければ http://localhost:8080 を前に付ける
    servers = get(doc, 'servers')
    srv = servers[0] if is_arr(servers) and len(servers) > 0 and is_obj(servers[0]) else None
    url = get(srv, 'url')
    url = url.strip() if is_str(url) else ''
    url = oa_fill_vars(url, conf, 'servers[0].url')
    variables = get(srv, 'variables')

    def var_default(m):
        d = get(get(variables, m.group(1)), 'default')
        return d if is_str(d) else m.group(0)
    url = PLACEHOLDER_RE.sub(var_default, url).rstrip('/')
    if re.match(r'^[Hh][Tt][Tt][Pp][Ss]?://', url):
        return url
    if url and not url.startswith('/'):
        url = '/' + url
    return 'http://localhost:8080' + url


def load_openapi(input_path, opts):
    # OpenAPI 3.0.x 定義（JSON）を読み込み、API一覧（Excel / CSV）と同じ項目の行に変換する
    # 戻り値: (行の一覧, OpenAPI 定義から求めた設定の既定値 [(設定キー, 値), ...])
    seed_text = opts.get('seed') or ''
    if seed_text and not re.match(r'^[0-9]{1,12}\Z', seed_text):
        raise ToolError('--seed は 0 以上の整数（12 桁以内）で指定してください: ' + seed_text)
    seed = int(seed_text) if seed_text else int.from_bytes(os.urandom(4), 'big') % 1000000000
    any_method = (opts.get('any_method') or 'GET').upper()
    if any_method not in ('GET', 'POST', 'PUT', 'PATCH', 'DELETE', 'HEAD', 'OPTIONS'):
        raise ToolError('--any-method は GET/POST/PUT/PATCH/DELETE/HEAD/OPTIONS のいずれかです: ' + any_method)
    placeholder_value = opts.get('placeholder_value') or '3000'
    if not re.match(r'^-?(0|[1-9][0-9]{0,14})([.][0-9]{1,12})?\Z', placeholder_value):
        raise ToolError('--placeholder-value は JSON の数値（例: 3000、-1、2.5。整数部 15 桁・小数部 12 桁以内）で指定してください: '
                        + placeholder_value)
    doc, placeholders = load_json(input_path, placeholder_value)
    if not is_obj(doc):
        raise GenError('OpenAPI 定義のルートが JSON オブジェクトではありません')
    if has(doc, 'swagger'):
        raise GenError('Swagger 2.0 形式には対応していません。OpenAPI 3.0.x 形式（"openapi": "3.0.3" など）に変換してください')
    ver = get(doc, 'openapi')
    if not is_str(ver) or not re.match(r'^3\.[0-9]+(\.[0-9]+)?', ver):
        raise GenError('"openapi" フィールドが 3.x ではありません（OpenAPI 3.0.x の JSON を指定してください）')
    if not ver.startswith('3.0'):
        warn('OpenAPI %s は 3.0.x 以外です。3.0.x と共通の項目のみを解釈します。' % ver)
    if not is_obj(get(doc, 'paths')):
        raise GenError('"paths" が定義されていません')
    info_obj = get(doc, 'info')
    title = get(info_obj, 'title')
    title = oa_fill_vars(title, opts.get('_oa_conf'), 'info.title') if is_str(title) else title
    title = title.strip() if is_str(title) and title.strip() else 'OpenAPI'
    api_version = get(info_obj, 'version')
    api_version = api_version.strip() if is_str(api_version) and api_version.strip() else '-'

    ctx = Ctx(doc, seed, {'required_only': bool(opts.get('required_only')), 'use_examples': bool(opts.get('use_examples')),
                          'any_method': any_method})
    ctx.api_key_set = bool(opts.get('_oa_api_key'))    # OpenAPI 用設定 API_HEADERS に API-Key（既定値以外）があれば build_operation で警告しない
    for ph_name, ph_lines in placeholders:
        ctx.warn(placeholder_warning(ph_name, ph_lines, placeholder_value))
    ops = collect_operations(ctx)
    if not ops:
        raise GenError('"paths" に呼び出し可能な操作（get/post/put/delete など）が 1 件もありません')
    records = []
    for idx, op_info in enumerate(ops):
        name = (('ANY(' + op_info['method'] + ')') if op_info['is_any'] else op_info['method']) + ' ' + op_info['path']
        op_id = get(op_info['op'], 'operationId')
        if is_str(op_id) and op_id:
            name += ' [' + op_id + ']'
        row = build_operation(ctx, op_info, 'API%02d_%s' % (idx + 1, name))
        row.update({'no': str(idx + 1), 'enabled': 'Y', 'name': name})
        rec = {'_row': idx + 1}
        for c in SPEC['api_columns']:
            rec[c['key']] = clean(row.get(c['key'], '')).strip()
        rec['_fp'] = oa_fingerprint(ctx, op_info)    # 更新（-u）用: 操作の定義の指紋
        records.append(rec)
    opts['_oa_needs_auth'] = needs_auth_header(ctx, ops)    # 認可の警告は設定ファイル（-s）を読んだ後に load_input で判定する
    opts['_oa_seed'] = seed
    for w in ctx.warnings:
        warn(w)
    if placeholders:
        info('引用符なしの ${...} に入れる数値は --placeholder-value で変更できます')
    info('API定義: OpenAPI %s（%s %s、%d 件）' % (ver, title, api_version, len(records)))
    info('テストデータの乱数シード: %d（--seed %d を指定すると同じテストデータを再生成できます）' % (seed, seed))
    defaults = [('base_url', oa_server_url(doc, opts.get('_oa_conf'))), ('output_file', safe_name(title) + '.jmx'),
                ('test_plan_name', '%s %s（OpenAPI %s）' % (title, api_version, ver))]
    return records, defaults


# ----------------------------------------------------------------------------- OpenAPI 用設定ファイル（-c）
# OpenAPI 定義（.json）から生成する場合だけ使う設定ファイル（設定 CSV と同じ形式）。設定ファイル・OpenAPI 定義の値より
# 優先して、次の設定に反映する（Windows 版 jmeter_api_tool.ps1 の -OpenApiConfig と同じ処理）
#   PROTOCOL / HOST / PORT → base_url の接続先（パス以降はそのまま）、HEADERS → common_headers、
#   TOKEN_DIR → access_token_dir / access_token_file（フォルダ直下の *.token のうち更新日時が最も新しいファイル）
#   PUBLIC_BASE_URL / NAME_PREFIX → OpenAPI 定義の ${public_base_url}（servers[0].url）/ ${name_prefix}（info.title）の値（oa_fill_vars）
#   API_HEADERS → 各 API のヘッダ（API一覧の「ヘッダ」。-c を指定しない場合・空欄は既定値。add_oa_api_headers）
TOKEN_PATTERN = '*.token'


def load_oa_config(path):
    if not os.path.isfile(path):
        raise ToolError('OpenAPI 用設定ファイルが見つかりません: ' + path)
    text, enc = read_text(path)
    values = load_settings_table([(i + 1, r) for i, r in enumerate(parse_csv(text))], os.path.basename(path))
    known = [c['key'] for c in SPEC['openapi_config']]
    conf = {}
    for k in sorted(values):
        if k.upper() in known:
            conf[k.upper()] = values[k]
        else:
            warn('OpenAPI 用設定: 不明な設定キー「%s」は無視します' % k)
    info('OpenAPI 用設定: %s' % path)
    return conf


def apply_oa_config(raw, conf, out_dir):
    # 接続先: 指定した項目だけ base_url を置き換える（空欄の PROTOCOL / HOST は OpenAPI 定義・設定ファイルの値のまま、空欄の PORT は未設定）
    proto, host, port = conf.get('PROTOCOL', '').lower(), conf.get('HOST', ''), conf.get('PORT', '')
    if proto not in ('', 'http', 'https'):
        warn('OpenAPI 用設定 PROTOCOL: 「%s」は http / https ではないため無視します' % conf['PROTOCOL'])
        proto = ''
    if host and not re.match(r'^(\[[0-9A-Fa-f:.]+\]|[A-Za-z0-9._-]+)\Z', host):
        warn('OpenAPI 用設定 HOST: 「%s」はホスト名に使えない文字を含むため無視します（例: api.example.com）' % host)
        host = ''
    if port and not (re.match(r'^[0-9]{1,5}\Z', port) and 1 <= int(port) <= 65535):
        warn('OpenAPI 用設定 PORT: 「%s」は 1～65535 の数値ではないため無視します' % port)
        port = ''
    if proto or host or port:
        base = raw.get('base_url', '')
        if not _URL_RE.match(base):
            base = 'https://' + base    # normalize_api と同じく、スキームが無い場合は https とみなす
        m = _URL_RE.match(base)
        p = parse_url(base)
        if m and p:
            raw['base_url'] = (proto or p['scheme']) + '://' + (host or p['host']) + (':' + port if port else '') + base[m.end(2):]
            info('設定 base_url: %s（OpenAPI 用設定の PROTOCOL / HOST / PORT を反映）' % raw['base_url'])
    # HTTP ヘッダ: 共通ヘッダに追加し、同じ名前のヘッダは置き換える。TOKEN_DIR があれば Authorization: Bearer ${ACCESS_TOKEN}
    add = parse_headers(conf.get('HEADERS', ''), 'OpenAPI 用設定 HEADERS')
    if conf.get('TOKEN_DIR') and 'authorization' not in [k.lower() for k, v in add]:
        add.append(('Authorization', 'Bearer ${ACCESS_TOKEN}'))
    if add:
        names = [k.lower() for k, v in add]
        kept = [x for x in norm_nl(raw.get('common_headers', '')).split('\n')
                if x.strip() and x.split(':', 1)[0].strip().lower() not in names]
        raw['common_headers'] = '\n'.join(kept + [k + ': ' + v for k, v in add])
        info('共通ヘッダ: %s（OpenAPI 用設定。設定 common_headers の同じ名前のヘッダは置き換え）' % ', '.join(k for k, v in add))
    # アクセストークン: JMeter の実行中、API を呼び出すたびに TOKEN_DIR 直下の *.token のうち最新のファイルを読み込む
    if conf.get('TOKEN_DIR'):
        raw['access_token_dir'] = conf['TOKEN_DIR']
        raw['access_token_file'] = TOKEN_PATTERN
        info('アクセストークン: %s 直下の %s のうち更新日時が最も新しいファイルを、JMeter の実行中に API の呼び出しごとに読み込みます'
             % (conf['TOKEN_DIR'], TOKEN_PATTERN))
        check_token_dir(conf['TOKEN_DIR'], TOKEN_PATTERN, out_dir)


def check_token_dir(dir_text, pattern, base_dir):
    # 生成時の確認: JMeter の実行時と同じ規則でフォルダを決め（\ と / は同じ扱い、相対パスは JMX のフォルダ基準、
    # ; で区切った場合は最初に存在するフォルダ）、現在いちばん新しいファイルを表示する
    dirs = []
    for d in dir_text.split(';'):
        d = d.strip().replace('\\', '/')
        if d:
            d = d[2:] if d.startswith('~/') else d
            dirs.append(os.path.normpath(d if os.path.isabs(d) else os.path.join(os.path.abspath(base_dir), d)))
    found = [d for d in dirs if os.path.isdir(d)]
    if not found:
        warn('アクセストークンのフォルダが見つかりません: %s（JMeter を実行する環境に作成してください。相対パスは JMX のフォルダ基準）'
             % ' / '.join(dirs))
        return
    try:
        files = [f for f in os.listdir(found[0]) if fnmatch.fnmatch(f, pattern) and os.path.isfile(os.path.join(found[0], f))]
        stamp = lambda f: (int(os.path.getmtime(os.path.join(found[0], f)) * 1000), f)
        newest = max(files, key=stamp) if files else None
    except OSError as e:
        warn('アクセストークンのフォルダを読み取れません: %s（%s）' % (found[0], e))
        return
    if newest is None:
        warn('アクセストークンのフォルダに %s のファイルがありません: %s（JMeter の実行までに配置してください）' % (pattern, found[0]))
        return
    info('アクセストークン: 現在の最新ファイル %s（更新日時 %s。%s は %d 件）'
         % (os.path.join(found[0], newest),
            datetime.datetime.fromtimestamp(stamp(newest)[0] / 1000.0).strftime('%Y-%m-%d %H:%M:%S'), pattern, len(files)))


def add_oa_api_headers(records, conf):
    # 各 API のヘッダ（API ごとの HTTP Header Manager）の初期値: OpenAPI 用設定 API_HEADERS（-c を指定しない場合・空欄は既定値、「なし」は付けない）。
    # OpenAPI 定義から求めた同じ名前のヘッダ（Accept など）は置き換え、API一覧の「ヘッダ」の後ろに追加する
    text = (conf or {}).get('API_HEADERS', '')
    where = 'OpenAPI 用設定 API_HEADERS'
    if not text:
        text = [c['default'] for c in SPEC['openapi_config'] if c['key'] == 'API_HEADERS'][0]
        where += ' の既定値'
    add = [] if parse_onoff(text) == 'OFF' else parse_headers(text, where)
    if not add:
        return
    names = [k.lower() for k, v in add]
    for rec in records:
        kept = [x for x in rec['headers'].split('\n') if x.strip() and x.split(':', 1)[0].strip().lower() not in names]
        rec['headers'] = '\n'.join(kept + [k + ': ' + v for k, v in add])
    info('各 API のヘッダ: %s（%s。OpenAPI 定義から求めた同じ名前のヘッダは置き換え）' % (', '.join(k for k, v in add), where))


def oa_api_key_set(text):
    # OpenAPI 用設定 API_HEADERS に API-Key（既定値の XXXXXXX 以外の値）があるか。ある場合は API-Key の警告を表示しない
    pat = r'(?im)^[ \t]*api-key[ \t]*:[ \t]*(.*?)[ \t]*$'
    m = re.search(pat, text)
    d = re.search(pat, [c['default'] for c in SPEC['openapi_config'] if c['key'] == 'API_HEADERS'][0])
    return bool(m and m.group(1) and m.group(1) != (d.group(1) if d else ''))


# ----------------------------------------------------------------------------- input loading
def _aliases():
    m = {}
    for c in SPEC['api_columns']:
        for a in [c['key'], c['header']] + c['aliases']:
            m.setdefault(normalize_key(a), c['key'])
    return m


def load_api_table(rows, source):
    alias = _aliases()
    header_at = -1
    colmap = {}
    for idx, (rn, cells) in enumerate(rows):
        found = {}
        for ci, v in enumerate(cells):
            k = alias.get(normalize_key(v))
            if k and k not in found:
                found[k] = ci
        if 'url' in found:
            header_at, colmap = idx, found
            break
    if header_at < 0:
        raise ToolError('%s: 見出し行（「URL」列を含む行）が見つかりません' % source)
    records = []
    for rn, cells in rows[header_at + 1:]:
        first = ''
        for v in cells:
            if v.strip():
                first = v.strip()
                break
        if first == '' or first.startswith('#'):
            continue
        rec = {'_row': rn}
        for key in [c['key'] for c in SPEC['api_columns']]:
            ci = colmap.get(key, -1)
            rec[key] = clean(cells[ci]).strip() if 0 <= ci < len(cells) else ''
        records.append(rec)
    return records


def load_settings_table(rows, source):
    kc = vc = -1
    start = -1
    for idx, (rn, cells) in enumerate(rows):
        kc = vc = -1
        for ci, v in enumerate(cells):
            nv = normalize_key(v)
            if kc < 0 and nv in ('設定キー', 'key', 'キー'):
                kc = ci
            if vc < 0 and nv in ('設定値', 'value', '値'):
                vc = ci
        if kc >= 0 and vc >= 0:
            start = idx
            break
    if start < 0:
        raise ToolError('%s: 見出し行（「設定キー」「設定値」）が見つかりません' % source)
    values = {}
    for rn, cells in rows[start + 1:]:
        key = cells[kc].strip() if kc < len(cells) else ''
        if not re.match(r'^[A-Za-z_][A-Za-z0-9_]*$', key):
            continue
        values[key] = clean(cells[vc]).strip() if vc < len(cells) else ''
    return values


def _find_sheet(sheets, names):
    want = [normalize_key(x) for x in names]
    for name, rows in sheets:
        if normalize_key(name) in want:
            return name, rows
    return None, None


def load_input(input_path, settings_path, opts=None):
    ext = os.path.splitext(input_path)[1].lower()
    raw = None
    oa_defaults = []
    oa_conf = None
    if (opts or {}).get('oa_config') and ext != '.json':
        warn('OpenAPI 用設定ファイル（-c）は OpenAPI 定義（.json）を入力する場合だけ使います（無視します）: ' + opts['oa_config'])
    if ext in ('.xlsx', '.xlsm'):
        sheets = xlsx_read(input_path)
        name, rows = _find_sheet(sheets, [TPL['api_sheet'], 'API', 'APIs', 'api_list'])
        if rows is None:
            for nm, rs in sheets:
                try:
                    load_api_table(rs, nm)
                    name, rows = nm, rs
                    break
                except ToolError:
                    continue
        if rows is None:
            raise ToolError('「%s」シートが見つかりません: %s' % (TPL['api_sheet'], input_path))
        records = load_api_table(rows, '%s（%sシート）' % (os.path.basename(input_path), name))
        if not settings_path:
            sname, srows = _find_sheet(sheets, [TPL['settings_sheet'], 'settings'])
            if srows is not None:
                raw = load_settings_table(srows, '%s（%sシート）' % (os.path.basename(input_path), sname))
                info('設定: 「%s」シートを使用します' % sname)
            else:
                warn('「%s」シートが無いため、設定はすべて既定値を使用します' % TPL['settings_sheet'])
        info('API定義: 「%s」シート（%d 件）' % (name, len(records)))
    elif ext in ('.csv', '.txt'):
        text, enc = read_text(input_path)
        rows = [(i + 1, r) for i, r in enumerate(parse_csv(text))]
        records = load_api_table(rows, os.path.basename(input_path))
        info('API定義: CSV（文字コード %s、%d 件）' % (enc, len(records)))
        if not settings_path:
            cand = os.path.splitext(input_path)[0] + '_settings.csv'
            if os.path.isfile(cand):
                settings_path = cand
            else:
                warn('設定 CSV（%s）が無いため、設定はすべて既定値を使用します' % os.path.basename(cand))
    elif ext == '.json':
        opts = opts or {}
        if opts.get('oa_config'):
            oa_conf = load_oa_config(opts['oa_config'])
            opts['_oa_auth'] = bool(oa_conf.get('TOKEN_DIR')) or bool(re.search(r'(?im)^\s*authorization\s*:', oa_conf.get('HEADERS', '') + '\n' + oa_conf.get('API_HEADERS', '')))
            opts['_oa_api_key'] = oa_api_key_set(oa_conf.get('API_HEADERS', ''))
            opts['_oa_conf'] = oa_conf
        records, oa_defaults = load_openapi(input_path, opts)
        if opts.get('update'):
            oa_update_rows(records, opts)               # 更新（-u）: 既存の API の番号とテストデータを引き継ぐ
        opts['_oa_rows'] = [dict(r) for r in records]   # 基準ファイル（.base.json）に保存する値（API_HEADERS の反映前）
        add_oa_api_headers(records, oa_conf)
        if not settings_path:
            cand = os.path.splitext(input_path)[0] + '_settings.csv'
            if os.path.isfile(cand):
                settings_path = cand
            else:
                info('設定 CSV（%s）が無いため既定値を使用します（接続先などは OpenAPI 定義の値）' % os.path.basename(cand))
    elif ext == '.xls':
        raise ToolError('.xls 形式（Excel 97-2003）には対応していません。Excel で .xlsx 形式で保存し直してください: ' + input_path)
    else:
        raise ToolError('対応していないファイル形式です（.xlsx / .csv / .json）: ' + input_path)
    if settings_path:
        if not os.path.isfile(settings_path):
            raise ToolError('設定ファイルが見つかりません: ' + settings_path)
        if settings_path.lower().endswith(('.xlsx', '.xlsm')):
            sheets = xlsx_read(settings_path)
            sname, srows = _find_sheet(sheets, [TPL['settings_sheet'], 'settings'])
            if srows is None:
                raise ToolError('「%s」シートが見つかりません: %s' % (TPL['settings_sheet'], settings_path))
            raw = load_settings_table(srows, os.path.basename(settings_path))
        else:
            text, enc = read_text(settings_path)
            raw = load_settings_table([(i + 1, r) for i, r in enumerate(parse_csv(text))], os.path.basename(settings_path))
        info('設定: %s' % settings_path)
    raw = raw or {}
    for key, value in oa_defaults:
        if not raw.get(key):
            raw[key] = value
            info('設定 %s: %s（設定ファイルで未指定のため OpenAPI 定義から設定）' % (key, value))
    if oa_conf is not None:
        apply_oa_config(raw, oa_conf, opts.get('out') or 'output')
    # 認可（security）の警告: OpenAPI 用設定（-c）・設定ファイル（-s）のどちらにも Authorization の指定が無い場合だけ表示する
    if ext == '.json' and opts.get('_oa_needs_auth') and not (
            opts.get('_oa_auth') or raw.get('access_token_dir') or re.search(r'(?im)^\s*authorization\s*:', raw.get('common_headers', ''))):
        warn('認可（security）が定義された API があります。Authorization ヘッダが必要な場合は、'
             '設定 common_headers（例: Authorization: Bearer ${AUTH_TOKEN}）か、OpenAPI 用設定ファイル（-c）の TOKEN_DIR で指定してください')
    return raw, records


def parse_bool(v):
    k = normalize_key(v)
    if k in ('true', 'yes', 'y', '1', 'on', 'はい', '有効', 'する'):
        return True
    if k in ('false', 'no', 'n', '0', 'off', 'いいえ', '無効', 'しない'):
        return False
    return None


def parse_onoff(v):
    k = normalize_key(v)
    if k in ('on', 'y', 'yes', 'true', '1', '有効', '生成'):
        return 'ON'
    if k in ('off', 'n', 'no', 'false', '0', 'なし', '生成しない', '不要'):
        return 'OFF'
    if k in ('disable', 'disabled', '無効', '無効化'):
        return 'DISABLE'
    return None


def resolve_settings(raw):
    known = set(s['key'] for s in SPEC['settings'])
    for k in sorted(raw):
        if k not in known:
            warn('設定: 不明な設定キー「%s」は無視します' % k)
    cfg = {}
    for s in SPEC['settings']:
        k, t = s['key'], s['type']
        v = raw.get(k, '')
        if v == '':
            v = s['default']
        if t == 'int':
            if v != '' and not re.match(r'^-?[0-9]+$', v):
                warn('設定 %s: 「%s」は数値ではないため既定値「%s」を使用します' % (k, v, s['default']))
                v = s['default']
        elif t == 'bool':
            b = parse_bool(v)
            if b is None:
                warn('設定 %s: 「%s」は true / false ではないため既定値「%s」を使用します' % (k, v, s['default']))
                b = parse_bool(s['default'])
            v = b
        elif t == 'onoff':
            o = parse_onoff(v)
            if o is None:
                warn('設定 %s: 「%s」は ON / OFF / DISABLE ではないため既定値「%s」を使用します' % (k, v, s['default']))
                o = parse_onoff(s['default'])
            v = o
        elif t == 'enum':
            lv = v.lower()
            if lv not in s['choices']:
                warn('設定 %s: 「%s」は %s のいずれでもないため既定値「%s」を使用します' % (k, v, ' / '.join(s['choices']), s['default']))
                lv = s['default']
            v = lv
        cfg[k] = v
    cfg['_common_headers'] = parse_headers(cfg['common_headers'], '設定 common_headers')
    if cfg['access_token_dir'] and 'authorization' not in [k.lower() for k, v in cfg['_common_headers']]:
        cfg['_common_headers'].append(('Authorization', 'Bearer ${ACCESS_TOKEN}'))
    cfg['_user_vars'] = parse_pairs_strict(cfg['user_variables'], '設定 user_variables')
    cfg['_cookies'] = parse_pairs_strict(cfg['cookies'], '設定 cookies')
    if cfg['duration'] and int(cfg['duration']) > 0 and cfg['loops'] != '-1':
        warn('設定: 実行時間（duration）を指定する場合は、ループ回数（loops）を -1 にしてください（現在: %s）' % cfg['loops'])
    return cfg


# ----------------------------------------------------------------------------- API normalisation
_URL_RE = re.compile(r'^([A-Za-z][A-Za-z0-9+.\-]*)://([^/?#]*)([^?#]*)(?:\?([^#]*))?(?:#.*)?$')
_VAR_RE = re.compile(r'^[A-Za-z_][A-Za-z0-9_]*$')
_UNRESERVED = set('ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-._~')
_PATH_SAFE = _UNRESERVED | set("/:@!$&'()*+,;=")
QUERY_IN_ARGS = ('GET', 'DELETE', 'OPTIONS')
NO_BODY = ('GET', 'DELETE', 'OPTIONS', 'HEAD')


def parse_url(url):
    m = _URL_RE.match(url)
    if not m:
        return None
    auth = m.group(2)
    host, port = auth, ''
    pm = re.match(r'^(.*):([0-9]+|\$\{[^}]*\})$', auth)
    if pm:
        host, port = pm.group(1), pm.group(2)
    if '@' in host:
        host = host.split('@')[-1]
    return {'scheme': m.group(1).lower(), 'host': host, 'port': port,
            'path': m.group(3) or '/', 'query': m.group(4) or ''}


def _close_brace(s, i):
    depth = 0
    j = i + 1
    while j < len(s):
        if s[j] == '{':
            depth += 1
        elif s[j] == '}':
            depth -= 1
            if depth == 0:
                return j
        j += 1
    return -1


def pct_encode(s, safe):
    out = []
    i, n = 0, len(s)
    while i < n:
        ch = s[i]
        if ch == '$' and i + 1 < n and s[i + 1] == '{':
            j = _close_brace(s, i)
            if j > 0:
                out.append(s[i:j + 1])
                i = j + 1
                continue
        if ch == '%' and re.match(r'^%[0-9A-Fa-f]{2}', s[i:i + 3]):
            out.append(s[i:i + 3])
            i += 3
            continue
        if ch in safe:
            out.append(ch)
        else:
            out.append(''.join('%%%02X' % b for b in bytearray(ch.encode('utf-8'))))
        i += 1
    return ''.join(out)


def needs_encode(v):
    return re.search(r'%[0-9A-Fa-f]{2}', v) is None


def split_pairs(text):
    out = []
    for line in norm_nl(text).split('\n'):
        for part in line.split('&'):
            p = part.strip()
            if not p:
                continue
            if '=' in p:
                k, v = p.split('=', 1)
                out.append((k.strip(), v.strip()))
            else:
                out.append((p, ''))
    return out


def parse_headers(text, where):
    out = []
    for line in norm_nl(text).split('\n'):
        s = line.strip()
        if not s:
            continue
        if ':' not in s or not s.split(':', 1)[0].strip():
            warn('%s: ヘッダは「名前: 値」の形式で記入してください（無視します）: %s' % (where, s))
            continue
        k, v = s.split(':', 1)
        out.append((k.strip(), v.strip()))
    return out


def parse_pairs_strict(text, where):
    out = []
    for line in norm_nl(text).split('\n'):
        s = line.strip()
        if not s:
            continue
        if '=' not in s:
            warn('%s: 「名前=値」の形式で記入してください（無視します）: %s' % (where, s))
            continue
        k, v = s.split('=', 1)
        k = k.strip()
        if not k:
            warn('%s: 名前が空です（無視します）: %s' % (where, s))
            continue
        out.append((k, v.strip()))
    return out


def parse_assign(text, where, what):
    out = []
    for line in norm_nl(text).split('\n'):
        s = line.strip()
        if not s:
            continue
        k, sep, v = s.partition('=')
        k, v = k.strip(), v.strip()
        if not sep or not _VAR_RE.match(k) or not v:
            warn('%s: %sは「変数名=式」の形式で記入してください（無視します）: %s' % (where, what, s))
            continue
        out.append((k, v))
    return out


def _is_form(body):
    segs = [x.strip() for x in re.split(r'[\n&]', body) if x.strip()]
    return bool(segs) and all(re.match(r'^[^=\s]+=', x) for x in segs)


def normalize_api(rec, index, cfg):
    where = '%d行目' % rec['_row']
    no = rec['no']
    if re.match(r'^[0-9]{1,6}$', no):
        aid = 'API%02d' % int(no)
    elif no:
        aid = 'API_' + re.sub(r'[^A-Za-z0-9]', '_', no)
    else:
        aid = 'API%02d' % (index + 1)
    where = '%s（%s）' % (where, aid)
    url = rec['url']
    if not url:
        warn('%s: URL が空のためスキップします' % where)
        return None
    enabled = True
    if rec['enabled']:
        e = parse_bool(rec['enabled'])
        if e is None:
            warn('%s: 有効「%s」は Y / N で記入してください（Y として扱います）' % (where, rec['enabled']))
        else:
            enabled = e
    method = rec['method'].upper()
    if not method:
        warn('%s: メソッドが空のため GET として扱います' % where)
        method = 'GET'
    if method not in ('GET', 'POST', 'PUT', 'PATCH', 'DELETE', 'HEAD', 'OPTIONS'):
        warn('%s: メソッド「%s」はそのまま設定します（GET / POST 以外は動作を確認してください）' % (where, method))
    full = url
    if url.startswith('/'):
        if not cfg['base_url']:
            warn('%s: URL が「/」で始まっていますが設定 base_url が空のためスキップします' % where)
            return None
        full = cfg['base_url'].rstrip('/') + url
    p = parse_url(full)
    if p is None:
        p = parse_url('https://' + full)
        if p is None:
            warn('%s: URL の形式が正しくないためスキップします: %s' % (where, url))
            return None
        warn('%s: URL にスキーム（https://）が無いため https として扱います' % where)
        full = 'https://' + full
    pairs = split_pairs(p['query']) + split_pairs(rec['query'])
    path = pct_encode(p['path'], _PATH_SAFE)
    args = []
    if method in QUERY_IN_ARGS:
        args = [(k, v, needs_encode(v)) for k, v in pairs]
    elif pairs:
        path += '?' + '&'.join(pct_encode(k, _UNRESERVED) + '=' + pct_encode(v, _UNRESERVED) for k, v in pairs)
    headers = parse_headers(rec['headers'], where)
    body = rec['body']
    bt = rec['body_type'].lower() or 'auto'
    if bt not in ('auto', 'json', 'form', 'raw', 'xml', 'multipart'):
        warn('%s: ボディ形式「%s」は auto / json / form / raw のいずれかで記入してください（auto として扱います）' % (where, rec['body_type']))
        bt = 'auto'
    kind = ''
    if body and method in NO_BODY:
        warn('%s: %s ではリクエストボディを使用できないため無視します' % (where, method))
        body = ''
    if body:
        if bt == 'auto':
            s = body.lstrip()
            if s.startswith('{') or s.startswith('['):
                kind = 'json'
            elif s.startswith('<'):
                kind = 'xml'
            elif _is_form(body):
                kind = 'form'
            else:
                kind = 'raw'
        else:
            kind = bt
    if kind in ('form', 'multipart'):
        args = [(k, v, needs_encode(v)) for k, v in split_pairs(body)]
    if kind == 'json':
        body = re.sub(r'("(?:[^"\\]|\\.)*")|\u3000', lambda m: m.group(1) or ' ', body)
    ctype = {'json': 'application/json', 'xml': 'application/xml'}.get(kind, '')
    if ctype:
        names = [k.lower() for k, v in headers + cfg['_common_headers']]
        if 'content-type' not in names:
            headers.append(('Content-Type', ctype))

    def num(v, label):
        if v == '':
            return -1
        if re.match(r'^[0-9]+$', v):
            return int(v)
        warn('%s: %s「%s」は数値で記入してください（無視します）' % (where, label, v))
        return -1
    wait = num(rec['wait_ms'], '待機時間')
    if wait < 0:
        wait = int(cfg['default_wait_ms']) if re.match(r'^[0-9]+$', cfg['default_wait_ms']) else 0
    max_ms = num(rec['max_time_ms'], '応答時間上限')
    statuses = []
    for c in [x for x in re.split(r'[,\s/]+', rec['expected_status']) if x]:
        if re.match(r'^[1-5][0-9][0-9]$', c):
            if c not in statuses:
                statuses.append(c)
        else:
            warn('%s: 期待ステータス「%s」は 3 桁の数字で記入してください（無視します）' % (where, c))
    texts = [x for x in norm_nl(rec['expected_text']).split('\n') if x.strip()]
    name = re.sub(r'\s+', ' ', rec['name']).strip() or (method + ' ' + p['path'])
    return {
        'row': rec['_row'], 'id': aid, 'label': aid + '_' + name, 'enabled': enabled, 'method': method,
        'url': full, 'scheme': p['scheme'], 'host': p['host'], 'port': p['port'], 'path': path,
        'args': args, 'body': body if kind in ('json', 'xml', 'raw') else '', 'kind': kind, 'headers': headers,
        'wait': wait, 'max_ms': max_ms, 'statuses': statuses, 'texts': texts,
        'regex': parse_assign(rec['regex_extract'], where, '正規表現抽出'),
        'json': parse_assign(rec['json_extract'], where, 'JSON抽出'),
        'pre': rec['pre_script'], 'post': rec['post_script'], 'comment': rec['comment'],
    }


# ----------------------------------------------------------------------------- JMX builder
def el(tag, attrs=None, children=None, text=None):
    return [tag, attrs or [], children or [], text]


def sp(name, value):
    return el('stringProp', [('name', name)], text='' if value is None else value)


def bp(name, value):
    return el('boolProp', [('name', name)], text='true' if value else 'false')


def ip(name, value):
    return el('intProp', [('name', name)], text=str(value))


def lp(name, value):
    return el('longProp', [('name', name)], text=str(value))


def cp(name, items):
    return el('collectionProp', [('name', name)], items)


def ep(name, etype, children, extra=None):
    return el('elementProp', [('name', name), ('elementType', etype)] + (extra or []), children)


def obj(tag, gui, cls, name, enabled, props, comment=''):
    kids = [sp('TestPlan.comments', comment)] if comment else []
    return el(tag, [('guiclass', gui), ('testclass', cls), ('testname', name), ('enabled', 'true' if enabled else 'false')],
              kids + props)


def node(elem, children=None):
    return [elem, children or []]


def pesc(v):
    return v.replace('\\', '\\\\').replace(',', '\\,')


def pfun(name, default):
    return '${__P(' + name + ',' + pesc(default) + ')}'


def _ser_elem(e, ind, out):
    tag, attrs, children, text = e
    pad = '  ' * ind
    a = ''.join(' %s="%s"' % (k, xesc(v)) for k, v in attrs)
    if children:
        out.append('%s<%s%s>\n' % (pad, tag, a))
        for c in children:
            _ser_elem(c, ind + 1, out)
        out.append('%s</%s>\n' % (pad, tag))
    elif text is None:
        out.append('%s<%s%s/>\n' % (pad, tag, a))
    else:
        out.append('%s<%s%s>%s</%s>\n' % (pad, tag, a, xesc(text), tag))


def _ser_nodes(nodes, ind, out):
    for elem, kids in nodes:
        _ser_elem(elem, ind, out)
        if kids:
            out.append('  ' * ind + '<hashTree>\n')
            _ser_nodes(kids, ind + 1, out)
            out.append('  ' * ind + '</hashTree>\n')
        else:
            out.append('  ' * ind + '<hashTree/>\n')


def jmx_text(plan):
    out = ['<?xml version="1.0" encoding="UTF-8"?>\n',
           '<jmeterTestPlan version="1.2" properties="5.0" jmeter="%s">\n' % TOOL['jmeter'], '  <hashTree>\n']
    _ser_nodes([plan], 2, out)
    out.append('  </hashTree>\n</jmeterTestPlan>\n')
    return ''.join(out)


def jsr223(tag, name, enabled, params, script, comment):
    return obj(tag, 'TestBeanGUI', tag, name, enabled,
               [sp('scriptLanguage', 'groovy'), sp('parameters', params), sp('filename', ''),
                sp('cacheKey', 'true'), sp('script', script)], comment)


def header_manager(name, headers, comment):
    items = [ep('', 'Header', [sp('Header.name', k), sp('Header.value', v)]) for k, v in headers]
    return obj('HeaderManager', 'HeaderPanel', 'HeaderManager', name, True, [cp('HeaderManager.headers', items)], comment)


def thread_group(tag, gui, name, enabled, threads, rampup, loops, scheduler, duration, on_error, comment):
    loop = ep('ThreadGroup.main_controller', 'LoopController',
              [bp('LoopController.continue_forever', False), sp('LoopController.loops', loops)],
              [('guiclass', 'LoopControlPanel'), ('testclass', 'LoopController'), ('testname', 'Loop Controller'), ('enabled', 'true')])
    return obj(tag, gui, tag, name, enabled,
               [sp('ThreadGroup.on_sample_error', on_error), loop, sp('ThreadGroup.num_threads', threads),
                sp('ThreadGroup.ramp_time', rampup), bp('ThreadGroup.scheduler', scheduler),
                sp('ThreadGroup.duration', duration), sp('ThreadGroup.delay', ''),
                bp('ThreadGroup.same_user_on_next_iteration', True)], comment)


def save_config(detail):
    fields = [('time', True), ('latency', True), ('timestamp', True), ('success', True), ('label', True),
              ('code', True), ('message', True), ('threadName', True), ('dataType', True), ('encoding', detail),
              ('assertions', True), ('subresults', detail), ('responseData', detail), ('samplerData', detail),
              ('xml', detail), ('fieldNames', True), ('responseHeaders', detail), ('requestHeaders', detail),
              ('responseDataOnError', False), ('saveAssertionResultsFailureMessage', True),
              ('assertionsResultsToSave', '0'), ('bytes', True), ('sentBytes', True), ('url', True),
              ('threadCounts', True), ('idleTime', True), ('connectTime', True)]
    kids = [el(k, text=(v if isinstance(v, str) else ('true' if v else 'false'))) for k, v in fields]
    return el('objProp', [], [el('name', text='saveConfig'), el('value', [('class', 'SampleSaveConfiguration')], kids)])


def collector(gui, name, enabled, filename, comment, errors_only=False, detail=False):
    return obj('ResultCollector', gui, 'ResultCollector', name, enabled,
               [bp('ResultCollector.error_logging', errors_only), save_config(detail), sp('filename', filename)], comment)


def assign_hosts(cfg, apis):
    hosts = []

    def add(scheme, host, port):
        for h in hosts:
            if h['scheme'] == scheme and h['host'] == host and h['port'] == port:
                return h
        sfx = '' if not hosts else '_%d' % (len(hosts) + 1)
        h = {'scheme': scheme, 'host': host, 'port': port,
             'pv': 'PROTOCOL' + sfx, 'hv': 'HOST' + sfx, 'portv': 'PORT' + sfx}
        hosts.append(h)
        return h
    if cfg['base_url']:
        p = parse_url(cfg['base_url'])
        if p and '${' not in p['scheme'] + p['host'] + p['port']:
            add(p['scheme'], p['host'], p['port'])
    for a in apis:
        if '${' in a['scheme'] + a['host'] + a['port']:
            a['hostref'] = None
        else:
            a['hostref'] = add(a['scheme'], a['host'], a['port'])
    return hosts


def api_node(api, cfg):
    N, CM = SPEC['names'], SPEC['comments']
    aid = api['id']
    h = api['hostref']
    props = []
    if api['body']:
        props.append(bp('HTTPSampler.postBodyRaw', True))
        props.append(ep('HTTPsampler.Arguments', 'Arguments', [cp('Arguments.arguments', [
            ep('', 'HTTPArgument', [bp('HTTPArgument.always_encode', False), sp('Argument.value', api['body']),
                                    sp('Argument.metadata', '=')])])]))
    else:
        items = [ep(k, 'HTTPArgument', [bp('HTTPArgument.always_encode', enc), sp('Argument.value', v),
                                        sp('Argument.metadata', '='), bp('HTTPArgument.use_equals', True),
                                        sp('Argument.name', k)]) for k, v, enc in api['args']]
        props.append(ep('HTTPsampler.Arguments', 'Arguments', [cp('Arguments.arguments', items)],
                        [('guiclass', 'HTTPArgumentsPanel'), ('testclass', 'Arguments'),
                         ('testname', 'User Defined Variables'), ('enabled', 'true')]))
    props += [sp('HTTPSampler.domain', '${%s}' % h['hv'] if h else api['host']),
              sp('HTTPSampler.port', '${%s}' % h['portv'] if h else api['port']),
              sp('HTTPSampler.protocol', '${%s}' % h['pv'] if h else api['scheme']),
              sp('HTTPSampler.contentEncoding', cfg['content_encoding']),
              sp('HTTPSampler.path', api['path']), sp('HTTPSampler.method', api['method']),
              bp('HTTPSampler.follow_redirects', cfg['follow_redirects']), bp('HTTPSampler.auto_redirects', False),
              bp('HTTPSampler.use_keepalive', cfg['keep_alive']), bp('HTTPSampler.DO_MULTIPART_POST', api['kind'] == 'multipart'),
              sp('HTTPSampler.embedded_url_re', ''), sp('HTTPSampler.connect_timeout', ''),
              sp('HTTPSampler.response_timeout', '')]
    comment = fill(CM['sampler'], {'ID': aid, 'METHOD': api['method'], 'URL': api['url']})
    if api['comment']:
        comment += '\n' + api['comment']
    kids = []
    ids = {'ID': aid}
    if api['headers']:
        kids.append(node(header_manager(fill(N['api_headers'], ids), api['headers'], CM['api_headers'])))
    if api['wait'] > 0:
        ms = str(api['wait'])
        kids.append(node(obj('ConstantTimer', 'ConstantTimerGui', 'ConstantTimer', fill(N['api_timer'], {'ID': aid, 'MS': ms}),
                             True, [sp('ConstantTimer.delay', ms)], fill(CM['api_timer'], {'MS': ms}))))
    if api['pre']:
        kids.append(node(jsr223('JSR223PreProcessor', fill(N['api_pre'], ids), True, '', api['pre'], CM['api_pre'])))
    for var, expr in api['regex']:
        template = '$1$' if re.search(r'(?<!\\)\((?!\?)', expr) else '$0$'
        kids.append(node(obj('RegexExtractor', 'RegexExtractorGui', 'RegexExtractor', fill(N['api_regex'], {'ID': aid, 'VAR': var}),
                             True, [sp('RegexExtractor.useHeaders', 'false'), sp('RegexExtractor.refname', var),
                                    sp('RegexExtractor.regex', expr), sp('RegexExtractor.template', template),
                                    sp('RegexExtractor.default', 'NOT_FOUND'), bp('RegexExtractor.default_empty_value', False),
                                    sp('RegexExtractor.match_number', '1')],
                             fill(CM['api_regex'], {'VAR': var}))))
    for var, path in api['json']:
        kids.append(node(obj('JSONPostProcessor', 'JSONPostProcessorGui', 'JSONPostProcessor',
                             fill(N['api_json'], {'ID': aid, 'VAR': var}), True,
                             [sp('JSONPostProcessor.referenceNames', var), sp('JSONPostProcessor.jsonPathExprs', path),
                              sp('JSONPostProcessor.match_numbers', '1'), sp('JSONPostProcessor.defaultValues', 'NOT_FOUND')],
                             fill(CM['api_json'], {'VAR': var, 'PATH': path}))))
    if api['post']:
        kids.append(node(jsr223('JSR223PostProcessor', fill(N['api_post'], ids), True, '', api['post'], CM['api_post'])))
    if api['statuses']:
        codes = api['statuses']
        ignore = any(int(c) >= 400 for c in codes)
        kids.append(node(obj('ResponseAssertion', 'AssertionGui', 'ResponseAssertion',
                             fill(N['api_status'], {'ID': aid, 'CODES': ','.join(codes)}), True,
                             [cp('Asserion.test_strings', [sp(str(java_hash(c)), c) for c in codes]),
                              sp('Assertion.custom_message', ''), sp('Assertion.test_field', 'Assertion.response_code'),
                              bp('Assertion.assume_success', ignore), ip('Assertion.test_type', 40 if len(codes) > 1 else 8)],
                             fill(CM['api_status'], {'CODES': ' / '.join(codes)}) + (CM['api_status_ignore'] if ignore else ''))))
    if api['texts']:
        kids.append(node(obj('ResponseAssertion', 'AssertionGui', 'ResponseAssertion', fill(N['api_text'], ids), True,
                             [cp('Asserion.test_strings', [sp(str(java_hash(t)), t) for t in api['texts']]),
                              sp('Assertion.custom_message', ''), sp('Assertion.test_field', 'Assertion.response_data'),
                              bp('Assertion.assume_success', False), ip('Assertion.test_type', 16)], CM['api_text'])))
    if api['max_ms'] > 0:
        ms = str(api['max_ms'])
        kids.append(node(obj('DurationAssertion', 'DurationAssertionGui', 'DurationAssertion',
                             fill(N['api_duration'], {'ID': aid, 'MS': ms}), True,
                             [sp('DurationAssertion.duration', ms)], fill(CM['api_duration'], {'MS': ms}))))
    return node(obj('HTTPSamplerProxy', 'HttpTestSampleGui', 'HTTPSamplerProxy', api['label'], api['enabled'], props, comment), kids)


def build_jmx(cfg, apis, ctx):
    N, CM, SC = SPEC['names'], SPEC['comments'], SPEC['scripts']
    hosts = assign_hosts(cfg, apis)
    on = lambda key: cfg[key] != 'OFF'
    enabled = lambda key: cfg[key] == 'ON'
    rd = pfun('RESULT_DIR', cfg['result_dir'])
    top = []
    # 共通変数
    args = []

    def arg(name, value, desc):
        return ep(name, 'Argument', [sp('Argument.name', name), sp('Argument.value', value),
                                     sp('Argument.desc', desc), sp('Argument.metadata', '=')])
    for h in hosts:
        args.append(arg(h['pv'], pfun(h['pv'], h['scheme']), fill(CM['udv_protocol'], {'VAR': h['pv']})))
        args.append(arg(h['hv'], pfun(h['hv'], h['host']), fill(CM['udv_host'], {'VAR': h['hv']})))
        args.append(arg(h['portv'], pfun(h['portv'], h['port']), fill(CM['udv_port'], {'VAR': h['portv']})))
    for name, value in cfg['_user_vars']:
        v = value if re.search(r'[${}()]', value) else pfun(name, value)
        args.append(arg(name, v, fill(CM['udv_user'], {'VAR': name})))
    top.append(node(obj('Arguments', 'ArgumentsPanel', 'Arguments', N['udv'], True, [cp('Arguments.arguments', args)], CM['udv'])))
    if on('http_defaults'):
        top.append(node(obj('ConfigTestElement', 'HttpDefaultsGui', 'ConfigTestElement', N['http_defaults'], enabled('http_defaults'), [
            ep('HTTPsampler.Arguments', 'Arguments', [cp('Arguments.arguments', [])],
               [('guiclass', 'HTTPArgumentsPanel'), ('testclass', 'Arguments'), ('testname', 'User Defined Variables'), ('enabled', 'true')]),
            sp('HTTPSampler.domain', ''), sp('HTTPSampler.port', ''), sp('HTTPSampler.protocol', ''),
            sp('HTTPSampler.contentEncoding', cfg['content_encoding']), sp('HTTPSampler.path', ''),
            sp('HTTPSampler.concurrentPool', '6'),
            sp('HTTPSampler.connect_timeout', pfun('CONNECT_TIMEOUT', cfg['connect_timeout'])),
            sp('HTTPSampler.response_timeout', pfun('RESPONSE_TIMEOUT', cfg['response_timeout']))], CM['http_defaults'])))
    if cfg['_common_headers']:
        top.append(node(header_manager(N['common_headers'], cfg['_common_headers'], CM['common_headers'])))
    if on('cookie_manager'):
        domain = cfg['cookie_domain'] or ('${%s}' % hosts[0]['hv'] if hosts else '')
        cookies = [el('elementProp', [('name', k), ('elementType', 'Cookie'), ('testname', k)], [
            sp('Cookie.value', v), sp('Cookie.domain', domain), sp('Cookie.path', '/'), bp('Cookie.secure', False),
            lp('Cookie.expires', 0), bp('Cookie.path_specified', True), bp('Cookie.domain_specified', True)])
            for k, v in cfg['_cookies']]
        top.append(node(obj('CookieManager', 'CookiePanel', 'CookieManager', N['cookie'], enabled('cookie_manager'), [
            cp('CookieManager.cookies', cookies), bp('CookieManager.clearEachIteration', cfg['cookie_clear_each_iteration']),
            bp('CookieManager.controlledByThreadGroup', False)],
            fill(CM['cookie'], {'CLEAR': 'true' if cfg['cookie_clear_each_iteration'] else 'false'}))))
    scheduler = bool(cfg['duration']) and int(cfg['duration']) > 0
    # setUp
    setup = []
    if on('setup_teardown'):
        params = 'THREADS=%s RAMP_UP=%s LOOPS=%s' % (pfun('THREADS', cfg['threads']), pfun('RAMP_UP', cfg['ramp_up']),
                                                    pfun('LOOPS', cfg['loops']))
        if scheduler:
            params += ' DURATION=' + pfun('DURATION', cfg['duration'])
        if hosts:
            params += ' HOST=${%s}' % hosts[0]['hv']
        params += ' RESULT_DIR=' + rd
        setup.append(node(jsr223('JSR223Sampler', N['setup_start'], enabled('setup_teardown'), params, SC['setup_start'], CM['setup_start'])))
    if on('excel_data'):
        params = pfun('EXCEL_FILE', cfg['excel_file']) + '|' + pfun('EXCEL_SHEET', cfg['excel_sheet'])
        setup.append(node(jsr223('JSR223Sampler', N['excel_loader'], enabled('excel_data'), params, SC['excel_loader'], CM['excel_loader'])))
    if setup:
        tg_on = any(k[0][1][3][1] == 'true' for k in setup)
        top.append(node(thread_group('SetupThreadGroup', 'SetupThreadGroupGui', N['setup_tg'], tg_on, '1', '1', '1', False, '',
                                     'continue', CM['setup_tg']), setup))
    # メイン処理
    main = []
    if on('counter'):
        main.append(node(obj('CounterConfig', 'CounterConfigGui', 'CounterConfig', fill(N['counter'], {'NAME': cfg['counter_name']}),
                             enabled('counter'),
                             [sp('CounterConfig.start', cfg['counter_start']), sp('CounterConfig.end', ''),
                              sp('CounterConfig.incr', cfg['counter_increment']), sp('CounterConfig.name', cfg['counter_name']),
                              sp('CounterConfig.format', cfg['counter_format']), bp('CounterConfig.per_user', cfg['counter_per_thread'])],
                             fill(CM['counter'], {'NAME': cfg['counter_name'],
                                                  'SCOPE': CM['counter_scope_thread'] if cfg['counter_per_thread'] else CM['counter_scope_global']}))))
    if on('random_variable'):
        main.append(node(obj('RandomVariableConfig', 'TestBeanGUI', 'RandomVariableConfig', fill(N['random'], {'NAME': cfg['random_name']}),
                             enabled('random_variable'),
                             [sp('maximumValue', cfg['random_max']), sp('minimumValue', cfg['random_min']), sp('outputFormat', ''),
                              bp('perThread', False), sp('randomSeed', ''), sp('variableName', cfg['random_name'])],
                             fill(CM['random'], {'NAME': cfg['random_name'], 'MIN': cfg['random_min'], 'MAX': cfg['random_max']}))))
    if on('csv_data'):
        main.append(node(obj('CSVDataSet', 'TestBeanGUI', 'CSVDataSet', N['csv'], enabled('csv_data'),
                             [sp('delimiter', cfg['csv_delimiter']), sp('fileEncoding', cfg['csv_encoding']),
                              sp('filename', pfun('CSV_FILE', cfg['csv_file'])), bp('ignoreFirstLine', ctx['csv_skip_header']),
                              bp('quotedData', True), bp('recycle', cfg['csv_recycle']), sp('shareMode', 'shareMode.' + cfg['csv_share_mode']),
                              bp('stopThread', cfg['csv_stop_thread']), sp('variableNames', ctx['csv_vars'])],
                             fill(CM['csv'], {'VARS': ctx['csv_vars']}))))
    if on('excel_data'):
        main.append(node(jsr223('JSR223PreProcessor', N['excel_row'], enabled('excel_data'), '', SC['excel_row'], CM['excel_row'])))
    if cfg['access_token_dir']:
        params = pfun('ACCESS_TOKEN_DIR', cfg['access_token_dir']) + '|' + pfun('ACCESS_TOKEN_FILE', cfg['access_token_file'])
        main.append(node(jsr223('JSR223PreProcessor', N['access_token'], True, params, SC['access_token'], CM['access_token'])))
    if on('common_preprocessor'):
        main.append(node(jsr223('JSR223PreProcessor', N['common_pre'], enabled('common_preprocessor'), '', SC['common_pre'], CM['common_pre'])))
    if on('common_postprocessor'):
        main.append(node(jsr223('JSR223PostProcessor', N['common_post'], enabled('common_postprocessor'), '', SC['common_post'], CM['common_post'])))
    if on('throughput_timer'):
        main.append(node(obj('ConstantThroughputTimer', 'TestBeanGUI', 'ConstantThroughputTimer', N['throughput'], enabled('throughput_timer'),
                             [ip('calcMode', 4), sp('throughput', pfun('THROUGHPUT', cfg['throughput_per_min']))],
                             fill(CM['throughput'], {'VALUE': cfg['throughput_per_min']}))))
    for api in apis:
        main.append(api_node(api, cfg))
    if on('debug_sampler'):
        main.append(node(obj('DebugSampler', 'TestBeanGUI', 'DebugSampler', N['debug'], enabled('debug_sampler'),
                             [bp('displayJMeterProperties', False), bp('displayJMeterVariables', True),
                              bp('displaySystemProperties', False)], CM['debug'])))
    top.append(node(thread_group('ThreadGroup', 'ThreadGroupGui', N['main_tg'], True, pfun('THREADS', cfg['threads']),
                                 pfun('RAMP_UP', cfg['ramp_up']), pfun('LOOPS', cfg['loops']), scheduler,
                                 pfun('DURATION', cfg['duration']) if scheduler else '', cfg['on_sample_error'],
                                 fill(CM['main_tg'], {'DURATION': CM['main_tg_duration'] if scheduler else ''})), main))
    # tearDown
    if on('setup_teardown'):
        top.append(node(thread_group('PostThreadGroup', 'PostThreadGroupGui', N['teardown_tg'], enabled('setup_teardown'), '1', '1', '1',
                                     False, '', 'continue', CM['teardown_tg']),
                        [node(jsr223('JSR223Sampler', N['teardown_end'], enabled('setup_teardown'), '', SC['teardown_end'], CM['teardown_end']))]))
    # リスナー
    if on('view_results_tree'):
        top.append(node(collector('ViewResultsFullVisualizer', N['view_tree'], enabled('view_results_tree'), '', CM['view_tree'])))
    if on('summary_report'):
        top.append(node(collector('SummaryReport', N['summary'], enabled('summary_report'), '', CM['summary'])))
    if on('aggregate_report'):
        top.append(node(collector('StatVisualizer', N['aggregate'], enabled('aggregate_report'), '', CM['aggregate'])))
    if on('jtl_writer'):
        top.append(node(collector('SimpleDataWriter', N['jtl'], enabled('jtl_writer'), rd + '/result_all.jtl',
                                  fill(CM['jtl'], {'DIR': cfg['result_dir']}))))
    if on('error_writer'):
        top.append(node(collector('SimpleDataWriter', N['error_xml'], enabled('error_writer'), rd + '/error_detail.xml',
                                  CM['error_xml'], errors_only=True, detail=True)))
    if on('result_log'):
        params = rd + '|' + pfun('LOG_RESPONSE', cfg['log_response']) + '|' + pfun('LOG_MAX_CHARS', cfg['log_max_chars'])
        top.append(node(jsr223('JSR223Listener', N['result_log'], enabled('result_log'), params, SC['result_listener'], CM['result_log'])))
    plan = obj('TestPlan', 'TestPlanGui', 'TestPlan', ctx['plan_name'], True, [
        bp('TestPlan.functional_mode', False), bp('TestPlan.tearDown_on_shutdown', True),
        bp('TestPlan.serialize_threadgroups', False),
        ep('TestPlan.user_defined_variables', 'Arguments', [cp('Arguments.arguments', [])],
           [('guiclass', 'ArgumentsPanel'), ('testclass', 'Arguments'), ('testname', 'User Defined Variables'), ('enabled', 'true')]),
        sp('TestPlan.user_define_classpath', '')],
        fill(CM['test_plan'], {'VERSION': VERSION, 'DATE': ctx['date'], 'INPUT': ctx['input']}))
    return jmx_text(node(plan, top)), hosts


# ----------------------------------------------------------------------------- commands
def _setting_default(key):
    for s in SPEC['settings']:
        if s['key'] == key:
            return s['default']
    return ''


def _resolve_data(path, out_dir):
    return path if os.path.isabs(path) else os.path.join(out_dir, path)


def prepare_data(cfg, out_dir, force):
    created = []
    for flag, key, builder in (('csv_data', 'csv_file', build_testdata_csv), ('excel_data', 'excel_file', build_testdata_xlsx)):
        if cfg[flag] == 'OFF':
            continue
        rel = cfg[key]
        full = _resolve_data(rel, out_dir)
        if rel.replace('\\', '/') == _setting_default(key) and (force or not os.path.exists(full)):
            builder(full)
            created.append(full)
        elif not os.path.exists(full):
            warn('データファイルが見つかりません: %s（JMX のフォルダからの相対パス。実行前に配置してください）' % full)
    return created


def csv_variables(cfg, out_dir):
    if cfg['csv_data'] == 'OFF':
        return '', True
    delim = cfg['csv_delimiter']
    delim = '\t' if delim == '\\t' else (delim[:1] or ',')
    full = _resolve_data(cfg['csv_file'], out_dir)
    first = None
    if os.path.isfile(full):
        text, enc = read_text(full)
        rows = parse_csv(text, delim)
        if rows:
            first = [c.strip() for c in rows[0]]
    names = [x.strip() for x in cfg['csv_variables'].split(',') if x.strip()]
    if names:
        skip = True if first is None else [x.lower() for x in first] == [x.lower() for x in names]
        return ','.join(names), skip
    if first and any(first):
        if not all(first):
            warn('CSV の 1 行目に空の列名があります: %s' % full)
        return ','.join(first), True
    warn('CSV の変数名を決められないため、サンプルの列名を使用します（設定 csv_variables で指定してください）')
    return ','.join(SPEC['sample_csv']['header']), True


def safe_file_name(name, default):
    base = os.path.basename(name.replace('\\', '/'))
    s = re.sub(r'[^A-Za-z0-9._-]', '_', base).strip('._')
    return s or default


def write_jmx(path, text):
    if os.path.exists(path):
        bak = '%s.%s.bak' % (path, timestamp())
        os.rename(path, bak)
        info('既存の JMX を退避しました: ' + bak)
    write_text(path, text)


def cmd_generate(opts):
    input_path = opts.get('input')
    if not input_path:
        raise ToolError('入力ファイルを -i で指定してください（例: -i input/api_definition.xlsx）')
    if not os.path.isfile(input_path):
        raise ToolError('入力ファイルが見つかりません: ' + input_path)
    upd = upd_prepare(opts) if opts.get('update') else None    # -u: 既存の JMX を差分だけ更新する
    out_dir = opts.get('out') or 'output'
    info('入力ファイル: ' + input_path)
    raw, records = load_input(input_path, opts.get('settings'), opts)
    cfg = resolve_settings(raw)
    if opts.get('jmx_name'):
        cfg['output_file'] = opts['jmx_name']
    apis = []
    for i, rec in enumerate(records):
        a = normalize_api(rec, i, cfg)
        if a:
            apis.append(a)
    if not apis:
        raise ToolError('有効な API 定義が 1 件もありません')
    seen = set()
    for a in apis:
        if a['id'] in seen:
            warn('API の番号（No）が重複しています: %s' % a['id'])
        seen.add(a['id'])
    if upd:
        return upd_run(upd, cfg, apis, opts)
    if not os.path.isdir(out_dir):
        os.makedirs(out_dir)
    for f in prepare_data(cfg, out_dir, opts.get('force')):
        info('サンプルデータを作成しました: ' + f)
    csv_vars, skip = csv_variables(cfg, out_dir)
    out_name = safe_file_name(cfg['output_file'], 'api_test.jmx')
    if not out_name.lower().endswith('.jmx'):
        out_name += '.jmx'
    if out_name != cfg['output_file']:
        info('出力ファイル名: %s（指定値「%s」から変換）' % (out_name, cfg['output_file']))
    ctx = {'input': os.path.basename(input_path), 'date': datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
           'csv_vars': csv_vars, 'csv_skip_header': skip, 'plan_name': cfg['test_plan_name']}
    outputs = []
    if opts.get('split'):
        stem = out_name[:-4]
        used = set()
        for a in apis:
            c = dict(ctx)
            c['plan_name'] = cfg['test_plan_name'] + ' - ' + a['label']
            text, hosts = build_jmx(cfg, [a], c)
            fname = '%s_%s' % (stem, a['id'])
            k = 2
            while fname.lower() in used:
                fname = '%s_%s_%d' % (stem, a['id'], k)
                k += 1
            used.add(fname.lower())
            path = os.path.join(out_dir, fname + '.jmx')
            write_jmx(path, text)
            outputs.append(path)
    else:
        text, hosts = build_jmx(cfg, apis, ctx)
        path = os.path.join(out_dir, out_name)
        write_jmx(path, text)
        outputs.append(path)
        write_base(path, text, opts)    # 更新（-u）の基準ファイル
    for p in outputs:
        write_linux_cmd(p)              # Linux サーバでの実行コマンド
    n_on = len([a for a in apis if a['enabled']])
    info('API: %d 件（有効 %d / 無効 %d）' % (len(apis), n_on, len(apis) - n_on))
    for p in outputs:
        info('JMX を生成しました: ' + p)
    print('')
    print('次の手順:')
    print('  GUI で確認 : jmeter -t "%s"' % outputs[0])
    print('  CLI で実行 : ./jmeter_api_tool.sh run -t "%s"' % outputs[0])
    if WARNINGS:
        print('')
        print('警告が %d 件あります（[WARN] を確認してください）' % len(WARNINGS))
    return 0


def _write_if_absent(path, force, builder):
    if os.path.exists(path) and not force:
        warn('既に存在するため作成しません（上書きは -f）: ' + path)
        return False
    builder(path)
    info('作成しました: ' + path)
    return True


def cmd_init(opts):
    out_dir = opts.get('out') or 'input'
    if not os.path.isdir(out_dir):
        os.makedirs(out_dir)
    force = opts.get('force')
    _write_if_absent(os.path.join(out_dir, 'api_definition.xlsx'), force, build_template_xlsx)
    _write_if_absent(os.path.join(out_dir, 'api_definition.csv'), force,
                     lambda p: write_text(p, to_csv(template_api_rows()), bom=True))
    _write_if_absent(os.path.join(out_dir, 'api_definition_settings.csv'), force,
                     lambda p: write_text(p, to_csv(template_settings_rows()), bom=True))
    _write_if_absent(os.path.join(out_dir, 'openapi_config.csv'), force,
                     lambda p: write_text(p, to_csv(template_settings_rows(SPEC['openapi_config'])), bom=True))
    print('')
    print('次の手順: テンプレートに API を記入し、generate コマンドで JMX を生成します。')
    print('  ./jmeter_api_tool.sh generate -i %s' % os.path.join(out_dir, 'api_definition.xlsx'))
    print('  ./jmeter_api_tool.sh generate -i openapi.json -c %s   # OpenAPI 定義から生成する場合' % os.path.join(out_dir, 'openapi_config.csv'))
    return 0


def cmd_manual(opts):
    out_dir = opts.get('out') or 'docs'
    if not os.path.isdir(out_dir):
        os.makedirs(out_dir)
    for p in build_manual(out_dir):
        info('作成しました: ' + p)
    return 0


def _disp_width(s):
    return sum(2 if (ord(ch) >= 0x1100 and not 0xFF61 <= ord(ch) <= 0xFF9F) else 1 for ch in s)


def _disp_pad(s, width):
    return s + ' ' * max(0, width - _disp_width(s))


def _disp_rjust(s, width):
    return ' ' * max(0, width - _disp_width(s)) + s


def cmd_summary(path):
    text, enc = read_text(path)
    stats = json.loads(text)
    keys = sorted(k for k in stats if k != 'Total') + (['Total'] if 'Total' in stats else [])
    heads = (('件数', 8), ('エラー', 7), ('エラー率', 9), ('平均ms', 9), ('90%ms', 9), ('最大ms', 9), ('件/秒', 10))
    print(_disp_pad('ラベル', 36) + ' '.join(_disp_rjust(h, w) for h, w in heads))
    for k in keys:
        s = stats[k]
        label = s.get('transaction', k)
        print(_disp_pad(label[:34], 36) + '%8d %7d %8.2f%% %9.1f %9.1f %9.1f %10.2f' % (
            s.get('sampleCount', 0), s.get('errorCount', 0), s.get('errorPct', 0.0), s.get('meanResTime', 0.0),
            s.get('pct1ResTime', 0.0), s.get('maxResTime', 0.0), s.get('throughput', 0.0)))
    total = stats.get('Total', {})
    return 0 if total.get('errorCount', 0) == 0 else 3


def cmd_jmeter_version(jar):
    try:
        with zipfile.ZipFile(jar) as z:
            mf = z.read('META-INF/MANIFEST.MF').decode('utf-8', 'replace')
        m = re.search(r'Implementation-Version:\s*(\S+)', mf)
        print(m.group(1) if m else '')
    except Exception:
        print('')
    return 0


# =============================================================================
#  追加機能（Windows 版 jmeter_api_tool.ps1 の同じ名前の処理と同一）
#   1. add-curl   : curl コマンドを解析して API一覧（Excel / CSV）に 1 行追加する
#   2. linux-cmd  : Linux サーバで JMeter の CLI を実行するコマンドと各パラメータの説明をテキストファイルに出力する
#   3. generate -u: 既存の JMX を、入力ファイル（Excel / CSV / OpenAPI）との差分箇所だけ更新し、更新箇所を表示する
# =============================================================================
# ----------------------------------------------------------------------------- 1. add-curl
CURL_SHORT = {'X': 'request', 'H': 'header', 'd': 'data', 'F': 'form', 'u': 'user', 'A': 'user-agent', 'e': 'referer',
              'b': 'cookie', 'T': 'upload-file', 'o': 'output', 'w': 'write-out', 'm': 'max-time', 'x': 'proxy', 'E': 'cert',
              'c': 'cookie-jar', 'C': 'continue-at', 'D': 'dump-header', 'K': 'config', 'P': 'ftp-port', 'Q': 'quote',
              'r': 'range', 't': 'telnet-option', 'U': 'proxy-user', 'y': 'speed-time', 'Y': 'speed-limit', 'z': 'time-cond',
              'G': 'get', 'I': 'head', ':': 'next'}
CURL_SHORT_VALUE = 'XHdFuAebTowmxEcCDKPQrtUyYz'
CURL_LONG_VALUE = set((
    'url request header data data-ascii data-raw data-binary data-urlencode json form form-string user user-agent referer cookie '
    'cookie-jar output write-out max-time connect-timeout retry retry-delay retry-max-time proxy proxy-user cacert capath cert '
    'cert-type key key-type pass dump-header resolve connect-to interface local-port range time-cond max-redirs limit-rate '
    'config trace trace-ascii stderr speed-limit speed-time noproxy dns-servers ciphers tls-max proto proto-redir unix-socket '
    'abstract-unix-socket expect100-timeout keepalive-time max-filesize oauth2-bearer upload-file url-query aws-sigv4 output-dir '
    'request-target variable hostpubmd5 pinnedpubkey crlfile engine netrc-file ftp-port quote telnet-option continue-at '
    'happy-eyeballs-timeout-ms preproxy proxy-header socks4 socks4a socks5 socks5-hostname doh-url mail-from mail-rcpt mail-auth '
    'login-options delegation ftp-method ftp-account parallel-max create-file-mode alt-svc hsts etag-save etag-compare '
    'tls13-ciphers curves proxy-cacert proxy-cert proxy-key proxy-pass sasl-authzid service-name tlsuser tlspassword '
    'tlsauthtype ip-tos').split())
CURL_FLAGS = set((
    'silent show-error verbose include insecure location location-trusted compressed fail fail-with-body globoff http1.0 '
    'http1.1 http2 http2-prior-knowledge http3 http3-only no-buffer no-keepalive path-as-is progress-bar remote-name '
    'remote-name-all tlsv1 tlsv1.0 tlsv1.1 tlsv1.2 tlsv1.3 ipv4 ipv6 raw tcp-nodelay tr-encoding no-progress-meter basic '
    'no-sessionid ssl ssl-reqd disable disable-epsv retry-all-errors retry-connrefused create-dirs junk-session-cookies '
    'styled-output suppress-connect-headers xattr no-alpn no-npn ssl-no-revoke ssl-revoke-best-effort cert-status false-start '
    'no-clobber remove-on-error parallel parallel-immediate remote-time netrc netrc-optional list-only append use-ascii manual '
    'proxytunnel version help sslv2 sslv3 ignore-content-length').split())
CURL_HINTS = {
    'proxy': 'プロキシは JMeter の実行時オプション -H ホスト -P ポートで指定してください',
    'proxy-user': 'プロキシは JMeter の実行時オプション -H ホスト -P ポートで指定してください',
    'preproxy': 'プロキシは JMeter の実行時オプション -H ホスト -P ポートで指定してください',
    'socks4': 'SOCKS プロキシは JMeter の HTTP リクエストでは指定できません',
    'socks4a': 'SOCKS プロキシは JMeter の HTTP リクエストでは指定できません',
    'socks5': 'SOCKS プロキシは JMeter の HTTP リクエストでは指定できません',
    'socks5-hostname': 'SOCKS プロキシは JMeter の HTTP リクエストでは指定できません',
    'cert': 'クライアント証明書は JMeter の system.properties（javax.net.ssl.keyStore）で設定してください',
    'key': 'クライアント証明書は JMeter の system.properties（javax.net.ssl.keyStore）で設定してください',
    'digest': 'この認証方式は JMeter の HTTP Authorization Manager で設定してください',
    'ntlm': 'この認証方式は JMeter の HTTP Authorization Manager で設定してください',
    'negotiate': 'この認証方式は JMeter の HTTP Authorization Manager で設定してください',
    'anyauth': 'この認証方式は JMeter の HTTP Authorization Manager で設定してください',
    'config': 'curl の設定ファイル（-K）は読み込みません',
    'next': '--next 以降の 2 つ目のリクエストは読み込みません（add-curl で追加するのは 1 行です）',
    'resolve': '名前解決の指定は JMeter の DNS Cache Manager で設定してください',
    'connect-to': '名前解決の指定は JMeter の DNS Cache Manager で設定してください',
    'unix-socket': 'UNIX ドメインソケットは JMeter の HTTP リクエストでは使えません',
    'abstract-unix-socket': 'UNIX ドメインソケットは JMeter の HTTP リクエストでは使えません',
    'aws-sigv4': 'AWS の署名は前処理スクリプトなどで作成してください',
    'request-target': 'リクエストターゲットの指定は JMeter の HTTP リクエストでは表現できません',
}


def curl_split(text):
    """curl コマンドの文字列を引数に分ける。bash の引用符（' " $'...'）と \\ のエスケープ、行末の \\ ^ ` による継続、
    Windows（cmd）の ^ のエスケープに対応する。; && | の後ろ（別のコマンド）は読まない"""
    s = norm_nl(text)
    if '^"' in s or re.search(r'\^[ \t]*\n', s):
        return _curl_split_cmd(s)
    toks, cur, have = [], [], False
    i, n = 0, len(s)
    while i < n:
        c = s[i]
        if c in ' \t\n':
            if have:
                toks.append(''.join(cur))
                cur, have = [], False
            i += 1
        elif c in '`^' and re.match(r'[ \t]*(\n|$)', s[i + 1:i + 80]):
            j = s.find('\n', i)
            i = n if j < 0 else j + 1
        elif c == '\\':
            if i + 1 < n and s[i + 1] != '\n':
                cur.append(s[i + 1])
                have = True
            i += 2
        elif c == "'":
            j = s.find("'", i + 1)
            j = n if j < 0 else j
            cur.append(s[i + 1:j])
            have = True
            i = j + 1
        elif c == '$' and s[i + 1:i + 2] == "'":
            v, i = _curl_ansi(s, i + 2)
            cur.append(v)
            have = True
        elif c == '"':
            i += 1
            while i < n and s[i] != '"':
                if s[i] == '\\' and i + 1 < n and s[i + 1] in '"\\$`\n':
                    if s[i + 1] != '\n':
                        cur.append(s[i + 1])
                    i += 2
                else:
                    cur.append(s[i])
                    i += 1
            i += 1
            have = True
        elif c in ';&|' and not have:
            warn('curl: 「%s」以降は別のコマンドのため読みません' % s[i:i + 2].strip())
            break
        elif c == '#' and not have:
            j = s.find('\n', i)
            i = n if j < 0 else j
        else:
            cur.append(c)
            have = True
            i += 1
    if have:
        toks.append(''.join(cur))
    return toks


def _curl_ansi(s, i):
    # $'...' の中身（i は ' の次の位置）。\xHH・\NNN はバイトとして集め、UTF-8 で読めない場合は Latin-1 として読む
    simple = {'a': 7, 'b': 8, 'e': 27, 'E': 27, 'f': 12, 'n': 10, 'r': 13, 't': 9, 'v': 11, '\\': 92, "'": 39, '"': 34, '?': 63}
    b = bytearray()
    n = len(s)
    while i < n and s[i] != "'":
        c = s[i]
        if c == '\\' and i + 1 < n:
            d = s[i + 1]
            m = (re.match(r'x([0-9A-Fa-f]{1,2})', s[i + 1:i + 4]) or re.match(r'([0-7]{1,3})', s[i + 1:i + 4])
                 or re.match(r'u([0-9A-Fa-f]{1,4})|U([0-9A-Fa-f]{1,8})', s[i + 1:i + 10]))
            if d in simple:
                b.append(simple[d])
                i += 2
            elif m and d in 'xX01234567':
                b.append(int(m.group(1), 16 if d in 'xX' else 8) & 255)
                i += 1 + len(m.group(0))
            elif m:
                cp = int(m.group(1) or m.group(2), 16)
                b += (chr(cp) if cp <= 0x10FFFF and not 0xD800 <= cp <= 0xDFFF else '\ufffd').encode('utf-8')
                i += 1 + len(m.group(0))
            elif d == 'c' and i + 2 < n:
                b.append(ord(s[i + 2]) & 31)
                i += 3
            else:
                b += ('\\' + d).encode('utf-8')
                i += 2
        else:
            b += c.encode('utf-8')
            i += 1
    try:
        v = b.decode('utf-8')
    except UnicodeDecodeError:
        v = b.decode('latin-1')
    return v, i + 1


def _curl_split_cmd(s):
    # Windows（cmd）形式: ^ のエスケープを外してから、Windows の引数の規則（" で囲む・\" は "）で分ける
    t = []
    i, n = 0, len(s)
    while i < n:
        if s[i] == '^':
            if i + 1 < n and s[i + 1] != '\n':
                t.append(s[i + 1])
            i += 2
        else:
            t.append(s[i])
            i += 1
    s = ''.join(t)
    toks, cur, have, inq = [], [], False, False
    i, n = 0, len(s)
    while i < n:
        c = s[i]
        if c in ' \t\n' and not inq:
            if have:
                toks.append(''.join(cur))
                cur, have = [], False
            i += 1
        elif c == '\\':
            j = i
            while j < n and s[j] == '\\':
                j += 1
            if j < n and s[j] == '"':
                cur.append('\\' * ((j - i) // 2))
                if (j - i) % 2:
                    cur.append('"')
                    j += 1
            else:
                cur.append(s[i:j])
            i = j
            have = True
        elif c == '"':
            if inq and s[i + 1:i + 2] == '"':
                cur.append('"')
                i += 2
            else:
                inq = not inq
                i += 1
            have = True
        elif c in '&|' and not inq and not have:
            warn('curl: 「%s」以降は別のコマンドのため読みません' % s[i:i + 2].strip())
            break
        else:
            cur.append(c)
            have = True
            i += 1
    if have:
        toks.append(''.join(cur))
    return toks


def curl_decode(s, plus=True):
    # %XX（UTF-8）を戻す（plus のときは + を空白にする）。UTF-8 として読めない場合は None
    b = bytearray()
    i = 0
    while i < len(s):
        c = s[i]
        if c == '+' and plus:
            b += b' '
            i += 1
        elif c == '%' and re.match(r'%[0-9A-Fa-f]{2}', s[i:i + 3]):
            b.append(int(s[i + 1:i + 3], 16))
            i += 3
        else:
            b += c.encode('utf-8')
            i += 1
    try:
        return b.decode('utf-8')
    except UnicodeDecodeError:
        return None


def curl_pairs(text):
    # 「名前=値&...」→ [(名前, 値)]（%XX と + を戻した値。戻せないものは (元の文字列, None)）
    out = []
    for part in text.split('&'):
        if part == '':
            continue
        k, sep, v = part.partition('=')
        dk, dv = curl_decode(k), curl_decode(v)
        out.append((part, None) if dk is None or dv is None else (dk, dv if sep else ''))
    return out


def curl_rows(pairs):
    # API一覧の「名前=値」（改行区切り）にする。区切り文字などを含む値は oa_row_pairs と同じく URL エンコードして書く
    return '\n'.join(k if v is None else oa_row_pairs([(k, v)]) for k, v in pairs)


def _curl_file(name, strip):
    try:
        with open(name, 'rb') as f:
            data = f.read()
    except (IOError, OSError):
        warn('curl: ファイルを読み込めません（この値は無視します）: ' + name)
        return None
    text = data.decode('utf-8', 'replace')
    if text.startswith('\ufeff'):
        text = text[1:]
    return text.replace('\r', '').replace('\n', '') if strip else text


def _curl_apply(st, opt, v, shown):
    # オプション 1 つを反映する。戻り値 False は「ここで読み終える」
    if opt == 'url':
        st['urls'].append(v)
    elif opt == 'request':
        st['method'] = v.strip().upper()
    elif opt == 'header':
        st['headers'].append(v)
    elif opt in ('data', 'data-ascii', 'data-binary', 'json'):
        if v.startswith('@'):
            v = _curl_file(v[1:], opt in ('data', 'data-ascii'))
            if v is None:
                return True
        st['data'].append(v)
        st['json'] = st['json'] or opt == 'json'
    elif opt == 'data-raw':
        st['data'].append(v)
    elif opt == 'data-urlencode':
        if '=' in v:
            name, content = v.split('=', 1)
        elif '@' in v:
            name, fname = v.split('@', 1)
            content = _curl_file(fname, False)
            if content is None:
                return True
        else:
            name, content = '', v
        st['data'].append((name + '=' if name else '') + oa_pct_encode(content, False))
    elif opt in ('form', 'form-string'):
        st['form'].append((opt, v))
    elif opt in ('get', 'head'):
        st[opt] = True
    elif opt in ('user', 'upload-file'):
        st[opt] = v
    elif opt == 'oauth2-bearer':
        st['bearer'] = v
    elif opt == 'user-agent':
        st['agent'] = v
    elif opt == 'referer':
        st['referer'] = re.sub(r';auto$', '', v)
    elif opt == 'cookie':
        st['cookies'].append(v)
    elif opt == 'url-query':
        st['query'].append(v)
    elif opt in CURL_HINTS:
        warn('curl: %s は使えません。%s（無視します）' % (shown, CURL_HINTS[opt]))
        return opt != 'next'
    elif opt not in CURL_LONG_VALUE and opt not in CURL_FLAGS and shown.startswith('--'):
        warn('curl: 未対応のオプション %s は無視します' % shown)
    return True


def curl_parse(tokens, api_name):
    """curl の引数の一覧 → API一覧の 1 行（dict）。JMeter で表現できないオプションは警告して無視する"""
    head = tokens[0].replace('\\', '/').split('/')[-1].lower() if tokens else ''
    if head not in ('curl', 'curl.exe'):
        raise ToolError('curl コマンドとして読めません（先頭が curl ではありません）: ' + (tokens[0] if tokens else '（空）'))
    st = {'method': None, 'urls': [], 'headers': [], 'data': [], 'form': [], 'query': [], 'cookies': [], 'json': False,
          'get': False, 'head': False, 'upload-file': None, 'user': None, 'bearer': None, 'agent': None, 'referer': None}
    args = tokens[1:]
    i = 0
    while i < len(args):
        a = args[i]
        i += 1
        if a.startswith('--') and len(a) > 2:
            v = None
            if a[2:] in CURL_LONG_VALUE:
                if i >= len(args):
                    warn('curl: %s に値がありません（無視します）' % a)
                    break
                v = args[i]
                i += 1
            if not _curl_apply(st, a[2:], v, a):
                break
        elif a.startswith('-') and len(a) > 1:
            j, go = 1, True
            while j < len(a) and go:
                c = a[j]
                if c in CURL_SHORT_VALUE:
                    v = a[j + 1:]
                    if v == '':
                        if i >= len(args):
                            warn('curl: -%s に値がありません（無視します）' % c)
                            go = False
                            break
                        v = args[i]
                        i += 1
                    go = _curl_apply(st, CURL_SHORT[c], v, '-' + c)
                    break
                if c in CURL_SHORT:
                    go = _curl_apply(st, CURL_SHORT[c], None, '-' + c)
                j += 1
            if not go:
                break
        elif a.lower() in ('curl', 'curl.exe'):
            warn('curl: 2 つ目の curl コマンド以降は読みません（add-curl で追加するのは 1 行です）')
            break
        else:
            st['urls'].append(a)
    return _curl_row(st, api_name)


def _curl_row(st, api_name):
    urls = st['urls']
    if not urls:
        raise ToolError('curl コマンドに URL がありません')
    http = [u for u in urls if re.match(r'(?i)^https?://', u)]
    url = http[0] if http else urls[0]
    if len(urls) > 1:
        warn('curl: URL のような引数が複数あるため %s を使います（ほかは無視します）: %s' % (url, ', '.join(u for u in urls if u != url)))
    if not re.match(r'^[A-Za-z][A-Za-z0-9+.\-]*://', url):
        url = 'http://' + url
        warn('curl: URL にスキームが無いため、curl と同じく http とみなします: ' + url)
    m = re.match(r'^([A-Za-z][A-Za-z0-9+.\-]*)://([^/?#]*)([^?#]*)(?:\?([^#]*))?', url)
    scheme = m.group(1).lower()
    if scheme not in ('http', 'https'):
        raise ToolError('curl: http / https 以外の URL には対応していません: ' + url)
    auth, userinfo = m.group(2), ''
    if '@' in auth:
        userinfo, auth = auth.rsplit('@', 1)
    path = m.group(3) or '/'
    query = curl_pairs(m.group(4)) if m.group(4) else []
    for q in st['query']:
        query += curl_pairs(q[1:]) if q.startswith('+') else [tuple(q.split('=', 1)) if '=' in q else (q, '')]
    method = st['method'] or ('HEAD' if st['head'] else 'GET' if st['get'] else 'PUT' if st['upload-file'] is not None
                              else 'POST' if st['data'] or st['form'] else 'GET')
    hdrs = []
    for h in st['headers']:
        k, sep, v = h.partition(':')
        if h.startswith('@'):
            warn('curl: ファイルから読むヘッダ（-H @ファイル）は読み込みません: ' + h)
        elif sep and k.strip():
            if v.strip():
                hdrs.append([k.strip(), v.strip()])
        elif h.rstrip().endswith(';') and h.rstrip()[:-1].strip():
            hdrs.append([h.rstrip()[:-1].strip(), ''])
        else:
            warn('curl: ヘッダは「名前: 値」の形式で指定してください（無視します）: ' + h)

    def has(name):
        return any(x[0].lower() == name.lower() for x in hdrs)

    def add(name, value):
        if not has(name):
            hdrs.append([name, value])
    user = st['user']
    if user is None and userinfo:
        user = curl_decode(userinfo, False) or userinfo
    if user is not None:
        if ':' in user:
            add('Authorization', 'Basic ' + base64.b64encode(user.encode('utf-8')).decode('ascii'))
        else:
            warn('curl: -u にパスワードが無いため Authorization ヘッダは作りません（curl では実行時に入力を求められます）')
    if st['bearer'] is not None:
        add('Authorization', 'Bearer ' + st['bearer'])
    if st['agent'] is not None:
        add('User-Agent', st['agent'])
    if st['referer'] is not None:
        add('Referer', st['referer'])
    ck = []
    for c in st['cookies']:
        if '=' in c:
            ck.append(c.strip().rstrip(';').strip())
        else:
            warn('curl: クッキーのファイル（-b ファイル名）は読み込みません: ' + c)
    if ck:
        add('Cookie', '; '.join(ck))
    for x in [x for x in hdrs if x[0].lower() in ('content-length', 'transfer-encoding')]:
        hdrs.remove(x)
        info('curl: %s ヘッダは JMeter が自動で設定するため追加しません' % x[0])
    body_type, body = '', ''
    data = '&'.join(st['data']) if st['data'] else None
    if st['get'] and data is not None:
        query += curl_pairs(data)
        data = None
    ctype = [x[1] for x in hdrs if x[0].lower() == 'content-type']
    if st['form']:
        if data is not None:
            warn('curl: -d と -F は同時に指定できないため、-F の値を使います')
        pairs = []
        for kind, v in st['form']:
            k, sep, val = v.partition('=')
            if not sep or not k:
                warn('curl: %s は「名前=値」の形式で指定してください（無視します）: %s' % ('-F' if kind == 'form' else '--form-string', v))
                continue
            if kind == 'form' and val.startswith('@'):
                warn('curl: -F %s はファイルの送信のため追加しません（JMeter の HTTP リクエストの「ファイルアップロード」タブで設定してください）' % v)
                continue
            if kind == 'form' and val.startswith('<'):
                val = _curl_file(val[1:].split(';', 1)[0], False)
                if val is None:
                    continue
            elif kind == 'form':
                val = re.sub(r';\s*(type|filename|headers|encoder)=[^;]*', '', val)
            pairs.append((k, val))
        body_type, body = 'multipart', oa_row_pairs(pairs)
        if ctype and 'multipart' in ctype[0].lower():
            hdrs = [x for x in hdrs if x[0].lower() != 'content-type']
            info('curl: Content-Type（multipart/form-data）は JMeter が boundary を付けて設定するため、ヘッダには追加しません')
    elif st['upload-file'] is not None:
        v = _curl_file(st['upload-file'], False)
        if v is not None:
            body_type, body = 'raw', v
    elif data is not None:
        if st['json']:
            add('Content-Type', 'application/json')
            add('Accept', 'application/json')
            ctype = [x[1] for x in hdrs if x[0].lower() == 'content-type']
        eff = (ctype[0] if ctype else 'application/x-www-form-urlencoded').lower()
        segs = [x for x in data.split('&') if x != '']
        if 'x-www-form-urlencoded' in eff and segs and all(re.match(r'^[^=\s]+=', x) for x in segs):
            body_type, body = 'form', curl_rows(curl_pairs(data))
        elif 'json' in eff:
            body_type, body = 'json', data
        else:
            body_type, body = 'raw', data
            if not ctype:
                hdrs.append(['Content-Type', 'application/x-www-form-urlencoded'])
    if body and method in NO_BODY:
        warn('curl: %s ではリクエストボディを送信しないため、JMX の生成時に無視されます（リクエストボディ列には記入します）' % method)
    row = {'no': '', 'enabled': 'Y', 'name': api_name or (method + ' ' + path), 'method': method,
           'url': scheme + '://' + auth + path, 'query': curl_rows(query),
           'headers': '\n'.join(k + ': ' + v for k, v in hdrs), 'body_type': body_type, 'body': body,
           'comment': 'curl コマンドから追加'}
    return dict((k, clean(v)) for k, v in row.items())


def table_layout(rows, source):
    # API一覧の見出し行を探す（load_api_table と同じ規則）。戻り値: (見出し行の位置, 列の対応 {キー: 列番号}, 見出し行の列数)
    alias = _aliases()
    for idx, (rn, cells) in enumerate(rows):
        found = {}
        for ci, v in enumerate(cells):
            k = alias.get(normalize_key(v))
            if k and k not in found:
                found[k] = ci
        if 'url' in found:
            return idx, found, len(cells)
    raise ToolError('%s: 見出し行（「URL」列を含む行）が見つかりません' % source)


def table_next_no(rows, hidx, colmap):
    # 次の No（数値の No の最大値 + 1。数値が無い場合は API の行数 + 1）
    nums, count = [], 0
    for rn, cells in rows[hidx + 1:]:
        first = ''
        for v in cells:
            if v.strip():
                first = v.strip()
                break
        if first == '' or first.startswith('#'):
            continue
        count += 1
        ci = colmap.get('no', -1)
        v = cells[ci].strip() if 0 <= ci < len(cells) else ''
        if re.match(r'^[0-9]{1,6}$', v):
            nums.append(int(v))
    return str(max(nums) + 1 if nums else count + 1)


def column_header(key):
    for c in SPEC['api_columns']:
        if c['key'] == key:
            return c['header']
    return key


def show_row(row, where):
    info('追加する行（%s）:' % where)
    w = max(_disp_width(c['header']) for c in SPEC['api_columns'])
    for c in SPEC['api_columns']:
        v = row.get(c['key'], '')
        if v == '':
            continue
        lines = v.split('\n')
        print('  %s : %s' % (_disp_pad(c['header'], w), lines[0]))
        for x in lines[1:]:
            print('  %s   %s' % (' ' * w, x))


def backup_copy(path, data):
    bak = '%s.%s.bak' % (path, timestamp())
    write_file(bak, data)
    info('元のファイルを退避しました: ' + bak)


def xlsx_sheet_parts(z, names):
    # ブック内のシート → [(シート名, XML のパス)]
    def load(name):
        return ET.fromstring(z.read(name)) if name in names else None
    wb, rels = load('xl/workbook.xml'), load('xl/_rels/workbook.xml.rels')
    if wb is None or rels is None:
        raise ToolError('Excel ファイルの形式が正しくありません')
    relmap = dict((r.get('Id'), r.get('Target') or '') for r in rels if _ln(r.tag) == 'Relationship')
    out = []
    for el in wb.iter():
        if _ln(el.tag) != 'sheet':
            continue
        rid = ''
        for k, v in el.attrib.items():
            if k.startswith('{') and _ln(k) == 'id':
                rid = v
        target = relmap.get(rid, '')
        part = target[1:] if target.startswith('/') else 'xl/' + target
        if target and part in names:
            out.append((el.get('name') or '', part))
    return out


_XL_SHEETDATA = re.compile(r'<((?:[A-Za-z_][\w.-]*:)?)sheetData\b[^>]*?(/?)>')
_XL_ROW = re.compile(r'<((?:[A-Za-z_][\w.-]*:)?)row\b((?:[^>"\']|"[^"]*"|\'[^\']*\')*?)(/>|>(.*?)</\1row>)', re.S)
_XL_CELL = re.compile(r'<((?:[A-Za-z_][\w.-]*:)?)c\b((?:[^>"\']|"[^"]*"|\'[^\']*\')*?)(/>|>(.*?)</\1c>)', re.S)
_XML_ATTR = re.compile(r'([A-Za-z_:][-A-Za-z0-9_:.]*)\s*=\s*(?:"([^"]*)"|\'([^\']*)\')')


def _xattrs(s):
    return dict((m.group(1), m.group(2) if m.group(2) is not None else m.group(3)) for m in _XML_ATTR.finditer(s))


def xlsx_put_row(xml, target, values):
    # シートの XML の target 行だけを書き換える（値は inlineStr。書式は同じ行の既存のセル、無ければ上の行の同じ列から引き継ぐ）
    m = _XL_SHEETDATA.search(xml)
    if not m:
        raise ToolError('シートの形式を読み取れないため追記できません')
    pfx = m.group(1)
    widths = {}
    for cm in re.finditer(r'<(?:[A-Za-z_][\w.-]*:)?col\b([^>]*)>', xml):
        a = _xattrs(cm.group(1))
        if re.match(r'^[0-9]+$', a.get('min', '')) and re.match(r'^[0-9]+$', a.get('max', '')) and \
                re.match(r'^[0-9]+(\.[0-9]+)?$', a.get('width', '')):
            for c in range(int(a['min']), min(int(a['max']), int(a['min']) + 100) + 1):
                widths[c] = int(float(a['width']))
    rows = []
    content_end = m.end()
    if not m.group(2):
        content_end = xml.find('</%ssheetData>' % pfx, m.end())
        if content_end < 0:
            raise ToolError('シートの形式を読み取れないため追記できません')
        for rm in _XL_ROW.finditer(xml, m.end(), content_end):
            r = _xattrs(rm.group(2)).get('r', '')
            if not re.match(r'^[0-9]+$', r):
                raise ToolError('行番号（r 属性）の無い行があるため追記できません（Excel で開いて保存し直してください）')
            rows.append((int(r), rm))
    styles, existing, insert_at = {}, None, None
    for r, rm in rows:
        if r < target:
            for cx in _XL_CELL.finditer(rm.group(4) or ''):
                a = _xattrs(cx.group(2))
                if a.get('r') and a.get('s') is not None:
                    styles[_col_index(a['r'])] = a['s']
        elif r == target:
            existing = rm
        elif insert_at is None:
            insert_at = rm.start()
    out, keep_style = {}, {}
    row_attrs = ' r="%d"' % target
    if existing is not None:
        row_attrs = re.sub(r'\s(ht|customHeight)\s*=\s*("[^"]*"|\'[^\']*\')', '', existing.group(2))
        for cx in _XL_CELL.finditer(existing.group(4) or ''):
            a = _xattrs(cx.group(2))
            if not a.get('r'):
                raise ToolError('セル番地（r 属性）の無いセルがあるため追記できません（Excel で開いて保存し直してください）')
            out[_col_index(a['r'])] = cx.group(0)
            keep_style[_col_index(a['r'])] = a.get('s')
    lines = 1
    for c in sorted(values):
        s = keep_style.get(c) if keep_style.get(c) is not None else styles.get(c)
        out[c] = '<%sc r="%s"%s t="inlineStr"><%sis><%st xml:space="preserve">%s</%st></%sis></%sc>' % (
            pfx, cref(target, c + 1), (' s="%s"' % s) if s is not None else '', pfx, pfx, xesc(values[c]), pfx, pfx, pfx)
        lines = max(lines, wrap_lines(values[c], widths.get(c + 1, 8)))
    row_attrs += ' ht="%s" customHeight="1"' % fmt_h10(row_h10(lines))
    new_row = '<%srow%s>%s</%srow>' % (pfx, row_attrs, ''.join(out[c] for c in sorted(out)), pfx)
    if m.group(2):
        xml = xml[:m.start()] + '<%ssheetData>%s</%ssheetData>' % (pfx, new_row, pfx) + xml[m.end():]
    elif existing is not None:
        xml = xml[:existing.start()] + new_row + xml[existing.end():]
    else:
        pos = insert_at if insert_at is not None else content_end
        xml = xml[:pos] + new_row + xml[pos:]
    dm = re.search(r'(<(?:[A-Za-z_][\w.-]*:)?dimension\b[^>]*?\bref=")([A-Z]+)([0-9]+)(?::([A-Z]+)([0-9]+))?"', xml)
    if dm:
        c2, r2 = (dm.group(4), int(dm.group(5))) if dm.group(4) else (dm.group(2), int(dm.group(3)))
        ref = '%s%s:%s%d' % (dm.group(2), dm.group(3), col_letter(max(_col_index(c2) + 1, max(values) + 1)), max(r2, target))
        xml = xml[:dm.start(2)] + ref + xml[dm.end(5) if dm.group(4) else dm.end(3):]
    return xml


def xlsx_append(path, row, dry):
    with open(path, 'rb') as f:
        data = f.read()
    try:
        z = zipfile.ZipFile(io.BytesIO(data))
    except Exception:
        raise ToolError('Excel ファイルを開けません（.xlsx 形式で保存してください）: ' + path)
    with z:
        infos = z.infolist()
        names = set(x.filename for x in infos)
        shared = []
        if 'xl/sharedStrings.xml' in names:
            for si in ET.fromstring(z.read('xl/sharedStrings.xml')):
                if _ln(si.tag) == 'si':
                    shared.append(_xl_text(si))
        parts = xlsx_sheet_parts(z, names)
        want = [normalize_key(x) for x in (TPL['api_sheet'], 'API', 'APIs', 'api_list')]
        chosen = [p for p in parts if normalize_key(p[0]) in want][:1]
        for p in parts:
            if chosen:
                break
            try:
                table_layout(_xl_rows(ET.fromstring(z.read(p[1])), shared), p[0])
                chosen = [p]
            except ToolError:
                continue
        if not chosen:
            raise ToolError('「%s」シートが見つかりません: %s' % (TPL['api_sheet'], path))
        name, part = chosen[0]
        raw = z.read(part)
        rows = _xl_rows(ET.fromstring(raw), shared)
        hidx, colmap, width = table_layout(rows, '%s（%sシート）' % (os.path.basename(path), name))
        last = rows[hidx][0]
        for rn, cells in rows:
            if rn > last and any(v.strip() for v in cells):
                last = rn
        target = last + 1
        row = dict(row)
        row['no'] = row['no'] or table_next_no(rows, hidx, colmap)
        values = {}
        for c in SPEC['api_columns']:
            v = row.get(c['key'], '')
            if v and c['key'] in colmap:
                values[colmap[c['key']]] = v
            elif v:
                warn('「%s」シートに「%s」列が無いため、その値は記入しません' % (name, c['header']))
        if any(len(v) > 32767 for v in values.values()):
            raise ToolError('Excel のセルに入る文字数（32767 文字）を超える値があります。CSV の API 一覧に追加してください')
        show_row(row, '%s の「%s」シート %d 行目' % (path, name, target))
        if dry:
            info('--dry-run のため書き込みません')
            return 0
        new_xml = xlsx_put_row(raw.decode('utf-8'), target, values).encode('utf-8')
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zo:
            for x in infos:
                zo.writestr(x, new_xml if x.filename == part else z.read(x.filename), compress_type=x.compress_type)
    backup_copy(path, data)
    write_file(path, buf.getvalue())
    info('1 行追加しました: %s（「%s」シート %d 行目）' % (path, name, target))
    return 0


def csv_append(path, row, dry):
    with open(path, 'rb') as f:
        data = f.read()
    text, enc = read_text(path)
    rows = [(i + 1, r) for i, r in enumerate(parse_csv(text))]
    hidx, colmap, width = table_layout(rows, os.path.basename(path))
    row = dict(row)
    row['no'] = row['no'] or table_next_no(rows, hidx, colmap)
    cells = [''] * max(width, max(colmap.values()) + 1)
    for c in SPEC['api_columns']:
        v = row.get(c['key'], '')
        if v and c['key'] in colmap:
            cells[colmap[c['key']]] = v
        elif v:
            warn('CSV に「%s」列が無いため、その値は記入しません' % c['header'])
    show_row(row, '%s の末尾' % path)
    if dry:
        info('--dry-run のため書き込みません')
        return 0
    codec, bom = {'UTF-8(BOM)': ('utf-8', 3), 'UTF-8': ('utf-8', 0), 'Shift_JIS(CP932)': ('cp932', 0),
                  'UTF-16LE': ('utf-16-le', 2), 'UTF-16BE': ('utf-16-be', 2)}[enc]
    body = data[bom:].decode(codec, 'replace')
    nl = '\r\n' if '\r\n' in body or body == '' else '\n'
    line = ('' if body == '' or body.endswith('\n') else nl) + ','.join(csv_field(v) for v in cells) + nl
    try:
        add = line.encode(codec)
    except UnicodeEncodeError:
        raise ToolError('CSV の文字コード（%s）で表せない文字が含まれています。CSV を UTF-8 で保存し直してください: %s' % (enc, path))
    backup_copy(path, data)
    write_file(path, data + add)
    info('1 行追加しました: %s（%d 件目の API）' % (path, len([1 for rn, cl in rows[hidx + 1:] if any(v.strip() for v in cl)]) + 1))
    return 0


def cmd_add_curl(opts, tokens):
    path = opts.get('input')
    if not path:
        raise ToolError('追加先の API 一覧（.xlsx / .csv）を -i で指定してください（例: -i input/api_definition.xlsx）')
    if not os.path.isfile(path):
        raise ToolError('ファイルが見つかりません: ' + path)
    given = [x for x in (opts.get('curl'), opts.get('curl_file'), tokens) if x]
    if len(given) != 1:
        raise ToolError('curl コマンドは --curl "curl ..."、--curl-file ファイル、-- curl ... のどれか 1 つで指定してください')
    if tokens:
        args = tokens
    else:
        if opts.get('curl_file') and not os.path.isfile(opts['curl_file']):
            raise ToolError('curl コマンドのファイルが見つかりません: ' + opts['curl_file'])
        args = curl_split(opts['curl'] if opts.get('curl') else read_text(opts['curl_file'])[0])
    row = curl_parse(args, (opts.get('name') or '').strip())
    info('curl コマンドを解析しました: %s %s' % (row['method'], row['url']))
    ext = os.path.splitext(path)[1].lower()
    if ext in ('.xlsx', '.xlsm'):
        return xlsx_append(path, row, opts.get('dry_run'))
    if ext in ('.csv', '.txt'):
        return csv_append(path, row, opts.get('dry_run'))
    raise ToolError('add-curl の -i には API 一覧の Excel（.xlsx）または CSV（.csv）を指定してください: ' + path)


# ----------------------------------------------------------------------------- 2. linux-cmd
LXC = SPEC['linux_cmd']


def lx_props(text):
    # JMX の ${__P(名前,既定値)} → {名前: 既定値}（出現順。既定値の \, と \\ は戻す）
    text = jx_unescape(text)
    props = OrderedDict()
    i = 0
    while True:
        i = text.find('${__P(', i)
        if i < 0:
            break
        j, depth, buf = i + 6, 0, []
        while j < len(text):
            c = text[j]
            if c == '\\' and j + 1 < len(text):
                buf.append(text[j:j + 2])
                j += 2
                continue
            if c == '(':
                depth += 1
            elif c == ')':
                if depth == 0:
                    break
                depth -= 1
            buf.append(c)
            j += 1
        arg = ''.join(buf)
        k = 0
        while k < len(arg) and arg[k] != ',':
            k += 2 if arg[k] == '\\' else 1
        name = arg[:k].strip()
        if re.match(r'^[A-Za-z0-9_.\-]+$', name) and name not in props:
            props[name] = re.sub(r'\\([$,\\])', r'\1', arg[k + 1:])
        i = j + 1
    return props


def lx_quote(v):
    if v == '' or re.match(r'^[A-Za-z0-9_./:@%+=,-]+$', v):
        return v
    return "'" + v.replace("'", "'\\''") + "'"


def lx_desc(name):
    d = LXC['props'].get(name)
    if d:
        return d
    m = re.match(r'^(PROTOCOL|HOST|PORT)_([0-9]+)$', name)
    if m:
        return LXC['props'][m.group(1)] + fill(LXC['nth_host'], {'N': m.group(2)})
    return LXC['user_desc']


def lx_text(jmx_path):
    with open(jmx_path, 'rb') as f:
        text = f.read().decode('utf-8', 'replace')
    props = lx_props(text)
    names = [k for k in LXC['order'] if k in props] + [k for k in props if k not in LXC['order']]
    m = re.search(r'<TestPlan\b[^>]*?\btestname="([^"]*)"', text)
    plan = jx_unescape(m.group(1)) if m else ''
    jmx = os.path.basename(jmx_path)
    stem = os.path.splitext(jmx)[0]
    files = [(jmx, LXC['file_jmx'])]
    for key in ('CSV_FILE', 'EXCEL_FILE'):
        v = props.get(key, '').replace('\\', '/')
        if v and '${' not in v and not re.match(r'^[A-Za-z]:', v):
            files.append((v, fill(LXC['file_abs' if v.startswith('/') else 'file_rel'], {'KEY': key, 'DESC': LXC['file_desc'][key]})))
    for d in props.get('ACCESS_TOKEN_DIR', '').split(';'):
        d = d.strip().replace('\\', '/')
        d = d[2:] if d.startswith('~/') else d
        if d and not re.match(r'^[A-Za-z]:', d) and not d.startswith('//'):
            files.append((d.rstrip('/') + '/' + props.get('ACCESS_TOKEN_FILE', ''),
                          fill(LXC['file_abs' if d.startswith('/') else 'file_rel'], {'KEY': 'ACCESS_TOKEN_DIR', 'DESC': LXC['file_desc']['ACCESS_TOKEN_DIR']})))
    w = max(_disp_width(f) for f, d in files) + 2
    flines = ['    ' + _disp_pad(f, w) + d for f, d in files]
    jargs = ['-J%s=%s' % (k, lx_quote(props[k])) for k in names if k != 'RESULT_DIR']
    cmd = ['jmeter -n -t %s' % lx_quote(jmx), '-l "$RUN_DIR/result.jtl"', '-j "$RUN_DIR/jmeter.log"', '-e -o "$RUN_DIR/report"',
           '-Jjmeter.save.saveservice.subresults=false', '-JRESULT_DIR="$RUN_DIR"'] + jargs
    plines = []
    for k in names:
        plines.append('  -J%s=%s' % (k, lx_quote(props[k]) if k != 'RESULT_DIR' else '"$RUN_DIR"'))
        plines.append('      ' + lx_desc(k) + fill(LXC['default_note'], {'VALUE': props[k] if props[k] != '' else LXC['empty']}))
    return fill(LXC['template'], {'JMX': jmx, 'STEM': stem, 'PLAN': plan, 'VERSION': VERSION,
                                  'DATE': datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                                  'FILES': '\n'.join(flines), 'CMD_ML': ' \\\n  '.join(cmd), 'CMD_1L': ' '.join(cmd),
                                  'PROPS': '\n'.join(plines) if plines else LXC['no_props']})


def write_linux_cmd(jmx_path, out_dir=None):
    out = os.path.join(out_dir or os.path.dirname(jmx_path), os.path.splitext(os.path.basename(jmx_path))[0] + LXC['suffix'])
    write_text(out, lx_text(jmx_path))
    info('Linux サーバでの実行コマンドを出力しました: ' + out)
    return out


def cmd_linux_cmd(opts):
    path = opts.get('jmx')
    if not path:
        raise ToolError('JMX を -t で指定してください（例: -t output/api_test.jmx）')
    if not os.path.isfile(path):
        raise ToolError('JMX ファイルが見つかりません: ' + path)
    if opts.get('out') and not os.path.isdir(opts['out']):
        os.makedirs(opts['out'])
    write_linux_cmd(path, opts.get('out'))
    return 0


# ----------------------------------------------------------------------------- 3. generate -u（既存の JMX の更新）
UPD = SPEC['update']
_JX_TOKEN = re.compile(r'<!--.*?-->|<!\[CDATA\[.*?\]\]>|<\?.*?\?>|<!DOCTYPE[^>]*>'
                       r'|<(/?)([A-Za-z_][-A-Za-z0-9_.:]*)((?:[^>"\']|"[^"]*"|\'[^\']*\')*?)(/?)>', re.S)
_JX_ENT = re.compile(r'&(lt|gt|amp|quot|apos|#[0-9]+|#[xX][0-9A-Fa-f]+);')


def jx_unescape(s):
    def f(m):
        e = m.group(1)
        if e[0] != '#':
            return {'lt': '<', 'gt': '>', 'amp': '&', 'quot': '"', 'apos': "'"}[e]
        try:
            c = int(e[2:], 16) if e[1] in 'xX' else int(e[1:])
            return chr(c) if c <= 0x10FFFF and not 0xD800 <= c <= 0xDFFF else m.group(0)
        except ValueError:
            return m.group(0)
    return _JX_ENT.sub(f, s)


def jx_parse(text, what):
    """XML（JMX）を、要素ごとに元の文字列での位置（s: 開始 / os: 開始タグの終わり / cs: 終了タグの始まり / e: 終了）を持つ木にする"""
    top = {'tag': '', 'attrs': [], 'kids': [], 'text': None, 's': 0, 'os': 0, 'cs': len(text), 'e': len(text)}
    stack = [top]
    for m in _JX_TOKEN.finditer(text):
        if m.group(2) is None:
            continue
        if m.group(1):
            node = stack.pop()
            if node is top or node['tag'] != m.group(2):
                raise ToolError('%s の XML が正しくありません（</%s> の位置: %d 文字目）' % (what, m.group(2), m.start()))
            node['cs'], node['e'] = m.start(), m.end()
            if not node['kids']:
                raw = text[node['os']:node['cs']]
                parts = re.split(r'<!\[CDATA\[(.*?)\]\]>', raw, flags=re.S)
                node['text'] = ''.join(p if k % 2 else jx_unescape(p) for k, p in enumerate(parts))
            stack[-1]['kids'].append(node)
            continue
        node = {'tag': m.group(2), 'attrs': [[a.group(1), jx_unescape(a.group(2) if a.group(2) is not None else a.group(3))]
                                             for a in _XML_ATTR.finditer(m.group(3))],
                'kids': [], 'text': None, 's': m.start(), 'os': m.end(), 'cs': -1, 'e': m.end()}
        if m.group(4):
            stack[-1]['kids'].append(node)
        else:
            stack.append(node)
    if len(stack) != 1 or len(top['kids']) != 1:
        raise ToolError('%s の XML が正しくありません（閉じていない要素があります）' % what)
    return top['kids'][0]


def xattr(node, name):
    for k, v in node['attrs']:
        if k == name:
            return v
    return None


def xser(node, pad, nl):
    # 要素を JMX と同じ書式（2 文字の字下げ）で書く
    a = ''.join(' %s="%s"' % (k, xesc(v)) for k, v in node['attrs'])
    if node['kids']:
        return ('%s<%s%s>%s' % (pad, node['tag'], a, nl) + ''.join(xser(c, pad + '  ', nl) for c in node['kids'])
                + '%s</%s>%s' % (pad, node['tag'], nl))
    if node['text'] is None:
        return '%s<%s%s/>%s' % (pad, node['tag'], a, nl)
    return '%s<%s%s>%s</%s>%s' % (pad, node['tag'], a, xesc(node['text']).replace('\n', nl), node['tag'], nl)    # 値の中の改行も JMX の改行コードに合わせる


def upd_name_res():
    # 要素名のひな形（SPEC names）→ 正規表現。@@ID@@ / @@VAR@@ 以外の値（待機時間・ステータスなど）は任意の文字列
    out = []
    for key, tpl in SPEC['names'].items():
        pat, pos = '', 0
        for m in re.finditer(r'@@([A-Z]+)@@', tpl):
            pat += re.escape(tpl[pos:m.start()])
            pat += {'ID': r'(?:API[0-9]+|API_[A-Za-z0-9_]+)', 'VAR': r'(?P<VAR>.+)'}.get(m.group(1), r'.*')
            pos = m.end()
        out.append((key, re.compile('^' + pat + re.escape(tpl[pos:]) + '$', re.S)))
    return out


UPD_NAME_RES = upd_name_res()
UPD_ATTRS = ('testname', 'enabled', 'guiclass', 'testclass')


def upd_api_id(el, name):
    for p in el['kids']:
        if xattr(p, 'name') == 'TestPlan.comments' and p['text']:
            m = re.match(r'^(API[0-9]+|API_[A-Za-z0-9_]+): ', p['text'])
            if m:
                return m.group(1)
    m = re.match(r'^(API[0-9]+)_', name) or re.match(r'^(API_[A-Za-z0-9_]+?)_', name)
    return m.group(1) if m else None


def upd_ident(te):
    # このツールが作った要素の識別子。利用者が追加した要素は None（更新・削除の対象外）
    el = te['el']
    name = xattr(el, 'testname') or ''
    if el['tag'] == 'TestPlan':
        return ('testplan', '', '')
    for key, rx in UPD_NAME_RES:
        m = rx.match(name)
        if m:
            return (key, el['tag'], m.groupdict().get('VAR') or '')
    if el['tag'] == 'HTTPSamplerProxy':
        aid = upd_api_id(el, name)
        if aid:
            return ('api', aid, name[len(aid) + 1:] if name.startswith(aid + '_') else name)
    return None


def te_list(ht):
    out, kids, i = [], (ht['kids'] if ht is not None else []), 0
    while i < len(kids):
        el = kids[i]
        if el['tag'] == 'hashTree':
            i += 1
            continue
        tree = kids[i + 1] if i + 1 < len(kids) and kids[i + 1]['tag'] == 'hashTree' else None
        te = {'el': el, 'ht': tree, 'kids': te_list(tree)}
        te['id'] = upd_ident(te)
        out.append(te)
        i += 2 if tree is not None else 1
    return out


def te_root(text, what):
    root = jx_parse(text, what)
    if root['tag'] != 'jmeterTestPlan':
        raise ToolError('%s は JMeter のテスト計画（JMX）ではありません' % what)
    hts = [k for k in root['kids'] if k['tag'] == 'hashTree']
    plans = [t for t in (te_list(hts[0]) if hts else []) if t['el']['tag'] == 'TestPlan']
    if not plans:
        raise ToolError('%s にテストプラン（TestPlan）がありません' % what)
    return plans[0]


def upd_api_key(te, rule):
    if rule == 'label':
        return xattr(te['el'], 'testname')
    if rule == 'id':
        return te['id'][1]
    if rule == 'name':
        return te['id'][2]
    return re.sub(r'\s*\[[^\]]*\]$', '', te['id'][2])


def upd_pair(xs, ys, openapi):
    # 要素の対応付け（xs の位置 → ys の位置）。HTTP リクエストはラベル → 番号・API名の順（OpenAPI は API名を先）
    res, used, first = {}, set(), {}
    for j, y in enumerate(ys):
        if y['id'] and y['id'][0] != 'api' and y['id'] not in first:
            first[y['id']] = j
    for i, x in enumerate(xs):
        j = first.get(x['id']) if x['id'] and x['id'][0] != 'api' else None
        if j is not None and j not in used:
            res[i] = j
            used.add(j)
    for rule in ('label', 'name', 'core', 'id') if openapi else ('label', 'id', 'name'):
        idx = {}
        for j, y in enumerate(ys):
            if j not in used and y['id'] and y['id'][0] == 'api' and upd_api_key(y, rule):
                idx.setdefault(upd_api_key(y, rule), []).append(j)
        for i, x in enumerate(xs):
            if i in res or not x['id'] or x['id'][0] != 'api':
                continue
            lst = idx.get(upd_api_key(x, rule)) or []
            while lst and lst[0] in used:
                lst.pop(0)
            if lst:
                res[i] = lst.pop(0)
                used.add(res[i])
    return res


def upd_triples(okids, bkids, tkids, openapi):
    # 戻り値: ([(JMX の要素, 基準の要素, 生成した要素), ...]（生成した要素の順）, [定義から無くなった JMX の要素, ...])
    bo, tb = upd_pair(bkids, okids, openapi), upd_pair(tkids, bkids, openapi)
    rest_o = [j for j in range(len(okids)) if j not in set(bo.values())]
    rest_t = [i for i in range(len(tkids)) if i not in tb]
    to = upd_pair([tkids[i] for i in rest_t], [okids[j] for j in rest_o], openapi)
    triples = []
    for i, t in enumerate(tkids):
        if i in tb:
            triples.append((okids[bo[tb[i]]] if tb[i] in bo else None, bkids[tb[i]], t))
        else:
            k = rest_t.index(i)
            triples.append((okids[rest_o[to[k]]] if k in to else None, None, t))
    tbv = set(tb.values())
    return triples, [okids[bo[b]] for b in range(len(bkids)) if b in bo and b not in tbv]


def upd_pkey(node):
    n = xattr(node, 'name')
    if n is not None:
        return n
    if node['tag'] == 'objProp':
        for c in node['kids']:
            if c['tag'] == 'name':
                return 'obj:' + (c['text'] or '')
    return node['tag']


def upd_props(el):
    out = OrderedDict()
    for a in UPD_ATTRS:
        v = xattr(el, a)
        if v is not None:
            out['@' + a] = v
    for c in el['kids']:
        out.setdefault(upd_pkey(c), c)
    return out


def upd_val(v):
    # 比較用の値（属性は文字列、プロパティはノード）。空の文字列・false のプロパティは「無い」と同じに扱う
    if v is None:
        return None
    if isinstance(v, str):
        return 'A' + v
    attrs = '|'.join('%s=%s' % (a, b) for a, b in v['attrs'] if a != 'name')
    if not v['kids']:
        t = norm_nl(v['text'] or '')
        if (v['tag'] == 'stringProp' and t == '') or (v['tag'] == 'boolProp' and t == 'false'):
            return None
        return '%s(%s)=%s' % (v['tag'], attrs, t)
    if v['tag'] == 'collectionProp':
        return '%s(%s)[%s]' % (v['tag'], attrs, '\x1f'.join(upd_val(c) or '' for c in v['kids']))
    items = sorted('%s\x1e%s' % (upd_pkey(c), upd_val(c)) for c in v['kids'] if upd_val(c) is not None)
    return '%s(%s){%s}' % (v['tag'], attrs, '\x1f'.join(items))


def upd_tval(te):
    own = sorted('%s\x1e%s' % (k, upd_val(v)) for k, v in upd_props(te['el']).items() if upd_val(v) is not None)
    return '\x1d'.join(own) + '\x1c[' + '\x1c'.join(upd_tval(c) for c in te['kids']) + ']'


def upd_line_span(text, s, e):
    # 要素が行全体を占める場合は、行頭の字下げ・行末の改行も含めた範囲
    ls = text.rfind('\n', 0, s) + 1
    le = text.find('\n', e)
    le = len(text) if le < 0 else le + 1
    if text[ls:s].strip() == '' and text[e:le].strip() == '':
        return ls, le, True
    return s, e, False


def upd_pad(text, pos):
    ls = text.rfind('\n', 0, pos) + 1
    return text[ls:pos] if text[ls:pos].strip() == '' else ''


class UpdCtx(object):
    def __init__(self, text, openapi, two_way):
        self.text, self.openapi, self.two_way = text, openapi, two_way
        self.nl = '\r\n' if '\r\n' in text else '\n'
        self.edits, self.changes = [], []
        self.kept = self.kept_deleted = self.silent = 0

    def edit(self, s, e, t):
        self.edits.append((s, e, len(self.edits), t))

    def put(self, node, new_full, pad):
        s, e, full = upd_line_span(self.text, node['s'], node['e'])
        self.edit(s, e, new_full if full else new_full[len(pad):len(new_full) - len(self.nl)])


def upd_set_attr(c, el, name, value):
    stag = c.text[el['s']:el['os']]
    for m in _XML_ATTR.finditer(stag):
        if m.group(1) == name:
            g = 2 if m.group(2) is not None else 3
            c.edit(el['s'] + m.start(g), el['s'] + m.end(g), xesc(value))
            return
    end = el['os'] - (2 if stag.endswith('/>') else 1)
    c.edit(end, end, ' %s="%s"' % (name, xesc(value)))


def upd_apply_prop(c, o, op, tp, k):
    el = o['el']
    if k.startswith('@'):
        if tp.get(k) is not None:
            upd_set_attr(c, el, k[1:], tp[k])
        return
    onode, tnode = op.get(k), tp.get(k)
    if tnode is None:
        s, e, full = upd_line_span(c.text, onode['s'], onode['e'])
        c.edit(s, e, '')
        return
    if onode is not None:
        pad = upd_pad(c.text, onode['s'])
        c.put(onode, xser(tnode, pad, c.nl), pad)
        return
    keys = list(tp.keys())
    anchor = None
    for prev in reversed(keys[:keys.index(k)]):
        if not prev.startswith('@') and op.get(prev) is not None:
            anchor = op[prev]
            break
    first = [x for x in el['kids']]
    pad = upd_pad(c.text, first[0]['s']) if first else upd_pad(c.text, el['s']) + '  '
    if anchor is not None:
        pos = upd_line_span(c.text, anchor['s'], anchor['e'])[1]
    elif first:
        pos = upd_line_span(c.text, first[0]['s'], first[0]['e'])[0]
    else:
        nl_pos = c.text.find('\n', el['os'])
        pos = nl_pos + 1 if 0 <= nl_pos < el['cs'] else el['os']
    c.edit(pos, pos, xser(tnode, pad, c.nl))


def upd_merge_props(c, o, b, t, path):
    key = t['id'][0] if t['id'] else ''
    op, tp = upd_props(o['el']), upd_props(t['el'])
    bp = upd_props(b['el']) if b is not None else op
    two = c.two_way or b is None
    whole = False
    for k in list(tp.keys()) + [k for k in bp if k not in tp]:
        ov, bv, tv = upd_val(op.get(k)), upd_val(bp.get(k)), upd_val(tp.get(k))
        if tv == bv:
            if ov != bv:
                c.kept += 1
            continue
        if ov == tv:
            continue
        if k == 'script' and key in UPD['protect'] and (two or ov != bv):
            c.changes.append({'kind': '維持', 'path': path, 'what': k, 'old': op.get(k), 'new': tp.get(k)})
            continue
        if o['el']['cs'] >= 0:
            upd_apply_prop(c, o, op, tp, k)
        elif not whole:
            whole = True        # 子の無い（<要素 .../>）要素は、生成した要素で丸ごと置き換える
            pad = upd_pad(c.text, o['el']['s'])
            c.put(o['el'], xser(t['el'], pad, c.nl), pad)
        if key == 'testplan' and k == 'TestPlan.comments':
            c.silent += 1
        else:
            c.changes.append({'kind': '競合' if not two and ov != bv else '変更', 'path': path, 'what': k,
                              'old': op.get(k), 'new': tp.get(k)})


def upd_ser_te(t, pad, nl):
    return xser(t['el'], pad, nl) + (xser(t['ht'], pad, nl) if t['ht'] is not None else pad + '<hashTree/>' + nl)


def upd_count(te):
    return sum(1 + upd_count(k) for k in te['kids'])


def upd_merge(c, o, b, t, path):
    upd_merge_props(c, o, b, t, path)
    triples, gone = upd_triples(o['kids'], b['kids'] if b is not None else [], t['kids'], c.openapi)
    for x in gone:
        s, e, full = upd_line_span(c.text, x['el']['s'], (x['ht'] or x['el'])['e'])
        c.edit(s, e, '')
        c.changes.append({'kind': '削除', 'path': path + [xattr(x['el'], 'testname') or x['el']['tag']], 'te': x})
    groups, prev, first_o = OrderedDict(), None, None
    for ox, bx, tx in triples:
        if ox is not None:
            first_o = first_o or ox
            upd_merge(c, ox, bx, tx, path + [xattr(ox['el'], 'testname') or ox['el']['tag']])
            prev = ox
        elif bx is not None and upd_tval(bx) == upd_tval(tx):
            c.kept_deleted += 1
        else:
            groups.setdefault(id(prev) if prev is not None else None, [prev, []])[1].append(tx)
    ht = o['ht']
    for anchor, adds in groups.values():
        for tx in adds:
            c.changes.append({'kind': '追加', 'path': path + [xattr(tx['el'], 'testname') or tx['el']['tag']], 'te': tx,
                              'note': 'JMeter で削除済みの要素ですが、定義が変わったため追加します' if
                              any(bx is not None and tx is tt for oo, bx, tt in triples) else ''})
        if ht is None:
            continue
        kids = [k for k in ht['kids']]
        pad = upd_pad(c.text, kids[0]['s']) if kids else upd_pad(c.text, ht['s']) + '  '
        body = ''.join(upd_ser_te(tx, pad, c.nl) for tx in adds)
        if anchor is not None:
            pos = upd_line_span(c.text, anchor['el']['s'], (anchor['ht'] or anchor['el'])['e'])[1]
            c.edit(pos, pos, body)
        elif first_o is not None:
            pos = upd_line_span(c.text, first_o['el']['s'], first_o['el']['e'])[0]
            c.edit(pos, pos, body)
        elif kids:
            pos = upd_line_span(c.text, kids[-1]['s'], kids[-1]['e'])[1]
            c.edit(pos, pos, body)
        elif ht['cs'] < 0:
            hp = upd_pad(c.text, ht['s'])
            c.edit(ht['s'], ht['e'], '<hashTree>' + c.nl + body + hp + '</hashTree>')
        else:
            c.edit(ht['os'], ht['cs'], c.nl + body + upd_pad(c.text, ht['s']))


def upd_apply(text, edits):
    out, pos = [], 0
    for s, e, seq, t in sorted(edits, key=lambda x: (x[0], x[1], x[2])):
        if s < pos:
            raise ToolError('JMX の更新箇所が重なっています（内部エラー）')
        out.append(text[pos:s])
        out.append(t)
        pos = e
    out.append(text[pos:])
    return ''.join(out)


def upd_text(v):
    # 表示用の値（ヘッダは「名前: 値」、パラメータは「名前=値」の行）
    if v is None:
        return ''
    if isinstance(v, str):
        return v
    if not v['kids']:
        return norm_nl(v['text'] or '')
    if v['tag'] == 'collectionProp':
        return '\n'.join(upd_item(x) for x in v['kids'])
    if v['tag'] == 'objProp':
        return UPD['save_config']
    return '\n'.join(['%s=%s' % (upd_pkey(x), norm_nl(x['text'] or '')) for x in v['kids'] if not x['kids']] +
                     [upd_text(x) for x in v['kids'] if x['kids']])


def upd_item(x):
    if not x['kids']:
        return norm_nl(x['text'] or '')
    p = dict((upd_pkey(k), k) for k in x['kids'])

    def tx(key):
        return norm_nl(p[key]['text'] or '') if key in p and not p[key]['kids'] else ''
    if 'Header.name' in p:
        return tx('Header.name') + ': ' + tx('Header.value')
    if 'Argument.value' in p:
        return (tx('Argument.name') + '=' + tx('Argument.value')) if tx('Argument.name') else tx('Argument.value')
    if 'Cookie.value' in p:
        return (xattr(x, 'name') or '') + '=' + tx('Cookie.value')
    return upd_text(x).replace('\n', '; ')


def upd_linediff(a, b):
    al, bl = (a.split('\n') if a else []), (b.split('\n') if b else [])
    rest, minus, plus = list(bl), [], []
    for x in al:
        if x in rest:
            rest.remove(x)
        else:
            minus.append('- ' + x)
    rest = list(al)
    for x in bl:
        if x in rest:
            rest.remove(x)
        else:
            plus.append('+ ' + x)
    lines = minus + plus or [UPD['order_only']]
    limit = UPD['max_lines']
    return lines if len(lines) <= limit else lines[:limit] + [fill(UPD['more_lines'], {'N': str(len(lines) - limit)})]


def upd_show(c, path, mode_note, dry):
    print('')
    print(fill(UPD['title'], {'PATH': path}))
    print(mode_note)
    prev = None
    for ch in c.changes:
        if (ch['kind'], ch['path']) != prev or 'te' in ch:
            print('[%s] %s' % (ch['kind'], ' / '.join(ch['path']) or UPD['plan_label']))
        prev = (ch['kind'], ch['path'])
        if 'te' in ch:
            n = upd_count(ch['te'])
            print('       %s%s' % (UPD['tags'].get(ch['te']['el']['tag'], ch['te']['el']['tag']),
                                   fill(UPD['children'], {'N': str(n)}) if n else ''))
            if ch.get('note'):
                print('       ' + ch['note'])
            continue
        label = UPD['labels'].get(ch['what'], ch['what'])
        ot, nt = upd_text(ch['old']), upd_text(ch['new'])
        if '\n' not in ot + nt and _disp_width(ot) + _disp_width(nt) <= 100:
            print('       %s: %s → %s' % (label, ot if ot != '' else UPD['none'], nt if nt != '' else UPD['none']))
        else:
            print('       %s:' % label)
            for line in upd_linediff(ot, nt):
                print('         ' + line)
        if ch['kind'] in UPD['notes']:
            print('       ' + UPD['notes'][ch['kind']])
    counts = dict((k, len([1 for x in c.changes if x['kind'] == k])) for k in ('追加', '変更', '削除', '競合', '維持'))
    print('----')
    print(fill(UPD['summary'], dict((UPD['count_keys'][k], str(v)) for k, v in counts.items())) +
          fill(UPD['summary_kept'], {'KEPT': str(c.kept), 'DELETED': str(c.kept_deleted)}))
    if dry:
        print(UPD['dry_run'])
    return counts['追加'] + counts['変更'] + counts['削除'] + counts['競合']


def read_base(jmx_path):
    p = jmx_path + UPD['base_suffix']
    if not os.path.isfile(p):
        return None, p
    try:
        obj = json.loads(read_text(p)[0])
        if not isinstance(obj, dict) or not isinstance(obj.get('jmx'), str):
            raise ValueError('jmx がありません')
        return obj, p
    except (ValueError, ToolError) as e:
        warn('基準ファイルを読めないため使いません: %s（%s）' % (p, e))
        return None, p


def write_base(jmx_path, jmx_text, opts):
    meta = OrderedDict([('tool', TOOL['name'] + ' ' + VERSION), ('note', UPD['base_note']),
                        ('input', os.path.basename(opts.get('input') or ''))])
    if opts.get('_oa_rows') is not None:
        meta['seed'] = str(opts.get('_oa_seed'))
        meta['ops'] = OrderedDict((r['name'], OrderedDict([('no', r['no']), ('fp', r['_fp']),
                                                           ('row', OrderedDict((k, r[k]) for k in UPD_ROW_KEYS))]))
                                  for r in opts['_oa_rows'])
    meta['jmx'] = jmx_text
    write_text(jmx_path + UPD['base_suffix'], to_json(meta, 0, True) + '\n')


UPD_ROW_KEYS = [c['key'] for c in SPEC['api_columns'] if c['key'] not in ('no', 'enabled', 'name')]


def oa_fingerprint(ctx, info):
    # OpenAPI の操作の定義（パス共通のパラメータ・参照先のスキーマを含む）と生成の条件の指紋。同じなら同じテストデータになる
    refs = OrderedDict()

    def walk(x):
        if is_obj(x):
            r = x.get('$ref')
            if is_str(r) and r not in refs:
                refs[r] = resolve_ref(ctx, r)
                walk(refs[r])
            for v in x.values():
                walk(v)
        elif is_arr(x):
            for v in x:
                walk(v)
    parts = [get(info['item'], 'parameters'), info['op']]
    walk(parts)
    data = [info['key'], info['path'], info['method'], parts, [[k, refs[k]] for k in sorted(refs)],
            [str(ctx.seed), ctx.required_only, ctx.use_examples]]
    return hashlib.sha256(to_json(data, 0, False).encode('utf-8')).hexdigest()[:16]


def oa_update_rows(records, opts):
    # 更新（-u）: 既存の API（API名が同じもの）は番号を引き継ぎ、定義が変わっていなければ前回のテストデータを使う
    base = opts.get('_upd_base') or {}
    ops = base.get('ops') if isinstance(base.get('ops'), dict) else {}
    ids = opts.get('_upd_ids') or {}
    reused, plan = 0, []
    for rec in records:
        st = ops.get(rec['name']) if isinstance(ops.get(rec['name']), dict) else {}
        if st.get('fp') == rec['_fp'] and isinstance(st.get('row'), dict):
            for k in UPD_ROW_KEYS:
                if isinstance(st['row'].get(k), str):
                    rec[k] = st['row'][k]
            reused += 1
        no = st.get('no') if isinstance(st.get('no'), str) else ''
        plan.append(no or ids.get(rec['name']) or ids.get(re.sub(r'\s*\[[^\]]*\]$', '', rec['name'])) or '')
    known = [int(x) for x in plan if re.match(r'^[0-9]{1,6}$', x)]
    known += [int(v['no']) for v in ops.values() if isinstance(v, dict) and re.match(r'^[0-9]{1,6}$', str(v.get('no')))]
    known += [int(v) for v in ids.values() if re.match(r'^[0-9]{1,6}$', v)]
    nxt = max(known) + 1 if known else 1
    used, kept = set(), 0
    for rec, no in zip(records, plan):
        if re.match(r'^[0-9]{1,6}$', no) and int(no) not in used:
            rec['no'] = str(int(no))
            used.add(int(no))
            kept += 1
        else:
            rec['no'] = ''
    for rec in records:
        if rec['no'] == '':
            while nxt in used:
                nxt += 1
            rec['no'] = str(nxt)
            used.add(nxt)
    info('OpenAPI 定義からの更新: 既存の API %d 件の番号を引き継ぎ、%d 件は前回のテストデータを使います（追加された API: %d 件）'
         % (kept, reused, len(records) - kept))


def upd_prepare(opts):
    path = opts['update']
    if opts.get('split'):
        raise ToolError('--split と -u は同時に指定できません')
    if not os.path.isfile(path):
        raise ToolError('更新する JMX が見つかりません: %s（新しく作る場合は -u を付けずに実行してください）' % path)
    for k, label in (('jmx_name', '-n'), ('out', '-o')):
        if opts.get(k):
            info('%s は -u（更新）では使いません（更新する JMX と同じフォルダ・ファイル名のまま更新します）' % label)
    with open(path, 'rb') as f:
        raw = f.read()
    try:
        text = raw.decode('utf-8')
    except UnicodeDecodeError:
        raise ToolError('JMX を UTF-8 として読めません: ' + path)
    o = te_root(text, path)
    if not [k for k in o['kids'] if k['id'] and k['id'][0] == 'main_tg']:
        raise ToolError('%s は jmeter_api_tool で生成した JMX ではありません（「%s」がありません）' % (path, SPEC['names']['main_tg']))
    base, base_path = read_base(path)
    opts['out'] = os.path.dirname(os.path.abspath(path))
    openapi = os.path.splitext(opts['input'])[1].lower() == '.json'
    if openapi:
        ids = {}

        def walk(te):
            for k in te['kids']:
                if k['id'] and k['id'][0] == 'api' and re.match(r'^API[0-9]{1,6}$', k['id'][1]):
                    num = str(int(k['id'][1][3:]))
                    ids.setdefault(k['id'][2], num)
                    ids.setdefault(re.sub(r'\s*\[[^\]]*\]$', '', k['id'][2]), num)
                walk(k)
        walk(o)
        opts['_upd_ids'], opts['_upd_base'] = ids, base
        if not opts.get('seed'):
            seed, src = str((base or {}).get('seed') or ''), '基準ファイル'
            if not re.match(r'^[0-9]{1,12}$', seed):
                m = re.search(r'乱数シード ([0-9]{1,12})', text)
                seed, src = (m.group(1) if m else ''), 'JMX '
            if seed:
                opts['seed'] = seed
                info('テストデータの乱数シード: %s（%sの値。--seed で変更できます）' % (seed, src))
    return {'path': path, 'raw': raw, 'text': text, 'o': o, 'base': base, 'base_path': base_path, 'openapi': openapi}


def upd_synth(te):
    return {'el': te['el'], 'ht': te['ht'], 'id': te['id'], 'kids': [upd_synth(k) for k in te['kids'] if k['id'] is not None]}


def upd_run(upd, cfg, apis, opts):
    out_dir = opts['out']
    if not opts.get('dry_run'):
        for f in prepare_data(cfg, out_dir, opts.get('force')):
            info('サンプルデータを作成しました: ' + f)
    csv_vars, skip = csv_variables(cfg, out_dir)
    ctx = {'input': os.path.basename(opts['input']), 'date': datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
           'csv_vars': csv_vars, 'csv_skip_header': skip, 'plan_name': cfg['test_plan_name']}
    t_text, hosts = build_jmx(cfg, apis, ctx)
    t = te_root(t_text, '生成した JMX')
    o, b, note = upd['o'], None, ''
    if upd['base'] is not None:
        try:
            b = te_root(upd['base']['jmx'], upd['base_path'])
        except ToolError as e:
            warn(str(e))
        head = lambda te: (upd_text(upd_props(te['el']).get('TestPlan.comments')) or '').split('\n')[0]
        if b is not None and head(b) != head(o):
            warn('基準ファイル（%s）はこの JMX の生成時のものではないため使いません' % upd['base_path'])
            b = None
    if b is not None:
        note = fill(UPD['mode_3way'], {'BASE': upd['base_path']})
    else:
        if upd['base'] is None:
            warn(fill(UPD['no_base'], {'BASE': upd['base_path']}))
        note = UPD['mode_2way']
        b = upd_synth(o)
    c = UpdCtx(upd['text'], upd['openapi'], upd['base'] is None or note == UPD['mode_2way'])
    upd_merge(c, o, b, t, [])
    n = upd_show(c, upd['path'], note, opts.get('dry_run'))
    if opts.get('dry_run'):
        return 0
    if n == 0:
        info('差分はありません（JMX は更新しません）')
        if upd['base'] is None or note == UPD['mode_2way']:
            write_base(upd['path'], t_text, opts)
        return 0
    new_text = upd_apply(upd['text'], c.edits)
    te_root(new_text, '更新後の JMX')
    bak = '%s.%s.bak' % (upd['path'], timestamp())
    write_file(bak, upd['raw'])
    info('更新前の JMX を退避しました: ' + bak)
    write_file(upd['path'], new_text.encode('utf-8'))
    write_base(upd['path'], t_text, opts)
    write_linux_cmd(upd['path'])
    info('JMX を更新しました: %s（%d 箇所）' % (upd['path'], n))
    if WARNINGS:
        print('')
        print('警告が %d 件あります（[WARN] を確認してください）' % len(WARNINGS))
    return 0


def parse_opts(args, valued, flags):
    opts = {}
    i = 0
    while i < len(args):
        a = args[i]
        if a in valued:
            if i + 1 >= len(args):
                raise ToolError('%s には値が必要です' % a)
            opts[valued[a]] = args[i + 1]
            i += 2
        elif a in flags:
            opts[flags[a]] = True
            i += 1
        else:
            raise ToolError('不明なオプションです: %s（./jmeter_api_tool.sh help を参照）' % a)
    return opts


def main(argv):
    if not argv:
        raise ToolError('コマンドを指定してください（init / generate / add-curl / linux-cmd / manual / run / help）')
    cmd, rest = argv[0], argv[1:]
    if cmd == 'init':
        return cmd_init(parse_opts(rest, {'-o': 'out', '--out': 'out'}, {'-f': 'force', '--force': 'force'}))
    if cmd == 'generate':
        return cmd_generate(parse_opts(rest, {'-i': 'input', '--input': 'input', '-s': 'settings', '--settings': 'settings',
                                              '-o': 'out', '--out': 'out', '--seed': 'seed', '--any-method': 'any_method',
                                              '--placeholder-value': 'placeholder_value', '-n': 'jmx_name', '--jmx-name': 'jmx_name',
                                              '-c': 'oa_config', '--openapi-config': 'oa_config', '-u': 'update', '--update': 'update'},
                                       {'--split': 'split', '-f': 'force', '--force': 'force', '--dry-run': 'dry_run',
                                        '--required-only': 'required_only', '--use-examples': 'use_examples'}))
    if cmd == 'manual':
        return cmd_manual(parse_opts(rest, {'-o': 'out', '--out': 'out'}, {}))
    if cmd == 'add-curl':
        k = rest.index('--') if '--' in rest else len(rest)
        return cmd_add_curl(parse_opts(rest[:k], {'-i': 'input', '--input': 'input', '--curl': 'curl', '--curl-file': 'curl_file',
                                                  '--name': 'name'}, {'--dry-run': 'dry_run'}), rest[k + 1:])
    if cmd == 'linux-cmd':
        return cmd_linux_cmd(parse_opts(rest, {'-t': 'jmx', '--jmx': 'jmx', '-o': 'out', '--out': 'out'}, {}))
    if cmd == '_summary':
        return cmd_summary(rest[0])
    if cmd == '_jmeter_version':
        return cmd_jmeter_version(rest[0])
    raise ToolError('不明なコマンドです: %s' % cmd)


if __name__ == '__main__':
    try:
        sys.exit(main(sys.argv[1:]))
    except ToolError as e:
        sys.stderr.write('[ERROR] %s\n' % e)
        sys.exit(1)
    except KeyboardInterrupt:
        sys.stderr.write('[ERROR] 中断しました\n')
        sys.exit(130)
__PYTHON_CORE__
}

cmd_run() {
  local jmx="" threads="" rampup="" loops="" duration="" outbase="result" heap="" jhome="" fail_on_error=0
  local -a props=() extra=()
  while [ $# -gt 0 ]; do
    case "$1" in
      -t|--jmx)          need_value "$@"; jmx="$2"; shift 2 ;;
      --threads)         need_value "$@"; threads="$2"; shift 2 ;;
      --rampup|--ramp-up) need_value "$@"; rampup="$2"; shift 2 ;;
      --loops)           need_value "$@"; loops="$2"; shift 2 ;;
      --duration)        need_value "$@"; duration="$2"; shift 2 ;;
      -J)                need_value "$@"; props+=("$2"); shift 2 ;;
      -J?*)              props+=("${1#-J}"); shift ;;
      -o|--out)          need_value "$@"; outbase="$2"; shift 2 ;;
      --heap)            need_value "$@"; heap="$2"; shift 2 ;;
      --jmeter-home)     need_value "$@"; jhome="$2"; shift 2 ;;
      --fail-on-error)   fail_on_error=1; shift ;;
      --)                shift; extra=("$@"); break ;;
      -h|--help)         usage; return 0 ;;
      *)                 die "不明なオプションです: $1（./jmeter_api_tool.sh help を参照）" ;;
    esac
  done
  [ -n "$jmx" ] || die "実行する JMX を -t で指定してください（例: -t output/api_test.jmx）"
  [ -f "$jmx" ] || die "JMX ファイルが見つかりません: $jmx"
  local v p
  for v in "$threads" "$rampup" "$loops" "$duration"; do
    [ -z "$v" ] || [[ "$v" =~ ^-?[0-9]+$ ]] || die "数値を指定してください: $v"
  done
  for p in "${props[@]+"${props[@]}"}"; do
    [[ "$p" == ?*=* ]] || die "-J は 名前=値 の形式で指定してください: $p"
  done
  if [ -n "$heap" ] && ! [[ "$heap" =~ ^[0-9]+[mMgG]$ ]]; then
    die "--heap は 512m / 2g の形式で指定してください: $heap"
  fi

  local home
  home="$(find_jmeter "$jhome")" || die "JMeter が見つかりません。環境変数 JMETER_HOME を設定するか --jmeter-home で指定してください。"
  info "JMeter: $home"
  check_java
  if find_python; then
    local jver
    jver="$(run_core _jmeter_version "$home/bin/ApacheJMeter.jar" 2>/dev/null)"
    if [ -n "$jver" ]; then
      info "JMeter バージョン: $jver"
      [[ "$jver" == 5.6* ]] || warn "このツールは JMeter 5.6.3 を想定しています（検出: $jver）"
    fi
  fi

  local name ts base rundir jmx_abs
  name="$(basename "$jmx")"
  name="${name%.jmx}"
  ts="$(date +%Y%m%d_%H%M%S)"
  mkdir -p "$outbase" || die "結果フォルダを作成できません: $outbase"
  base="$(cd "$outbase" && pwd)"
  rundir="$base/${name}_${ts}"
  mkdir -p "$rundir" || die "結果フォルダを作成できません: $rundir"
  jmx_abs="$(cd "$(dirname "$jmx")" && pwd)/$(basename "$jmx")"

  local -a args=(-n -t "$jmx_abs" -l "$rundir/result.jtl" -j "$rundir/jmeter.log" -e -o "$rundir/report"
                 "-JRESULT_DIR=$rundir" "-Jjmeter.save.saveservice.subresults=false")
  [ -n "$threads" ] && args+=("-JTHREADS=$threads")
  [ -n "$rampup" ] && args+=("-JRAMP_UP=$rampup")
  [ -n "$loops" ] && args+=("-JLOOPS=$loops")
  [ -n "$duration" ] && args+=("-JDURATION=$duration")
  for p in "${props[@]+"${props[@]}"}"; do
    args+=("-J$p")
  done
  args+=("${extra[@]+"${extra[@]}"}")
  if [ -n "$heap" ]; then
    export HEAP="-Xms$heap -Xmx$heap -XX:MaxMetaspaceSize=256m"
    info "ヒープ: $HEAP"
  fi

  local -a launcher=("$home/bin/jmeter")
  [ -x "$home/bin/jmeter" ] || launcher=(sh "$home/bin/jmeter")
  info "結果フォルダ: $rundir"
  printf '[INFO] 実行コマンド: jmeter'
  printf ' %q' "${args[@]}"
  printf '\n\n'

  "${launcher[@]}" "${args[@]}"
  local rc=$?
  echo

  local summary_rc=0
  if [ -f "$rundir/report/statistics.json" ] && [ -n "$PYTHON_BIN" ]; then
    echo "==== 集計結果（API ごと） ===="
    run_core _summary "$rundir/report/statistics.json"
    summary_rc=$?
    echo
  elif [ -f "$rundir/result.jtl" ]; then
    local total errors
    total=$(( $(wc -l < "$rundir/result.jtl") - 1 ))
    errors=$(awk -F',' 'NR>1 && $8=="false"' "$rundir/result.jtl" | wc -l)
    echo "==== 集計結果 ==== 件数: $total / エラー（概算）: $errors"
    [ "$errors" -eq 0 ] || summary_rc=3
  fi

  echo "結果フォルダ : $rundir"
  [ -f "$rundir/report/index.html" ] && echo "  HTML レポート : $rundir/report/index.html"
  [ -f "$rundir/result.jtl" ] && echo "  JTL           : $rundir/result.jtl"
  [ -f "$rundir/jmeter.log" ] && echo "  JMeter ログ   : $rundir/jmeter.log"
  [ -f "$rundir/api_result.log" ] && echo "  結果ログ      : $rundir/api_result.log"
  if [ -f "$rundir/error_detail.xml" ] && grep -q '<httpSample\|<sample' "$rundir/error_detail.xml"; then
    echo "  エラー詳細    : $rundir/error_detail.xml"
  fi

  if [ "$rc" -ne 0 ]; then
    err "JMeter が異常終了しました（終了コード $rc）。jmeter.log の末尾:"
    [ -f "$rundir/jmeter.log" ] && tail -n 20 "$rundir/jmeter.log" >&2
    exit "$rc"
  fi
  if [ "$summary_rc" -eq 3 ]; then
    warn "エラーのリクエストがあります。HTML レポート・結果ログを確認してください。"
    [ "$fail_on_error" -eq 1 ] && exit 2
  fi
  return 0
}

main() {
  local cmd="${1:-help}"
  [ $# -gt 0 ] && shift
  case "$cmd" in
    init|generate|manual|add-curl|linux-cmd) run_core "$cmd" "$@" ;;
    run)                  cmd_run "$@" ;;
    help|-h|--help)       usage ;;
    version|-v|--version) echo "jmeter_api_tool $TOOL_VERSION" ;;
    *)                    err "不明なコマンドです: $cmd"; echo; usage; return 1 ;;
  esac
}

main "$@"
exit $?
