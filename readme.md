# TimeTree シフト自動取得 → Excel書き込みツール セットアップ手順書

**対象スクリプト:** `timetree_auto.py`

---

## 概要・できること

1. **Chrome の自動操作**：TimeTree へ自動でログインします。
2. **カレンダー選択**：指定したカレンダーを選択します。
3. **予定の取得・解析**：指定タイトルの予定を1件ずつ開き、開始/終了時刻を取得します。
4. **Excel 出力**：労働時間（分）を計算し、最終行に合計時間（h）を付与して既存の Excel ファイルの指定シートに書き込みます。

---

## 動作環境

* **OS**: Windows 10 / 11
* **Python**: 3.9 以上（3.11 / 3.12 推奨）
* **ブラウザ**: Google Chrome（最新版）
* **その他**: インターネット接続環境、書き込み先の Excel ファイル（`.xlsx`）が事前に存在すること

---

## モジュール一覧 (`import`)

### 標準ライブラリ（インストール不要）
```python
import time
from datetime import datetime, time as dtime
import os
import re
import sys
import traceback
```

### 外部ライブラリ（`pip` でのインストールが必要）
```python
import pandas as pd
from openpyxl import load_workbook
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
```

> **必要な pip パッケージ:** `pandas`, `openpyxl`, `selenium`

---

## セットアップ手順（初回のみ）

### 【手順1】 Python をインストールする
1. [Python 公式ダウンロードページ](https://www.python.org/downloads/) から Python をダウンロードします。
2. インストーラーの最初の画面で **「Add python.exe to PATH」** に必ずチェックを入れてから **「Install Now」** をクリックします。
3. コマンドプロンプトを開き、以下のコマンドで動作確認を行います。
   ```cmd
   python --version
   ```
   `Python 3.x.x` と表示されれば完了です。

---

### 【手順2】 Google Chrome をインストールする
既に入っていれば不要です。最新版に更新しておいてください。  
*(Chrome メニュー → ヘルプ → Google Chrome について)*

---

### 【手順3】 必要なライブラリをインストールする
コマンドプロンプト（または PowerShell）で以下を実行します。

```cmd
pip install pandas openpyxl selenium
```

`pip` が見つからない場合はこちらを実行してください：
```cmd
python -m pip install pandas openpyxl selenium
```

#### インストール確認
```cmd
pip list
```
一覧に `pandas`, `openpyxl`, `selenium` が表示されれば OK です。

> **Note:**  
> Selenium 4.6 以降は「Selenium Manager」が ChromeDriver を自動用意するため、`chromedriver` の手動ダウンロードは不要です。（古い Selenium をお使いの場合は `pip install --upgrade selenium` で更新してください）

---

### 【手順4】 スクリプトを配置する
`timetree_shift.py` を任意のフォルダに保存します。  
*例: `C:\Users\Risi\Documents\timetree\timetree_shift.py`*

---

### 【手順5】 スクリプトの設定項目を編集する
`timetree_shift.py` の冒頭にある【設定項目】を環境に合わせて書き換えます。

```python
TARGET_CALENDAR_NAME = ""
# → TimeTree で選択したいカレンダー名 (カレンダー一覧のアイコンの aria-label と完全一致させる)

TARGET_SHIFT_TITLE = ""
# → 抽出したい予定のタイトル (完全一致)

OUTPUT_EXCEL_PATH = r"C:\Users\user\Downloads\時間自動連動_請求書・作業実績.xlsx"
# → 書き込み先の既存 Excel ファイルのパス (パスの前の r は消さないこと)

TARGET_SHEET_NAME = "集計シート"
# → 書き込み先のシート名 (無ければ自動作成されます)

ZOOM = 0.5
# → ブラウザの表示倍率。0.5 = 50%。予定が全部読み取れない場合は 0.4 などに下げてください。
```

---

### 【手順6】 ログイン情報を設定する（2通りから選択）

#### 方法A: 環境変数を使う（共有する場合に推奨）
パスワードがファイルに残らない設定方法です。コマンドプロンプトで以下を実行します。（実行後、新しくコマンドプロンプトを開き直してください）

```cmd
setx TIMETREE_EMAIL "あなたのメールアドレス"
setx TIMETREE_PASSWORD "あなたのパスワード"
```

#### 方法B: スクリプトに直接書く（共有しない場合に推奨）
Windows の環境変数に残したくない場合の方法です。`timetree_shift.py` の以下の行の `""` 内に入力します。

```python
TIMETREE_EMAIL = os.environ.get("TIMETREE_EMAIL", "ここに入力")
TIMETREE_PASSWORD = os.environ.get("TIMETREE_PASSWORD", "ここに入力")
```
> **注意:** この方法を選択した場合、ファイルを他人に渡す・公開する際は必ずパスワードを削除してください。

---

### 【手順7】 書き込み先の Excel ファイルを用意する
- `OUTPUT_EXCEL_PATH` で指定した場所に `.xlsx` ファイルを配置してください。（ファイルが存在しないとエラーになります）
- **実行前に、その Excel ファイルを必ず閉じておいてください。**（開いたままだと警告が出て終了します）

---

## 実行方法

1. 同梱されている `.py` ファイルを実行（開く）します。
2. Chrome が自動で開き、以下の手順で処理が進みます。
   - **[ステップ1/4]** ログイン
   - **[ステップ2/4]** カレンダー選択
   - **[ステップ3/4]** 予定詳細の解析
   - **[ステップ4/4]** Excel への書き込み
3. 完了後、画面に `「エンターキーを押すと終了します」` と表示されます。結果を確認の上、Enter キーを押すとブラウザが閉じます。

> **注意:** 実行中は Chrome を手動で操作しないでください。（カレンダーが自動的に選択されない場合を除く）

---

## Excel への書き込み仕様

| 列 | 項目 |
|---|---|
| **A列** | 日付 |
| **B列** | 予定タイトル |
| **C列** | 開始時間 |
| **D列** | 終了時間 |
| **E列** | 労働時間（分） |
| **最終行** | A列に「合計」、D列に合計時間（例: `42.5h`） |

> **注意事項:**
> - 実行のたびに**シートの A〜E 列は一度クリアされてから上書き**されます。A〜E 列に手入力したデータは消去されるためご注意ください。
> - F列以降や、別シートの内容は変更されません。
> - 取得できるのは TimeTree で表示中の「月」の予定のみです。

---

## トラブルシューティング

| 症状 / エラーメッセージ | 原因と対処法 |
|---|---|
| `\'python\' は、内部コマンドまたは外部コマンド...` | Python インストール時に「Add python.exe to PATH」のチェック漏れです。Python を再インストールしてください。 |
| `ModuleNotFoundError: No module named 'xxx'` | 必要なライブラリが不足しています。【手順3】の `pip install` を実行してください。 |
| Chrome が起動しない / ドライバー関連のエラー | ・Chrome を最新版に更新してください。<br>・`pip install --upgrade selenium` を実行してください。<br>・社内・学内ネットワークの場合はドライバーの自動取得制限を確認してください。 |
| `メールアドレス/パスワードが未設定です` | 【手順6】を実施してください。環境変数を設定した場合は、コマンドプロンプトを開き直す必要があります。 |
| `ファイルが開かれています` | 書き込み先の Excel ファイルが開かれています。ファイルを閉じてから再実行してください。 |
| `ファイルが存在しません` | `OUTPUT_EXCEL_PATH` のパスが正しいか、先頭の `r"..."` の `r` が消えていないか確認してください。 |
| `カレンダー『○○』が見つかりません` | 実行時にログ表示される「検出されたカレンダー」一覧を確認し、`TARGET_CALENDAR_NAME` を完全一致させてください。 |
| 予定が「0件」と表示される / 一部しか取れない | ・`TARGET_SHIFT_TITLE` が完全一致しているか確認してください。<br>・`ZOOM` を `0.4` などに下げて再実行してみてください。<br>・対象カレンダーが既に表示中の場合、クリック操作で非表示化している可能性があります。手動で表示切り替えを行ってから実行してください。 |
| 労働時間が 0 分になる | 日時フォーマットが想定外の可能性があります。ログの「解析成功」行に表示される時刻を確認してください。 |
| ログインで止まる | 2段階認証や CAPTCHA が有効なアカウントでは自動ログインが機能しない場合があります。 |

---

## セキュリティに関する注意

- パスワードをスクリプト内に直接記述した場合、そのファイルをメール・チャット・GitHub 等で共有しないでください。（無断での二次配布は許可されていません）
- 誤ってパスワードを含むコードを共有してしまった場合は、速やかに TimeTree のパスワードを変更してください。
- 他者と共有しない環境であれば、方法B（スクリプト直書き）の利用を推奨します。

---

## 補足・免責事項

- TimeTree の画面構造（HTML）が変更された場合、正常に動作しなくなる可能性があります。その場合はスクリプト内の以下のセレクタを最新の HTML に合わせて更新してください。
  - `compact-calendar-list-item`（カレンダー選択）
  - `event-date-time-start` / `event-date-time-end`（日時取得）
- 本ツールは個人利用を想定しています。TimeTree の利用規約の範囲内でご使用ください。
- 本ツールの権利は作成者に帰属します。**無断での二次配布、転売、および商業目的での利用の一切を禁止します。**
  商用利用する際は、事前に権利者（`fortroottvrc@gmail.com`）へ連絡の上、契約に基づいてご利用ください。
