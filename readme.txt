==============================================================
 TimeTree シフト自動取得 → Excel書き込みツール  セットアップ手順書
 対象スクリプト: timetree_shift.py
==============================================================

■ このツールでできること
--------------------------------------------------------------
 1. Chrome を自動操作して TimeTree にログインする
 2. 指定したカレンダーを選択する
 3. 指定タイトルの予定を1件ずつ開き、開始/終了時刻を取得する
 4. 労働時間(分)を計算し、最終行に合計時間(h)を付けて
    既存のExcelファイルの指定シートに書き込む


■ 動作環境
--------------------------------------------------------------
 - OS            : Windows 10 / 11
 - Python        : 3.9 以上 (3.11 / 3.12 推奨)
 - ブラウザ      : Google Chrome (最新版)
 - インターネット接続
 - 書き込み先の Excel ファイル (.xlsx) が事前に存在すること


■ import 一覧
--------------------------------------------------------------
 【標準ライブラリ】(インストール不要)
   import time
   from datetime import datetime, time as dtime
   import os
   import re
   import sys
   import traceback

 【外部ライブラリ】(pip でインストールが必要)
   import pandas as pd
   from openpyxl import load_workbook
   from selenium import webdriver
   from selenium.webdriver.common.by import By
   from selenium.webdriver.common.keys import Keys
   from selenium.webdriver.chrome.options import Options
   from selenium.webdriver.support.ui import WebDriverWait
   from selenium.webdriver.support import expected_conditions as EC
   from selenium.common.exceptions import TimeoutException

 ※ 必要な pip パッケージは次の3つです。
     pandas / openpyxl / selenium


■ セットアップ手順(初回のみ)
--------------------------------------------------------------

 【手順1】 Python をインストールする
   1) https://www.python.org/downloads/ から Python をダウンロード
   2) インストーラーの最初の画面で
      「Add python.exe to PATH」に必ずチェックを入れてから
      「Install Now」をクリック
   3) コマンドプロンプトを開いて確認
        python --version
      → 「Python 3.x.x」と表示されればOK

 【手順2】 Google Chrome をインストールする
   すでに入っていれば不要です。最新版に更新しておいてください。
   (Chrome メニュー → ヘルプ → Google Chrome について)

 【手順3】 必要なライブラリをインストールする
   コマンドプロンプト(またはPowerShell)で次を実行します。

        pip install pandas openpyxl selenium

   pip が見つからない場合:

        python -m pip install pandas openpyxl selenium

   インストール確認:

        pip list

   → pandas / openpyxl / selenium が表示されればOK

   ※ Selenium 4.6 以降は「Selenium Manager」が ChromeDriver を
     自動で用意するため、chromedriver の手動ダウンロードは不要です。
     (古い Selenium を使っている場合は下記コマンドで更新)

        pip install --upgrade selenium

 【手順4】 スクリプトを配置する
   timetree_shift.py を任意のフォルダに保存します。
   例) C:\Users\Risi\Documents\timetree\timetree_shift.py

 【手順5】 スクリプトの設定項目を編集する
   timetree_shift.py の冒頭にある【設定項目】を書き換えます。

   TARGET_CALENDAR_NAME = ""
       → TimeTree で選択したいカレンダー名
         (カレンダー一覧のアイコンの aria-label と完全一致させる)

   TARGET_SHIFT_TITLE = ""
       → 抽出したい予定のタイトル(完全一致)

   OUTPUT_EXCEL_PATH = r"C:\Users\user\Downloads\時間自動連動_請求書・作業実績.xlsx"
       → 書き込み先の既存Excelファイルのパス
         (パスの前の r は消さないこと)

   TARGET_SHEET_NAME = "集計シート"
       → 書き込み先のシート名 (無ければ自動作成されます)

   ZOOM = 0.5
       → ブラウザの表示倍率。0.5 = 50%。
         予定が全部読み取れない場合は 0.4 などに下げてください。

 【手順6】 ログイン情報を設定する(2通りから選択)

   ■ 方法A: 環境変数を使う(共有する場合に推奨・パスワードがファイルに残らない)
     コマンドプロンプトで次を実行(1回のみ。実行後、
     新しくコマンドプロンプトを開き直すこと)

        setx TIMETREE_EMAIL "あなたのメールアドレス"
        setx TIMETREE_PASSWORD "あなたのパスワード"

   ■ 方法B: スクリプトに直接書く(共有しない場合に推奨・パスワードがWindowsの環境変数に残らない)
     timetree_shift.py の以下の行の "" の中に入力します。

        TIMETREE_EMAIL = os.environ.get("TIMETREE_EMAIL", "ここに入力")
        TIMETREE_PASSWORD = os.environ.get("TIMETREE_PASSWORD", "ここに入力")

     ※ この場合、ファイルを他人に渡す/公開する際は
       必ずパスワードを消してください。

 【手順7】 書き込み先の Excel ファイルを用意する
   - OUTPUT_EXCEL_PATH で指定した場所に .xlsx ファイルを置く
     (ファイルが存在しないとエラーになります)
   - 実行前に、そのExcelファイルを必ず閉じておく
     (開いたままだと警告が出て終了します)


■ 実行方法
--------------------------------------------------------------
 1) 同梱されている.pyファイルを開く

 2) Chrome が自動で開き、次の順に処理が進みます。
      [ステップ1/4] ログイン
      [ステップ2/4] カレンダー選択
      [ステップ3/4] 予定詳細の解析
      [ステップ4/4] Excelへの書き込み

 3) 完了後「エンターキーを押すと終了します」と表示されます。
    結果を確認してから Enter を押すとブラウザが閉じます。

 ※ 実行中は Chrome を手動で操作しないでください。(カレンダーが自動的に選択されない場合を除く)


■ Excel への書き込み内容
--------------------------------------------------------------
 A列: 日付
 B列: 予定タイトル
 C列: 開始時間
 D列: 終了時間
 E列: 労働時間(分)

 最終行: A列に「合計」、D列に合計時間(例: 42.5h)

 ※ 実行のたびに、シートのA〜E列は一度クリアされてから
   上書きされます。A〜E列に手入力したデータは消えるので注意。
 ※ F列以降や他のシートの内容は変更されません。
 ※ 取得できるのは、TimeTree で表示中の「月」の予定のみです。


■ トラブルシューティング
--------------------------------------------------------------
 Q. 「'python' は、内部コマンドまたは外部コマンド...」と出る
 A. Python インストール時に「Add python.exe to PATH」を
    入れていません。Python を再インストールしてください。

 Q. 「ModuleNotFoundError: No module named 'xxx'」と出る
 A. 手順3を実行してください。
        pip install pandas openpyxl selenium

 Q. Chrome が起動しない / ドライバー関連のエラーが出る
 A. ・Chrome を最新版に更新する
    ・pip install --upgrade selenium を実行する
    ・社内/学内ネットワークの場合、ドライバーの自動取得が
      制限されていないか確認する

 Q. 「メールアドレス/パスワードが未設定です」と出る
 A. 手順6を実施してください。環境変数を設定した場合は、
    コマンドプロンプトを開き直してから実行します。

 Q. 「ファイルが開かれています」と出て終了する
 A. 書き込み先のExcelを閉じてから再実行してください。

 Q. 「ファイルが存在しません」と出る
 A. OUTPUT_EXCEL_PATH のパスが正しいか確認してください。
    先頭の r"..." の r が消えていないかも確認。

 Q. 「カレンダー『○○』が見つかりません」と出る
 A. 画面に表示された「検出されたカレンダー」の一覧を確認し、
    TARGET_CALENDAR_NAME をその名前に合わせてください。

 Q. 予定が「0件」と表示される / 一部しか取れない
 A. ・TARGET_SHIFT_TITLE が予定タイトルと完全一致しているか確認
    ・ZOOM を 0.4 など小さくして再実行
    ・カレンダー一覧のボタンは「表示/非表示」の切り替え式の
      可能性があります。対象カレンダーがすでに表示中の場合、
      クリックで非表示になることがあります。
      その場合は、いったん手動で対象カレンダーのみ表示した状態に
      してから実行するか、開発者に相談してください。

 Q. 労働時間が 0 分になる
 A. 日時の表示形式が想定外の可能性があります。
    実行ログの「解析成功」行に表示される開始/終了時間を
    確認してください。

 Q. ログインで止まる (2段階認証・CAPTCHA など)
 A. 自動ログインは、追加認証が有効なアカウントでは
    動作しない場合があります。


■ セキュリティに関する注意
--------------------------------------------------------------
 - パスワードをスクリプトに直接書いた場合、そのファイルを
   メール・チャット・GitHub 等で共有しないでください。
　 そもそも下記の通り無断での二次配布を許可していません。
 - すでにパスワードを含むコードを共有してしまった場合は、
   TimeTree のパスワードを変更することを推奨します。
 - 共有しない前提なら(方法B)の利用を推奨します。


■ 補足
--------------------------------------------------------------
 - TimeTree の画面構成(HTML)が更新されると、要素の指定が
   合わなくなり動作しなくなることがあります。
   その場合はスクリプト内のセレクタ
     ・compact-calendar-list-item (カレンダー選択)
     ・event-date-time-start / event-date-time-end (日時取得)
   を、最新のHTMLに合わせて修正してください。
 - 本ツールは個人利用を想定しています。TimeTree の利用規約の
   範囲内でご利用ください。
 - 本ツールの権利は作成者に帰属します。無断での二次配布、転売、および商業目的での利用の一切を禁止します。
　 商用利用する際には権利者(fortroottvrc@gmail.com)に事前に連絡の上、契約上でのみ利用を許可します。
==============================================================
