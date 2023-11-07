# 必要なライブラリのインポート
# -----------------------------------------------------

import os
import re
import time
import logging
import requests
import pandas as pd
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from selenium.webdriver.support.ui import Select
from webdriver_manager.chrome import ChromeDriverManager
from oauth2client.service_account import ServiceAccountCredentials
import gspread
import json


# 各処理の関数
# -----------------------------------------------------

# loggerを使用するための関数
def setup_logging():

    logger = logging.getLogger('LoggingTest')    # ログの出力名を設定
    logger.setLevel(logging.INFO)                # ログレベル:INFOを指定

    sh = logging.StreamHandler() # ログのコンソール出力の設定
    logger.addHandler(sh)

    formatter = logging.Formatter('%(asctime)s [%(levelname)s]: %(message)s')  # ログの出力形式の設定
    sh.setFormatter(formatter)

    return logger

# Spreadsheetへ接続するための関数
def connect_spreadsheet(sheet_name):
    with open('config/spreadsheet.json', 'r') as f:
        data = json.load(f)

    # 認証情報を指定してクライアントを作成
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    credentials = ServiceAccountCredentials.from_json_keyfile_name(data["json_file"], scope)
    client = gspread.authorize(credentials)

    if sheet_name == "master":
        spread_sheet_key = data["connect"]
    elif sheet_name == "data":
        spread_sheet_key = data["connect"]
    else:
        raise ValueError("Invalid sheet_name. Use 'master' or 'data'.")

    # シート名に基づいてワークシートを取得
    worksheet = client.open_by_key(spread_sheet_key).worksheet(sheet_name)

    return worksheet

# Spreadsheetへ書き込みするための関数
def update_spreadsheet(worksheet, dataframe):
    # DataFrameのデータを2次元リストとして取得
    rows = dataframe.values.tolist()
    # ヘッダーを含むリストを取得
    data = [dataframe.columns.values.tolist()] + rows

    # スプレッドシートを更新するために一度全てのデータを削除
    worksheet.clear()
    # updateメソッドを使って一括でデータを書き込む
    worksheet.update('A1', data)




# コンフィグ
# -----------------------------------------------------

# ロギングのセットアップを行い、ロガーを取得する
logger = setup_logging()

# 要素を取得するまでの待機時間
timeout = 3


# リスト
# -----------------------------------------------------

# 取得した情報を代入するリスト
keywords = []
filenames = []

# 取得に失敗したASINを格納するリスト
failed_list = []

# DataFrameへ代入するためのリスト
filename_list = []
keyword_list = []
price_list = []
point_list = []
shipping_list = []


# メイン処理
# -----------------------------------------------------

logger.info('データの取得を開始します。')

start_row = 2  # 開始行
worksheet_import = connect_spreadsheet("data")  # "import"シートに接続

# 取得する列のリスト
columns = ['A','B','C','D','E']

# 各列のデータを格納するためのリスト
data = {}

# 指定した列ごとにデータを取得
for col in columns:
    data[col] = worksheet_import.col_values(gspread.utils.a1_to_rowcol(col + '1')[1])[start_row-1:]

# 各列のデータを変数に格納
jan = data['A']
asin = data['B']
price = data['C']
point = data['D']
shipping = data['E']
actual_price = []

# 各リストの要素を一つずつ取り出して合計を計算する

for i in range(len(price)):
    sum = int(price[i]) + int(point[i]) + int(shipping[i])
    actual_price.append(str(sum))


# 商品情報をデータフレームにまとめる
df = pd.DataFrame({
    "JAN": jan,
    "ASIN": asin,
    "売価": price,
    "ポイント": point,
    "送料": shipping,
    "実売価": actual_price
})


# スプレッドシートに接続
worksheet = connect_spreadsheet('data')

# スプレッドシートを更新
update_spreadsheet(worksheet, df)


# # Excelファイルに書き込む
# excel_filename = 'get-amazon-image/export/output.xlsx'
# df.to_excel(excel_filename, index=False)

# if not failed_list:
#     logger.info('%s 件の処理が完了しました。', len(keywords))
# else:
#     success = len(keywords) - len(failed_list)
#     logger.info('%s 件の処理が完了しました。（成功：%s 件 / 失敗：%s 件）', len(keywords), success, len(failed_list))

#     # ファイルに箇条書きで出力する
#     with open('get-amazon-image/export/failed_list.txt', 'w') as file:
#         for jan in failed_list:
#             file.write(f'{jan}\n')
