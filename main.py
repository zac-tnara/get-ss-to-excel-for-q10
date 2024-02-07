# -*- coding: utf-8 -*-
#! python3

"""main.py

       * get_ss_to_excel_for_q10 用 main

"""

import datetime
import os
import pathlib
import pickle
import re
import sys
import time
from tkinter import messagebox

import gspread
import win32com.client
from dotenv import load_dotenv
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow

from config import config
from package import modules

load_dotenv()

logger = config.logger
logger_f = config.logger_f

price_survey_spreadsheet_id = config.PRICE_SURVEY_SPREADSHEET_ID

##- データ反映するExcel情報
# 見出し行番
h_row = config.HEADER_ROW

# 列番 ※アルファベット
item_cd_col = config.ITEM_CODE_COL
inven_col = config.INVENTRY_COL
cost_col = config.COST_COL
postage_col = config.POSTAGE_COL


def main():
    # Excelマクロよりpython実行テスト用

    logger.info("価格調査SSより商品データ取得の開始")
    # 開始時間を記録
    start_time = datetime.datetime.now()
    # print(f"Start time: {start_time.strftime('%H:%M:%S')}")

    ####################################

    # アクティブなExcelシートを取得する
    xl = win32com.client.Dispatch("Excel.Application")
    if xl.ActiveSheet is None:
        print("Excelが開かれていないため処理を終了します", end="")
        for _ in range(0, 3):
            print(".", end="", flush=True)
            time.sleep(0.5)
        sys.exit()

    # アクティブシート名を取得する
    active_sheet_name = xl.ActiveSheet.Name
    print(f"アクティブシート名: {active_sheet_name}")
    if config.EXCEL_TARGET_SHEET_NAME != active_sheet_name:
        print("対象シートがアクティブでないため処理を終了します", end="")
        for _ in range(0, 3):
            print(".", end="", flush=True)
            time.sleep(0.5)
        sys.exit()

    active_sheet = xl.Worksheets(active_sheet_name)

    print("しばらくお待ちください...")

    ####################################

    # 認証情報とスコープを設定
    SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
    creds = None

    # トークン.pickleが存在する場合は、それを読み込む
    token_path = pathlib.Path(config.GOOGLE_CREDS_FILE_PATH).joinpath(
        "data", "token.pickle"
    )
    if token_path.exists():
        with open(token_path, "rb") as token:
            creds = pickle.load(token)

    # 資格情報が無い場合は、ユーザー認証フローを実行
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            # creds_file_path = pathlib.Path(
            #     "C:\\Users\\t.nara-pc\\develop2\\python\\get_ss_to_excel_for_q10\\"
            # ).joinpath("data", "key", f'{os.getenv("GOOGLE_CREDS")}.json')

            # 通販共有フォルダを指定 ※ras接続でないと使用できない様に
            creds_file_path = pathlib.Path(config.GOOGLE_CREDS_FILE_PATH).joinpath(
                "data", "key", f'{os.getenv("GOOGLE_CREDS")}.json'
            )

            # ファイルパスが存在しない場合はメッセージを表示
            if not creds_file_path.exists():
                # messagebox.showerror("Error", "指定されたファイルパスはありません")
                print("指定されたファイルパスはありません", end="")
                for _ in range(0, 3):
                    print(".", end="", flush=True)
                    time.sleep(0.5)
                sys.exit()

            flow = InstalledAppFlow.from_client_secrets_file(creds_file_path, SCOPES)
            creds = flow.run_local_server(port=0)

        # 次回のために資格情報を保存
        with open(token_path, "wb") as token:
            pickle.dump(creds, token)

    client = gspread.authorize(creds)
    # 取得先シート名リスト
    sheet_list = [
        "MICHELIN",
        "BRIDGESTONE",
        "GOODYEAR",
        "YOKOHAMA",
        "DUNLOP",
        "CONTINENTAL",
        "FALKEN",
    ]
    item_info_dic = modules.get_item_info_from_price_survey_sheet(client, sheet_list)

    # 価格調査シートより取得した商品情報の送料別(税抜き)項目（postage 数式）から送料の値のみ抜き出し、辞書を更新
    # 数式は「=(L5-525)/1.1」の形式、こちらより「525」の部分を取得する
    pattern = re.compile(r"(?<=-)\d{3,}")

    item_info_dic = {
        key: {**value, "postage": match.group(0)}
        for key, value in item_info_dic.items()
        if (match := pattern.search(value.get("postage")))
    }

    # pprint.pprint(item_info_dic)

    ##- 取得したデータを基に、アクティブなエクセルへそれぞれの情報を書き込んでいく
    # Excelのアルファベットカラムを数値カラムに変換する。
    column = modules.excel_column_to_number(item_cd_col)
    # Excel の定数を win32com.client.constants で取得

    # xlUp = win32com.client.constants.xlUp
    xlUp = -4162
    # アクティブシート最下行の取得
    last_row = active_sheet.Cells(active_sheet.Rows.Count, column).End(xlUp).Row
    print(f"Excel{active_sheet_name}シートの最終行: {last_row}")

    # エクセルより、入力のあるセルの範囲を指定して値を取得する
    cell_values = active_sheet.Range(
        f"{item_cd_col}{h_row+1}:{item_cd_col}{last_row}"
    ).Value

    # タプルをリストへ変更 ※要素を調整
    cell_values = [
        str(value[0]).replace(".0", "") if value[0] is not None else ""
        for value in cell_values
    ]

    # シート反映用辞書を作成
    item_info2_dic = {
        i: {
            "product_code": l_v if l_v else "",
            "inventry": item_info_dic[l_v]["inventry"] if l_v else "",
            "cost": item_info_dic[l_v]["cost"] if l_v else "",
            "postage": item_info_dic[l_v]["postage"] if l_v else "",
        }
        for i, l_v in enumerate(cell_values)
    }

    # 列の値のリストを用意する
    inventry_values, cost_values, postage_values = zip(
        *[
            (value["inventry"], value["cost"], value["postage"])
            for value in item_info2_dic.values()
        ]
    )
    inventry_values = [[item] for item in inventry_values]
    cost_values = [[item] for item in cost_values]
    postage_values = [[item] for item in postage_values]

    # 列の範囲値を設定する
    length = len(cell_values) + 2

    print("Excelシートへ値の書き込み中... ※Excelの操作はしないでください")
    range_expression = lambda col: f"{col}{h_row+1}:{col}{length}"
    active_sheet.Range(range_expression(inven_col)).Value = inventry_values
    active_sheet.Range(range_expression(cost_col)).Value = cost_values
    active_sheet.Range(range_expression(postage_col)).Value = postage_values

    # 処理終了時間を各列の見出しへコメント設定
    finish_time = datetime.datetime.now()
    formatted_time = finish_time.strftime("%H:%M:%S")
    comment_text = f"データ更新時刻: {formatted_time}"

    for cell in [f"{inven_col}{h_row}", f"{cost_col}{h_row}", f"{postage_col}{h_row}"]:
        modules.update_comment(active_sheet.Range(cell), comment_text)

    # 対象エクセルの保存
    xl.ActiveWorkbook.Save()
    logger.info("ブックを保存しました。")

    ####################################

    # print(f"Finished time: {finish_time.strftime('%H:%M:%S')}")
    print("処理を終了します。", end="")
    for _ in range(0, 3):
        print(".", end="", flush=True)
        time.sleep(0.5)
    sys.exit()


if __name__ == "__main__":
    main()
