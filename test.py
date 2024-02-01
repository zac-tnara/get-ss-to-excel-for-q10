# -*- coding: utf-8 -*-
#! python3

"""test.py

       * 各種テスト用

"""

import datetime
import os
import pathlib
import pickle
import sys
from tkinter import messagebox

from config import config
from dotenv import load_dotenv
from google.auth.transport.requests import Request

# from google.oauth2.service_account import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

# from googleapiclient.errors import HttpError
# from package import modules

load_dotenv()

logger = config.logger
logger_f = config.logger_f

price_survey_spreadsheet_id = config.PRICE_SURVEY_SPREADSHEET_ID


def main():
    # pythonファイルexe化テスト

    # 開始時間を記録
    start_time = datetime.datetime.now()
    print(f"Start time: {start_time.strftime('%H:%M:%S')}")

    ####################################

    # messagebox.showinfo("実行中です")

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
                messagebox.showerror("Error", "指定されたファイルパスはありません")
                sys.exit()

            flow = InstalledAppFlow.from_client_secrets_file(creds_file_path, SCOPES)
            creds = flow.run_local_server(port=0)

        # 次回のために資格情報を保存
        with open(token_path, "wb") as token:
            pickle.dump(creds, token)

    # 資格情報を使用してサービスを構築
    service = build("sheets", "v4", credentials=creds)

    # スプレッドシートIDと範囲を指定
    SPREADSHEET_ID = price_survey_spreadsheet_id
    RANGE_NAME = "MICHELIN!B5:C6"

    # APIリクエストを実行してデータを取得
    sheet = service.spreadsheets()
    result = (
        sheet.values().get(spreadsheetId=SPREADSHEET_ID, range=RANGE_NAME).execute()
    )
    values = result.get("values", [])

    if not values:
        print("No data found.")
    else:
        print("Name, Major:")
        for row in values:
            # カラム名とデータを印刷
            print(f"{row[0]}, {row[1]}")

    ####################################

    # 終了時間を記録
    finish_time = datetime.datetime.now()
    print(f"Finished time: {finish_time.strftime('%H:%M:%S')}")

    messagebox.showinfo("Info", f"価格調査SSより取得したデータ\n{values}")


if __name__ == "__main__":
    main()
