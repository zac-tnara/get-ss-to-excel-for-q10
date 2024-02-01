# -*- coding: utf-8 -*-
#! python3
# config.py - 各種設定用

import logging.config
import os

from dotenv import load_dotenv

load_dotenv()


######################

##- デバッグモード設定
# is_debug = True
is_debug = False

######################

# # ログ設定ファイルの読み込み
log_ini_path = ".\\config\\logging.ini"

# ログ設定ファイルの読み込み
logging.config.fileConfig(log_ini_path)

# log_level = "debugLog"
log_level = "infoLog"
# log_level = "baseLog"
logger = logging.getLogger(log_level)
logger_f = logging.getLogger("fileLog")


##-- 環境変数を参照 --##

# 共有フォルダ 各種ファイル格納パス
GOOGLE_CREDS_FILE_PATH = r"\\192.168.11.30\共有\通販共有\__dev\get_ss_to_excel_for_q10"


## Google workspace 設定情報
# 価格調査スプレッドシートID
PRICE_SURVEY_SPREADSHEET_ID = "1RlxwHHLfNR99zP8YqG2m41qabe5C_xKVhrj5vm-wc6k"

## SS情報
# 見出し行番
SS_HEADER_ROW = 4

# 列番
SS_ITEM_CODE_COL = "B"
SS_INVENTRY_COL = "K"
SS_COST_COL = "I"
SS_POSTAGE_COL = "M"


## Excel情報
# データ更新対象Excelシート名
EXCEL_TARGET_SHEET_NAME = "テスト_価格調査"
# EXCEL_TARGET_SHEET_NAME ="価格調査"

# 見出し行番
HEADER_ROW = 2

# 列番
ITEM_CODE_COL = "B"
INVENTRY_COL = "I"
COST_COL = "AS"
POSTAGE_COL = "AT"
