# -*- coding: utf-8 -*-
#! python3

"""modules.py

       * ユーザ定義関数用

"""
from tqdm import tqdm

from config import config

logger = config.logger
logger_f = config.logger_f

## SS情報
# 見出し行番
ss_h_row = config.SS_HEADER_ROW

# 列番 ※アルファベット
ss_item_code_col = config.SS_ITEM_CODE_COL
ss_inven_col = config.SS_INVENTRY_COL
ss_cost_col = config.SS_COST_COL
ss_postage_col = config.SS_POSTAGE_COL


def get_item_info_from_price_survey_sheet(client, sheet_list):
    try:
        ##- 価格調査シート(Y-R-Au)より、auPayのlotno, lotno-set を取得し辞書データにする
        # 通販価格調査データ_Yahoo-Rakuten-Aupay
        spreadsheet_id = config.PRICE_SURVEY_SPREADSHEET_ID
        workbook = client.open_by_key(spreadsheet_id)

        ##- 各シートが存在しているかチェックとデータの取得
        item_info_dic = {}
        existing_worksheet = None  # 判定変数の初期化

        print("価格調査シートよりデータ取得開始")

        # スプレッドシート中のシートオブジェクトを取得
        worksheets = client.open_by_key(spreadsheet_id).worksheets()
        pbar = tqdm(
            sheet_list,
            desc="Get item info from price survey sheet",
            ncols=130,  # 表示全体の幅
            unit="ss",
        )
        # データ取得対象シートの存在チェック
        for sheet_name in pbar:
            for w in worksheets:
                if w.title == sheet_name:
                    existing_worksheet = w
                    break

            postfix = f"{existing_worksheet}"
            # logger.debug(postfix)

            if not existing_worksheet:
                logger.warning(f"{sheet_name}が存在していないため、スキップします")
                continue

            pbar.set_postfix_str(postfix)

            # データの取得
            sheet = workbook.worksheet(sheet_name)

            # 対象シート、B(商品コード)・K(在庫)・I(最終原価)・M(送料別(税抜き))の値を取得し、item_info_dicに格納する
            all_rows_for_fomura = sheet.get_all_values(value_render_option="FORMULA")
            all_rows_for_value = sheet.get_all_values()
            for row_f, row_v in zip(
                all_rows_for_fomura[ss_h_row:], all_rows_for_value[ss_h_row:]
            ):
                item_code_col_num = excel_column_to_number(ss_item_code_col) - 1
                code_value = row_f[item_code_col_num]
                if code_value:
                    code_value = str(code_value).replace(" ", "").replace("　", "")
                    cost_col_num = excel_column_to_number(ss_cost_col) - 1
                    cost_value = (
                        str(row_v[cost_col_num])
                        .replace(" ", "")
                        .replace("　", "")
                        .replace(",", "")
                    )
                    inven_col_num = excel_column_to_number(ss_inven_col) - 1
                    inventry_value = (
                        str(row_v[inven_col_num])
                        .replace(" ", "")
                        .replace("　", "")
                        .replace(",", "")
                    )
                    postage_col_num = excel_column_to_number(ss_postage_col) - 1
                    postage_value = row_f[postage_col_num]
                    if code_value != "":
                        item_info_dic[code_value] = {
                            "cost": cost_value,
                            "inventry": inventry_value,
                            "postage": postage_value,
                        }

        return item_info_dic
    except Exception as e:
        logger.warning(f"エラーが発生しました: {str(e)}")
        return None


def excel_column_to_number(column):
    """
    Excelのアルファベットカラムを数値カラムに変換する。

    Args:
        カラム (str)：エクセルのアルファベット列。

    Returns:
        int: 列番号。

    Example:
        excel_column_to_number('A')  # Output: 1
        excel_column_to_number('Z')  # Output: 26
        excel_column_to_number('AA')  # Output: 27
    """
    result = 0
    for i in range(len(column)):
        result *= 26
        result += ord(column[i]) - ord("A") + 1
    return result


def update_comment(cell, comment_text):
    """
    与えられたcomment_textでセルのコメントを更新する。

    Args:
        cell (Cell): コメントを更新するセル・オブジェクト。
        comment_text (str): コメントとして設定するテキスト。

    Returns:
        None
    """
    comment = cell.Comment
    if comment is not None:
        comment.Delete()
    cell.AddComment(comment_text)
