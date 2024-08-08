# ローカルとドライブでセールピックアップアイテムを取得する
# 必要なライブラリのインポート
import csv
import datetime
from dateutil.relativedelta import relativedelta
from google.oauth2 import service_account
from googleapiclient.discovery import build
import gradio as gr
from gradio_calendar import Calendar
import glob
import gspread
from gspread_formatting import Border, Borders
from gspread_formatting import CellFormat, Color, format_cell_range
import io
import numpy as np
import os
import pandas as pd
import random
import string
import sys

# 定数の定義
DICT_CREDENTIAL = {
    "type": "service_account",
    "project_id": "digital-publishing-386907",
    "private_key_id": "28224218544058a0c55b36" +
                      "e2cd35b20371201599",
    "private_key": "-----BEGIN PRIVATE KEY-----\n" +
                   "MIIEvAIBADANBgkqhkiG9w0BAQEFAASCBKYwggSiAgEAAoIBAQC3+" +
                   "xyLHx+eVT1b\nET3b7eN4XlETmrRvBtnYAT7Zjayn688fP7gbQj/" +
                   "45oQDoo0FNv4IWeup0vV48Z4x\nPcvaa6U32lDufj32yJze/lTUPax" +
                   "D1rVHN3Wdb9dRIiU2StjbJFV8t1ZUPYOZ6Qqc\n" +
                   "zRjS5DNiJUs0aixAt" +
                   "QhTDWvqxcTT9Bo8bI/ZeWSO9bcWz+dtZxnyqEy/p0rV2rPi\n" +
                   "E6hSUuW7Uh/MfdxRxga4Ix02+A7jU2uf8jddmXB7hIWy1i8+" +
                   "TU4FuVmH6e4LLgu0\n" +
                   "xdOF/h2cuj10IENv8+X5m99tAd8uuf3UN" +
                   "so8sj3AvCLkur263Ux80x9BqaZyjLxI\n" +
                   "E2N7CX+tAgMBAAECggEADiyJaWpvbCpS/" +
                   "MYaRuxP3wEdK+56Qid79vk5l1cj2xf+\n" +
                   "sACftXHoLcvMld8bEDDJZ2lOD5pSEQxETLTfFKfAZcoq/" +
                   "AS7z1xrQX7EmElcESnk\n" +
                   "c2UhaYypQPXpegJQLKni8CXLv9exYNUkXSor5GtyTfhjj9k" +
                   "yKZYI2yUokEDGRjHh\n" +
                   "ttvJXZNJogHXBVWzj7Vamt3coedsciPtZ" +
                   "733xE8A1B0PFCvnaKoRhjMs1InNEJav\n" +
                   "Tkcqa16Gxi3Rw1daZ1Fn8gjaqifuDYM5N" +
                   "ttCSyNIjk1vPoPCsSHKn5YgHQgWJga4\n" +
                   "a7D1k+jW9RkgjgdtI0T7GE2CHJleEmGDp" +
                   "cL3sYqGvQKBgQDaUkU5yVmS1y3EnVAA\n" +
                   "aGsHKrc0Fe2KdTT/BOSNlevsXiPxI3QAP" +
                   "0o9jckwQmpKdReVoVBb4tBi0BJ3wRON\n" +
                   "WAsNACVE0/UEQYrpu6PwQhx8XXzeVRxoO" +
                   "rlDKLMBVdxl6j0bsJwfgguNbPlVCAE/\n" +
                   "HEyAZ+DYGHN1FHEAJVbs2ZyirwKBgQDXu" +
                   "6Itr1mCQEzwo/PDqJuewtkJqq+k0TB0\n" +
                   "CA9WpytXG6187ylrE60cQRTdz0FGxL71o" +
                   "qIGWOseXbAEQH8O4In8S/bqnm2tbbLQ\n" +
                   "Vkz6UKEZsBr3loWlUONwZeNbKsIjHgLEg" +
                   "fGB+G/irbQ4IMTk9CV9rjD6CT6TVroH\n" +
                   "UH1qlpBKYwKBgF1irYvPTcpa0o/0fmD+S" +
                   "TGymtTjwEzmX7np3N2XUGg1yIgAE0F7\n" +
                   "0QTNXk6PSin5NhJiAx6awWpS+GNTKkrea" +
                   "zOvaUGsrHSamJHsGm7NyKOF1cDAhTss\n" +
                   "S0yn3xHmKTVK4cKzY8SyesCO6YPuvaHCO" +
                   "BMA3BNzOgfNq5xVXH5Jgw+vAoGAW7V0\n" +
                   "GB+22VwkWRgZhE+k+DS0txtMV7Bl/K2Ad" +
                   "8HQ9tLZSYcSAGb47E3uZOy6Py9cTme4\n" +
                   "oSIjsWD6dpREbzqc7hgM+2gmD9fWcCJ/z" +
                   "tl/4r+udxoR7lkYlqt5n0PqC6uyWX8z\n" +
                   "/6BxT9ewCTxE91+ioG7wexp683+mzX02E" +
                   "5218SkCgYBs/5gxhEOolSsgncalSB3A\n" +
                   "HfPXLDyksw72IxVHCLcjfXW5BqOxLd6J0" +
                   "9ciAcuORrQ+ehI5Dkdck2UzzJDdQg7p\n" +
                   "vWCYNxy9nqz0ybWL6JU7TkTbAP+QQLGcp" +
                   "1IrRC70cSnxKVi2B+mmOUn+2bvNlZ3x\n" +
                   "PbXXUgDWLbKGyBiu2r04dQ==\n-----END PRIVATE KEY-----\n",
    "client_email": "saller-picker@digital-publishing-386907" +
                    ".iam.gserviceaccount.com",
    "client_id": "100885449778972574403",
    "auth_uri": "https://accounts.google.com/o/oauth2/auth",
    "token_uri": "https://oauth2.googleapis.com/token",
    "auth_provider_x509_cert_url": "https://www.googleapis.com/" +
                                   "oauth2/v1/certs",
    "client_x509_cert_url": "https://www.googleapis.com/robot/" +
                            "v1/metadata/x509/" +
                            "saller-picker%40digital-publishing-386907." +
                            "iam.gserviceaccount.com",
    "universe_domain": "googleapis.com"
    }
FILE_ID_TABLE = "1nik5SofcKbJgNRlDqRzWnuQwoSl1ZFUdcWFOKXC9M4c"
DIRECTORY_SHEET_OUT = "1kq3716S0RYLroAYROBLVGWaUnHTPdMd3"
DIRECTORY_ID_MEDIADO_CASH = "1k1wnTeSAngZWi596LFtqEpYlaxvCvE6Z"
DIRECTORY_ID_MEDIADO_PAPER = "1NmCNHYD9TT52slwfrYH05m6vRPzfSO-S"
DIRECTORY_PATH_MEDIADO_CASH = "/content/drive/Shareddrives/"\
    "2. 社外秘-OSR/2.OS/5.価格訴求/2.割引/1.電子/3.自動掃き出し/tsv_メディアドゥ_金額"
DIRECTORY_PATH_MEDIADO_PAPER = "/content/drive/Shareddrives/"\
    "2. 社外秘-OSR/2.OS/5.価格訴求/2.割引/1.電子/3.自動掃き出し/tsv_メディアドゥ_冊数"


# Google APIの認証
def get_credential():
    scopes = [
            'https://www.googleapis.com/auth/spreadsheets',
            'https://www.googleapis.com/auth/drive'
        ]
    obj_credential = service_account.Credentials.\
        from_service_account_info(DICT_CREDENTIAL, scopes=scopes)
    return obj_credential


# ファイル一覧の取得
def get_files_from_directory(obj_credential, str_directory_id,
                             str_file_query: str = ""):
    service = build("drive", "v3", credentials=obj_credential)
    list_item = []
    str_query = f"'{str_directory_id}' in parents and trashed = false "
    if not str_file_query == "":
        str_query += f"and '{str_file_query}' in name"
    while True:
        results = service.files().list(
            q=str_query,
            fields="nextPageToken, files(id, name)"
        ).execute()
        list_item += results.get("files", [])
        next_page_token = results.get("nextPageToken", None)
        if next_page_token is None:
            break
    return list_item


# Drive上のtsvを二次元配列に変換する
def get_list_from_tsv_id(obj_credential, file_id: str):
    # APIクライアントの作成
    service = build("drive", "v3", credentials=obj_credential)
    # ファイルのメタデータを取得
    request = service.files().get_media(fileId=file_id)
    file_content = request.execute()
    # メタデータの読み込み
    tsv_content = file_content.decode("utf-8")
    # 二次元配列に変換
    reader = csv.reader(io.StringIO(tsv_content), delimiter="\t")
    rows = [row for row in reader]
    return rows


# 変換テーブルの取得
def get_table(obj_credential):
    # Google SpreadSheetの認証
    gc = gspread.authorize(credentials=obj_credential)
    # 変換テーブルのあるシートを開く
    workbook = gc.open_by_key(FILE_ID_TABLE)
    worksheet = workbook.worksheet("セール自動化用リスト")
    # シートの内容を取得
    list_value_all = worksheet.get_all_values()
    list_header = list_value_all[0]
    list_content = list_value_all[1:]
    # IDが格納されている列番号の取得
    int_column_index_item = list_header.index("品番")
    int_column_index_20id = list_header.index("２０桁コンテンツID")
    int_column_index_mbj = list_header.index("MBJ商品ID")
    int_column_index_yodobashi = list_header.index("""ヨドバシID
(SKU)""")
    int_column_index_genre = list_header.index("ジャンル")
    int_column_index_bool_sale = list_header.index("割引")
    int_column_conent_name = list_header.index("コンテンツ名")
    int_column_index_price = list_header.index("価格")
    int_column_index_author = list_header.index("著者名")
    # 変換テーブルの定義
    dict_table = {}
    dict_table[list_header[int_column_index_20id]] = {}
    dict_table[list_header[int_column_index_mbj]] = {}
    dict_table[list_header[int_column_index_yodobashi]] = {}
    dict_table[list_header[int_column_index_genre]] = {}
    dict_table[list_header[int_column_index_bool_sale]] = []
    dict_table[list_header[int_column_conent_name]] = {}
    dict_table[list_header[int_column_index_item]] = {}
    dict_table[list_header[int_column_index_price]] = {}
    # 変換テーブルに各アイテムの情報を入れていく
    for row in list_content:
        if True:
            # 品番の取得
            itemcode = str(row[int_column_index_item])
            # 各ストア商品IDの取得
            itemcode_20id = row[int_column_index_20id]
            itemcode_mbj = row[int_column_index_mbj]
            itemcode_yodobashi = row[int_column_index_yodobashi]
            price = row[int_column_index_price]
            author = row[int_column_index_author]
            dict_table["２０桁コンテンツID"][itemcode_20id] = itemcode
            dict_table["MBJ商品ID"][itemcode_mbj] = itemcode
            dict_table["""ヨドバシID
(SKU)"""][itemcode_yodobashi] = itemcode
            dict_table["品番"][itemcode] = {
                list_header[int_column_index_20id]: itemcode_20id,
                list_header[int_column_index_mbj]: itemcode_mbj,
                list_header[int_column_index_yodobashi]: itemcode_yodobashi,
                list_header[int_column_index_price]: price,
                list_header[int_column_index_author]: author
            }
            # ジャンルの取得
            genre = row[int_column_index_genre]
            if genre in dict_table[list_header[int_column_index_genre]]:
                dict_table["ジャンル"][genre].append(itemcode)
            else:
                dict_table["ジャンル"][genre] = [itemcode]
            # 割引の可不可の取得
            bool_sale = str(row[int_column_index_bool_sale])
            if bool_sale == "1":
                dict_table["割引"].append(itemcode)
            dict_table["コンテンツ名"][itemcode] = row[int_column_conent_name]
    return dict_table


# メディアドゥの売上クロス集計を縦持ちのデータフレームに変換する
# ローカルの場合、サービスアカウントからtsvの一覧を取得する
def get_df_mediado_cash(obj_credential, str_tsv_id,
                        dict_table_20id: dict[str:str]):
    # tsvの中身を二次元配列で取得
    list_raw = get_list_from_tsv_id(obj_credential, str_tsv_id)
    # ヘッダーと中身の取得
    list_header = list_raw[0]
    list_content = list_raw[1:]
    # ２０桁コンテンツIDと書店グループ名、セールの有無がどこの列に入っているかを取得
    int_column_index_item20id = list_header.index("２０桁コンテンツID")
    int_column_index_shop_name = list_header.index("書店グループ名")
    int_column_index_campaign = list_header.index("キャンペーン")
    # 日付が何列目から入っているかを取得
    int_column_index_date = 0
    for cell in list_header:
        if "/" in cell:
            break
        else:
            int_column_index_date += 1
    # データフレーム用のリスト
    list_itemcode = []
    list_shop_name = []
    list_campaign = []
    list_sold = []
    list_date = []
    # リストの中身を取得していく
    for row in list_content:
        item20id = row[int_column_index_item20id]
        if item20id in dict_table_20id:
            itemcode = dict_table_20id[item20id]
            shop_name = row[int_column_index_shop_name]
            campaign = row[int_column_index_campaign]
            for cellcnt in range(int_column_index_date, len(row)):
                date = list_header[cellcnt]
                if "/" in date:
                    sold = row[cellcnt]
                    list_itemcode.append(itemcode)
                    list_shop_name.append(shop_name)
                    list_campaign.append(campaign)
                    list_date.append(date)
                    list_sold.append(int(sold))
    # データフレームに変換
    df_cash = pd.DataFrame({
        "商品コード": list_itemcode,
        "書店グループ名": list_shop_name,
        "キャンペーン": list_campaign,
        "売上金額": list_sold,
        "日付": list_date
    })
    return df_cash


def get_df_mediado_paper(obj_credential, str_tsv_id, dict_table_20id):
    list_raw_paper = get_list_from_tsv_id(obj_credential, str_tsv_id)
    list_header_paper = list_raw_paper[0]
    int_column_index_item20id_paper = list_header_paper.index("２０桁コンテンツID")
    int_column_index_shoo_name_paper = list_header_paper.index("書店グループ名")
    for i in range(len(list_header_paper)):
        header = list_header_paper[i]
        if "/" in header:
            int_column_index_datePaper = i
            break
    list_content_paper = list_raw_paper[1:]
    list_itemcode_paper = []
    list_shop_name_paper = []
    list_sold_paper = []
    list_date_paper = []
    for row in list_content_paper:
        item20id = row[int_column_index_item20id_paper]
        if item20id in dict_table_20id:
            itemcode = dict_table_20id[item20id]
            shop_name = row[int_column_index_shoo_name_paper]
            for cellcnt in range(int_column_index_datePaper, len(row)):
                date = list_header_paper[cellcnt]
                if "/" in date:
                    sold = row[cellcnt]
                    list_itemcode_paper.append(itemcode)
                    list_shop_name_paper.append(shop_name)
                    list_date_paper.append(date)
                    list_sold_paper.append(int(sold))
    df_paper = pd.DataFrame(
        {
            "商品コード": list_shop_name_paper,
            "書店グループ名": list_shop_name_paper,
            "売上冊数": list_sold_paper,
            "日付": list_date_paper
        }
    )
    return df_paper


# メディアドゥ売上リストの取得
def get_df_mediado_all(obj_credential, list_tsv_cash,
                       list_tsv_paper, dict_table_20id):
    list_df_cash = []
    for tsv in list_tsv_cash:
        str_tsv_id = tsv["id"]
        df_mediado_cash = get_df_mediado_cash(obj_credential, str_tsv_id,
                                              dict_table_20id)
        list_df_cash.append(df_mediado_cash)
    df_cash = pd.concat(list_df_cash, ignore_index=True)
    list_df_paper = []
    for tsv in list_tsv_paper:
        str_tsv_id = tsv["id"]
        df_mediado_paper = get_df_mediado_paper(obj_credential, str_tsv_id,
                                                dict_table)
        list_df_paper.append(df_mediado_paper)
    df_paper = pd.concat(list_df_paper, ignore_index=True)
    df_out = pd.merge(df_cash, df_paper,
                      on=["商品コード", "書店グループ名", "日付"],
                      how="left")
    df_out["単価"] = df_out["売上金額"] / df_out["売上冊数"]
    for i in range(len(df_out)):
        itemcode = df_out["商品コード"][i]
        price = dict_table["品番"][itemcode]["価格"]
        while "," in price:
            price = price.replace(",", "")
        price = int(price)
        tanka = df_out["単価"][i]
        if tanka == np.inf:
            if i > 1:
                if df_out["商品コード"][i] == df_out["商品コード"][i-1]:
                    df_out["キャンペーン"][i] = df_out["キャンペーン"][i-1]
        elif price > tanka:
            df_out["キャンペーン"][i] = "1"
    return df_out


# グローバル環境の場合、globからtsvの一覧を取得する
# 金額の場合
def get_df_mediado_cash_glob(str_tsv_file, dict_table_20id):
    list_raw = []
    with open(str_tsv_file, "r", encoding="utf_8_sig")as f:
        reader = csv.reader(f, delimiter="\t")
        for row in reader:
            list_raw.append(row)
    # ヘッダーと中身の取得
    list_header = list_raw[0]
    list_content = list_raw[1:]
    # ２０桁コンテンツIDと書店グループ名、セールの有無がどこの列に入っているかを取得
    int_column_index_item20id = list_header.index("２０桁コンテンツID")
    int_column_index_shop_name = list_header.index("書店グループ名")
    int_column_index_campaign = list_header.index("キャンペーン")
    # 日付が何列目から入っているかを取得
    int_column_index_date = 0
    for cell in list_header:
        if "/" in cell:
            break
        else:
            int_column_index_date += 1
    # データフレーム用のリスト
    list_itemcode = []
    list_shop_name = []
    list_campaign = []
    list_sold = []
    list_date = []
    # リストの中身を取得していく
    for row in list_content:
        item20id = row[int_column_index_item20id]
        if item20id in dict_table_20id:
            itemcode = dict_table_20id[item20id]
            shop_name = row[int_column_index_shop_name]
            campaign = row[int_column_index_campaign]
            for cellcnt in range(int_column_index_date, len(row)):
                date = list_header[cellcnt]
                if "/" in date:
                    sold = row[cellcnt]
                    list_itemcode.append(itemcode)
                    list_shop_name.append(shop_name)
                    list_campaign.append(campaign)
                    list_date.append(date)
                    list_sold.append(int(sold))
    # データフレームに変換
    df_cash = pd.DataFrame({
        "商品コード": list_itemcode,
        "書店グループ名": list_shop_name,
        "キャンペーン": list_campaign,
        "売上金額": list_sold,
        "日付": list_date
    })
    return df_cash


# 冊数の場合
def get_df_mediado_paper_glob(str_tsv_file, dict_table_20id):
    list_raw_paper = []
    with open(str_tsv_file, "r", encoding="utf_8_sig")as f:
        reader = csv.reader(f, delimiter="\t")
        for row in reader:
            list_raw_paper.append(row)
    list_header_paper = list_raw_paper[0]
    int_column_index_item20id_paper = list_header_paper.index("２０桁コンテンツID")
    int_column_index_shoo_name_paper = list_header_paper.index("書店グループ名")
    for i in range(len(list_header_paper)):
        header = list_header_paper[i]
        if "/" in header:
            int_column_index_datePaper = i
            break
    list_content_paper = list_raw_paper[1:]
    list_itemcode_paper = []
    list_shop_name_paper = []
    list_sold_paper = []
    list_date_paper = []
    for row in list_content_paper:
        item20id = row[int_column_index_item20id_paper]
        if item20id in dict_table_20id:
            itemcode = dict_table_20id[item20id]
            shop_name = row[int_column_index_shoo_name_paper]
            for cellcnt in range(int_column_index_datePaper, len(row)):
                date = list_header_paper[cellcnt]
                if "/" in date:
                    sold = row[cellcnt]
                    list_itemcode_paper.append(itemcode)
                    list_shop_name_paper.append(shop_name)
                    list_date_paper.append(date)
                    list_sold_paper.append(int(sold))
    df_paper = pd.DataFrame(
        {
            "商品コード": list_shop_name_paper,
            "書店グループ名": list_shop_name_paper,
            "売上冊数": list_sold_paper,
            "日付": list_date_paper
        }
    )
    return df_paper


def get_df_mediado_all_glob(dict_table_20id):
    # 戻れるようにカレントディレクトリを保存しておく
    str_path_current = os.getcwd()
    # 金額の取得
    os.chdir(DIRECTORY_PATH_MEDIADO_CASH)
    list_file_tsv_cash = glob.glob("*.tsv")
    list_df_cash = [get_df_mediado_cash_glob(str_tsv_file, dict_table_20id)
                    for str_tsv_file in list_file_tsv_cash]
    df_cash = pd.concat(list_df_cash, ignore_index=True)
    # 冊数の取得
    os.chdir(DIRECTORY_PATH_MEDIADO_PAPER)
    list_file_tsv_paper = glob.glob("*.tsv")
    list_df_paper = [get_df_mediado_paper_glob(str_tsv_file, dict_table_20id)
                     for str_tsv_file in list_file_tsv_paper]
    df_paper = pd.concat(list_df_paper, ignore_index=True)
    df_out = pd.merge(df_cash, df_paper,
                      on=["商品コード", "書店グループ名", "日付"],
                      how="left")
    df_out["単価"] = df_out["売上金額"] / df_out["売上冊数"]
    for i in range(len(df_out)):
        itemcode = df_out["商品コード"][i]
        price = dict_table["品番"][itemcode]["価格"]
        while "," in price:
            price = price.replace(",", "")
        price = int(price)
        tanka = df_out["単価"][i]
        if tanka == np.inf:
            if i > 1:
                if df_out["商品コード"][i] == df_out["商品コード"][i-1]:
                    df_out["キャンペーン"][i] = df_out["キャンペーン"][i-1]
        elif price > tanka:
            df_out["キャンペーン"][i] = "1"
    os.chdir(str_path_current)
    return df_out


# セール推奨アイテムの一覧を取得する
def pick_itemcode_for_sale(df_log, list_shop_name: list[str],
                           start_date: str, end_date: str,
                           except_code: list[str] = [],
                           choice_code: list[str] = []):
    # セールの開始日と終了日をdatetime型で取得する
    start_date_datetime = datetime.datetime.strptime(start_date, "%Y/%m/%d")
    end_date_datetime = datetime.datetime.strptime(end_date, "%Y/%m/%d")

    # セールに含まないアイテムを除く
    df_log = df_log[~df_log["商品コード"].isin(except_code)]
    str_queryshop_name = '('
    for shop_name in list_shop_name:
        str_queryshop_name += f'"{shop_name}"'
        str_queryshop_name += ','
    str_queryshop_name += ')'
    df_log_shop = df_log.query(f'書店グループ名 in {str_queryshop_name}')

    # 過去の売上一覧から商品コードを取得する
    list_itemcode = set(df_log_shop["商品コード"].tolist())
    if not choice_code == []:
        list_itemcode = choice_code
    list_itemcode_except_all = except_code
    for shop_name in list_shop_name:
        # kindleの場合、セール開始日から過去30日間にセールがあった商品を除外する
        if "kindle" in shop_name:
            date_lastsale_check = start_date_datetime - relativedelta(days=30)
            date_check = date_lastsale_check
            str_date_queried = '['
            while not date_check ==\
                    start_date_datetime:
                strDate = datetime.datetime.\
                    strftime(date_check, format="%Y/%m%d")
                str_date_queried += f'"{strDate}",'
            str_date_queried += ']'
            df_log_shop_last = df_log_shop.query(f'日付 in {str_date_queried}')
            df_log_shop_last_sale = df_log_shop_last.query("キャンペーン == 1")
            list_itemcode_except = set(df_log_shop_last_sale["商品コード"].tolist())
            for itemcode in list_itemcode_except:
                list_itemcode_except_all.append(itemcode)
        # kindle以外の場合過去56日間でセール日数が28日以内のアイテムを取得する
        # 8週間の日付を文字列で取得する
        else:
            span_sale = (end_date_datetime-start_date_datetime).days
            start_date_datetime_ago =\
                start_date_datetime - relativedelta(days=56)
            list_str_date = []
            datetime_date = start_date_datetime_ago
            while not datetime_date == start_date_datetime:
                list_str_date.append(
                    datetime.datetime.strftime(datetime_date, "%Y/%m/%d"))
                datetime_date = datetime_date + relativedelta(days=1)
            # 日付が過去8週間以内になっているショップのデータフレームを取得する
            str_date_queried = '['
            for strDate in list_str_date:
                str_date_queried += f'"{strDate}",'
            str_date_queried += ']'
            df_log_last = df_log_shop.query(f'日付 in {str_date_queried}')
            # セールがあった＝キャンペーンが1になっているログを抽出
            df_log_last_sale = df_log_last.query("キャンペーン == 1")
            # 列の重複を削除するためにmeanをかける
            df_log_last_sale = df_log_last_sale.groupby(
                ["商品コード", "書店グループ名", "日付"]).\
                agg({"売上金額": "sum", "キャンペーン": "mean"})
            # セールの日数を集計する（キャンペーンがある日はキャンペーンに1が入っているのでsumでよい）
            df_log_last_sale_cnt = df_log_last_sale.\
                groupby(["商品コード", "書店グループ名", "日付"]).agg("sum")
            # 4週間=28日以上セールがあるアイテムをピックアップする
            df_log_last_sale_except = df_log_last_sale_cnt.query("キャンペーン>=28")
            df_log_last_sale_except = df_log_last_sale_except.reset_index()
            list_itemcode_except = set(
                df_log_last_sale_except["商品コード"].tolist())
            # セールする日数でも同様に集計する
            end_date_datetime_ago = end_date_datetime - relativedelta(days=56)
            list_str_date_2 = []
            cnt_date = end_date_datetime_ago
            while not cnt_date == end_date_datetime:
                str_date_2 = datetime.datetime.strftime(cnt_date, "%Y/%m/%d")
                list_str_date_2.append(str_date_2)
                cnt_date = cnt_date + relativedelta(days=1)
            str_date_queried_2 = '['
            for strdate in list_str_date_2:
                str_date_queried_2 += f'"{strdate}",'
            str_date_queried_2 += ']'
            df_log_last_2 = df_log.query(f"日付 in {str_date_queried_2}")
            df_log_last_sale_2 = df_log_last_2.query("キャンペーン == 1")
            # 列の重複を削除するためにmeanをかける
            df_log_last_sale_2 = df_log_last_sale_2.\
                groupby(["商品コード", "書店グループ名", "日付"]).\
                agg({"売上金額": "sum", "キャンペーン": "mean"})
            # セールの日数を集計する（キャンペーンがある日はキャンペーンに1が入っているのでsumでよい）
            df_log_last_sale_2_cnt = df_log_last_sale_2.\
                groupby(["商品コード", "書店グループ名", "日付"]).agg("sum")
            # セールの最終日から数えて4週間=28日以上セールにあるアイテムをピックアップする
            df_log_last_sale_2_except = df_log_last_sale_2_cnt.\
                query(f"キャンペーン>={span_sale}")
            df_log_last_sale_2_except = df_log_last_sale_2_except.reset_index()
            list_itemcode_except2 = set(
                df_log_last_sale_2_except["商品コード"].tolist())
            for itemcode in list_itemcode_except:
                list_itemcode_except_all.append(itemcode)
            for itemcode in list_itemcode_except2:
                list_itemcode_except_all.append(itemcode)
    # 上記のアイテムを除いてセールに使えるアイテムの一覧を出す
    list_itemcode_out = []
    for itemCode in list_itemcode:
        if itemCode not in list_itemcode_except_all:
            list_itemcode_out.append(itemCode)
    return list_itemcode_out


# 指定されたアイテムコード群の商品のうち、売上金額が高い順にn個商品をピックアップし、売上金額も出力する
def pick_sales_log(df_log, list_shop_name: list[str],
                   list_item: list[str], item_count: int ,date_span: int, goal: int):
    # データフレームから指定された書店グループ名に入るものをピックアップする
    df_log_shop = df_log[df_log["書店グループ名"].isin(list_shop_name)]
    df_log_shop["日付_datetime"] = pd.to_datetime(df_log_shop["日付"],
                                                format="%Y/%m/%d")
    df_log_shop_item = df_log_shop[df_log_shop["商品コード"].isin(list_item)]
    # データフレームを最新30日間でピックアップする
    datetime_last = max(df_log_shop_item["日付_datetime"])
    datetime_count_first = datetime_last - pd.Timedelta(days=30)
    df_log_shop_item_date = df_log_shop_item[df_log_shop_item["日付_datetime"]
                                             >= datetime_count_first]
    del df_log_shop_item_date["日付_datetime"]
    del df_log_shop_item_date["日付"]
    # データフレームをアイテムごとにgroupbyする
    df_log_agg = df_log_shop_item_date.groupby(["商品コード"]).agg("sum")
    df_log_agg = df_log_agg.reset_index()
    df_log_agg = df_log_agg.sort_values("売上金額", ascending=False).reset_index()
    # データフレームにセールになった場合の売上期間の予測を入れる（仮に日売れが1.2倍になるとしておく）
    df_log_agg["売上金額_予測"] = (df_log_agg["売上金額"] / 30 *date_span * 1.2).round().astype(int)
    if df_log_agg[:item_count]["売上金額"].sum() < goal:
        df_log_out = df_log_agg[:item_count]
    elif len(df_log_agg) > item_count:
        # 上位20%,中位60%,下位20%がバランスよく入るようにする
        top_20_count = int(0.2*item_count)
        mid_60_count = int(0.6 * item_count)
        bottom_20_count = item_count - top_20_count - mid_60_count
        df_log_top = df_log_agg.iloc[:top_20_count]
        df_log_mid = df_log_agg[top_20_count: top_20_count+mid_60_count]
        df_log_bottom = df_log_agg[top_20_count + mid_60_count:]
        df_log_out = pd.concat([
            df_log_top.sample(top_20_count),
            df_log_mid.sample(mid_60_count),
            df_log_bottom.sample(bottom_20_count)
            ])
        
        # 売上金額予測がゴールをぎりぎり超えるように調整
        # ループが長くなりすぎないようにするためのループカウント
        timeout = 0
        while (df_log_out["売上金額_予測"].sum() < goal or df_log_out["売上金額_予測"].sum() > goal*2) and timeout < 20:
            timeout += 1
            if df_log_out["売上金額_予測"].sum() < goal:
                df_remaining = df_log_top[~df_log_top["商品コード"].isin(df_log_out["商品コード"])]
                if len(df_remaining) == 0:
                    df_remaining = df_log_mid[~df_log_mid["商品コード"].isin(df_log_out["商品コード"])]
                    if len(df_remaining) == 0:
                        df_remaining = df_log_bottom[~df_log_bottom["商品コード"].isin(df_log_out["商品コード"])]
            else:
                df_remaining = df_log_mid[~df_log_mid["商品コード"].isin(df_log_out["商品コード"])]
                if len(df_remaining) == 0:
                    df_remaining = df_log_bottom[~df_log_bottom["商品コード"].isin(df_log_out["商品コード"])]
                    if len(df_remaining) == 0:
                        df_remaining = df_log_top[~df_log_top["商品コード"].isin(df_log_out["商品コード"])]
            additional_item = df_remaining.sample(1)
            df_log_out = pd.concat([df_log_out,additional_item])
            df_log_out = df_log_out.sort_values("売上金額",ascending=False, ignore_index=True)
            if len(df_log_out) > item_count:
                if df_log_out["売上金額_予測"].sum() > goal*2:
                    df_log_out = df_log_out.iloc[-item_count:]
                else:
                    df_log_out = df_log_out.iloc[:item_count]
    else:
        df_log_out = df_log_agg
    df_log_out = df_log_out.sort_values("売上金額",ascending=False,ignore_index=True)
    df_log_out = df_log_out.reset_index()
    listout = []
    for i in range(len(df_log_out)):
        itemcode = df_log_out["商品コード"][i]
        soldsum = df_log_out["売上金額"][i]
        soldpredict = df_log_out["売上金額_予測"][i]
        listout.append({
            "itemcode": itemcode,
            "soldsum": soldsum,
            "soldpredict":soldpredict
        })
    return listout


# セールのピックアップ
def sale_picker(df_log, list_shop_name: list[str],
                start_date: str, end_date: str,
                item_cnt: int, dict_table,
                genre: str = "すべて",
                goal: int = 0):
    # 例外リストの取得
    list_code_except = dict_table["割引"]
    if genre == "すべて":
        list_code_chosen = []
    else:
        list_code_chosen = dict_table["ジャンル"][genre]
    list_str_itemcode = pick_itemcode_for_sale(df_log, list_shop_name,
                                               start_date, end_date,
                                               except_code=list_code_except,
                                               choice_code=list_code_chosen)
    int_data_span = (datetime.datetime.strptime(end_date, "%Y/%m/%d") - datetime.datetime.strptime(start_date, "%Y/%m/%d")).days
    print(int_data_span)
    list_sales_log = pick_sales_log(df_log, list_shop_name,
                                    list_str_itemcode, item_cnt,
                                    int_data_span,goal)
    dict_table_content_name = dict_table["コンテンツ名"]
    for int_log_cnt in range(len(list_sales_log)):
        sales_log = list_sales_log[int_log_cnt]
        itemcode = sales_log["itemcode"]
        content_name = dict_table_content_name[itemcode]
        list_sales_log[int_log_cnt]["contentname"] = content_name
    return list_sales_log


# インターフェースの定義
# ジャンル一覧のドロップボックス
def dropdown_genre(dict_table):
    dict_genre = dict_table["ジャンル"]
    list_genre = ["すべて"] + [genre for genre in dict_genre]
    dropdown = gr.Dropdown(
        choices=list_genre,
        multiselect=False,
        value="すべて",
        label="ジャンル"
    )
    return dropdown


# 書店グループ名一覧のドロップボックス
def dropdown_shop_name(df_log):
    df_shop = df_log.copy()
    del df_shop["商品コード"]
    del df_shop["日付"]
    df_shop = df_shop.groupby("書店グループ名").agg("sum").reset_index()
    df_shop = df_shop.sort_values("売上金額", ascending=False)
    list_shop_name = [shop_name for shop_name in
                      set(df_shop["書店グループ名"].tolist())]
    dropdown = gr.Dropdown(
        choices=list_shop_name,
        multiselect=True,
        label="書店グループ名"
    )
    return dropdown


# 検索の機能
def sale_picker_interface(shop_list: list[str], genre: str,
                          start_date, end_date,
                          item_cnt,goal):
    start_date_str = f"{start_date.year}/{start_date.month}/{start_date.day}"
    end_date_str = f"{end_date.year}/{end_date.month}/{end_date.day}"
    list_item = sale_picker(df_mediado_all, shop_list,
                            start_date_str, end_date_str,
                            int(item_cnt), dict_table, genre,goal)
    list_index = []
    list_itemcode = []
    list_content_name = []
    list_item_sold = []
    list_item_sold_predict = []
    list_item20id = []
    list_item_mbj = []
    list_item_yodobashi = []
    for itemcnt in range(len(list_item)):
        item = list_item[itemcnt]
        itemcode = item["itemcode"]
        content_name = item["contentname"]
        soldsum = item["soldsum"]
        sold_predict = item["soldpredict"]
        itemcode_table = dict_table["品番"][itemcode]
        item20id = itemcode_table["２０桁コンテンツID"]
        item_mbj = itemcode_table["MBJ商品ID"]
        item_yodobashi = itemcode_table["""ヨドバシID
(SKU)"""]
        list_index.append(itemcnt + 1)
        list_itemcode.append(itemcode)
        list_content_name.append(content_name)
        list_item_sold.append(soldsum)
        list_item_sold_predict.append(sold_predict)
        list_item20id.append(item20id)
        list_item_mbj.append(item_mbj)
        list_item_yodobashi.append(item_yodobashi)
    df_out = pd.DataFrame(
        {
            "No": list_index,
            "商品コード": list_itemcode,
            "過去30日間売上": list_item_sold,
            "セール期間予測売上":list_item_sold_predict,
            "コンテンツ名": list_content_name,
            "２０桁コンテンツID": list_item20id,
            "MBJID": list_item_mbj,
            "ヨドバシID": list_item_yodobashi
        }
    )
    df_sum = pd.DataFrame(
        {
            "商品点数": [len(list_index)],
            "過去30日間売上": [sum(list_item_sold)],
            "セール期間予測売上": [sum(list_item_sold_predict)]
        } 
    )
    return df_out,df_sum


# データフレームのコンバート機能
def convert_data_frame(df, format: str, start_date, end_date):
    if format == "メディアドゥ":
        list_content_id = df["２０桁コンテンツID"].tolist()
    elif format == "MBJ":
        list_content_id = df["MBJID"].tolist()
    elif format == "ヨドバシ":
        list_content_id = df["ヨドバシID"].tolist()
    datetime_today = datetime.datetime.now(
        datetime.timezone(datetime.timedelta(hours=9)))
    str_datetime = str(datetime_today.year*100000000 +
                       datetime_today.month*1000000 +
                       datetime_today.day * 10000 +
                       datetime_today.hour*100 + datetime_today.minute)
    obj_credential = get_credential()
    gc = gspread.authorize(credentials=obj_credential)
    sh = gc.create(f"【{format}様】セールアイテム一覧_{str_datetime}", DIRECTORY_SHEET_OUT)
    list_content_name = df["コンテンツ名"].tolist()
    list_itemcode = df["商品コード"].tolist()
    dict_table = get_table(obj_credential)
    listout = [[
        "", "＜作品リスト＞", "", "", "", "",
        "", "", "", "", "", "", "", "", "",
        "", "", ""],
        ["", "出版社名", "キャンペーン種別（施策内容）", "TID", "タイトル名", "発行形態",
         "CID", "コンテンツ名", "巻番号", "著者名", "通常価格（税抜）", "キャンペーン価格（税抜）",
         "試し読み増量", "開始日", "終了日", "備考", "コピーライト", "JDCN", "出版社経理コード"]]
    for rowcnt in range(len(list_content_id)):
        content_name = list_content_name[rowcnt]
        contentId = list_content_id[rowcnt]
        itemcode = list_itemcode[rowcnt]
        author = dict_table["品番"][str(itemcode)]["著者名"]
        price = dict_table["品番"][str(itemcode)]["価格"]
        while "," in price:
            price = price.replace(",", "")
        price = int(price)
        half_price = int(price / 2)
        start_date_str = datetime.datetime.strftime(start_date, "%m/%d/%Y")
        end_date_str = datetime.datetime.strftime(end_date, "%m/%d/%Y")
        row = ["", "ディスカヴァー・トゥエンティワン", "割引", "",
               content_name, "書籍", contentId, content_name, "",
               author, price, half_price, "",
               start_date_str, end_date_str, "", "", "", ""]
        listout.append(row)
    ws = sh.sheet1
    ws.update(f"A1:S{str(len(listout)+1)}", listout)
    ws.format("B1", {
        "textFormat": {
            "fontSize": 27
        }
    })
    ws.format("B2:S2", {
        "backgroundColor": {
            "red": 0.31,
            "green": 0.5,
            "blue": 0.74
        },
        "textFormat": {
            "foregroundColor": {
                "red": 1.0,
                "green": 1.0,
                "blue": 1.0
            },
            "fontSize": 14,
            "bold": True
        }
    })
    b = Border("SOLID", Color(0, 0, 0, 0))
    fmt = CellFormat(borders=Borders(b, b, b, b))
    format_cell_range(ws, f"B2:S{len(listout)}", fmt)
    message = f"https://docs.google.com/spreadsheets/d/{sh.id}"
    return message


# パスワード生成用の関数
def generate_random_pass(n):
    randlst = [random.choice(string.ascii_letters + string.digits)
               for i in range(n)]
    return ''.join(randlst)


if __name__ == "__main__":
    # コマンドラインアーギュメントの取得
    argv = sys.argv
    share = argv[1] == "--share"
    inbrowser = argv[2] == "--inbrowser"
    admin = argv[3] == "--admin"
    env = argv[4]
    if env not in ["--local", "--global"]:
        env = "--local"
    print("Google Driveに接続しています...")
    obj_credential = get_credential()
    print("商品リストを取得しています...")
    dict_table = get_table(obj_credential)
    dict_table_20id = dict_table["２０桁コンテンツID"]
    print("売上データを取得しています...")
    if env == "--local":
        list_tsv_cash = get_files_from_directory(obj_credential,
                                                 DIRECTORY_ID_MEDIADO_CASH)
        list_tsv_paper = get_files_from_directory(obj_credential,
                                                  DIRECTORY_ID_MEDIADO_PAPER)

        df_mediado_all = get_df_mediado_all(obj_credential,
                                            list_tsv_cash, list_tsv_paper,
                                            dict_table_20id)
    elif env == "--global":
        df_mediado_all = get_df_mediado_all_glob(dict_table_20id)
    dfAll = df_mediado_all.copy()
    # 全体のインターフェースの定義
    with gr.Blocks() as interface:
        # コンポーネント
        with gr.Row():
            dropdown_genre_box = dropdown_genre(dict_table)
            dropdown_shop_name_box = dropdown_shop_name(df_mediado_all)
        with gr.Row():
            calendar_start_date = Calendar(type="datetime", label="開始日付")
            calendar_end_date = Calendar(type="datetime", label="終了日付")
        with gr.Row():
            input_item_cnt = gr.Number(label="アイテム点数")
            input_item_goal = gr.Number(label="売上金額ゴール")
        button_sale_pick = gr.Button(value="セールアイテムを出力")
        data_frame_sum = gr.DataFrame(
            headers=["商品点数", "過去30日間売上", "セール期間予測売上"],
            label="合計")
        data_frame_items = gr.DataFrame(
            headers=["No", "商品コード", "過去30日間売上", "セール期間予測売上",
                     "コンテンツ名",
                     "２０桁コンテンツID", "MBJID", "ヨドバシID"],
            label="セール推奨アイテム")
        
        with gr.Row():
            dropdown_format = gr.Dropdown(
                choices=["メディアドゥ", "MBJ", "ヨドバシ"],
                multiselect=False,
                label="出力フォーマット")
            button_output = gr.Button(value="スプレッドシートにして出力")
        url_output = gr.Text(label="出力URL")
        # イベントリスナー
        button_sale_pick.click(fn=sale_picker_interface,
                               inputs=[dropdown_shop_name_box,
                                       dropdown_genre_box,
                                       calendar_start_date, calendar_end_date,
                                       input_item_cnt,input_item_goal],
                               outputs=[data_frame_items,data_frame_sum],
                               api_name="sale_picker")
        button_output.click(fn=convert_data_frame,
                            inputs=[data_frame_items, dropdown_format,
                                    calendar_start_date, calendar_end_date],
                            outputs=url_output)

    if admin:
        randompass = generate_random_pass(10)
        print("user:discover")
        print(f"pass:{randompass}")
        interface.launch(share=share, inbrowser=inbrowser,
                         auth=("discover", randompass))
    else:
        interface.launch(share=share, inbrowser=inbrowser)
