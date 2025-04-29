import os
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials

# 【1】Google スプレッドシートの設定  
SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/1aBeKQdcxraefE2a0ThaJmpoSMGWAborpiUmOJAkJ1YM/edit?gid=0#gid=0"
CSV_FILES = {
    "日次レポート": "enhance_utf8(日次レポート).csv",
    "日次レポート (マイナビ)": "enhance_utf8(日次レポート_マイナビ).csv"
}

# 【2】Google API 認証情報の設定  
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
CREDENTIALS_FILE = "enhance-453402-b143062ab037.json"  # 認証JSONファイル

def authenticate_google():
    """Google API に認証し、スプレッドシートへ接続"""
    try:
        creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES)
        client = gspread.authorize(creds)
        return client
    except Exception as e:
        print(f"[ERROR] Google 認証に失敗: {e}")
        exit(1)

def get_last_row(worksheet):
    """シートの最終行を取得"""
    data = worksheet.get_all_values()
    return len(data) + 1  # 最終行の1つ下

def upload_csv_to_sheet(csv_file, sheet_name, client):
    """CSV を指定のスプレッドシートにアップロード（重複防止対応）"""
    if not os.path.exists(csv_file):
        print(f"[WARNING] {csv_file} が見つかりません。スキップします。")
        return

    print(f"[INFO] {csv_file} を {sheet_name} にアップロード開始")

    # CSV を読み込み（重複排除）
    df = pd.read_csv(csv_file, encoding="utf-8-sig").drop_duplicates()

    # **E列を確保するために空列を追加**
    df["ダミーE列"] = ""

    # **カラム順を確実に維持**
    try:
        df = df[["日付", "広告枠名", "Imp", "Click", "ダミーE列", "ネット"]]
    except KeyError as e:
        print(f"[ERROR] {csv_file} に必要なカラムがありません: {e}")
        return

    # スプレッドシートに接続
    try:
        spreadsheet = client.open_by_url(SPREADSHEET_URL)
        worksheet = spreadsheet.worksheet(sheet_name)
    except Exception as e:
        print(f"[ERROR] スプレッドシートにアクセスできません: {e}")
        return

    # 最終行を取得（既存データ取得）
    last_row = get_last_row(worksheet)
    existing_data = worksheet.get_all_values()
    
    # 直前のデータを取得（過去データと比較）
    if len(existing_data) > 1:  # データがある場合
        last_existing_date = existing_data[-1][0]  # A列の最後の日付
        csv_date = df.iloc[0, 0]  # CSV の最初の行の日付

        if last_existing_date == csv_date:
            print(f"[WARNING] {sheet_name} の最終行が既に {csv_date} であるためスキップ")
            return  # 既にアップロードされている場合はスキップ

    # **データを確実にF列へ配置**
    data = df.values.tolist()

    if not data:
        print(f"[WARNING] {csv_file} にデータがありません。スキップします。")
        return

    # **デバッグ用**
    print("[DEBUG] スプレッドシートへアップロードするデータ:")
    for row in data[:5]:  # 最初の5行だけ表示
        print(row)

    # **アップロード**20250317に格子処理を追加
    try:
        worksheet.append_rows(data, value_input_option='USER_ENTERED')  # 確実にデータを追加
        print(f"[INFO] {sheet_name} にデータをアップロード完了")

        # **格子線を適用する範囲を計算**
        new_data_start = last_row  # **アップロードされた最初の行**
        new_data_end = last_row + len(data) - 1  # **アップロードされた最終行**

        # **格子線をA～G列に適用**
        sheet_id = worksheet._properties['sheetId']
        requests = [{
            "updateBorders": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": new_data_start - 1,  # 0-based index
                    "endRowIndex": new_data_end,  # 最終行の次の行
                    "startColumnIndex": 0,  # A列
                    "endColumnIndex": 7  # G列
                },
                "top": {"style": "SOLID", "width": 1, "color": {"red": 0, "green": 0, "blue": 0}},
                "bottom": {"style": "SOLID", "width": 1, "color": {"red": 0, "green": 0, "blue": 0}},
                "left": {"style": "SOLID", "width": 1, "color": {"red": 0, "green": 0, "blue": 0}},
                "right": {"style": "SOLID", "width": 1, "color": {"red": 0, "green": 0, "blue": 0}},
                "innerHorizontal": {"style": "SOLID", "width": 1, "color": {"red": 0, "green": 0, "blue": 0}},
                "innerVertical": {"style": "SOLID", "width": 1, "color": {"red": 0, "green": 0, "blue": 0}}
            }
        }]

        # **E列とG列に関数を追加（表示形式を指定）**
        for row in range(new_data_start, new_data_end + 1):
            # E列に関数を追加（小数第2位表示）
            requests.append({
                "updateCells": {
                    "range": {
                        "sheetId": sheet_id,
                        "startRowIndex": row - 1,
                        "endRowIndex": row,
                        "startColumnIndex": 4,  # E列
                        "endColumnIndex": 5
                    },
                    "rows": [{
                        "values": [{
                            "userEnteredValue": {
                                "formulaValue": f"=IFERROR(D{row}/C{row},\"\")"
                            },
                            "userEnteredFormat": {
                                "numberFormat": {
                                    "type": "NUMBER",
                                    "pattern": "0.00%"
                                }
                            }
                        }]
                    }],
                    "fields": "userEnteredValue,userEnteredFormat.numberFormat"
                }
            })
            # G列に関数を追加（小数第1位表示）
            requests.append({
                "updateCells": {
                    "range": {
                        "sheetId": sheet_id,
                        "startRowIndex": row - 1,
                        "endRowIndex": row,
                        "startColumnIndex": 6,  # G列
                        "endColumnIndex": 7
                    },
                    "rows": [{
                        "values": [{
                            "userEnteredValue": {
                                "formulaValue": f"=IFERROR(F{row}/C{row}*1000,\"\")"
                            },
                            "userEnteredFormat": {
                                "numberFormat": {
                                    "type": "NUMBER",
                                    "pattern": "\"¥\"#,##0.0"
                                }
                            }
                        }]
                    }],
                    "fields": "userEnteredValue,userEnteredFormat.numberFormat"
                }
            })


        spreadsheet.batch_update({"requests": requests})


        print("[INFO] アップロードした範囲に格子線を適用")

    except gspread.exceptions.APIError as e:
        print(f"[ERROR] スプレッドシートへの書き込みに失敗: {e}")
        return


    # **A列（"日付"）のセルを右寄せに設定**
    try:
        fmt = {
            "horizontalAlignment": "RIGHT",
            "numberFormat": {
                "type": "DATE"
            }
        }
        sheet_id = worksheet._properties['sheetId']
        requests = [{
            "repeatCell": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": last_row - 1,
                    "endRowIndex": last_row - 1 + len(data),
                    "startColumnIndex": 0,
                    "endColumnIndex": 1
                },
                "cell": {
                    "userEnteredFormat": fmt
                },
                "fields": "userEnteredFormat(horizontalAlignment,numberFormat)"
            }
        }]
        spreadsheet.batch_update({"requests": requests})
        print("[INFO] A列の日付を右寄せに設定")
    except Exception as e:
        print(f"[ERROR] A列の日付フォーマット設定に失敗: {e}")

    # **CSV 削除**
    try:
        os.remove(csv_file)
        print(f"[INFO] {csv_file} を削除しました")
    except Exception as e:
        print(f"[ERROR] CSV 削除に失敗: {e}")


def main():
    """メイン処理"""
    print("[INFO] Google スプレッドシートへのアップロード処理を開始")

    # 認証
    client = authenticate_google()

    # 各 CSV を対応するシートにアップロード
    for sheet_name, csv_file in CSV_FILES.items():
        upload_csv_to_sheet(csv_file, sheet_name, client)

    print("[INFO] すべての処理が完了しました")

if __name__ == "__main__":
    main()
