import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime, timedelta

# Google スプレッドシートの設定
SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/1aBeKQdcxraefE2a0ThaJmpoSMGWAborpiUmOJAkJ1YM/edit?gid=757156917#gid=757156917"
SHEET_NAMES = ["日次レポート", "日次レポート (マイナビ)"]

# Google API 認証情報
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
CREDENTIALS_FILE = "enhance-453402-b143062ab037.json"

def authenticate_google():
    """Google API に認証し、スプレッドシートへ接続"""
    try:
        creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES)
        client = gspread.authorize(creds)
        print("[INFO] Google API 認証に成功しました")
        return client
    except Exception as e:
        print(f"[ERROR] Google 認証に失敗: {e}")
        exit(1)

def parse_date(value):
    """A列の値を日付オブジェクトに変換（フォーマットの揺れに対応）"""
    if not value:
        return None

    value = value.strip()  # 空白削除

    # **デバッグ用**
    print(f"[DEBUG] A列の日付解析: '{value}'")

    date_formats = ["%Y/%m/%d", "%Y/%-m/%-d", "%m/%d", "%-m/%-d"]
    today_year = datetime.today().year  # 年がない場合の補完

    for fmt in date_formats:
        try:
            parsed_date = datetime.strptime(value, fmt)
            if parsed_date.year == 1900:  # 年がない場合の補完
                parsed_date = parsed_date.replace(year=today_year)
            return parsed_date
        except ValueError:
            continue  

    return None  # 変換できなかった場合

def delete_old_data(worksheet):
    """異常な過去データ（前日より古いデータ）を削除"""
    print(f"[INFO] {worksheet.title} のデータを取得中...")

    # **データ取得**
    data = worksheet.get_all_values()
    if not data or len(data) < 3:
        print("[WARNING] シートにデータがありません")
        return  # ✅ **シートが空なら処理を終了**

    # **処理基準日（昨日）を設定**
    today = datetime.today()
    yesterday_str = (today - timedelta(days=1)).strftime("%Y/%m/%d")
    yesterday_obj = datetime.strptime(yesterday_str, "%Y/%m/%d")

    print(f"[DEBUG] 処理基準日（前日）: {yesterday_str}")

    # **基準日（最も下にある前日の日付）を見つける**
    last_valid_index = None
    for i in range(len(data) - 1, 2, -1):  # **3行目から最終行まで逆順検索**
        row_date = parse_date(data[i][0])
        if row_date and row_date == yesterday_obj:
            last_valid_index = i
            break

    # **基準日が見つからなかった場合**
    if last_valid_index is None:
        print(f"[ERROR] {worksheet.title}: {yesterday_str} のデータが見つかりません")
        return  # ✅ **処理を終了**

    # **基準日が複数行ある場合の最後の行を見つける**
    last_valid_index_end = last_valid_index
    for i in range(last_valid_index + 1, len(data)):
        row_date = parse_date(data[i][0])
        if row_date and row_date == yesterday_obj:
            last_valid_index_end = i  # **基準日の最後の行を更新**
        else:
            break  # **異なる日付になったら終了**

    print(f"[DEBUG] 基準日（{yesterday_str}）の最終行を特定: {last_valid_index_end + 1} 行目")

    # **削除対象の行は、基準日の "最後の行" の1つ下（+1）からではなく、2つ下（+2）から最終行まで**
    rows_to_delete = list(range(last_valid_index_end + 2, len(data)))  # **1-based index**

    if rows_to_delete:
        print(f"[INFO] 削除対象行: {rows_to_delete}")

        # **バッチ処理で削除**
        sheet_id = worksheet._properties['sheetId']
        requests = [
            {"deleteDimension": {"range": {
                "sheetId": sheet_id,
                "dimension": "ROWS",
                "startIndex": row - 1,  # **0-based index**
                "endIndex": row
            }}} for row in sorted(rows_to_delete, reverse=True)
        ]

        worksheet.spreadsheet.batch_update({"requests": requests})
        print(f"[INFO] {len(rows_to_delete)} 行の異常データを削除しました")
    else:
        print("[INFO] 削除対象の異常データはありませんでした")



def main():
    """メイン処理"""
    print("[INFO] スプレッドシートの異常値削除処理を開始")
    client = authenticate_google()
    spreadsheet = client.open_by_url(SPREADSHEET_URL)

    for sheet_name in SHEET_NAMES:
        worksheet = spreadsheet.worksheet(sheet_name)
        delete_old_data(worksheet)

    print("[INFO] スプレッドシートの異常値削除処理が完了しました")

if __name__ == "__main__":
    main()
