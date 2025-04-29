import gspread
from google.oauth2.service_account import Credentials
import os
from datetime import datetime

# 【1】Google スプレッドシートの設定
SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/1aBeKQdcxraefE2a0ThaJmpoSMGWAborpiUmOJAkJ1YM/edit?gid=0#gid=0"
SHEET_NAMES = ["日次レポート", "日次レポート (マイナビ)"]

# 【2】Google API 認証情報の設定
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
CREDENTIALS_FILE = "enhance-453402-b143062ab037.json"

def authenticate_google():
    """Google API に認証し、スプレッドシートへ接続"""
    creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES)
    client = gspread.authorize(creds)
    return client

def get_last_filled_row(worksheet):
    """シート内のデータが入っている最終行を取得（下から上にチェック）"""
    data = worksheet.get_all_values()
    if not data:
        return 0  # データなし

    for i in range(len(data) - 1, -1, -1):
        if any(cell.strip() for cell in data[i]):  # どこかに値が入っていたら
            return i + 1  # 1-based index
    return 0  # すべて空白なら0

def get_total_rows(worksheet):
    """スプレッドシートの総行数を取得"""
    return worksheet.row_count

def check_previous_run_date(sheet_name):
    """行追加の前回実行日.txt を読み込み、90日経過しているか確認"""
    last_run_date_file = "行追加の前回実行日.txt"
    
    # 初回実行時の処理
    if not os.path.exists(last_run_date_file):
        current_date = datetime.today().strftime('%Y/%m/%d')
        with open(last_run_date_file, 'w') as f:
            f.write(f"{sheet_name}:{current_date}\n")
        print(f"[INFO] 初回実行日を {current_date} に設定しました。")
        return True  # 初回実行として処理を実行

    # 前回実行日をファイルから読み込む
    with open(last_run_date_file, 'r') as f:
        lines = f.readlines()

    # 各シートの日付を確認
    for line in lines:
        if ":" in line:  # フォーマットチェック
            sheet, last_run_date_str = line.strip().split(":")
            if sheet == sheet_name:
                last_run_date = datetime.strptime(last_run_date_str, '%Y/%m/%d')
                current_date = datetime.today()

                # 以下で90日経過しているかを確認(ここで実行の日を調整する)
                if (current_date - last_run_date).days >= 10:
                    print(f"[INFO] {sheet_name} は10日経過しているため処理を実行します。")
                    return True  # 90日経過しているので処理を実行
                else:
                    print(f"[INFO] {sheet_name} の前回実行日から90日未満のため、処理をスキップします。")
                    return False  # 90日未経過の場合、処理をスキップ
        else:
            print(f"[WARNING] {line.strip()} は不正なフォーマットです。")

    # 日付が記載されていなければ処理を実行
    current_date = datetime.today().strftime('%Y/%m/%d')
    print(f"[INFO] {sheet_name} の前回実行日が記載されていないため処理を実行します。")
    return True

   
def update_last_run_date(sheet_name):
    """行追加の前回実行日.txt を更新"""
    last_run_date_file = "行追加の前回実行日.txt"
    current_date = datetime.today().strftime('%Y/%m/%d')

    # 既存のデータを読み込む
    with open(last_run_date_file, 'r') as f:
        lines = f.readlines()

    # シートの実行日を更新
    with open(last_run_date_file, 'w') as f:
        found = False
        for line in lines:
            sheet, _ = line.strip().split(":")
            if sheet == sheet_name:
                f.write(f"{sheet_name}:{current_date}\n")
                found = True
            else:
                f.write(line)
        
        if not found:
            f.write(f"{sheet_name}:{current_date}\n")
    
    print(f"[INFO] {sheet_name} の行追加の前回実行日を {current_date} に更新しました。")

def add_100_rows_with_format(worksheet, sheet_name):
    """値が入っている最終行の1つ下から100行追加"""
    
    # 90日経過しているか確認
    if not check_previous_run_date(sheet_name):
        return  # 90日未経過の場合、処理をスキップ

    if worksheet is None:
        print(f"[ERROR] {sheet_name}: ワークシートが取得できませんでした。処理をスキップします。")
        return

    last_filled_row = get_last_filled_row(worksheet)  # 値が入っている最終行
    total_rows = get_total_rows(worksheet)  # 現在のスプレッドシートの行数

    print(f"[DEBUG] 取得した最終行: {last_filled_row}")
    
    # 空白行のチェックを削除
    # 100行未満であれば常に100行を追加するように変更

    
    # ワークシートの全データを取得
    try:
        all_data = worksheet.get_all_values()
    except Exception as e:
        print(f"[ERROR] {sheet_name}: データ取得時にエラーが発生: {e}")
        return

    # 最終行の1つ下の行の状態をチェック
    next_row_index = last_filled_row + 1  # ここで next_row_index を定義

    # 次の行がシートの範囲内であることを確認
    if next_row_index <= total_rows:
        next_row_data = all_data[next_row_index - 1] if next_row_index <= len(all_data) else []

        if not any(next_row_data):  # すべて空なら削除
            worksheet.delete_rows(next_row_index)
            print(f"[INFO] {sheet_name} の最終行の1つ下の完全な空行を削除しました。")

    # シートの最大行数を超えないように調整
    if last_filled_row + 102 > total_rows:
        worksheet.add_rows((last_filled_row + 102) - total_rows)
        print(f"[INFO] {sheet_name} のシートに不足分の行を追加しました。")

    # 最終行の1つ下に格子線付きの1行を追加
    grid_row = [[""] * 7]  # A~G列（7列分）の空白行
    worksheet.insert_rows(grid_row, row=last_filled_row + 1)
    print(f"[INFO] {sheet_name} に格子線付きの1行を追加しました。")

    # その下に100行追加
    empty_rows = [[""] * 7 for _ in range(100)]
    worksheet.insert_rows(empty_rows, row=last_filled_row + 2)
    print(f"[INFO] {sheet_name} に 100 行を追加しました。")

    # シートIDを取得
    sheet_id = worksheet._properties['sheetId']
    
    # requestsリストを初期化
    requests = []

    # 関数を適用する範囲を拡張
    for row in range(last_filled_row + 1, last_filled_row + 102):
        requests.append({
            "updateCells": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": row - 1,
                    "endRowIndex": row,
                    "startColumnIndex": 4,
                    "endColumnIndex": 5
                },
                "rows": [{
                    "values": [{"userEnteredValue": {"formulaValue": f"=IFERROR(D{row}/C{row},\"\")"}}]
                }],
                "fields": "userEnteredValue"
            }
        })
        requests.append({
            "updateCells": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": row - 1,
                    "endRowIndex": row,
                    "startColumnIndex": 6,
                    "endColumnIndex": 7
                },
                "rows": [{
                    "values": [{"userEnteredValue": {"formulaValue": f"=IFERROR(F{row}/C{row}*1000,\"\")"}}]
                }],
                "fields": "userEnteredValue"
            }
        })

    # A〜G列のすべてのセルに格子を適用
    requests.append({
        "updateBorders": {
            "range": {
                "sheetId": sheet_id,
                "startRowIndex": last_filled_row,
                "endRowIndex": last_filled_row + 101,
                "startColumnIndex": 0,
                "endColumnIndex": 7
            },
            "top": {"style": "SOLID", "width": 1, "color": {"red": 0, "green": 0, "blue": 0}},
            "bottom": {"style": "SOLID", "width": 1, "color": {"red": 0, "green": 0, "blue": 0}},
            "left": {"style": "SOLID", "width": 1, "color": {"red": 0, "green": 0, "blue": 0}},
            "right": {"style": "SOLID", "width": 1, "color": {"red": 0, "green": 0, "blue": 0}},
            "innerHorizontal": {"style": "SOLID", "width": 1, "color": {"red": 0, "green": 0, "blue": 0}},
            "innerVertical": {"style": "SOLID", "width": 1, "color": {"red": 0, "green": 0, "blue": 0}}
        }
    })

    # リクエストを一括適用
    worksheet.spreadsheet.batch_update({"requests": requests})

    # 行追加の前回実行日を更新
    update_last_run_date(sheet_name)
    
    print(f"[INFO] {sheet_name} の行追加処理が完了しました。")
    current_date = datetime.today().strftime('%Y/%m/%d')
    print(f"[INFO] {sheet_name} の行追加の前回実行日を {current_date} に更新しました。")

def main():
    """メイン処理"""
    print("[INFO] Google スプレッドシートの処理を開始")

    # 調整された処理
    client = authenticate_google()

    # スプレッドシートを開く
    spreadsheet = client.open_by_url(SPREADSHEET_URL)

    # 日次レポートと日次レポート (マイナビ) のシートを処理
    for sheet_name in SHEET_NAMES:
        try:
            worksheet = spreadsheet.worksheet(sheet_name)
            print(f"[INFO] {sheet_name} ワークシートを正常に取得しました。")
            add_100_rows_with_format(worksheet, sheet_name)
        except Exception as e:
            print(f"[ERROR] {sheet_name} のワークシート取得に失敗しました: {e}")

    print("[INFO] すべての処理が完了しました")


if __name__ == "__main__":
    main()
