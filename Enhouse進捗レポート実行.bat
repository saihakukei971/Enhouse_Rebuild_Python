@echo off
cd /d %~dp0

:: Python スクリプトの実行 - 広告枠ID取得とCSV出力
echo [INFO] Enhouse_01_広告枠ID取得とCSV出力.py を実行中...
python Enhouse_01_広告枠ID取得とCSV出力.py
echo [INFO] Enhouse_01_広告枠ID取得とCSV出力.py 完了

:: Python スクリプトの実行 - CSVデータをスプレッドシートにアップロード
echo [INFO] Enhouse_02_CSVデータをスプレッドシートにアップロード.py を実行中...
python Enhouse_02_CSVデータをスプレッドシートにアップロード.py
echo [INFO] Enhouse_02_CSVデータをスプレッドシートにアップロード.py 完了

:: 3分（180秒）の待機処理（アップロード直後の異常データ追加を待つ）
echo [INFO] 3分待機中（異常データが追加されないか確認）
timeout /t 180
echo [INFO] 3分待機完了

:: Python スクリプトの実行 - 異常値削除処理
echo [INFO] Enhouse_03_異常値削除.py を実行中...
python Enhouse_03_異常値削除.py
echo [INFO] Enhouse_03_異常値削除.py 完了

:: 3分（180秒）の待機処理（削除後の異常データ追加を待つ）
echo [INFO] 3分待機中（異常データが追加されないか再確認）
timeout /t 180
echo [INFO] 3分待機完了

:: Python スクリプトの実行 - 行と関数の自動追加処理
echo [INFO] Enhouse_04_行と関数の自動追加.py を実行中...
python Enhouse_04_行と関数の自動追加.py
echo [INFO] Enhouse_04_行と関数の自動追加.py 完了

exit
