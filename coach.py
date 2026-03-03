import os
import json
import garth
import gspread
from garminconnect import Garmin
from google.oauth2.service_account import Credentials

GARMIN_HASH = os.environ.get("GARMIN_HASH")
GCP_CREDENTIALS_JSON = os.environ.get("GCP_CREDENTIALS")
SHEET_NAME = "Garmin_Running_Data"

def format_pace(speed_mps):
    """將公尺/秒轉換為 分:秒/公里 的配速格式"""
    if not speed_mps or speed_mps <= 0:
        return "0:00"
    pace_seconds = 1000 / speed_mps
    minutes = int(pace_seconds // 60)
    seconds = int(pace_seconds % 60)
    return f"{minutes}:{seconds:02d}"

def main():
    try:
        print("🔄 1. 正在連線至 Garmin...")
        garth.client.loads(GARMIN_HASH)
        garmin_client = Garmin()
        garmin_client.garth = garth.client
        
        print("🔍 2. 正在尋找最後一筆跑步紀錄...")
        activities = garmin_client.get_activities(0, 10)
        latest_run = None
        
        for act in activities:
            if 'running' in act.get('activityType', {}).get('typeKey', '').lower():
                latest_run = act
                break

        if not latest_run:
            print("✅ 最近的紀錄中找不到任何跑步資料。")
            return
            
        act_id = latest_run.get('activityId')
        act_name = latest_run.get('activityName')
        
        print(f"🎉 找到最新跑步紀錄: {act_name}")
        
        splits = garmin_client.get_activity_splits(act_id)
        laps_data = splits.get('lapDTOs', []) if splits else []

        # 定義所有要抓取的欄位 Key
        lap_keys = [
            'lapIndex', 'intensityType', 'startTimeGMT', 'distance', 'duration', 'movingDuration', 
            'elapsedDuration', 'elevationGain', 'elevationLoss', 'maxElevation', 'minElevation',
            'averageSpeed', 'averageMovingSpeed', 'maxSpeed', 'calories', 'bmrCalories',
            'averageHR', 'maxHR', 'averageRunCadence', 'maxRunCadence',
            'averageTemperature', 'maxTemperature', 'minTemperature',
            'averagePower', 'maxPower', 'minPower', 'normalizedPower', 'totalWork',
            'groundContactTime', 'groundContactBalanceLeft', 'strideLength',
            'verticalOscillation', 'verticalRatio',
            'maxVerticalSpeed', 'maxRespirationRate', 'avgRespirationRate',
            'directWorkoutComplianceScore', 'avgGradeAdjustedSpeed',
            'stepSpeedLoss', 'stepSpeedLossPercent',
            'startLatitude', 'startLongitude', 'endLatitude', 'endLongitude',
            'wktStepIndex', 'wktIndex', 'messageIndex'
        ]
        
        fieldnames = ["Activity_ID", "Activity_Name", "Lap_Avg_Pace_Formatted", "Lap_GAP_Pace_Formatted"] + lap_keys

        # 準備要寫入 Google Sheets 的資料矩陣 (List of Lists)
        rows_to_insert = []
        if laps_data:
            for lap in laps_data:
                row_list = [
                    str(act_id), # 確保 ID 是字串格式
                    act_name,
                    format_pace(lap.get('averageSpeed', 0)),
                    format_pace(lap.get('avgGradeAdjustedSpeed', 0))
                ]
                for key in lap_keys:
                    row_list.append(lap.get(key, ""))
                rows_to_insert.append(row_list)
        else:
            print("⚠️ 這筆活動沒有逐圈 (Lap) 資料。")
            return

        print("☁️ 3. 正在連線至 Google Sheets 並寫入資料...")
        # 設定 Google API 憑證
        scopes = [
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive"
        ]
        creds_dict = json.loads(GCP_CREDENTIALS_JSON)
        creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
        client = gspread.authorize(creds)
        
        # 開啟試算表與工作表
        sheet = client.open(SHEET_NAME).sheet1
        
        # 讀取現有資料
        existing_data = sheet.get_all_values()
        
        # 檢查是否為空表，如果是，先寫入 Header
        if not existing_data:
            sheet.append_row(fieldnames)
            existing_ids = []
        else:
            # 🟢 防呆機制：確保這一個 row 裡面真的有資料 (長度 > 0)，才去抓 row[0]
            existing_ids = [str(row[0]) for row in existing_data if len(row) > 0]
            
        # 檢查這筆 Activity_ID 是否已經寫入過，避免重複執行時重複寫入
        if str(act_id) in existing_ids:
            print(f"✅ 發現重複：活動 ID {act_id} 已經存在於試算表中，跳過寫入。")
            return

        # 批次寫入所有圈數資料
        sheet.append_rows(rows_to_insert)
        
        print(f"✅ 大功告成！最新資料已成功同步至 Google 試算表 [{SHEET_NAME}]。")

    except Exception as e:
        print(f"❌ 腳本執行失敗：{e}")

if __name__ == "__main__":
    main()
