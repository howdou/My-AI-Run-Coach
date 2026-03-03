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

        # 定義所有要抓取的欄位 Key (Garmin JSON 原本的 Key)
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
        
        # 🟢 將 Header 轉換為全中文顯示
        fieldnames = [
            "活動 ID", "活動名稱", "單圈平均配速", "單圈平均坡度校正配速",
            "圈數索引 (lapIndex)", "強度類型 (intensityType)", "開始時間 GMT (startTimeGMT)", 
            "距離 公尺 (distance)", "持續時間 秒 (duration)", "移動時間 秒 (movingDuration)", 
            "總經過時間 秒 (elapsedDuration)", "總爬升 公尺 (elevationGain)", "總下降 公尺 (elevationLoss)", 
            "最高海拔 公尺 (maxElevation)", "最低海拔 公尺 (minElevation)",
            "平均速度 m/s (averageSpeed)", "平均移動速度 m/s (averageMovingSpeed)", "最高速度 m/s (maxSpeed)", 
            "卡路里 (calories)", "基礎代謝卡路里 (bmrCalories)",
            "平均心率 (averageHR)", "最高心率 (maxHR)", "平均步頻 (averageRunCadence)", "最高步頻 (maxRunCadence)",
            "平均溫度 (averageTemperature)", "最高溫度 (maxTemperature)", "最低溫度 (minTemperature)",
            "平均功率 (averagePower)", "最大功率 (maxPower)", "最小功率 (minPower)", "標準化功率 (normalizedPower)", "總作功 (totalWork)",
            "觸地時間 ms (groundContactTime)", "觸地時間平衡-左腳 % (groundContactBalanceLeft)", "步幅 cm (strideLength)",
            "垂直震幅 cm (verticalOscillation)", "移動效率/垂直比例 % (verticalRatio)",
            "最大垂直速度 m/s (maxVerticalSpeed)", "最大呼吸率 (maxRespirationRate)", "平均呼吸率 (avgRespirationRate)",
            "訓練符合度分數 (directWorkoutComplianceScore)", "平均坡度校正速度 m/s (avgGradeAdjustedSpeed)",
            "步速損失 (stepSpeedLoss)", "步速損失百分比 (stepSpeedLossPercent)",
            "起點緯度 (startLatitude)", "起點經度 (startLongitude)", "終點緯度 (endLatitude)", "終點經度 (endLongitude)",
            "訓練步驟索引 (wktStepIndex)", "訓練索引 (wktIndex)", "訊息索引 (messageIndex)"
        ]

        # 準備要寫入 Google Sheets 的資料矩陣 (List of Lists)
        rows_to_insert = []
        if laps_data:
            for lap in laps_data:
                row_list = [
                    str(act_id), 
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
        scopes = [
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive"
        ]
        creds_dict = json.loads(GCP_CREDENTIALS_JSON)
        creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
        client = gspread.authorize(creds)
        
        sheet = client.open(SHEET_NAME).sheet1
        
        existing_data = sheet.get_all_values()
        sheet.append_row(fieldnames)
        if not existing_data:
            sheet.append_row(fieldnames)
            existing_ids = []
        else:
            existing_ids = [str(row[0]) for row in existing_data if len(row) > 0]
            
        if str(act_id) in existing_ids:
            print(f"✅ 發現重複：活動 ID {act_id} 已經存在於試算表中，跳過寫入。")
            return

        sheet.append_rows(rows_to_insert)
        
        print(f"✅ 大功告成！最新資料已成功同步至 Google 試算表 [{SHEET_NAME}]。")

    except Exception as e:
        print(f"❌ 腳本執行失敗：{e}")

if __name__ == "__main__":
    main()
