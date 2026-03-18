import os
import json
import garth
import gspread
import requests
from garminconnect import Garmin
from google.oauth2.service_account import Credentials

GARMIN_HASH = os.environ.get("GARMIN_HASH")
GCP_CREDENTIALS_JSON = os.environ.get("GCP_CREDENTIALS")
LINE_CHANNEL_ACCESS_TOKEN = os.environ.get("LINE_CHANNEL_ACCESS_TOKEN")
LINE_USER_ID = os.environ.get("LINE_USER_ID")
SHEET_NAME = "Garmin_Running_Data"

def send_line_message(message_text):
    """使用 LINE Messaging API 發送推播通知"""
    if not LINE_CHANNEL_ACCESS_TOKEN or not LINE_USER_ID:
        print("未設定 LINE_CHANNEL_ACCESS_TOKEN 或 LINE_USER_ID，略過通知。")
        return
        
    url = "https://api.line.me/v2/bot/message/push"
    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {LINE_CHANNEL_ACCESS_TOKEN}"
    }
    data = {
        "to": LINE_USER_ID,
        "messages": [
            {
                "type": "text",
                "text": message_text
            }
        ]
    }
    try:
        response = requests.post(url, headers=headers, json=data)
        if response.status_code != 200:
            print(f"⚠️ LINE 通知發送失敗: {response.status_code}, {response.text}")
    except Exception as e:
        print(f"⚠️ LINE 通知發送發生錯誤: {e}")

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
        print("☁️ 1. 正在連線至 Google Sheets 取得最後一筆 ID...")
        scopes = [
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive"
        ]
        creds_dict = json.loads(GCP_CREDENTIALS_JSON)
        creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
        client = gspread.authorize(creds)
        sheet = client.open(SHEET_NAME).sheet1
        
        existing_data = sheet.get_all_values()
        latest_sheet_id = 0
        
        # 尋找第一欄 (Activity ID) 的最大值，若無資料或只有表頭則預設為 0
        if len(existing_data) > 1:
            for row in existing_data[1:]:
                if len(row) > 0 and str(row[0]).isdigit():
                    latest_sheet_id = max(latest_sheet_id, int(row[0]))
                    
        print(f"📌 目前 Sheet 中最大的 Activity ID 為: {latest_sheet_id}")

        print("🔄 2. 正在連線至 Garmin...")
        garth.client.loads(GARMIN_HASH)
        garmin_client = Garmin()
        garmin_client.garth = garth.client
        
        print("🔍 3. 正在尋找新的跑步紀錄...")
        activities = garmin_client.get_activities(0, 30)
        new_runs = []
        
        for act in activities:
            if 'running' in act.get('activityType', {}).get('typeKey', '').lower():
                act_id = int(act.get('activityId', 0))
                if act_id > latest_sheet_id:
                    new_runs.append(act)

        if not new_runs:
            print("✅ 目前沒有比 Sheet 內更新的跑步資料，跳過同步。")
            return
            
        print(f"🎉 發現 {len(new_runs)} 筆新跑步紀錄！準備抓取詳細資料...")
        new_runs.reverse()

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

        rows_to_insert = []
        synced_names = [] 
        
        for act in new_runs:
            act_id = act.get('activityId')
            act_name = act.get('activityName')
            print(f"  - 處理中: {act_name} (ID: {act_id})")
            
            splits = garmin_client.get_activity_splits(act_id)
            laps_data = splits.get('lapDTOs', []) if splits else []
            
            if not laps_data:
                print(f"    ⚠️ 此活動沒有逐圈 (Lap) 資料，已略過。")
                continue
                
            synced_names.append(act_name)
                
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

        if rows_to_insert:
            print("☁️ 4. 正在批次寫入至 Google Sheets...")
            if not existing_data:
                sheet.append_row(fieldnames)
                
            sheet.append_rows(rows_to_insert)
            
            success_msg = f"✅ Garmin 同步成功！\n新增了 {len(new_runs)} 筆跑步紀錄：\n" + "、".join(synced_names)
            print(success_msg)
            send_line_message(success_msg)

    except Exception as e:
        error_msg = f"❌ Garmin 腳本執行失敗：\n{e}"
        print(error_msg)
        send_line_message(error_msg)
        raise e

if __name__ == "__main__":
    main()
