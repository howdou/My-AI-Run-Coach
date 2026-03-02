import os
import json
import garth
import requests
from garminconnect import Garmin
from google import genai 
from datetime import datetime, timezone, timedelta  # 🌟 新增時間模組

GARMIN_HASH = os.environ.get("GARMIN_HASH")
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY")
DISCORD_WEBHOOK_URL = os.environ.get("DISCORD_WEBHOOK_URL")

LAST_ID_FILE = "last_activity_id.txt"

def send_discord_notify(message):
    chunks = [message[i:i+1900] for i in range(0, len(message), 1900)]
    for chunk in chunks:
        response = requests.post(DISCORD_WEBHOOK_URL, json={"content": chunk})
        if response.status_code not in [200, 204]:
            raise Exception(f"Discord 傳送失敗，錯誤碼: {response.status_code}, 內容: {response.text}")

def main():
    try:
        print("🔄 1. 正在連線至 Garmin...")
        garth.client.loads(GARMIN_HASH)
        garmin_client = Garmin()
        garmin_client.garth = garth.client
        
        last_id = None
        if os.path.exists(LAST_ID_FILE):
            with open(LAST_ID_FILE, "r") as f:
                last_id = f.read().strip()

        print("🔍 2. 正在比對新紀錄...")
        activities = garmin_client.get_activities(0, 200)
        new_records = []
        
        for act in activities:
            if str(act.get('activityId')) == last_id:
                break 
            new_records.append(act)

        if not new_records:
            print("✅ 目前沒有新的運動紀錄。")
            return
        
        print(f"🎉 發現 {len(new_records)} 筆新紀錄！正在執行資料瘦身...")
        payloads = []
        act_names = []
        for act in new_records:
            act_id = act.get('activityId')
            act_names.append(act.get('activityName'))
            summary = garmin_client.get_activity(act_id)
            splits = garmin_client.get_activity_splits(act_id)
            
            slim_act = {
                "name": act.get('activityName'),
                "distance_m": act.get('distance', 0),
                "duration_s": act.get('duration', 0),
                "elevation_gain_m": act.get('elevationGain', 0),
                "avg_hr": summary.get('averageHR') or act.get('averageHR') or summary.get('averageHeartRateInBeatsPerMinute') or 0,
                "max_hr": summary.get('maxHR') or act.get('maxHR') or summary.get('maxHeartRateInBeatsPerMinute') or 0,
                "avg_cadence": summary.get('averageRunningCadenceInStepsPerMinute') or act.get('averageRunningCadenceInStepsPerMinute') or 0,
                "avg_stride_length": summary.get('averageStrideLength') or act.get('averageStrideLength') or 0,
                "avg_vertical_oscillation": summary.get('averageVerticalOscillation') or act.get('averageVerticalOscillation') or 0,
                "avg_ground_contact_time": summary.get('averageGroundContactTime') or act.get('averageGroundContactTime') or 0,
                "laps": [{"distance_m": lap.get('distance', 0), 
                          "duration_s": lap.get('duration', 0), 
                          "avg_hr": lap.get('averageHR') or lap.get('averageHeartRateInBeatsPerMinute') or 0,
                          "avg_cadence": lap.get('averageRunningCadenceInStepsPerMinute') or lap.get('averageRunCadence') or 0,
                          "avg_vertical_oscillation": lap.get('averageVerticalOscillation') or 0,
                          "avg_ground_contact_time": lap.get('averageGroundContactTime') or 0
                         } for lap in splits.get('lapDTOs', [])] if splits else []
            }
            payloads.append(slim_act)
            
        names_str = "、".join(act_names)
        print(f"🧠 3. 正在呼叫 Gemini API 綜合分析 [{names_str}]...")
        ai_client = genai.Client(api_key=GEMINI_API_KEY)
        
        # 🌟 取得台灣時間 (UTC+8) 的當下日期
        tw_tz = timezone(timedelta(hours=8))
        today_str = datetime.now(tw_tz).strftime("%Y年%m月%d日")
        
        # 🧠 提示詞升級：加入今天日期，啟動賽事倒數邏輯！
        prompt = f"""
        今天是 {today_str}。你是一位專業的越野跑與馬拉松教練。這是我最新累積的 {len(new_records)} 筆 Garmin 運動數據：{names_str}。
        考量我 161 cm 的身高，請深入分析我的「高階跑步動態與經濟性」：綜合評估步頻 (avg_cadence)、步距 (avg_stride_length)、垂直震幅 (avg_vertical_oscillation) 與 觸地時間 (avg_ground_contact_time, 單位:毫秒) 的搭配是否流暢，是否有過多能量浪費在上下彈跳或觸地過久。
        
        請根據今天的日期，推算距離我 4 月 12 日的 30km 越野賽（1721m 爬升）及 4 月 26 日的半馬還有多少時間，並給予符合當前訓練週期的步態微調與體能分配建議。
        另外，為了避免高強度賽事引發腸胃不適，請建議好消化、不脹氣的賽中補給，並提供利用鈣、鎂等補充品幫助肌肉放鬆的賽後恢復策略。
        
        ⚠️ 限制：排版適合 Discord 閱讀（多用條列式與 Emoji），總字數盡量控制在 2000 字內。
        數據：{json.dumps(payloads, ensure_ascii=False)}
        """
        
        response = ai_client.models.generate_content(model='gemini-3.1-pro-preview', contents=prompt)
        
        print("📱 4. 正在將報告發送至 Discord...")
        send_discord_notify(f"🏃‍♂️ **AI 教練綜合分析報告 ({today_str})：{names_str}**\n\n{response.text}")
        
        with open(LAST_ID_FILE, "w") as f:
            f.write(str(new_records[0].get('activityId')))
        print("✅ 大功告成！書籤已更新。")

    except Exception as e:
        print(f"❌ AI 教練執行失敗：{e}")
        if "Discord 傳送失敗" not in str(e):
            try:
                send_discord_notify(f"❌ AI 教練執行失敗：{e}")
            except:
                pass

if __name__ == "__main__":
    main()
