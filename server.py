from flask import Flask, request, jsonify
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
import re

app = Flask(__name__)

# 🔹 Google Sheets API 認証
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name("credentials.json", scope)
client = gspread.authorize(creds)

# 🔹 使用するスプレッドシート（Google Sheets）
spreadsheet = client.open("勤怠管理票_デモ")  # 📌 ここはあなたのスプレッドシート名
sheet = spreadsheet.sheet1  # 📌 最初のシートを使用

# 🔹 LINE Webhook のエンドポイント
@app.route("/webhook", methods=["POST"])
def webhook():
    data = request.json
    if "events" not in data:
        return "No events", 400

    for event in data["events"]:
        if event["type"] == "message" and event["message"]["type"] == "text":
            user_id = event["source"]["userId"]  # 📌 LINE ID を取得
            message = event["message"]["text"]
            timestamp = datetime.now().strftime("%H:%M:%S")

            # 🔹 メッセージ解析（正規表現で判定）
            if re.match(r".+、出勤", message):
                name = message.split("、")[0]
                update_excel(user_id, name, "氏名", timestamp, "出勤時間")

            elif re.match(r".+、到着", message):
                company = message.split("、")[0]
                update_excel(user_id, company, "得意先名", timestamp, "到着時間")

            elif re.match(r"業務終了、備考(.+)", message):
                note = re.findall(r"備考(.+)", message)[0]
                update_excel(user_id, "業務終了", "業務終了時間", timestamp, "業務終了時間")
                update_excel(user_id, note, "備考", "", "備考")

    return "OK", 200

# 🔹 Google Sheets のデータ更新
def update_excel(user_id, value, col1_name, col1_value, col2_name):
    headers = sheet.row_values(1)  # シートのヘッダー（1行目）

    col1 = headers.index(col1_name) + 1  # 列番号を取得
    col2 = headers.index(col2_name) + 1

    # 🔹 ユーザーごとの行を検索
    records = sheet.get_all_records()
    row = None
    for i, record in enumerate(records, start=2):
        if record.get("LINE ID") == user_id:
            row = i
            break

    # 🔹 新しいユーザーなら追加
    if row is None:
        row = len(records) + 2
        sheet.update_cell(row, 1, user_id)  # LINE ID を記入

    # 🔹 データを記入
    sheet.update_cell(row, col1, value)
    sheet.update_cell(row, col2, col1_value)

if __name__ == "__main__":
    app.run(port=5000)
