import os
import json
import time
import threading
from dotenv import load_dotenv
from flask import Flask, jsonify, request
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import requests

# === KONFIGURASI ===
load_dotenv()
SPREADSHEET_ID = '1d07pl6fcepVqpaaH3wZHKvq8cedUVoIzev9ci8I_GxM'
TELEGRAM_BOT_TOKEN = os.environ["TELEGRAM_BOT_TOKEN"]

# === AUTENTIKASI GOOGLE SHEETS ===
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
credentials_info = json.loads(os.environ["GOOGLE_CREDS_JSON"])
credentials_info["private_key"] = credentials_info["private_key"].replace("\\n", "\n")
creds = ServiceAccountCredentials.from_json_keyfile_dict(credentials_info, scope)
client = gspread.authorize(creds)

# === INISIALISASI FLASK ===
app = Flask(__name__)

# === GET NAMA SHEET SEBENARNYA ===
def get_real_sheet_name(requested_name):
    try:
        spreadsheet = client.open_by_key(SPREADSHEET_ID)
        for sheet in spreadsheet.worksheets():
            if sheet.title.lower() == requested_name.lower():
                return sheet.title
    except Exception as e:
        print("❌ Gagal mendapatkan daftar sheet:", e)
    return None

# === AMBIL DATAFRAME ===
def get_dataframe(sheet_name):
    try:
        sheet = client.open_by_key(SPREADSHEET_ID).worksheet(sheet_name)
        data = sheet.get_all_values()
        headers = data[0]
        rows = data[1:]
        return pd.DataFrame(rows, columns=headers)
    except Exception as e:
        print(f"❌ Gagal baca sheet '{sheet_name}':", e)
        return None

# === SIMPAN KE EXCEL ===
def dataframe_to_excel(df, filename):
    try:
        df.to_excel(filename, index=False)
    except Exception as e:
        print(f"❌ Gagal simpan Excel: {e}")

# === SIMPAN KE GAMBAR ===
def dataframe_to_image(df, filename):
    try:
        height = max(len(df) * 0.5, 2)
        width = max(len(df.columns) * 1.5, 4)
        fig, ax = plt.subplots(figsize=(width, height))
        ax.axis('off')
        table = ax.table(cellText=df.values, colLabels=df.columns, loc='center', cellLoc='center')
        table.auto_set_font_size(False)
        table.set_fontsize(10)
        fig.tight_layout()
        plt.savefig(filename, dpi=200)
        plt.close()
    except Exception as e:
        print(f"❌ Gagal simpan gambar: {e}")

# === KIRIM FILE KE TELEGRAM ===
def send_telegram_file(chat_id, file_path, file_type='photo'):
    if not os.path.exists(file_path) or os.path.getsize(file_path) == 0:
        print(f"❌ File tidak valid: {file_path}")
        return
    url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendPhoto" if file_type == 'photo' \
        else f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendDocument"
    with open(file_path, 'rb') as f:
        file_data = {'photo' if file_type == 'photo' else 'document': f}
        response = requests.post(url, data={'chat_id': chat_id}, files=file_data)
    print(f"📤 Kirim ke {chat_id} | Status: {response.status_code}")

# === KIRIM PESAN TEKS KE TELEGRAM ===
def send_telegram_message(chat_id, message):
    url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendMessage"
    response = requests.post(url, data={
        "chat_id": chat_id,
        "text": message
    })
    print(f"📝 Kirim pesan ke {chat_id} | Status: {response.status_code}")

# === ENDPOINT FLASK UNTUK HANDLE REQUEST ===
@app.route('/<nama_sheet>', methods=['GET'])
def kirim_data_sheet(nama_sheet):
    chat_id = request.args.get("chat_id")
    from_bot = request.args.get("from_bot", "false").lower() == "true"
    real_name = get_real_sheet_name(nama_sheet)

    if not real_name:
        if from_bot and chat_id:
            send_telegram_message(chat_id, "❌ Sheet tidak tersedia.")
        return jsonify({"status": "error", "message": "Sheet tidak tersedia."}), 404

    df = get_dataframe(real_name)
    if df is None or df.empty:
        if from_bot and chat_id:
            send_telegram_message(chat_id, f"❌ Sheet '{real_name}' kosong atau gagal dibaca.")
        return jsonify({"status": "error", "message": f"Sheet '{real_name}' kosong atau gagal dibaca"}), 500

    if from_bot and chat_id:
        if len(df) > 25 or len(df.columns) > 25:
            file_path = f"{real_name}.xlsx"
            dataframe_to_excel(df, file_path)
            send_telegram_file(chat_id, file_path, file_type='document')
        else:
            file_path = f"{real_name}.png"
            dataframe_to_image(df, file_path)
            send_telegram_file(chat_id, file_path, file_type='photo')

    return jsonify({"status": "success", "message": f"Data dari sheet '{real_name}' berhasil dikirim"})

# === POLLING TELEGRAM BOT ===
def polling_bot():
    print("🤖 Memulai polling Telegram...")
    offset = None
    last_commands = {}

    while True:
        try:
            url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/getUpdates"
            if offset:
                url += f"?offset={offset}"
            res = requests.get(url)
            data = res.json()

            if data["ok"]:
                for update in data["result"]:
                    offset = update["update_id"] + 1
                    if "message" in update:
                        message = update["message"]
                        chat_id = str(message["chat"]["id"])
                        text = message.get("text", "").strip()

                        if text.startswith("/"):
                            nama_sheet = text[1:].strip().lower()
                            if last_commands.get(chat_id) == nama_sheet:
                                continue  # Skip duplicate
                            last_commands[chat_id] = nama_sheet

                            print(f"📥 Permintaan dari {chat_id}: /{nama_sheet}")
                            flask_url = f"http://localhost:5000/{nama_sheet}?from_bot=true&chat_id={chat_id}"
                            try:
                                requests.get(flask_url)
                            except Exception as e:
                                print("❌ Gagal request ke Flask:", e)
            time.sleep(1)
        except Exception as e:
            print("❌ Error polling:", e)
            time.sleep(5)

# === JALANKAN APLIKASI ===
if __name__ == '__main__':
    threading.Thread(target=polling_bot, daemon=True).start()
    app.run(debug=False, port=5000)
