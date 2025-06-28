import os
import json
from dotenv import load_dotenv
load_dotenv()

import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd
import matplotlib.pyplot as plt
import time
import requests


# === KONFIGURASI ===
SPREADSHEET_ID = '1d07pl6fcepVqpaaH3wZHKvq8cedUVoIzev9ci8I_GxM'
SHEET_NAME = 'PIVOT 2'
TELEGRAM_BOT_TOKEN = os.environ["TELEGRAM_BOT_TOKEN"]
TELEGRAM_CHAT_ID = os.environ["TELEGRAM_CHAT_ID"]

# === AUTENTIKASI GOOGLE SHEETS ===
# === AUTENTIKASI GOOGLE SHEETS ===
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]

credentials_info = json.loads(os.environ["GOOGLE_CREDS_JSON"])
credentials_info["private_key"] = credentials_info["private_key"].replace("\\n", "\n")

creds = ServiceAccountCredentials.from_json_keyfile_dict(credentials_info, scope)
client = gspread.authorize(creds)
sheet = client.open_by_key(SPREADSHEET_ID).worksheet(SHEET_NAME)


# === FUNGSI BACA DATA SEBAGAI DATAFRAME ===
def get_dataframe():
    data = sheet.get_all_values()
    headers = data[0]
    rows = data[1:]
    return pd.DataFrame(rows, columns=headers)

# === FUNGSI KONVERSI DATAFRAME KE GAMBAR PNG ===
def dataframe_to_image(df, filename='update.png'):
    fig, ax = plt.subplots(figsize=(len(df.columns) * 1.5, len(df) * 0.5))
    ax.axis('off')
    table = ax.table(cellText=df.values, colLabels=df.columns, loc='center', cellLoc='center')
    table.auto_set_font_size(False)
    table.set_fontsize(10)
    fig.tight_layout()
    plt.savefig(filename, dpi=200)
    plt.close()

# === FUNGSI KIRIM GAMBAR KE TELEGRAM ===
def send_telegram_photo(photo_path):
    url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendPhoto"
    with open(photo_path, 'rb') as photo:
        response = requests.post(url, data={'chat_id': TELEGRAM_CHAT_ID}, files={'photo': photo})
    if response.status_code != 200:
        print(f"❌ Gagal kirim gambar: {response.text}")
    else:
        print("✅ Gambar berhasil dikirim ke Telegram")

# === LOOP KIRIM SETIAP 1 JAM ===
while True:
    try:
        df = get_dataframe()
        dataframe_to_image(df, 'update.png')
        print("📤 Mengirim gambar update setiap 1 jam...")
        send_telegram_photo('update.png')
        time.sleep(3600)  # tidur 1 jam
    except Exception as e:
        print("❌ Error:", e)
        time.sleep(3600)
