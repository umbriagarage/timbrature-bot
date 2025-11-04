import telebot
from telebot import types
from datetime import datetime
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import time
import math
import json
import os

# =============== CONFIGURAZIONE ===============
TOKEN = "8072993225:AAEbp_P3ew50N0ZkenHqW63EDhCg5DennQk"
SPREADSHEET_ID = "1vnnsZj87OnK7ZqkmcH97ttcQ8AXYzdKLbXHHWsS6rpU"
WORKSHEET_NAME = "TIMBRATURE"
OWNER_EMAIL = "umbriagarage@gmail.com"
TIMEZONE = "Europe/Rome"

# =============== AUTORIZZAZIONI ===============
ALLOWED_USER_IDS = {80821293, 1829982561}

# =============== GEOFENCE ===============
DEFAULT_OFFICE = {"lat": 43.463360, "lon": 12.238560}  # Via Biturgense 74
GEOFENCE_METERS = 100
GEOFENCE_FILE = "geofence.json"

# Stato
pending_action = {}           # {user_id: {"azione": "...", "ts": time.time()}}
waiting_office_location = set()

# =============== UTIL: sede calibrata ===============
def load_office_coords():
    if os.path.exists(GEOFENCE_FILE):
        try:
            with open(GEOFENCE_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                if "lat" in data and "lon" in data:
                    return float(data["lat"]), float(data["lon"])
        except Exception:
            pass
    return DEFAULT_OFFICE["lat"], DEFAULT_OFFICE["lon"]

def save_office_coords(lat, lon):
    with open(GEOFENCE_FILE, "w", encoding="utf-8") as f:
        json.dump({"lat": lat, "lon": lon}, f)

# =============== GOOGLE SHEETS ===============
def gsheets_client():
    scope = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = ServiceAccountCredentials.from_json_keyfile_name("credentials.json", scope)
    return gspread.authorize(creds)

def get_or_create_spreadsheet(gc):
    try:
        sh = gc.open_by_key(SPREADSHEET_ID)
        if OWNER_EMAIL and "@" in OWNER_EMAIL:
            try:
                sh.share(OWNER_EMAIL, perm_type="user", role="writer")
            except Exception as e:
                print(f"⚠️ Impossibile condividere il file a {OWNER_EMAIL}: {e}")
    except gspread.SpreadsheetNotFound:
        print("❌ Errore: file non trovato. Controlla l'ID o le condivisioni (service account).")
        raise

    try:
        ws = sh.worksheet(WORKSHEET_NAME)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=WORKSHEET_NAME, rows=2000, cols=10)
        ws.update(
            range_name="A1:E1",
            values=[["Nome", "Azione", "Data", "Ora", "Conferma"]]
        )
    return sh, ws

def append_timbratura(ws, nome, azione, when, conferma=""):
    data_str = when.strftime("%d/%m/%Y")
    ora_str = when.strftime("%H:%M:%S")
    ws.append_row([nome, azione, data_str, ora_str, conferma], value_input_option="USER_ENTERED")

# =============== HELPER ===============
def now_local():
    try:
        return datetime.now(ZoneInfo(TIMEZONE))
    except ZoneInfoNotFoundError:
        return datetime.now()

def haversine_m(lat1, lon1, lat2, lon2):
    R = 6371000.0
    from math import radians, sin, cos, atan2, sqrt
    phi1 = radians(lat1); phi2 = radians(lat2)
    dphi = radians(lat2 - lat1)
    dlmb = radians(lon2 - lon1)
    a = sin(dphi/2)**2 + cos(phi1)*cos(phi2)*sin(dlmb/2)**2
    c = 2*atan2(sqrt(a), sqrt(1-a))
    return R*c

def check_auth(msg):
    return (not ALLOWED_USER_IDS) or (msg.from_user.id in ALLOWED_USER_IDS)

# =============== TELEGRAM BOT ===============
bot = telebot.TeleBot(TOKEN, parse_mode="HTML")

def main_keyboard():
    kb = types.ReplyKeyboardMarkup(resize_keyboard=True)
    kb.row("🕗 ENTRATA", "🏁 USCITA")
    kb.row("⚙️ /setsede")
    return kb

def location_keyboard(for_action):
    kb = types.ReplyKeyboardMarkup(resize_keyboard=True, one_time_keyboard=True)
    kb.add(types.KeyboardButton(text=f"📡 Invia posizione per {for_action}", request_location=True))
    return kb

@bot.message_handler(commands=["start", "help"])
def start_cmd(msg):
    if not check_auth(msg):
        bot.reply_to(msg, "❌ Non sei autorizzato a usare questo bot.")
        return
    lat, lon = load_office_coords()
    bot.reply_to(
        msg,
        "👋 Benvenuto nel bot timbrature.\n"
        "Scegli un’azione e invia la posizione.\n"
        f"📍 Sede valida entro {GEOFENCE_METERS} m dalla posizione salvata.",
        reply_markup=main_keyboard(),
    )

@bot.message_handler(commands=["adminlink"])
def adminlink_cmd(msg):
    if not check_auth(msg):
        bot.reply_to(msg, "❌ Non sei autorizzato a usare questo bot.")
        return
    gc = gsheets_client()
    sh, _ = get_or_create_spreadsheet(gc)
    bot.reply_to(msg, f"🔗 <b>Link file timbrature</b>:\n{sh.url}")

@bot.message_handler(commands=["setsede"])
def set_sede_cmd(msg):
    if not check_auth(msg):
        bot.reply_to(msg, "❌ Non sei autorizzato a usare questo bot.")
        return
    waiting_office_location.add(msg.from_user.id)
    kb = types.ReplyKeyboardMarkup(resize_keyboard=True, one_time_keyboard=True)
    kb.add(types.KeyboardButton(text="📡 Invia posizione per impostare la sede", request_location=True))
    bot.reply_to(
        msg,
        "⚙️ Calibrazione sede: invia la posizione dal punto esatto di timbratura.",
        reply_markup=kb,
    )

# ======= AZIONE: solo ENTRATA e USCITA =======
@bot.message_handler(func=lambda m: m.text in ["🕗 ENTRATA", "🏁 USCITA"])
def choose_action(msg):
    if not check_auth(msg):
        bot.reply_to(msg, "❌ Non sei autorizzato a usare questo bot.")
        return
    azione = msg.text.replace("🕗 ", "").replace("🏁 ", "")
    pending_action[msg.from_user.id] = {"azione": azione, "ts": time.time()}
    bot.send_message(
        msg.chat.id,
        f"➡️ {azione} selezionata.\nTocca il pulsante per inviare la posizione.",
        reply_markup=location_keyboard(azione),
    )

# ======= POSIZIONE: gestisce sia /setsede che l'azione =======
@bot.message_handler(content_types=["location"])
def handle_location(msg):
    if not check_auth(msg):
        bot.reply_to(msg, "❌ Non sei autorizzato a usare questo bot.")
        return

    # Calibrazione sede
    if msg.from_user.id in waiting_office_location:
        if not msg.location:
            waiting_office_location.discard(msg.from_user.id)
            bot.reply_to(msg, "⚠️ Posizione non valida. /setsede per riprovare.", reply_markup=main_keyboard())
            return
        save_office_coords(msg.location.latitude, msg.location.longitude)
        waiting_office_location.discard(msg.from_user.id)
        bot.reply_to(msg, "✅ Sede aggiornata.", reply_markup=main_keyboard())
        return

    # Azione con geofence
    info = pending_action.get(msg.from_user.id)
    if not info:
        bot.reply_to(msg, "ℹ️ Prima scegli un’azione (ENTRATA/USCITA).", reply_markup=main_keyboard())
        return

    if time.time() - info.get("ts", 0) > 180:
        pending_action.pop(msg.from_user.id, None)
        bot.reply_to(msg, "⏱️ Posizione arrivata troppo tardi: ripeti l’azione.", reply_markup=main_keyboard())
        return

    if not msg.location:
        pending_action.pop(msg.from_user.id, None)
        bot.reply_to(msg, "⚠️ Posizione non valida. Riprova l’azione.", reply_markup=main_keyboard())
        return

    office_lat, office_lon = load_office_coords()
    dist = haversine_m(msg.location.latitude, msg.location.longitude, office_lat, office_lon)
    if dist > GEOFENCE_METERS:
        pending_action.pop(msg.from_user.id, None)
        bot.reply_to(msg, "❌ Non sei in sede. Timbratura non valida.", reply_markup=main_keyboard())
        return

    nome = msg.from_user.first_name or "Sconosciuto"
    azione = info["azione"]
    now = now_local()

    try:
        gc = gsheets_client()
        _, ws = get_or_create_spreadsheet(gc)
        append_timbratura(ws, nome, azione, now, conferma="IN SEDE")
        bot.reply_to(msg, f"✅ {azione} registrata — {now.strftime('%d/%m/%Y %H:%M:%S')}", reply_markup=main_keyboard())
    except Exception as e:
        bot.reply_to(msg, f"⚠️ Errore nel salvataggio: <code>{e}</code>")
    finally:
        pending_action.pop(msg.from_user.id, None)

# =============== AVVIO ===============
if __name__ == "__main__":
    print("🚀 Bot timbrature attivo. Premi CTRL+C per uscire.")
    while True:
        try:
            bot.infinity_polling(timeout=60, long_polling_timeout=60)
        except Exception as e:
            print(f"Polling error: {e}. Retry tra 5s...")
            time.sleep(5)
