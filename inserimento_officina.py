import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
from gspread.utils import rowcol_to_a1
import re  # per adattare i riferimenti di riga nelle formule

# === CONFIGURAZIONE ===
SHEET_NAME = "Officina"
WORKSHEET_NAME = "MAGGIO25"
CRED_FILE = "progetto-457616-4f11433a44b5.json"

# === AUTENTICAZIONE ===
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name(CRED_FILE, scope)
client = gspread.authorize(creds)
sheet = client.open(SHEET_NAME).worksheet(WORKSHEET_NAME)

# === DATA OGGI ===
oggi = datetime.now().strftime("%d/%m/%Y")

# === RACCOLTA DATI ===
print("\n📋 INSERIMENTO NUOVA LAVORAZIONE\n")

meccanico = input("🧑‍🔧 Nome del meccanico: ").strip()
servizio = input("🔧 Lavoro/servizio eseguito: ").strip()
incasso = input("💰 Incasso (€): ").strip()
fattura = input("🧾 Numero fattura (o N/D): ").strip()
costi = input("💸 Costi (€): ").strip()

# === CERCA LA PRIMA RIGA CON LA DATA DI OGGI ===
colonna_date = sheet.col_values(1)
riga_inserimento = None

for i, valore in enumerate(colonna_date, start=1):
    if valore == oggi:
        valori_riga = sheet.row_values(i)
        servizio_cell = valori_riga[2] if len(valori_riga) > 2 else ""
        if servizio_cell.strip() == "":
            riga_inserimento = i
        else:
            riga_inserimento = i + 1
            sheet.insert_row([""] * sheet.col_count, riga_inserimento)

            # === Copia e adatta le formule dalla riga sopra ===
            for col in range(1, sheet.col_count + 1):
                cell_above = rowcol_to_a1(i, col)
                formula = sheet.acell(cell_above, value_render_option='FORMULA').value
                if formula and str(formula).startswith("="):
                    pattern = r'([A-Z]+)' + str(i)
                    replacement = lambda m: m.group(1) + str(riga_inserimento)
                    formula_adattata = re.sub(pattern, replacement, formula)
                    sheet.update_acell(rowcol_to_a1(riga_inserimento, col), formula_adattata)

            # Inserisce la data nella nuova riga
            sheet.update_acell(f"A{riga_inserimento}", oggi)
        break

# === SE LA DATA NON ESISTE, AGGIUNGI IN FONDO ===
if riga_inserimento is None:
    riga_inserimento = len(colonna_date) + 1
    sheet.update_acell(f"A{riga_inserimento}", oggi)

# === INSERISCI I DATI ===
sheet.update_acell(f"B{riga_inserimento}", meccanico)
sheet.update_acell(f"C{riga_inserimento}", servizio)
sheet.update_acell(f"D{riga_inserimento}", incasso)
sheet.update_acell(f"E{riga_inserimento}", fattura)
sheet.update_acell(f"H{riga_inserimento}", costi)

print(f"\n✅ Lavorazione registrata correttamente nella riga {riga_inserimento} per la data {oggi}!\n")
