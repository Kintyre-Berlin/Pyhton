import os
import time
from datetime import datetime
import win32com.client as win32

# === HILFSFUNKTION: Excel sicher initialisieren (inkl. makepy fallback) ===
def safe_excel_dispatch():
    try:
        print("📦 Versuche Excel COM-Interface zu laden...")
        return win32.gencache.EnsureDispatch("Excel.Application")
    except Exception as e:
        print(f"⚠️ Fehler bei EnsureDispatch: {e}")
        print("🔧 Versuche jetzt, makepy manuell auszuführen...")
        from win32com.client import makepy
        makepy.GenerateFromTypeLibSpec("Microsoft Excel xx.x Object Library")  # Wird automatisch aufgelöst
        # Jetzt nochmal versuchen:
        return win32.gencache.EnsureDispatch("Excel.Application")

# === DYNAMISCHER PFAD ZU DOWNLOADS ===
user_profile = os.path.expanduser("~")
downloads_path = os.path.join(user_profile, "Downloads")
iqy_path = os.path.join(downloads_path, "query.iqy")

# === PRÜFEN OB DIE .IQY-DATEI EXISTIERT ===
if not os.path.exists(iqy_path):
    raise FileNotFoundError(f"❌ .iqy-Datei nicht gefunden: {iqy_path}")

# === DATEINAME MIT HEUTIGEM DATUM ===
date_string = datetime.now().strftime("%Y%m%d")
save_name = f"{date_string} - Entity ID lookup failed.xlsx"
save_path = os.path.join(downloads_path, save_name)

# === E-MAIL-VORBEREITUNG MIT DATEINAME IM BETREFF & TEXT ===
empfaenger = "empfaenger@example.com"  # <-- ANPASSEN!
mail_betreff = f"Entity ID Lookup Report – {save_name}"
mail_text = (
    f"Hallo,\n\n"
    f"Anbei die gefilterte Datei: {save_name}\n"
    "Gefiltert nach 'Downloaded = OPEN' und 'EntityName = Entity ID lookup failed'.\n\n"
    "Grüße\nTino"
)

# === EXCEL STARTEN & .IQY LADEN ===
excel = safe_excel_dispatch()
excel.Visible = True

print("📂 Öffne .iqy-Datei...")
wb = excel.Workbooks.Open(iqy_path)
ws = wb.Sheets(1)

# === WARTEN BIS DIE DATEN WIRKLICH GELADEN SIND ===
print("⏳ Warte auf Daten...")
max_wait = 30
elapsed = 0
while ws.UsedRange.Rows.Count < 2 and elapsed < max_wait:
    time.sleep(1)
    elapsed += 1
    print(f"  ... {elapsed}s vergangen")

if ws.UsedRange.Rows.Count < 2:
    raise Exception("❌ Daten wurden nicht geladen – prüf die .iqy-Datei oder SharePoint-Verbindung.")

print("✅ Daten erfolgreich geladen.")

# === FILTER SETZEN ===
print("🔍 Filter setzen...")
ws.Range("A1").AutoFilter(Field=2, Criteria1="OPEN")  # Spalte B
ws.Range("A1").AutoFilter(Field=4, Criteria1="Entity ID lookup failed")  # Spalte D

# === SPEICHERN ALS XLSX ===
print(f"💾 Speichere Datei als: {save_name}")
wb.SaveAs(save_path, FileFormat=51)  # 51 = .xlsx
wb.Close(SaveChanges=False)
excel.Quit()

# === OUTLOOK-MAIL ERSTELLEN & ANZEIGEN ===
print("📧 Erstelle E-Mail...")
outlook = win32.Dispatch("Outlook.Application")
mail = outlook.CreateItem(0)
mail.To = empfaenger
mail.Subject = mail_betreff
mail.Body = mail_text

if os.path.exists(save_path):
    mail.Attachments.Add(save_path)
else:
    raise FileNotFoundError(f"❌ Datei nicht gefunden: {save_path}")

mail.Display()  # ← Zum Prüfen und manuell Absenden

print("✅ Mail erstellt und angezeigt. Klick auf 'Senden', wenn alles passt.")
