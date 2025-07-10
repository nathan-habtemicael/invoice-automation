import imaplib
import email
import pdfplumber
import os
import re
import openpyxl
from openpyxl.styles import Font
from datetime import datetime
from openpyxl.styles import Font, PatternFill, Border, Side
from gpt4all import GPT4All

        # Anzahl logischer Kerne herausfinden
num_threads = os.cpu_count()
print("Logische Kerne (inkl. Hyperthreading):", num_threads)


        # GPT4All-Modell laden
model_path = r"C:\Users\nhabt\OneDrive\PycharmProjects\invoice-automation\Phi-3-mini-4k-instruct-q4.gguf"
ki = GPT4All(model_path, n_threads= num_threads-1)


        # UID-Handling
def lade_uid_max(dateiname="uid_max.txt"):
    try:
        with open(dateiname, "r") as f:
            return int(f.read().strip())        #strip() entfernt lehrzeichenartige Zeichen vor und nach dem String
    except:
        return 0


def speichere_uid_max(uid, dateiname="uid_max.txt"):
    with open(dateiname, "w") as f:
        f.write(str(uid))


        # E-Mail-Suche vorbereiten
def search_email(uid_max, criteria):
    search_strings = []
    for criteria_dict in criteria:
        search_parts = []
        for key, value in criteria_dict.items():
            search_parts.append(f'{key} "{value}"')
        search_parts.append(f'UID {uid_max + 1}:*')
        search_strings.append(f'({" ".join(search_parts)})')
    return search_strings


        # Excel-Datei erstellen
def create_excel_table(alle_rechnungen, excel_file="Invoices.xlsx"):
    if os.path.exists(excel_file):
        # Bestehende Datei laden
        wb = openpyxl.load_workbook(excel_file)
        ws = wb.active
        start_row = ws.max_row + 1  # neue Daten darunter schreiben
    else:
        # Neue Datei erstellen
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Invoices"

        # Überschriften schreiben
        ws.cell(row=1, column=1, value="Datum")
        ws.cell(row=1, column=2, value="Firma")
        ws.cell(row=1, column=3, value="Rechnungsnummer")
        ws.cell(row=1, column=4, value="Zwischensumme")
        ws.cell(row=1, column=5, value="Gesamtsumme")

        # Spaltenbreite definieren
        ws.column_dimensions['A'].width = 12
        ws.column_dimensions['B'].width = 20
        ws.column_dimensions['C'].width = 25
        ws.column_dimensions['D'].width = 17
        ws.column_dimensions['E'].width = 17

        # Überschriften formatieren (fett + hellgrün)
        fill_color = PatternFill(fill_type="solid", fgColor="8EE53F")
        for col in range(1, 6):
            cell = ws.cell(row=1, column=col)
            cell.font = Font(bold=True, size=12)
            cell.fill = fill_color

        start_row = 2  # ab zweiter Zeile Daten einfügen

        # Neue Daten anhängen
    for row_num, rechnung in enumerate(alle_rechnungen, start=start_row):
        ws.cell(row=row_num, column=1,
                value=rechnung.get("Datum").strftime("%d.%m.%Y") if rechnung.get("Datum") else "")
        ws.cell(row=row_num, column=2, value=rechnung.get("Firma", ""))
        ws.cell(row=row_num, column=3, value=rechnung.get("Rechnungsnummer", ""))

        zelle_zwischensumme = ws.cell(row=row_num, column=4, value=rechnung.get("Zwischensumme", 0.0))
        zelle_zwischensumme.number_format = '#,##0.00 "€"'

        zelle_gesamtsumme = ws.cell(row=row_num, column=5, value=rechnung.get("Gesamtbetrag", 0.0))
        zelle_gesamtsumme.number_format = '#,##0.00 "€"'


    dünne_linien = Side(border_style="thin", color="000000")
    max_row = ws.max_row
    max_col = 5

    for row in range(1, max_row + 1):
        for col in range(1, max_col + 1):
            cell = ws.cell(row=row, column=col)
            cell.border = Border(
                left=dünne_linien,
                right=dünne_linien,
                top=dünne_linien,
                bottom=dünne_linien
            )

    wb.save(excel_file)
    print(f"Excel-Datei '{excel_file}' wurde aktualisiert.")

        # Daten extrahieren
def extract_invoice_data_ki(text):
    prompt = f"""
       Extrahiere die folgenden Rechnungsdaten aus dem Text im EXAKT angegebenen Format. Gib NUR die Daten aus, ohne jegliche zusätzliche Erklärungen, Beispiele, Überschriften oder Kommentare. Wenn ein Wert nicht gefunden wird, lasse das Feld leer.

       Format:
       Datum: TT.MM.JJJJ
       Rechnungsnummer: <Zeichenkette ohne Leerzeichen>
       Firma: <Zeichenkette>
       Zwischensumme: <Zahl mit 2 Nachkommastellen, Punkt als Dezimaltrennzeichen>
       Gesamtsumme: <Zahl mit 2 Nachkommastellen, Punkt als Dezimaltrennzeichen>

       Text:
       {text[:1500]}
    """
    antwort = ki.generate(prompt=prompt, max_tokens=300, temp=0.1, streaming=False)
    print("Antwort der KI:\n", antwort)

    try:
        return {
            "Datum": datetime.strptime(re.search(r"Datum:\s*(\d{2}\.\d{2}\.\d{4})", antwort).group(1),
                                       "%d.%m.%Y") if "Datum:" in antwort else None,
            "Rechnungsnummer": re.search(r"Rechnungsnummer:\s*(\S+)", antwort).group(
                1) if "Rechnungsnummer:" in antwort else "",
            "Firma": re.search(r"Firma:\s*(.+)", antwort).group(1).strip() if "Firma:" in antwort else "",
            "Zwischensumme": float(
                re.search(r"Zwischensumme:\s*([\d\.]+)", antwort).group(1)) if "Zwischensumme:" in antwort else 0.0,
            "Gesamtbetrag": float(
                re.search(r"Gesamtsumme:\s*([\d\.]+)", antwort).group(1)) if "Gesamtsumme:" in antwort else 0.0
        }
    except Exception as e:
        print("Fehler beim Parsen der KI-Antwort:", e)
        return None


        # Anhang verarbeiten
def process_attachment(part, save_dir="anhänge"):
    dateiname = part.get_filename() or f"anhang_{id(part)}.bin"
    os.makedirs(save_dir, exist_ok=True)
    filepath = os.path.join(save_dir, dateiname)

    with open(filepath, 'wb') as f:
        f.write(part.get_payload(decode=True))
    print(f"Anhang {dateiname} gespeichert")

    try:
        text = ''
        if dateiname.endswith('.pdf'):
            with pdfplumber.open(filepath) as pdf:
                text = ''.join(page.extract_text() or '' for page in pdf.pages)
        elif dateiname.endswith('.txt'):
            with open(filepath, 'r', encoding='utf-8') as f:
                text = f.read()
        elif dateiname.endswith('.html'):
            with open(filepath, 'r', encoding='utf-8') as f:
                text = re.sub(r'<[^>]+>', '', f.read()).strip()
        else:
            print(f"Unbekanntes Format: {dateiname}")
            return None

        rechnung = extract_invoice_data_ki(text)
        if rechnung:
            alle_rechnungen.append(rechnung)
            print(f"Rechnungsdaten extrahiert: {rechnung}")
        return rechnung

    except Exception as e:
        print(f"Fehler beim Verarbeiten von {dateiname}: {e}")
        return None


        # Hauptprogramm
server = "imap.gmail.com"
email_user = "test.max.mustermann01@gmail.com"
email_pass = "hxaj odpb fauv znmd"
email_format = 'RFC822'
criteria = [{"FROM": "nhabtemicael@gmail.com"}, {"FROM": "service@paypal.de"}]
save_dir = "anhänge"
alle_rechnungen = []
uid_max = lade_uid_max()

mail = imaplib.IMAP4_SSL(server)
mail.login(email_user, email_pass)
mail.select("INBOX")
print("IMAP-Server verbunden")

search_str = search_email(uid_max, criteria)

for search in search_str:
    result, data = mail.uid('search', None, search)
    uid_list = data[0].decode().split() if data[0] else []

    if not uid_list:
        print(f"Keine neuen E-Mails für {search}")
        continue

    for single_uid in uid_list:
        uid_int = int(single_uid)
        if uid_int > uid_max:
            uid_max = uid_int  # Max UID aktualisieren

        result, msg_data = mail.uid('fetch', single_uid, email_format)
        raw_email = msg_data[0][1]
        email_message = email.message_from_bytes(raw_email)

        for part in email_message.walk():
            if part.get_content_disposition() == 'attachment':
                process_attachment(part, save_dir)

# Excel speichern
if alle_rechnungen:
    create_excel_table(alle_rechnungen)
else:
    print("Keine Rechnungen gefunden")

# UID speichern und Verbindung schließen
speichere_uid_max(uid_max)
mail.logout()
print("Verbindung geschlossen")
