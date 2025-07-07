import imaplib
import email
import pdfplumber
import os
import re
import openpyxl
from openpyxl.styles import Font
from datetime import datetime

from gpt4all import GPT4All
model_path = r"C:\Users\nhabt\OneDrive\PycharmProjects\invoice-automation\Phi-3-mini-4k-instruct-q4.gguf"
ki = GPT4All(model_path)

def search_email(uid_max, criteria):
    """Erstellt Suchstrings für mehrere Absender."""
    search_strings = []
    for criteria_dict in criteria:
        search_parts = []
        for key, value in criteria_dict.items():
            search_parts.append(f'{key} "{value}"')
        search_parts.append(f'UID {uid_max + 1}:*')
        search_strings.append(f'({" ".join(search_parts)})')
    return search_strings


from openpyxl.styles import Font, PatternFill, Border, Side

def create_excel_table(alle_rechnungen, excel_file="Invoices.xlsx"):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Invoices"

    # Überschriften
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

    # Daten eintragen mit Euro hinten
    for row_num, rechnung in enumerate(alle_rechnungen, start=2):
        ws.cell(row=row_num, column=1,
                value=rechnung.get("Datum").strftime("%d.%m.%Y") if rechnung.get("Datum") else "")
        ws.cell(row=row_num, column=2, value=rechnung.get("Firma", ""))
        ws.cell(row=row_num, column=3, value=rechnung.get("Rechnungsnummer", ""))

        zelle_zwischensumme = ws.cell(row=row_num, column=4, value=rechnung.get("Zwischensumme", 0.0))
        zelle_zwischensumme.number_format = '#,##0.00 "€"'

        zelle_gesamtsumme = ws.cell(row=row_num, column=5, value=rechnung.get("Gesamtbetrag", 0.0))
        zelle_gesamtsumme.number_format = '#,##0.00 "€"'

    # Überschriften formatieren (fett + hellgrün)
    fill_color = PatternFill(fill_type="solid", fgColor="8EE53F")
    for col in range(1, 6):
        cell = ws.cell(row=1, column=col)
        cell.font = Font(bold=True, size=12)
        cell.fill = fill_color

    # Dünner Rahmen innen und außen
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
    print(f"Excel-Datei '{excel_file}' wurde gespeichert.")




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
       {text[:4000]}
       """
    antwort = ki.generate(prompt=prompt, max_tokens=300, temp=0.1, streaming=False)
    print("Antwort der KI:\n", antwort)

    # Antwort parsen mit Regex
    try:
        datum_match = re.search(r"Datum:\s*(\d{2}\.\d{2}\.\d{4})", antwort)
        rechnungsnr_match = re.search(r"Rechnungsnummer:\s*(\S*)", antwort)
        firma_match = re.search(r"Firma:\s*(.+)", antwort)
        zwischensumme_match = re.search(r"Zwischensumme:\s*([\d\.]+)", antwort)
        gesamtsumme_match = re.search(r"Gesamtsumme:\s*([\d\.]+)", antwort)

        rechnung = {
            "Datum": datetime.strptime(datum_match.group(1), "%d.%m.%Y") if datum_match else None,
            "Rechnungsnummer": rechnungsnr_match.group(1) if rechnungsnr_match else "",
            "Firma": firma_match.group(1).strip() if firma_match else "",
            "Zwischensumme": float(zwischensumme_match.group(1)) if zwischensumme_match else 0.0,
            "Gesamtbetrag": float(gesamtsumme_match.group(1)) if gesamtsumme_match else 0.0
        }
        return rechnung

    except Exception as e:
        print("Fehler beim Parsen der KI-Antwort:", e)
        return None


def process_attachment(part, save_dir="anhänge"):
    """Speichert und verarbeitet Anhänge (PDF, TXT, HTML)."""
    # Dateinamen holen oder generieren
    dateiname = part.get_filename() or f"anhang_{id(part)}.bin"
    if not os.path.exists(save_dir):
        os.makedirs(save_dir)
    filepath = os.path.join(save_dir, dateiname)

    # Inhalt speichern
    dateiinhalt = part.get_payload(decode=True)
    if dateiinhalt:
        with open(filepath, 'wb') as anhang_datei:
            anhang_datei.write(dateiinhalt)
        print(f"Anhang {dateiname} gespeichert")

        # Inhalt verarbeiten
        try:
            text = ''
            if dateiname.endswith('.pdf'):
                with pdfplumber.open(filepath) as pdf:
                    text = ''.join(page.extract_text() or '' for page in pdf.pages)
                    print(f"Inhalt der PDF {dateiname}:\n{text or 'Kein Text gefunden'}\n")
            elif dateiname.endswith('.txt'):
                with open(filepath, 'r', encoding='utf-8') as anhang_datei:
                    text = anhang_datei.read()
                    print(f"Inhalt der TXT-Datei {dateiname}:\n{text or 'Kein Text gefunden'}\n")
            elif dateiname.endswith('.html'):
                with open(filepath, 'r', encoding='utf-8') as anhang_datei:
                    html = anhang_datei.read()
                    text = re.sub(r'<[^>]+>', '', html).strip()
                    print(f"Inhalt der HTML-Datei {dateiname}:\n{text or 'Kein Text gefunden'}\n")
            else:
                print(f"Unbekanntes Format: {dateiname}")
                return None

            # Rechnungsdaten extrahieren
            rechnung = extract_invoice_data_ki(text)
            if rechnung:
                alle_rechnungen.append(rechnung)
                print(f"Rechnungsdaten extrahiert: {rechnung}")
            return rechnung

        except Exception as e:
            print(f"Fehler beim Lesen der Datei {dateiname}: {e}")
            return None



# Verbindungseinstellungen
server = "imap.gmail.com"
email_user = "test.max.mustermann01@gmail.com"
email_pass = "hxaj odpb fauv znmd"
email_format = 'RFC822'
alle_rechnungen = []
criteria = [
    {"FROM": "nhabtemicael@gmail.com"},
    {"FROM": "service@paypal.de"}
]
uid_max = 0
save_dir = "anhänge"

# Verbindung zum Server aufbauen
mail = imaplib.IMAP4_SSL(server)
mail.login(email_user, email_pass)
mail.select("INBOX")
print("IMAP-Server aktualisiert")

# Suchstrings bauen
search_str = search_email(uid_max, criteria)
print("Suchstring:", search_str)

# Suche ausführen
for search in search_str:
    result, data = mail.uid('search', None, search)
    print("Suchergebnis:", data)                                                       # date enthält alle UIDs
    print("Resultat:", result)

    # UIDs dekodieren
    uid = data[0].decode().split() if data[0] else []                                   #uids werden von byte stings -> liste[strings]

    if not uid:
        print(f"Keine neuen E-Mails für {search} gefunden")
        continue

    # Über jede UID iterieren
    for single_uid in uid:
        result, msg_data = mail.uid('fetch', single_uid, email_format)    #Code holt Rohdaten der Mail
        raw_email = msg_data[0][1]

        # E-Mail parsen
        email_message = email.message_from_bytes(raw_email)                              # Rohdaten aus der Mail wird zu strukturieren Daten herausgelesen

        # Anhänge speichern und verarbeiten
        for part in email_message.walk():
            if part.get_content_disposition() == 'attachment':
                process_attachment(part, save_dir)

# Excel-Tabelle erstellen
if alle_rechnungen:
    create_excel_table(alle_rechnungen)
else:
    print("Keine Rechnungsdaten zum Speichern in Excel gefunden")

# Verbindung schließen
mail.logout()
print("Verbindung geschlossen")


