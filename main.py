import imaplib
import email
import pdfplumber
import os
import re
import datetime
import openpyxl
from openpyxl.styles import Font


def extract_invoice_data(text):
    """
    Extrahiert Rechnungsdaten aus dem Text (Datum, Nummer, Firma/Händler, Beträge).
    """
    datum_match = re.search(r'Datum[:\s]*([0-9]{2}\.[0-9]{2}\.[0-9]{4})', text)
    rechnungsnr_match = re.search(r'Rechnungsnummer[:\s]*(\S+)', text)
    firma_match = re.search(r'(?:Firma|Unternehmen|Verkäufer|Händler)[:\s]*(.+?)\n', text)
    zwischensumme_match = re.search(r'Zwischensumme[:\s]*([\d.,]+)', text)
    gesamtbetrag_match = re.search(r'(?:Gesamtbetrag)[:\s]*([\d.,]+)', text)

    if not datum_match:
        return None  # Kein gültiger Datensatz gefunden

    try:
        datum = datetime.datetime.strptime(datum_match.group(1), "%d.%m.%Y")
    except ValueError:
        return None

    return {
        "Datum": datum,
        "Rechnungsnummer": rechnungsnr_match.group(1) if rechnungsnr_match else '',
        "Firma": firma_match.group(1).strip() if firma_match else 'Unbekannt',
        "Zwischensumme": zwischensumme_match.group(1) if zwischensumme_match else '',
        "Gesamtbetrag": gesamtbetrag_match.group(1) if gesamtbetrag_match else ''
    }


def create_excel_table(alle_rechnungen, excel_file="rechnungen.xlsx"):
    """
    Erstellt eine Excel-Tabelle mit Rechnungen, gruppiert nach Jahr und Monat.
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Rechnungen"

    # Kopfzeile
    headers = ["Datum", "Rechnungsnummer", "Firma", "Zwischensumme", "Gesamtbetrag"]
    ws.append(headers)
    for col in ws[1]:
        col.font = Font(bold=True)

    # Rechnungen nach Jahr und Monat sortieren
    alle_rechnungen.sort(key=lambda x: x["Datum"])

    current_year = None
    current_month = None
    row = 2  # Start nach der Kopfzeile

    for rechnung in alle_rechnungen:
        year = rechnung["Datum"].year
        month = rechnung["Datum"].month

        # 4 Zeilen Abstand bei neuem Jahr
        if current_year != year:
            if current_year is not None:
                row += 4
            current_year = year
            current_month = None

        # 2 Zeilen Abstand bei neuem Monat
        if current_month != month:
            if current_month is not None:
                row += 2
            current_month = month

        # Rechnungsdaten einfügen
        ws.append([
            rechnung["Datum"].strftime("%d.%m.%Y"),
            rechnung["Rechnungsnummer"],
            rechnung["Firma"],
            rechnung["Zwischensumme"],
            rechnung["Gesamtbetrag"]
        ])
        row += 1

    # Spaltenbreite anpassen
    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = max_length + 2
        ws.column_dimensions[column].width = adjusted_width

    wb.save(excel_file)
    print(f"Excel-Tabelle gespeichert: {excel_file}")


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
            rechnung = extract_invoice_data(text)
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