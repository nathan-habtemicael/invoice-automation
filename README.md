# Invoice Automation
This project automates the extraction of invoice data from email attachments and stores it in a structured Excel file. The processing is handled by a locally running AI model called "Phi-3-mini-4k-instruct-q4", eliminating the need for any external cloud services.
The application is developed in Python and uses various libraries to access emails via IMAP, extract text from PDF and text files, and create well-formatted Excel spreadsheets. 
The AI model is responsible for identifying and extracting key invoice information from the documents.


# Practical Use Case
This tool is designed to help users automatically collect and organize invoice data received via email. Instead of manually opening each invoice and entering details into a spreadsheet, this program handles the process automatically.
It connects to the user's email, downloads new invoices, extracts important information such as invoice number, date, company name, and total amount using an AI model, and compiles everything into a clear Excel file. Because the AI runs locally on the user's computer, no data is sent to external servers, ensuring privacy and security.
This automation saves time, reduces errors, and provides a convenient overview of invoices, which is particularly useful for managing subscription payments or business expenses.


# Technologies Used
- Python 3.10 or higher
- IMAP protocol for email processing
- PDF-Plumber for text extraction from PDF files
- OpenPyXL for creating and formatting Excel files
- Local AI model "Phi-3-mini-4k-instruct-q4" for data extraction

The AI model must be downloaded separately and stored locally. The path to the model can be configured within the source code. The project supports processing email attachments in PDF, TXT, and HTML formats. Additionally, it only processes new emails to avoid duplicate extraction.
