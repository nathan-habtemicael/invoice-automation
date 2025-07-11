# Invoice-automation
This project automates the extraction of invoice data from email attachments and stores it in a structured Excel file. The processing is done using a locally running AI model called "Phi-3-mini-4k-instruct-q4," eliminating the need for any external cloud services.

The application is developed in Python and utilizes various libraries to access emails via IMAP, extract text from PDF and text files, and create and format Excel spreadsheets. The AI model handles the extraction of relevant invoice information from the documents.

# Technologies used:
- Python 3.10 or higher
- IMAP protocol for email processing
- PDF-Plumber for text extraction from PDF files
- OpenPyXL for creating and formatting Excel files
- Local AI model "Phi-3-mini-4k-instruct-q4" for data extraction

The AI model must be downloaded separately and made available locally. The path to the model can be configured within the source code. The project supports processing email attachments in PDF, TXT, and HTML formats. Additionally, it only processes new emails to avoid duplicate extraction.
