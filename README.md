# Logistics Reporting Automation – PT PELNI Internship Project

This project was developed during my internship at **PT PELNI (Persero)** to automate logistics data processing and reporting for container transport operations.

The system transforms raw shipping manifest data into structured operational reports and assists operational staff in automating transaction input into the internal CIS system.

The automation pipeline integrates **PDF data extraction, system automation, and reporting generation**.

---

## Project Overview

The automation system performs the following workflow:

1. Extract shipping manifest data from PDF documents
2. Clean and structure logistics data using Python
3. Automate container transaction input to the CIS system
4. Generate structured logistics recap reports in Google Sheets

This project helps reduce manual work, improve reporting accuracy, and accelerate logistics monitoring processes.

---

## Automation Pipeline

Shipping Manifest PDF
↓
Python Data Extraction (pdfplumber + pandas)
↓
Structured Dataset (CSV / Excel)
↓
Selenium Automation (CIS System Input)
↓
Google Apps Script (Logistics Recap Report)

---

## Project Components

### 1. Manifest PDF Parser

Script for extracting and processing shipping manifest data from PDF documents.

Features:

- Extract table data from shipping manifest PDFs
- Clean and normalize shipper and consignee information
- Process container grouping based on remarks
- Convert extracted data into structured CSV and Excel reports
- Automatically format Excel output for operational use

Main script:
python/manifest_pdf_parser.py

---

### 2. CIS Transaction Automation

Automation script using **Selenium WebDriver** to assist container transaction input into the internal CIS system.

Features:

- Automated dropdown selection and form input
- File upload automation (container photos and supporting documents)
- Progress tracking and ETA estimation
- Error detection and failure logging
- Automated result logging into Excel

Main script:
python/cis_transaction_automation.py

---

### 3. Google Sheets Logistics Recap Automation

Google Apps Script used to automatically generate logistics recap reports from processed manifest data stored in Google Sheets.

Features:

- Automatically groups logistics data by voyage and destination port
- Aggregates shipper quantities and freight revenue
- Generates formatted recap tables for each voyage
- Applies automatic styling, coloring, and totals calculation
- Produces operational reports ready for logistics monitoring

Main script:
apps-script/generate_rekap_logistik.gs

---

## Demo Video

A short demonstration of the automation workflow can be viewed here:

https://youtube.com/B9I63Ns0F4k

The demo shows:

- PDF manifest parsing
- Automated transaction input process
- Generated logistics recap reports

---

## Technologies Used

Python  
Pandas  
pdfplumber  
Selenium WebDriver  
Regular Expressions  
OpenPyXL  
Google Apps Script  

---

## Repository Structure
pelni-logistics-report-automation

python/
manifest_pdf_parser.py
cis_transaction_automation.py

apps-script/
generate_rekap_logistik.gs

requirements.txt
README.md

---

## Installation

Clone this repository:
git clone https://github.com/giffaniriz25feb/pelni-logistics-report-automation.git

Install required Python libraries:
pip install -r requirements.txt

---

## Disclaimer

This repository is shared for **educational and portfolio purposes only**.

The original datasets used in this project contain **internal operational data from PT PELNI** and therefore cannot be publicly shared.

All sensitive operational data, documents, and system credentials have been excluded from this repository.
