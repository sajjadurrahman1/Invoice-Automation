Invoice Automation & Reporting (Google Workspace)
Overview

This project demonstrates an automated invoice processing workflow built using Google Sheets, Gmail, Google Drive, and Google Apps Script.
The goal is to simulate a real-world commercial invoicing process by automatically collecting invoice data, storing documents, and enabling structured analysis and reporting.

The project focuses on data accuracy, process digitalization, and reporting support, reflecting common operational challenges in commercial and logistics environments.

Key Features

Automatically imports PDF invoice attachments from Gmail emails labeled “Invoices”

Saves invoice PDFs to a dedicated Google Drive folder

Extracts invoice metadata from file names and stores it in a centralized Invoice Database (Google Sheets)

Tracks Gmail Message IDs to prevent duplicate imports

Enables pivot tables and dashboards for invoice status and trend analysis

Supports manual refresh and validation of invoice data

Technologies Used

Google Sheets – invoice database, analysis, and dashboards

Gmail – invoice source (email attachments & labels)

Google Drive – document storage

Google Apps Script – automation logic

Pivot Tables & Charts – data visualization and reporting

Project Structure

importInvoiceAttachments.gs – Imports invoice PDFs, extracts data, and updates the database

fillGmailIdColumn.gs – Backfills and maintains Gmail Message IDs for traceability

Invoice Database (Google Sheet) – Centralized data store for invoices

Dashboard – Pivot tables and charts for analysis

Use Case

This project simulates how automated data collection and reporting can support:

Commercial invoicing workflows

Data accuracy and traceability

Operational decision-making

Digital transformation initiatives

Author

Sajjadur Rahman
Information Engineering Student
GitHub: https://github.com/sajjadurrahman1
