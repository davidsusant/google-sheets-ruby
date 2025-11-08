# Google Sheets Manager - Ruby Service Account Script

A Ruby script for creating and configuring Google Sheets using a service account. Perfect for QA automation, test result reporting, and automated spreadsheet management.

## Features

- ✅ Create new spreadsheets programmatically
- ✅ Add multiple sheets to a spreadsheet
- ✅ Write data to specific ranges
- ✅ Format header rows (bold, colored, frozen)
- ✅ Auto-resize columns to fit content
- ✅ Add dropdown validation (status fields, etc.)
- ✅ Add checkbox validation
- ✅ Apply conditional formatting (color-coded statuses)
- ✅ Retrieve spreadsheet metadata

## Prerequisites

- Ruby 3.0 or higher
- A Google Cloud Project
- A Service Account with Google Sheets API enabled

## Setup

### 1. Create a Service Account

1. Go to [Google Cloud Console](https://console.cloud.google.com/)
2. Create a new project or select an existing one
3. Enable the Google Sheets API:
   - Navigate to "APIs & Services" > "Library"
   - Search for "Google Sheets API"
   - Click "Enable"
4. Create a Service Account:
   - Navigate to "APIs & Services" > "Credentials"
   - Click "Create Credentials" > "Service Account"
   - Fill in the service account details
   - Click "Create and Continue"
   - Skip optional steps and click "Done"
5. Create a JSON key:
   - Click on the newly created service account
   - Go to the "Keys" tab
   - Click "Add Key" > "Create new key"
   - Choose "JSON" format
   - Click "Create" (the key file will download automatically)
6. Save the downloaded JSON file as ```credentials.json``` in your project directory

### 2. Install Dependencies



## Usage

### Basic Usage

Run the example script:

```bash
ruby google_sheets_manager.rb
```

This will create a sample test results spreadsheets with formatted headers, dropdowns, and conditional formatting.

