# Document Redaction Add-in (Word)

This Office Add-in provides a task pane with a single button that:

1. Enables Track Changes
2. Inserts a confidentiality header: **CONFIDENTIAL DOCUMENT**
3. Redacts sensitive information in the document body

## Features

### 1) Redact Sensitive Information
The add-in scans the entire document body text and replaces sensitive values with a redaction marker.

**Required patterns**
- Email addresses
- Phone numbers
- Social Security Numbers (SSNs)

**Additional patterns**
- Credit card numbers
- Dates of birth (DOB)
- Common structured IDs
- SSN last-4 digits when explicitly referenced as SSN / social security


### 2) Add Confidential Header 
On each run, the add-in ensures the primary header contains:

**CONFIDENTIAL DOCUMENT**


### 3) Enable Tracking Changes
Track Changes is enabled only when the Word JavaScript API requirement set **WordApi 1.5** is supported:

- `Office.context.requirements.isSetSupported("WordApi", "1.5")`

All header insertion and redactions occur after enabling tracking (when supported), so Word logs the modifications.

## Tech / Constraints
- TypeScript
- Runs in Word desktop and Word on the web
- Self-written CSS only (no external UI libraries)

## Run Locally

### Install
```bash
npm install
```

Start
```bash
npm start
```

This will:

Compile TypeScript in watch mode

Start an HTTPS dev server on port 3000

Attempt to sideload the add-in into Word

HTTPS certificate setup (Word Desktop)

Word desktop requires HTTPS. If you see certificate errors, install the development certificates and restart npm start:

```bash
npx office-addin-dev-certs install
```

Testing

Open the provided Document-To-Be-Redacted.docx

Open the add-in task pane

Click the button to add the header and redact sensitive content


Publisher: Feneel Doshi

Copyright: © 2025 Feneel Doshi. All rights reserved.
