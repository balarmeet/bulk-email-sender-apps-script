# 📧 Bulk Email Automation System (Google Apps Script)

## 🚀 Overview

A production-ready bulk email automation system built using **Google Apps Script** that integrates **Google Sheets, Google Docs, Google Drive, and Gmail** to send personalized HTML emails at scale.

This tool allows users to manage email campaigns directly from Google Sheets, use Google Docs as dynamic templates, embed inline images automatically, attach PDF files, and track email status with detailed error logging.

---

## 🏗 Architecture Overview

Google Sheet (Data + Configuration)  
        ↓  
Google Apps Script  
        ↓  
Google Docs (Template → HTML Conversion)  
        ↓  
Google Drive API (Image Export + Attachments)  
        ↓  
Gmail (MailApp.sendEmail)

---

## 📊 Google Sheet Structure

### 🔧 Configuration Section

| Cell | Purpose | Description |
|------|----------|-------------|
| B1 | Email Subject | Subject line for all outgoing emails |
| B2 | Google Doc ID | Template Document ID used for HTML email body |

---

### 🧪 Test Email Row (Row 7)

Used to test the template before bulk sending.

| Column | Field | Description |
|--------|--------|-------------|
| A7 | Email | Test recipient email |
| B7 | Name | Replaces `{{Name}}` |
| C7 | AttachmentId | Optional Google Drive File ID |
| D7 | v1 | Replaces `{{v1}}` |
| E7 | v2 | Replaces `{{v2}}` |
| F7 | v3 | Replaces `{{v3}}` |
| G7 | Sent Status | YES / ERROR |
| H7 | Date Sent | Auto timestamp |
| I7 | Error Message | Error details if failed |

---

### 📩 Bulk Email Data (Starting from Row 10)

| Column | Header | Purpose |
|--------|--------|----------|
| A | Email | Recipient email address |
| B | Name | Replaces `{{Name}}` |
| C | AttachmentId | Google Drive File ID (converted to PDF) |
| D | v1 | Dynamic variable |
| E | v2 | Dynamic variable |
| F | v3 | Dynamic variable |
| G | Sent Status | YES / ERROR |
| H | Date Sent | Auto-generated timestamp |
| I | Error Details | Stores row-level errors |

---

## 🧩 Template Personalization

The Google Doc template supports the following placeholders:

| Placeholder | Value Source |
|-------------|--------------|
| `{{Name}}` | Column B |
| `{{v1}}` | Column D |
| `{{v2}}` | Column E |
| `{{v3}}` | Column F |

### Example Template
Hello {{Name}},

Your invoice number is {{v1}}.
Due date: {{v2}}.
Amount: {{v3}}.


---

## ✨ Key Features

- Personalized email templates
- HTML email conversion from Google Docs
- Inline image processing using CID
- Optional PDF attachment support
- Test email functionality
- Row-level error tracking
- Automated timestamp updates
- Gmail rate-limit handling (2-second delay)
- Custom Spreadsheet UI menu

---

## 🛠 Technologies Used

- Google Apps Script
- Gmail (MailApp)
- Google Drive API
- Google Docs API
- Google Sheets
- JavaScript (ES5)

---

## 🔐 Required OAuth Scopes
https://www.googleapis.com/auth/script.

https://www.googleapis.com/auth/spreadsheets

https://www.googleapis.com/auth/documents

https://www.googleapis.com/auth/drive

https://www.googleapis.com/auth/gmail.send


---

## ⚙️ How It Works

1. Enter subject in cell B1  
2. Enter Google Doc template ID in cell B2  
3. Add recipient data starting from row 10  
4. Use custom menu → "Send Bulk Emails"  
5. Script:
   - Converts Google Doc to HTML
   - Processes inline images
   - Replaces placeholders
   - Sends emails
   - Updates status and timestamps

---

## 📈 Technical Highlights

- OAuth token handling for secure API requests
- HTML extraction from Google Docs
- Inline image transformation using Content-ID (CID)
- Dynamic placeholder replacement engine
- Attachment conversion to PDF
- Error isolation per row
- Spreadsheet-driven configuration system

---

## 🎯 Use Cases

- HR recruitment campaigns  
- Invoice distribution  
- Marketing outreach  
- Client notifications  
- Event invitations  

---

## 💼 Resume Description (Short Version)

Developed a bulk email automation system using Google Apps Script integrated with Gmail, Sheets, Docs, and Drive APIs. Implemented dynamic HTML template processing with inline image embedding, PDF attachment handling, and row-level error tracking with Gmail rate-limit compliance.

---

## 🚀 Future Improvements

- Retry failed emails feature
- Email quota monitoring
- Logging dashboard sheet
- Web-based UI using HTMLService
- Scheduled sending via time-driven triggers

---

## ⭐ Why This Project Matters

This project demonstrates:

- API integration expertise
- Automation system design
- Production-level error handling
- Template engine development
- Real-world business problem solving
- Strong understanding of Google Workspace ecosystem
