# 📧 Residency Email Automation Tool

This is a Flask web app that automates the process of sending personalized Letters of Interest to medical residency programs via Gmail.

---

## 🚀 Features

- Upload Excel file with contact info
- Customize email subject and body using placeholders like `{{Name}}` and `{{Email}}`
- Emails are sent one by one with random delays (throttling)
- Log file with success/failure status is generated and emailed to the sender
- Downloadable log file + failed email report
- Loading spinner while sending

---

## 📂 Folder Structure

email-automation-mvp/
├── app.py
├── templates/
│ ├── index.html
│ ├── preview.html
│ └── log_download.html
├── uploads/ # Uploaded Excel files
├── sent_logs/ # Output logs with status
├── .gitignore
└── README.md


---

## 🛠️ Setup Instructions

1. **Clone the repo**  
```bash
git clone https://github.com/yourusername/email-automation-mvp.git
cd email-automation-mvp
