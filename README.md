# Jira Report Automation in Python

![GitHub release (latest by date)](https://img.shields.io/github/v/release/tk6209/JiraReportAutomationPython)

This project automates the download of a CSV report from JIRA and the execution of an Excel macro to process and update a daily dashboard.

## 🔧 Requirements

- Python 3.7 (32-bit)
- [GeckoDriver](https://github.com/mozilla/geckodriver/releases)
- Mozilla Firefox installed
- Excel with macros enabled
- System environment variable set: `cred_anyconn` (contains your JIRA password)

## 📦 Installation

```bash
pip install -r requirements.txt
▶️ Running the Automation
You can run the main script directly with Python:

bash
Copiar
Editar
python OnePage.py
Or via the batch file:

bash
Copiar
Editar
run.bat
⚙️ Automation Architecture
This project integrates Python, VBScript, and Excel to automate the download, processing, and refresh of a JIRA report.

🔁 Flow Overview
OnePage.py (Python + Selenium):

Logs into the JIRA portal using your Windows username and the password stored in the environment variable cred_anyconn.

Downloads the CSV report.

Renames the file to raw.csv.

Calls the run.bat script.

run.bat:

Executes the script.vbs, passing the path to the Excel macro-enabled workbook.

script.vbs (VBScript):

Opens Excel.

Runs the macro named refresh.

Closes Excel automatically.

📊 Diagram
pgsql
Copiar
Editar
+----------------------+
|     OnePage.py       |
|  (Python + Selenium) |
+----------------------+
            |
            v
     Downloads CSV file
            |
            v
     Renames to raw.csv
            |
            v
       Executes run.bat
            |
            v
+----------------------+
|       run.bat        |
+----------------------+
            |
            v
     Calls script.vbs
            |
            v
+----------------------+
|      script.vbs      |
|  (VBScript + Excel)  |
+----------------------+
            |
            v
  Opens .xlsm workbook
  Runs "refresh" macro
  Closes Excel
📁 Project Structure
arduino
Copiar
Editar
JiraReportAutomationPython/
│
├── README.md
├── .gitignore
├── requirements.txt
├── OnePage.py
├── OnePage.bak.py
├── OnePage.spec
├── run.bat
├── script.vbs
├── run_command.txt
├── assets/
│   └── python.ico
├── macros/
│   └── Dashboard of Operations.xlsm
└── docs/
    └── tutorial.md
📝 Notes
Make sure the macro refresh exists and works properly inside Dashboard of Operations.xlsm.

The cred_anyconn environment variable must be configured in your system with your JIRA password.

Firefox and GeckoDriver must be installed and available in your system's PATH.

The tempAutomation folder (used for downloads) must exist in your user directory.

📄 License
MIT License (or replace with your preferred license).
