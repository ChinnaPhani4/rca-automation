# RCA Automation with Python 🛠️

This project automates the analysis and reporting of incident RCA data from weekly ServiceNow exports.

## 🚀 Features

- Detects recurring incident issues from short descriptions
- Flags RCAs and tracks RCA owners vs ticket assignees
- Generates visual reports (Excel + chart image)
- Sends email via Outlook (optional)

## 📂 Structure

- `analyze_rca.py`: Parses and analyzes RCA data
- `send_rca_report_outlook.py`: Sends reports via Outlook
- `data/sample_incidents.xlsx`: Sample ServiceNow export
- `output/`: Generated reports (excluded from git)

## 🛠 Dependencies

Install these using pip:
```bash
pip install pandas matplotlib seaborn openpyxl pywin32
