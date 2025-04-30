# 📦 Supply Chain Orders Management System

A professional-grade Python-based tool for managing, analyzing, and automating supply chain order processes.
The system enables full integration with a dummy REST API (via JSON), supports Excel inputs/outputs, and provides real-time monitoring, alerting, and reporting.

## 📌 Features
- 🔁 Bi-directional API integration
- 📊 Excel automation (read/write)
- 📁 Versioned backups
- 📈 Summary reports and dashboard (Excel + PDF)
- 📬 Email alerts on failure
- 🧪 Simulation mode
- 🧩 Process monitoring and logs

## 🗂️ Folder Structure
SupplyChainOrdersAPI/
├── main.py
├── config.py
├── utils.py
├── data/
│   ├── orders.xlsx
│   ├── help.txt
│   ├── backups/
│   ├── logs/
│   └── *.xlsx/*.pdf

## 🔧 Prerequisites
- Python 3.13+
- Install: pip install -r requirements.txt

## 🚀 How to Run
python main.py

## 📬 Email
- Alerts sent to: stein3101@gmail.com
- Uses Gmail SMTP (SSL)

## 📄 Author
Guy Stein
v1.0 – April 30, 2025
