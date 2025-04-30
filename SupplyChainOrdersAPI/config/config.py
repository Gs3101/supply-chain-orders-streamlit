import os

# נתיבי קבצים
EXCEL_FILE_PATH = os.path.join("data", "orders.xlsx")
LOG_FILE_PATH = os.path.join("logs", "operations.log")
BACKUP_FOLDER = os.path.join("data", "backups")
PDF_OUTPUT_PATH = os.path.join("data", "Business_Summary_Report.pdf")
DASHBOARD_PATH = os.path.join("data", "dashboard.xlsx")
ORDER_SUMMARY_PATH = os.path.join("data", "order_summary.xlsx")
CHART_IMAGE_PATH = os.path.join("data", "top_products_chart.png")

# פרטי מייל
SENDER_EMAIL = "stein3101@gmail.com"
SENDER_PASSWORD = "cmfg lcrj pyyd nnzt"  # סיסמת אפליקציה
ALERT_EMAIL = "stein3101@gmail.com"

# API
API_BASE_URL = "https://dummyjson.com/carts"

# Cleanup
MAX_LOG_DAYS = 7
MAX_BACKUPS = 5
