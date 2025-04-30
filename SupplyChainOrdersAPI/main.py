import sys
import time
import os
import pandas as pd
import requests
import smtplib
import shutil
from datetime import datetime
from email.message import EmailMessage
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.drawing.image import Image
import matplotlib.pyplot as plt
from fpdf import FPDF

# === CONFIG ===
EXCEL_FILE_PATH = os.path.join("data", "orders.xlsx")
LOG_FILE_PATH = os.path.join("logs", "operations.log")
BACKUP_FOLDER = os.path.join("data", "backups")
PDF_OUTPUT_PATH = os.path.join("data", "Business_Summary_Report.pdf")
DASHBOARD_PATH = os.path.join("data", "dashboard.xlsx")
ORDER_SUMMARY_PATH = os.path.join("data", "order_summary.xlsx")
CHART_IMAGE_PATH = os.path.join("data", "top_products_chart.png")
SENDER_EMAIL = "stein3101@gmail.com"
SENDER_PASSWORD = "cmfg lcrj pyyd nnzt"
ALERT_EMAIL = "stein3101@gmail.com"
API_BASE_URL = "https://dummyjson.com/carts"
MAX_LOG_DAYS = 7
MAX_BACKUPS = 5
ALERTS_FILE_PATH = "alerts.xlsx"

# === TOOLS ===
def get_current_timestamp():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

def create_backup():
    os.makedirs(BACKUP_FOLDER, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_filename = f"orders_backup_{timestamp}.xlsx"
    backup_path = os.path.join(BACKUP_FOLDER, backup_filename)
    shutil.copy(EXCEL_FILE_PATH, backup_path)
    print(f"📁 Backup created: {backup_path}")
    return backup_path

def send_alert_email(subject, body):
    try:
        msg = EmailMessage()
        msg['Subject'] = subject
        msg['From'] = SENDER_EMAIL
        msg['To'] = ALERT_EMAIL
        msg.set_content(body)
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
            smtp.login(SENDER_EMAIL, SENDER_PASSWORD)
            smtp.send_message(msg)
        print("📧 Alert email sent successfully.")
    except Exception as e:
        print(f"❌ Failed to send alert email: {e}")

def cleanup_old_files():
    print("🧹 Cleaning old backups and logs...")
    if os.path.exists(LOG_FILE_PATH):
        with open(LOG_FILE_PATH, "r") as f:
            lines = f.readlines()
        new_lines = []
        for line in lines:
            parts = line.split(",")
            if len(parts) > 0:
                try:
                    timestamp = datetime.strptime(parts[0], "%Y-%m-%d %H:%M:%S")
                    if (datetime.now() - timestamp).days <= MAX_LOG_DAYS:
                        new_lines.append(line)
                except:
                    continue
        with open(LOG_FILE_PATH, "w") as f:
            f.writelines(new_lines)

    backups = sorted([
        os.path.join(BACKUP_FOLDER, f)
        for f in os.listdir(BACKUP_FOLDER)
        if f.endswith(".xlsx")
    ], key=os.path.getmtime, reverse=True)
    for old_file in backups[MAX_BACKUPS:]:
        os.remove(old_file)
        print(f"🗑️ Deleted old backup: {old_file}")

# === API ===
def get_orders():
    try:
        response = requests.get(API_BASE_URL)
        response.raise_for_status()
        return response.json().get("carts", [])
    except Exception as e:
        print(f"❌ API Error (GET): {e}")
        return []

def create_order(data):
    try:
        response = requests.post(API_BASE_URL, json=data)
        response.raise_for_status()
        return response.json()
    except Exception as e:
        print(f"❌ API Error (CREATE): {e}")

def update_order(order_id, data):
    try:
        response = requests.put(f"{API_BASE_URL}/{order_id}", json=data)
        response.raise_for_status()
        return response.json()
    except Exception as e:
        print(f"❌ API Error (UPDATE): {e}")

def delete_order(order_id):
    try:
        response = requests.delete(f"{API_BASE_URL}/{order_id}")
        response.raise_for_status()
        return response.status_code == 200
    except Exception as e:
        print(f"❌ API Error (DELETE): {e}")

def check_api_health():
    try:
        response = requests.get(API_BASE_URL, timeout=5)
        if response.status_code == 200:
            print("✅ API is reachable.")
        else:
            print(f"⚠️ API returned status code {response.status_code}")
    except Exception as e:
        print(f"❌ API Health Check Failed: {e}")

# === Excel ===
def load_orders_from_excel():
    return pd.read_excel(EXCEL_FILE_PATH, sheet_name='Orders')

def get_updates_from_excel():
    df = load_orders_from_excel()
    return df[df['update_required'].astype(str).str.lower() == 'yes']

def get_deletions_from_excel():
    df = load_orders_from_excel()
    return df[df['delete_flag'].astype(str).str.lower() == 'yes']

def get_new_orders_from_excel():
    df = load_orders_from_excel()
    return df[df['order_id'].isnull()]

def save_orders_to_excel(df_orders, df_items, df_orders_and_items):
    with pd.ExcelWriter(EXCEL_FILE_PATH, engine='openpyxl') as writer:
        df_orders.to_excel(writer, sheet_name='Orders', index=False)
        df_items.to_excel(writer, sheet_name='Order_Items', index=False)
        df_orders_and_items.to_excel(writer, sheet_name='Orders_and_Items', index=False)

def fetch_orders_from_api():
    try:
        print(f"[{get_current_timestamp()}] 📥 Fetching orders from API...")
        orders = get_orders()
        if not orders:
            raise Exception("No orders returned from API")

        create_backup()
        df_orders_raw = pd.DataFrame(orders)
        df_orders_raw.rename(columns={"id": "order_id"}, inplace=True)

        # שמירה של products כמחרוזת JSON בטבלת Orders
        df_orders_raw["products"] = df_orders_raw["products"].apply(lambda p: json.dumps(p, ensure_ascii=False))

        if 'update_required' not in df_orders_raw.columns:
            df_orders_raw['update_required'] = ""
        if 'delete_flag' not in df_orders_raw.columns:
            df_orders_raw['delete_flag'] = ""

        # === יצירת טבלת פריטים מפורקת ===
        order_items = []
        for _, row in df_orders_raw.iterrows():
            order_id = row["order_id"]
            try:
                products = json.loads(row["products"])
                for product in products:
                    order_items.append({
                        "order_id": order_id,
                        "product_id": product.get("id"),
                        "title": product.get("title"),
                        "price": product.get("price"),
                        "quantity": product.get("quantity"),
                        "total": product.get("total")
                    })
            except Exception as p_err:
                print(f"⚠️ Failed to parse products for order {order_id}: {p_err}")
                continue

        df_items = pd.DataFrame(order_items)
        df_orders_and_items = pd.merge(df_orders_raw, df_items, on='order_id', how='inner')
        save_orders_to_excel(df_orders_raw, df_items, df_orders_and_items)

        print(f"[{get_current_timestamp()}] ✅ Orders fetched and saved.")
        log_operation("Fetch Orders", "Success")
    except Exception as e:
        handle_error("Fetch Orders", e)


import json

def process_excel_changes(simulation_mode=False):
    try:
        print(f"[{get_current_timestamp()}] 🔍 Processing Excel changes...")

        df = pd.read_excel(EXCEL_FILE_PATH, sheet_name='Orders')

        updates = df[df['update_required'].astype(str).str.lower() == 'yes']
        deletions = df[df['delete_flag'].astype(str).str.lower() == 'yes']
        new_orders = df[df['order_id'].isnull()]

        print(f"🛠️ Updates: {len(updates)}, Deletions: {len(deletions)}, New Orders: {len(new_orders)}")

        if simulation_mode:
            print("🧪 Simulation mode: No changes sent to API.")
            return

        # === עדכון הזמנות קיימות ===
        for _, row in updates.iterrows():
            try:
                products = json.loads(row["products"]) if pd.notna(row["products"]) else []
                update_order(int(row["order_id"]), {
                    "total": row["total"],
                    "discountedTotal": row["discountedTotal"],
                    "totalProducts": row["totalProducts"],
                    "totalQuantity": row["totalQuantity"],
                    "products": products
                })
                log_operation("Update Order", "Success", int(row["order_id"]))
            except Exception as inner_e:
                handle_error("Update Order JSON", inner_e)

        # === מחיקת הזמנות ===
        for _, row in deletions.iterrows():
            delete_order(int(row["order_id"]))
            log_operation("Delete Order", "Success", int(row["order_id"]))

        # === יצירת הזמנות חדשות ===
        for _, row in new_orders.iterrows():
            try:
                products = json.loads(row["products"]) if pd.notna(row["products"]) else []
                create_order({
                    "userId": int(row["userId"]),
                    "products": products
                })
                log_operation("Create Order", "Success")
            except Exception as inner_e:
                handle_error("Create Order JSON", inner_e)

        print(f"[{get_current_timestamp()}] ✅ All Excel changes processed successfully.")
    except Exception as e:
        handle_error("Process Excel Changes", e)

def generate_order_summary():
    df = pd.read_excel(EXCEL_FILE_PATH, sheet_name='Orders_and_Items')
    summary = {
        "Total Orders": df['order_id'].nunique(),
        "Total Products Sold": df['quantity'].sum(),
        "Total Revenue": df['total_x'].sum(),
        "Average Order Value": df.groupby('order_id')['total_x'].first().mean(),
        "Average Products per Order": df.groupby('order_id')['product_id'].count().mean(),
        "Average Revenue per Product": df['total_y'].mean()

    }
    summary_df = pd.DataFrame([summary])
    with pd.ExcelWriter(ORDER_SUMMARY_PATH, engine='openpyxl') as writer:
        summary_df.to_excel(writer, sheet_name='Orders_And_Products_Summary', index=False)
    print(f"✅ order_summary.xlsx saved at {ORDER_SUMMARY_PATH}")


from openpyxl.styles import Font

def generate_order_summary_report():
    try:
        df = pd.read_excel(EXCEL_FILE_PATH, sheet_name='Orders_and_Items')
        summary = {
            "Total Orders": df['order_id'].nunique(),
            "Total Products Sold": df['quantity'].sum(),
            "Total Revenue": df['total_x'].sum(),
            "Average Order Value": df.groupby('order_id')['total_x'].first().mean(),
            "Average Products per Order": df.groupby('order_id')['product_id'].count().mean(),
            "Average Revenue per Product": df['total_y'].mean()
        }

        wb = Workbook()
        ws = wb.active
        ws.title = "Business_Summary"

        # Metadata
        ws.append(["Generated At:", get_current_timestamp()])
        ws.append([])

        # Header
        ws.append(["Metric", "Value"])
        for key, value in summary.items():
            ws.append([key, round(value, 2) if isinstance(value, (int, float)) else value])

        # עיצוב: הדגשת כותרות
        for cell in ws["A3:B3"][0]:
            cell.font = Font(bold=True)

        # שמירה
        output_path = os.path.join("data", "business_summary_report.xlsx")
        wb.save(output_path)
        print(f"✅ business_summary_report.xlsx saved at {output_path}")
        log_operation("Generate Order Summary Report", "Success")

    except Exception as e:
        handle_error("Generate Order Summary Report", e)


def run_process_monitoring_summary():
    try:
        print("📊 Generating process monitoring summary...")

        if not os.path.exists(LOG_FILE_PATH):
            raise Exception("Log file not found.")

        # קריאת לוג
        df_logs = pd.read_csv(LOG_FILE_PATH, names=["timestamp", "operation", "status", "order_id"])
        df_logs["timestamp"] = pd.to_datetime(df_logs["timestamp"])

        summary = df_logs.groupby(["operation", "status"]).size().reset_index(name="count")

        # ניתוח לפי משתמשים אם ניתן (אם order_id != null)
        df_by_user = df_logs[df_logs["order_id"].notna()]
        df_by_user = df_by_user.groupby(["operation", "order_id"]).size().reset_index(name="actions_per_order")

        with pd.ExcelWriter("data/monitoring_summary.xlsx", engine='openpyxl') as writer:
            summary.to_excel(writer, sheet_name="Activity_Summary", index=False)
            df_by_user.to_excel(writer, sheet_name="Order_Activity", index=False)

        print("✅ monitoring_summary.xlsx saved in data/")
        log_operation("Monitoring Summary", "Success")

    except Exception as e:
        handle_error("Monitoring Summary", e)


def generate_excel_dashboard():
    df = pd.read_excel(EXCEL_FILE_PATH, sheet_name='Orders_and_Items')

    orders_stats = df.groupby('order_id').agg({
        'total_x': 'first',
        'total_y': 'sum',
        'quantity': 'sum',
        'userId': 'first'
    }).reset_index().rename(columns={'total_y': 'total'})

    products_stats = df.groupby('title').agg({
        'quantity': 'sum',
        'total_y': 'sum'
    }).rename(columns={'total_y': 'total'}).sort_values(by='quantity', ascending=False).reset_index()

    plt.figure(figsize=(8,6))
    products_stats[:10].plot(kind='bar', x='title', y='quantity', legend=False)
    plt.title('Top 10 Products by Quantity')
    plt.xlabel('Product')
    plt.ylabel('Quantity Sold')
    plt.tight_layout()
    plt.savefig(CHART_IMAGE_PATH)
    plt.close()

    wb = Workbook()
    ws1 = wb.active
    ws1.title = "Orders_Stats"
    for r in dataframe_to_rows(orders_stats, index=False, header=True):
        ws1.append(r)

    ws2 = wb.create_sheet("Products_Stats")
    for r in dataframe_to_rows(products_stats, index=False, header=True):
        ws2.append(r)

    ws3 = wb.create_sheet("Charts")
    img = Image(CHART_IMAGE_PATH)
    img.width = 800
    img.height = 600
    ws3.add_image(img, "A1")

    ws4 = wb.create_sheet("Dashboard")
    ws4["A1"] = "Total Orders"
    ws4["B1"] = df['order_id'].nunique()
    ws4["A2"] = "Total Products Sold"
    ws4["B2"] = df['quantity'].sum()
    ws4["A3"] = "Total Revenue"
    ws4["B3"] = df['total_x'].sum()
    ws4["A4"] = "Avg Order Value"
    ws4["B4"] = round(df.groupby('order_id')['total_x'].first().mean(), 2)
    ws4["A5"] = "Avg Products per Order"
    ws4["B5"] = round(df.groupby('order_id')['product_id'].count().mean(), 2)

    wb.save(DASHBOARD_PATH)
    print(f"✅ dashboard.xlsx saved at {DASHBOARD_PATH}")
12

def generate_pdf_report():
    df = pd.read_excel(EXCEL_FILE_PATH, sheet_name='Orders_and_Items')
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", "B", 16)
    pdf.cell(0, 10, "Business Summary Report", ln=True, align="C")
    pdf.set_font("Arial", size=12)
    pdf.cell(0, 10, f"Total Orders: {df['order_id'].nunique()}", ln=True)
    pdf.cell(0, 10, f"Total Products Sold: {df['quantity'].sum()}", ln=True)
    pdf.cell(0, 10, f"Total Revenue: {df['total_x'].sum():,.2f}", ln=True)
    pdf.cell(0, 10, f"Avg Order Value: {df.groupby('order_id')['total_x'].first().mean():,.2f}", ln=True)
    pdf.cell(0, 10, f"Avg Products per Order: {df.groupby('order_id')['product_id'].count().mean():.2f}", ln=True)
    if os.path.exists(CHART_IMAGE_PATH):
        pdf.image(CHART_IMAGE_PATH, w=180)
    pdf.output(PDF_OUTPUT_PATH)
    print(f"✅ PDF saved at {PDF_OUTPUT_PATH}")

# === Alerts + Logs ===
def view_last_alerts():
    try:
        if os.path.exists(ALERTS_FILE_PATH):
            df = pd.read_excel(ALERTS_FILE_PATH)
            print(df.tail(5))
        else:
            print("⚠️ No alerts file found.")
    except Exception as e:
        print(f"❌ Failed to view alerts: {e}")

def log_operation(operation, status, order_id=None):
    timestamp = get_current_timestamp()
    os.makedirs(os.path.dirname(LOG_FILE_PATH), exist_ok=True)
    with open(LOG_FILE_PATH, "a") as f:
        f.write(f"{timestamp},{operation},{status},{order_id or ''}\n")

def log_alert(context, exception):
    timestamp = get_current_timestamp()
    print(f"[{timestamp}] ⚠️ ALERT: {context}: {exception}")
    alert = pd.DataFrame([{"timestamp": timestamp, "context": context, "error": str(exception)}])
    if os.path.exists(ALERTS_FILE_PATH):
        df = pd.read_excel(ALERTS_FILE_PATH)
        df = pd.concat([df, alert], ignore_index=True)
    else:
        df = alert
    df.to_excel(ALERTS_FILE_PATH, index=False)
    send_alert_email(f"ALERT: {context}", str(exception))

def handle_error(context, exception):
    log_operation(context, "ERROR")
    log_alert(context, exception)

# === Display Menu ===
def display_menu():
    print("="*55)
    print(" Supply Chain Orders Manager - Main Menu")
    print("="*55)
    print("1. Fetch Orders from API")
    print("2. Process Excel Changes")
    print("3. Create Manual Backup")
    print("4. Run Business Data Summary Report")
    print("5. Run Process Monitoring Summary Report")
    print("6. Health Check")
    print("7. View Last Alerts")
    print("8. Full Synchronization")
    print("9. Simulation Mode (Dry Run)")
    print("10. Clean Old Logs and Backups")
    print("11. 📄 Generate Order Summary Report (Excel)")
    print("12. 📊 Generate Excel Dashboard Report")
    print("13. 📝 Generate PDF Business Report")
    print("14. About / Help")
    print("15. 🚪 Exit")
    print("="*55)

# === Main Function ===
def main():
    while True:
        display_menu()
        choice = input("Select option (1–15): ").strip()
        try:
            if choice == "1":
                fetch_orders_from_api()
            elif choice == "2":
                process_excel_changes()
            elif choice == "3":
                create_backup()
            elif choice == "4":
                generate_order_summary()
            elif choice == "5":
                run_process_monitoring_summary()
            elif choice == "6":
                check_api_health()
            elif choice == "7":
                view_last_alerts()
            elif choice == "8":
                fetch_orders_from_api()
                process_excel_changes()
                generate_order_summary()
            elif choice == "9":
                process_excel_changes(simulation_mode=True)
            elif choice == "10":
                cleanup_old_files()
            elif choice == "11":
                generate_order_summary_report()
            elif choice == "12":
                generate_excel_dashboard()
            elif choice == "13":
                generate_pdf_report()


            elif choice == "14":
                print("📚 Supply Chain Orders Management System v1.0")
                print(f"Time: {get_current_timestamp()}")
                print("Developed by: Guy Stein")
                print("\n🔹 Description:")
                print("This tool automates the management, analysis, and reporting of supply chain order data.")
                print("\n🔹 Features:")
                print("- Bi-directional integration with dummy REST API")
                print("- Excel input/output with automated processing")
                print("- Summary reports, dashboards, PDF generation, monitoring logs")
                print("- Alerts, backups, health checks, and simulation mode")
                print("\n🔹 Technologies:")
                print("- Python 3.13, Pandas, OpenPyXL, Matplotlib, FPDF, SMTP, REST")
                print("\n🔹 Future Additions:")
                print("- Enhanced anomaly detection, real-time API integration, user login")

                help_lines = [
                    "📚 Supply Chain Orders Management System v1.0",
                    f"Time: {get_current_timestamp()}",
                    "Developed by: Guy Stein",
                    "",
                    "🔹 Description:",
                    "This tool automates the management, analysis, and reporting of supply chain order data.",
                    "",
                    "🔹 Features:",
                    "- Bi-directional integration with dummy REST API",
                    "- Excel input/output with automated processing",
                    "- Summary reports, dashboards, PDF generation, monitoring logs",
                    "- Alerts, backups, health checks, and simulation mode",
                    "",
                    "🔹 Technologies:",
                    "- Python 3.13, Pandas, OpenPyXL, Matplotlib, FPDF, SMTP, REST",
                    "",
                    "🔹 Future Additions:",
                    "- Enhanced anomaly detection, real-time API integration, user login"
                ]

                with open("data/help.txt", "w", encoding="utf-8") as f:
                    for line in help_lines:
                        f.write(line + "\n")


            elif choice == "15":
                print("👋 Goodbye!")
                sys.exit(0)
            else:
                print("❌ Invalid choice. Try again.")
            time.sleep(1)
        except Exception as e:
            handle_error("Main Menu", e)

# === Start Program ===
if __name__ == "__main__":
    main()
