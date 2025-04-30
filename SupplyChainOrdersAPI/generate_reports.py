import os
import pandas as pd
import matplotlib.pyplot as plt
from openpyxl import Workbook
from openpyxl.drawing.image import Image
from openpyxl.utils.dataframe import dataframe_to_rows
from fpdf import FPDF
from config import (
    EXCEL_FILE_PATH, ORDER_SUMMARY_PATH, DASHBOARD_PATH,
    CHART_IMAGE_PATH, PDF_OUTPUT_PATH
)

# === דוח order_summary.xlsx ===
def generate_order_summary():
    df = pd.read_excel(EXCEL_FILE_PATH, sheet_name='Orders_and_Items')

    summary = {
        "Total Orders": df['order_id'].nunique(),
        "Total Products Sold": df['quantity'].sum(),
        "Total Revenue": df['total_x'].sum(),
        "Average Order Value": df.groupby('order_id')['total_x'].first().mean(),
        "Average Products per Order": df.groupby('order_id')['product_id'].count().mean(),
        "Average Revenue per Product": df['total'].mean()
    }

    summary_df = pd.DataFrame([summary])
    with pd.ExcelWriter(ORDER_SUMMARY_PATH, engine='openpyxl') as writer:
        summary_df.to_excel(writer, sheet_name='Orders_And_Products_Summary', index=False)

    print(f"✅ order_summary.xlsx saved at {ORDER_SUMMARY_PATH}")

# === דוח dashboard.xlsx עם גרף ===
def generate_excel_dashboard():
    df = pd.read_excel(EXCEL_FILE_PATH, sheet_name='Orders_and_Items')

    orders_stats = df.groupby('order_id').agg({
        'total_x': 'first',
        'total': 'sum',
        'quantity': 'sum',
        'userId': 'first'
    }).reset_index()

    products_stats = df.groupby('title').agg({
        'quantity': 'sum',
        'total': 'sum'
    }).sort_values(by='quantity', ascending=False).reset_index()

    # גרף עמודות
    plt.figure(figsize=(8,6))
    products_stats[:10].plot(kind='bar', x='title', y='quantity', legend=False)
    plt.title('Top 10 Products by Quantity')
    plt.xlabel('Product')
    plt.ylabel('Quantity Sold')
    plt.tight_layout()
    plt.savefig(CHART_IMAGE_PATH)
    plt.close()

    # כתיבה לקובץ Excel
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

# === יצירת PDF ===
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

    pdf.cell(0, 10, f"Average Order Value: {df.groupby('order_id')['total_x'].first().mean():,.2f}", ln=True)
    pdf.cell(0, 10, f"Average Products per Order: {df.groupby('order_id')['product_id'].count().mean():.2f}", ln=True)

    if os.path.exists(CHART_IMAGE_PATH):
        pdf.image(CHART_IMAGE_PATH, w=180)

    pdf.output(PDF_OUTPUT_PATH)
    print(f"✅ PDF saved at {PDF_OUTPUT_PATH}")
