import streamlit as st
import pandas as pd
from datetime import datetime
from main import (
    fetch_orders_from_api, process_excel_changes, create_backup,
    generate_order_summary, run_process_monitoring_summary,
    check_api_health, view_last_alerts, cleanup_old_files,
    generate_order_summary_report, generate_excel_dashboard,
    generate_pdf_report
)

# הגדרת עמוד
st.set_page_config(page_title="Supply Chain Orders App", layout="wide")

st.title("📦 Supply Chain Orders Management System")
st.caption("Developed by Guy Stein | Version 1.0")
st.divider()

status = st.empty()

# פונקציית הרצה עם תצוגת סטטוס
def run(label, func, *args, **kwargs):
    with st.spinner(f"🔄 Running: {label}..."):
        try:
            func(*args, **kwargs)
            status.success(f"✅ {label} completed successfully!")
        except Exception as e:
            status.error(f"❌ Error in {label}: {e}")

# תפריט צד - מופרד לפי קטגוריות
with st.sidebar:
    st.title("📋 Menu")

    st.subheader("🔧 Operations")
    if st.button("📥 Fetch Orders from API"):
        run("Fetch Orders", fetch_orders_from_api)
    if st.button("📄 Process Excel Changes"):
        run("Process Excel Changes", process_excel_changes)
    if st.button("💾 Create Manual Backup"):
        run("Create Backup", create_backup)

    st.subheader("📈 Reports & Analysis")
    if st.button("📊 Run Business Summary Report"):
        run("Generate Summary", generate_order_summary)
    if st.button("📋 Run Monitoring Summary"):
        run("Monitoring Summary", run_process_monitoring_summary)
    if st.button("📄 Generate Order Summary Report"):
        run("Order Summary Report", generate_order_summary_report)
    if st.button("📈 Generate Dashboard"):
        run("Dashboard", generate_excel_dashboard)
    if st.button("📝 Generate PDF Report"):
        run("PDF Report", generate_pdf_report)

    st.subheader("🧰 System & Monitoring")
    if st.button("🩺 API Health Check"):
        run("API Health Check", check_api_health)
    if st.button("🔍 View Last Alerts"):
        try:
            df = pd.read_excel("alerts.xlsx")
            st.subheader("🔔 Last Alerts")
            st.dataframe(df.tail(10))
        except Exception as e:
            st.warning(f"⚠️ Could not read alerts file: {e}")
    if st.button("🧪 Simulation Mode (No Save)"):
        run("Simulation Mode", process_excel_changes, simulation_mode=True)
    if st.button("🧹 Clean Old Logs / Backups"):
        run("Cleanup", cleanup_old_files)

    st.subheader("ℹ️ System")
    if st.button("ℹ️ About / Help"):
        st.markdown(f"""
        **Supply Chain Orders App v1.0**

        - Author: Guy Stein  
        - Date: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}  
        - Developed using Python 3.13 + Streamlit  
        - Features: API integration, Excel automation, alerts, dashboards and more
        """)

    if st.button("🚪 Exit"):
        st.warning("You can close this tab or stop the app from the terminal.")
