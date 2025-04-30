import pandas as pd
from config import EXCEL_FILE_PATH

# === טעינת גיליון Orders ===
def load_orders_from_excel():
    return pd.read_excel(EXCEL_FILE_PATH, sheet_name='Orders')

# === קבלת שורות לעדכון (update_required = yes) ===
def get_updates_from_excel():
    df = load_orders_from_excel()
    return df[df['update_required'].astype(str).str.lower() == 'yes']

# === קבלת שורות למחיקה (delete_flag = yes) ===
def get_deletions_from_excel():
    df = load_orders_from_excel()
    return df[df['delete_flag'].astype(str).str.lower() == 'yes']

# === קבלת שורות חדשות (order_id חסר) ===
def get_new_orders_from_excel():
    df = load_orders_from_excel()
    return df[df['order_id'].isnull()]

# === שמירת שלושת הגיליונות לאקסל ===
def save_orders_to_excel(df_orders, df_items, df_orders_and_items):
    with pd.ExcelWriter(EXCEL_FILE_PATH, engine='openpyxl') as writer:
        df_orders.to_excel(writer, sheet_name='Orders', index=False)
        df_items.to_excel(writer, sheet_name='Order_Items', index=False)
        df_orders_and_items.to_excel(writer, sheet_name='Orders_and_Items', index=False)
