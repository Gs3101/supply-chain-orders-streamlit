import os
import smtplib
import shutil
from datetime import datetime, timedelta
from email.message import EmailMessage
from config import (
    SENDER_EMAIL, SENDER_PASSWORD, ALERT_EMAIL,
    EXCEL_FILE_PATH, BACKUP_FOLDER, LOG_FILE_PATH,
    MAX_LOG_DAYS, MAX_BACKUPS
)

# === תאריך ושעה נוכחיים ===
def get_current_timestamp():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

# === יצירת גיבוי לקובץ אקסל ===
def create_backup():
    os.makedirs(BACKUP_FOLDER, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_filename = f"orders_backup_{timestamp}.xlsx"
    backup_path = os.path.join(BACKUP_FOLDER, backup_filename)
    shutil.copy(EXCEL_FILE_PATH, backup_path)
    print(f"📁 Backup created: {backup_path}")
    return backup_path

# === שליחת מייל התראה ===
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

# === ניקוי לוגים וגיבויים ישנים ===
def cleanup_old_files():
    print("🧹 Cleaning old backups and logs...")

    # ניקוי לוגים ישנים
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

    # ניקוי גיבויים ישנים
    backups = sorted([
        os.path.join(BACKUP_FOLDER, f)
        for f in os.listdir(BACKUP_FOLDER)
        if f.endswith(".xlsx")
    ], key=os.path.getmtime, reverse=True)
    for old_file in backups[MAX_BACKUPS:]:
        os.remove(old_file)
        print(f"🗑️ Deleted old backup: {old_file}")
