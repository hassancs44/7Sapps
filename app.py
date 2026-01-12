from flask import Flask, request, jsonify, render_template, send_from_directory, session, redirect
from flask_cors import CORS
import pandas as pd
import os
from datetime import datetime
import requests
import re
# ================================================================
# 🧱 قاعدة البيانات الذكية SEVENS (مزامنة ثنائية Excel ↔ SQLite)
# ================================================================
import sqlite3
import pandas as pd
import os
import time
import threading

import os
import ssl
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from dotenv import load_dotenv
import os
import ssl
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from dotenv import load_dotenv

load_dotenv()

SMTP_SERVER = os.getenv("SMTP_SERVER", "mail.sevens.sa")
SMTP_USER = os.getenv("SMTP_USER", "ticket.support@sevens.sa")
SMTP_PASSWORD = os.getenv("SMTP_PASSWORD")
SMTP_PORT = int(os.getenv("SMTP_PORT", "587"))

def send_html_email_via_company(to_email: str, subject: str, html_body: str):
    if not SMTP_PASSWORD:
        raise Exception("SMTP_PASSWORD is not set in .env")

    msg = MIMEMultipart("alternative")
    msg["To"] = to_email
    msg["From"] = f"SEVENS System <{SMTP_USER}>"
    msg["Subject"] = subject
    msg.attach(MIMEText(html_body, "html", "utf-8"))

    try:
        print(f"[SMTP] Connecting to {SMTP_SERVER}:{SMTP_PORT} as {SMTP_USER}")
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT, timeout=20) as server:
            server.ehlo()
            context = ssl.create_default_context()
            server.starttls(context=context)
            server.ehlo()
            server.login(SMTP_USER, SMTP_PASSWORD)
            server.sendmail(SMTP_USER, [to_email], msg.as_string())
        print(f"[SMTP] ✅ Email sent to {to_email}")
        return True

    except Exception as e:
        print("[SMTP] ❌ Failed to send email")
        import traceback
        traceback.print_exc()
        raise


BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_SQLITE = os.path.join(BASE_DIR, "data", "sevens.db")
DATA_DIR = os.path.join(BASE_DIR, "data")
USERS_XLSX = os.path.join(DATA_DIR, "database.xlsx")
REQUESTS_XLSX = os.path.join(DATA_DIR, "requests.xlsx")
CHATS_XLSX = os.path.join(DATA_DIR, "chat_messages.xlsx")
# ✅ تعريف المسارات الأساسية في أعلى الملف
DATA_DIR = os.path.join(BASE_DIR, "data")
os.makedirs(DATA_DIR, exist_ok=True)

DB_PATH = os.path.join(DATA_DIR, "database.xlsx")         # ملف المستخدمين
REQUESTS_PATH = os.path.join(DATA_DIR, "requests.xlsx")   # ملف الطلبات
CHAT_PATH = os.path.join(DATA_DIR, "chat_messages.xlsx")  # ملف دردشات الطلبات
REQUESTS_SHEET = "الطلبات جميع"


# ==== إنشاء الجداول الأساسية ====
def init_sqlite():
    conn = sqlite3.connect(DB_SQLITE)
    c = conn.cursor()

    c.execute("""
    CREATE TABLE IF NOT EXISTS users (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT,
        role TEXT,
        password TEXT,
        email TEXT UNIQUE,
        department TEXT,
        status TEXT DEFAULT 'نشط',
        created_at TEXT DEFAULT CURRENT_TIMESTAMP
    )
    """)

    c.execute("""
    CREATE TABLE IF NOT EXISTS requests (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        req_id TEXT UNIQUE,
        date TEXT,
        title TEXT,
        description TEXT,
        sender_dept TEXT,
        receiver_dept TEXT,
        status TEXT,
        assigned_to TEXT,
        updated_by TEXT,
        duration TEXT,
        file_name TEXT,
        created_at TEXT DEFAULT CURRENT_TIMESTAMP
    )
    """)

    c.execute("""
    CREATE TABLE IF NOT EXISTS chats (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        req_id TEXT,
        sender_name TEXT,
        department TEXT,
        message TEXT,
        file_name TEXT,
        timestamp TEXT DEFAULT CURRENT_TIMESTAMP
    )
    """)

    c.execute("""
    CREATE TABLE IF NOT EXISTS logs (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        event TEXT,
        user TEXT,
        department TEXT,
        details TEXT,
        created_at TEXT DEFAULT CURRENT_TIMESTAMP
    )
    """)

    conn.commit()
    conn.close()
    print("✅ SQLite structure ready")


# ==== تشغيل أولي ====
if not os.path.exists(DB_SQLITE):
    print("🆕 Creating SEVENS database...")
    init_sqlite()

else:
    print("ℹ️ SEVENS database found — syncing now...")


# 🔹 المسار الحالي (جذر المشروع)
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# 🔹 تعريف Flask مع المسار الصحيح للقوالب
app = Flask(
    __name__,
    template_folder="templates",
    static_folder="static"
)
app.secret_key = "SEVENS-SECRET-2025"
CORS(app, resources={r"/api/*": {"origins": "*"}})


# ✅ تعريف المجلد الأساسي للمشروع
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# ========== 🧩 Google Drive Backup Integration ==========
from googleapiclient.discovery import build
from google.oauth2 import service_account
from googleapiclient.http import MediaFileUpload, MediaIoBaseDownload
import json

CONFIG_PATH = os.path.join(BASE_DIR, "config.json")

def load_config():
    """تحميل إعدادات النسخ الاحتياطي"""
    if not os.path.exists(CONFIG_PATH):
        default_conf = {
            "backup_mode": "local",  # أو "drive"
            "google_drive_folder_id": "",
            "service_key_path": os.path.join("data", "sevens-service-key.json")
        }
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(default_conf, f, ensure_ascii=False, indent=2)
        return default_conf
    with open(CONFIG_PATH, "r", encoding="utf-8") as f:
        return json.load(f)

CONFIG = load_config()
# ✅ فحص اتصال Google Drive عند التشغيل
try:
    key_data = os.environ.get("GOOGLE_SERVICE_KEY", "").strip()
    if key_data:
        creds = service_account.Credentials.from_service_account_info(
            json.loads(key_data),
            scopes=["https://www.googleapis.com/auth/drive"]
        )

        service = build("drive", "v3", credentials=creds)
        about = service.about().get(fields="user").execute()
        user_email = about["user"]["emailAddress"]
        print(f"✅ Google Drive connected successfully as: {user_email}")
    else:
        print("⚠️ GOOGLE_SERVICE_KEY not found (Drive backup disabled).")
except Exception as e:
    print("❌ Google Drive connection test failed:", e)

def upload_to_drive(file_path):
    """رفع ملف إلى Google Drive (مع نقل الملكية إلى صاحب الحساب الحقيقي)"""
    try:
        if CONFIG.get("backup_mode") != "drive":
            print("🟡 النسخ الاحتياطي المحلي مفعل (لم يتم الرفع إلى Drive).")
            return

        key_data = os.environ.get("GOOGLE_SERVICE_KEY", "").strip()
        if not key_data:
            print("⚠️ لم يتم العثور على GOOGLE_SERVICE_KEY في البيئة.")
            return

        service_key = json.loads(key_data)
        creds = service_account.Credentials.from_service_account_info(
            service_key,
            scopes=["https://www.googleapis.com/auth/drive"]
        )
        service = build("drive", "v3", credentials=creds)

        folder_id = CONFIG.get("google_drive_folder_id")
        file_name = os.path.basename(file_path)
        file_metadata = {"name": file_name, "parents": [folder_id]}
        media = MediaFileUpload(file_path, resumable=True)

        # رفع الملف
        uploaded = service.files().create(
            body=file_metadata,
            media_body=media,
            fields="id"
        ).execute()

        file_id = uploaded.get("id")

        # نقل الملكية إلى حسابك الشخصي
        service.permissions().create(
            fileId=file_id,
            body={
                "type": "user",
                "role": "owner",
                "emailAddress": "sevensitapp@gmail.com"
            },
            transferOwnership=True
        ).execute()

        print(f"✅ Backup uploaded & transferred ownership: {file_name}")
    except Exception as e:
        print("❌ upload_to_drive error:", e)

# ✅ إنشاء مجلد الرفع
UPLOAD_DIR = os.path.join(BASE_DIR, "uploads")
os.makedirs(UPLOAD_DIR, exist_ok=True)

# ✅ مسار ملف دردشة الطلبات
CHAT_PATH = os.path.join(BASE_DIR, "chat_messages.xlsx")

UPLOAD_DIR = os.path.join(BASE_DIR, "uploads")
os.makedirs(UPLOAD_DIR, exist_ok=True)

# ✅ مسار ملف دردشة الطلبات
CHAT_PATH = os.path.join(BASE_DIR, "chat_messages.xlsx")

from googleapiclient.errors import HttpError

def download_from_drive(file_name, local_path):
    """📥 تنزيل ملف من Google Drive إلى السيرفر"""
    try:
        key_data = os.environ.get("GOOGLE_SERVICE_KEY", "").strip()
        if not key_data:
            print("⚠️ لم يتم العثور على GOOGLE_SERVICE_KEY في البيئة.")
            return False

        service_key = json.loads(key_data)
        creds = service_account.Credentials.from_service_account_info(
            service_key,
            scopes=["https://www.googleapis.com/auth/drive"]
        )
        service = build("drive", "v3", credentials=creds)

        folder_id = CONFIG.get("google_drive_folder_id")
        query = f"'{folder_id}' in parents and name='{file_name}' and trashed=false"

        results = service.files().list(q=query, fields="files(id, name, modifiedTime)").execute()
        files = results.get("files", [])
        if not files:
            print(f"⚠️ لم يتم العثور على الملف {file_name} في Google Drive.")
            return False

        # 📄 تنزيل أحدث نسخة (الأحدث من حيث وقت التعديل)
        file_id = sorted(files, key=lambda x: x["modifiedTime"], reverse=True)[0]["id"]

        request = service.files().get_media(fileId=file_id)
        with open(local_path, "wb") as f:
            downloader = MediaIoBaseDownload(f, request)
            done = False
            while not done:
                status, done = downloader.next_chunk()
                if status:
                    print(f"⬇️ تحميل {file_name}: {int(status.progress() * 100)}%")
        print(f"✅ تم تنزيل الملف: {file_name}")
        return True

    except HttpError as e:
        print(f"❌ Google API error أثناء تنزيل {file_name}: {e}")
        return False
    except Exception as e:
        print(f"❌ خطأ أثناء تنزيل {file_name}: {e}")
        return False


def load_chats():
    """تحميل سجل دردشات الطلبات من ملف Excel أو إنشاؤه إن لم يوجد"""
    if not os.path.exists(CHAT_PATH):
        df = pd.DataFrame(columns=['رقم الطلب', 'المرسل', 'القسم', 'الرسالة', 'الملف', 'الوقت'])
        df.to_excel(CHAT_PATH, index=False)
        print("✅ Created chat_messages.xlsx")
        return df
    try:
        df = pd.read_excel(CHAT_PATH)
        # تنظيف الأعمدة وتوحيد الأسماء
        df.columns = [str(c).strip() for c in df.columns]
        for col in ['رقم الطلب', 'المرسل', 'القسم', 'الرسالة', 'الملف', 'الوقت']:
            if col not in df.columns:
                df[col] = ''
        return df
    except Exception as e:
        print("❌ load_chats error:", e)
        return pd.DataFrame(columns=['رقم الطلب', 'المرسل', 'القسم', 'الرسالة', 'الملف', 'الوقت'])

def normalize_arabic(text):
    if not isinstance(text, str):
        text = str(text)

    text = text.strip()

    # إزالة مخفي
    text = text.replace('\u200f','').replace('\u200e','')

    # توحيد ألف والهمزات
    text = re.sub(r'[إأآا]', 'ا', text)


    # تاء مربوطة
    text = text.replace('ة','ة')

    # كلمة إدارة
    text = text.replace('اداره','ادارة')
    text = text.replace('ادارة','ادارة')
    text = text.replace('ادره','ادارة')
    text = text.replace('الاداره','ادارة')
    text = text.replace('الادارة','ادارة')

    return text

# ============== إعدادات عامة ==============
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
REQUESTS_SHEET = "الطلبات جميع"
EXPORT_DIR = os.path.join(BASE_DIR, "exports")
os.makedirs(EXPORT_DIR, exist_ok=True)

## مفتاح واجهة OpenRouter API  (احصل عليه من https://openrouter.ai)
OPENROUTER_API_KEY = "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"

# ✅ استرجاع النسخ الاحتياطية قبل التشغيل (في حال الملفات مفقودة)
for fpath, fname in [
    (DB_PATH, "database.xlsx"),
    (REQUESTS_PATH, "requests.xlsx"),
    (CHAT_PATH, "chat_messages.xlsx"),
]:
    if not os.path.exists(fpath):
        print(f"📥 الملف {fname} مفقود، سيتم استرجاعه من Google Drive...")
        download_from_drive(fname, fpath)

# ============== دوال مساعدة ==============
def ensure_excel_exists():
    if not os.path.exists(DB_PATH):
        users_cols = ['الاسم', 'الصلاحية', 'كلمة المرور', 'البريد الإلكتروني', 'القسم']
        pd.DataFrame(columns=users_cols).to_excel(DB_PATH, index=False)
        print("✅ Created users DB")

    if not os.path.exists(REQUESTS_PATH):
        req_cols = ['رقم الطلب', 'التاريخ', 'العنوان', 'الوصف', 'القسم المرسل',
                    'القسم المستلم', 'الحالة', 'الموظف المعين', 'آخر تحديث بواسطة', 'الوقت', 'الملف']
        pd.DataFrame(columns=req_cols).to_excel(REQUESTS_PATH, index=False, sheet_name=REQUESTS_SHEET)
        print("✅ Created requests DB")
    else:
        print("📂 Excel files already exist ✅")

# ✅ استدعِها مرة واحدة عند بدء التشغيل
ensure_excel_exists()


def normalize_columns(df):
    def clean(name):
        name = str(name).strip()
        name = name.replace("\u200f", "").replace("\u200e", "")

        # ❗ لا نغير كلمة "مؤرشف" نهائياً
        if name.replace(" ", "") in ["مؤرشف", "مؤرشفه", "ارشيف", "الارشيف"]:
            return "مؤرشف"

        # باقي الأعمدة فقط
        name = name.replace("إ", "ا").replace("أ", "ا").replace("آ", "ا")
        name = name.replace("ـ", "")
        name = name.replace("  ", " ")
        return name.strip()

    df.columns = [clean(c) for c in df.columns]
    return df



def load_users():
    try:
        df = pd.read_excel(DB_PATH)

        # إزالة الأعمدة المكررة
        df = remove_duplicate_columns(df)
        df = pd.read_excel(DB_PATH)

        # ⭐ عمود إجبار تغيير كلمة المرور لأول مرة
        if "force_reset" not in df.columns:
            df["force_reset"] = df["force_reset"].astype(str)
            df["force_reset"] = (
                df["force_reset"]
                .str.replace(".0", "", regex=False)
                .str.replace(".00", "", regex=False)
                .str.strip()
            )

            df["force_reset"] = "1"

        # 🔹 تنظيف الأعمدة من أي رموز أو فراغات غريبة
        df.columns = (
            df.columns
            .astype(str)
            .str.replace('\u200f', '', regex=True)
            .str.replace('\u200e', '', regex=True)
            .str.replace(' ', '', regex=True)
            .str.strip()
        )

        # ✅ توحيد أسماء الأعمدة مهما كانت كتابتها
        rename_map = {
            'الاسم': 'الاسم',
            'الاسمالكامل': 'الاسم',
            'الاسم_الكامل': 'الاسم',
            'الا سم': 'الاسم',
            'الإسم': 'الاسم',
            'اسم': 'الاسم',

            'البريدالإلكتروني': 'البريد الإلكتروني',
            'البريدالالكتروني': 'البريد الإلكتروني',
            'البريدالالكترونى': 'البريد الإلكتروني',
            'الايميل': 'البريد الإلكتروني',
            'email': 'البريد الإلكتروني',
            'ايميل': 'البريد الإلكتروني',

            'القسم': 'القسم',
            'القسم_الموظف': 'القسم',
            'ادارة': 'القسم',

            'الصلاحيه': 'الصلاحية',
            'الوظيفة': 'الصلاحية',
            'role': 'الصلاحية'
        }

        # 🧩 إعادة التسمية بناءً على التطابق الجزئي (حتى لو ناقص حرف)
        for col in list(df.columns):
            normalized = re.sub(r'[إأآا]', 'ا', col).replace(' ', '').lower()
            for k, v in rename_map.items():
                if re.sub(r'[إأآا]', 'ا', k).replace(' ', '').lower() in normalized:
                    df.rename(columns={col: v}, inplace=True)

        # ✅ التأكد أن كل الأعمدة المهمة موجودة حتى لو ناقصة
        for col in [
            'الاسم',
            'البريد الإلكتروني',
            'القسم',
            'الصلاحية',
            'كلمة المرور',
            'الشركة',
            'الفرع'
        ]:

            if col not in df.columns:
                df[col] = ''
        # ⭐ إضافة عمود الأقسام الأخرى إذا غير موجود
        if "الأقسام الأخرى" not in df.columns:
            df["الأقسام الأخرى"] = ""
        else:
            df["الأقسام الأخرى"] = df["الأقسام الأخرى"].astype(str).fillna("").str.strip()

        # ⭐ إضافة ودعم عمود apps (صلاحيات الأنظمة)
        if "apps" not in df.columns:
            df["apps"] = ""
        else:
            df["apps"] = df["apps"].astype(str).fillna("").str.strip()

        return normalize_department_names(df)
    except Exception as e:
        print("❌ load_users error:", e)
        return pd.DataFrame()

def get_user_all_departments():
    """إرجاع كل أقسام المستخدم (القسم الأساسي + الأقسام الأخرى) بعد التطبيع"""
    try:
        users_df = load_users()
        if users_df.empty:
            return []

        # استخراج أعمدة البريد - القسم - الأقسام الأخرى
        email_col = next((c for c in users_df.columns if "بريد" in c or "email" in c.lower()), None)
        dept_col  = next((c for c in users_df.columns if "قسم" in c), None)
        extra_col = next((c for c in users_df.columns if "أخرى" in c or "اخرى" in c), None)

        if not email_col or not dept_col:
            return []

        # المستخدم الحالي
        user_email = session.get("user", {}).get("email", "").strip().lower()
        users_df[email_col] = users_df[email_col].astype(str).str.lower().str.strip()

        row = users_df[users_df[email_col] == user_email]
        if row.empty:
            return []

        row = row.iloc[0]

        # -----------------------------
        # القسم الأساسي
        # -----------------------------
        all_depts = []
        main_dept = str(row.get(dept_col, "")).strip()
        if main_dept:
            all_depts.append(normalize_arabic(main_dept))

        # -----------------------------
        # الأقسام الأخرى (مقسمة بفواصل)
        # -----------------------------
        if extra_col:
            raw_extra = str(row.get(extra_col, "")).strip()

            if raw_extra:
                raw_extra = raw_extra.replace("\u200f", "").replace("\u200e", "")
                raw_extra = raw_extra.replace(" ،", "،").replace("، ", "،")
                raw_extra = raw_extra.replace(" ,", ",").replace(", ", ",")

                raw_extra = re.sub(r"\s*,\s*", ",", raw_extra)
                raw_extra = re.sub(r"\s*،\s*", "،", raw_extra)

                raw_extra = raw_extra.replace("،", ",")

                parts = [p.strip() for p in raw_extra.split(",") if p.strip()]
                for p in parts:
                    all_depts.append(normalize_arabic(p))

        # إزالة التكرار
        all_depts = list(dict.fromkeys(all_depts))

        return all_depts

    except Exception as e:
        print("get_user_all_departments error:", e)
        return []


def normalize_department_names(df):
    """توحيد أسماء الأقسام داخل قاعدة المستخدمين"""
    if 'القسم' in df.columns:
        df['القسم'] = (
            df['القسم']
            .astype(str)
            .str.strip()
            .str.replace('\u200f','', regex=True)
            .str.replace('\u200e','', regex=True)
            .str.replace('  ',' ', regex=True)
            .str.replace('الادارة','إدارة', regex=False)
        )
    return df
def remove_duplicate_columns(df):
    """إزالة الأعمدة المكررة بعد تنظيفها"""
    seen = set()
    new_cols = []
    drop_idx = []

    for idx, col in enumerate(df.columns):
        clean = (
            str(col)
            .strip()
            .replace("\u200f", "")
            .replace("\u200e", "")
        )
        clean = re.sub(r"[إأآا]", "ا", clean)

        if clean in seen:
            drop_idx.append(idx)
        else:
            seen.add(clean)
            new_cols.append(col)

    # حذف الأعمدة المكررة
    if drop_idx:
        df = df.drop(df.columns[drop_idx], axis=1)

    return df

def load_requests():
    try:
        if not os.path.exists(REQUESTS_PATH):
            return pd.DataFrame()

        df = pd.read_excel(REQUESTS_PATH, dtype=str)

        # ============================================================
        # 1) إزالة الأعمدة المكررة (قبل أي شيء)
        # ============================================================
        df = remove_duplicate_columns(df)

        # ============================================================
        # 2) تنظيف أسماء الأعمدة من المحارف الخفية
        # ============================================================
        df.columns = (
            df.columns
            .str.strip()
            .str.replace("\u200f", "")
            .str.replace("\u200e", "")
        )

        # ============================================================
        # 3) خريطة توحيد أسماء الأعمدة
        # ============================================================
        rename_map = {
            "الحاله": "الحالة",
            "اخر تحديث بواسطه": "آخر تحديث بواسطة",
            "اخر تحديث بواسطة": "آخر تحديث بواسطة",
            "اخر تحديث": "آخر تحديث بواسطة",
            "بدا التنفيذ بواسطه": "بدأ التنفيذ بواسطة",
            "اغلق بواسطه": "أغلق بواسطة",
            "القسم المستلم ": "القسم المستلم",
            "القسم المرسل ": "القسم المرسل",
        }

        for old, new in rename_map.items():
            if old in df.columns:
                df.rename(columns={old: new}, inplace=True)

        # ============================================================
        # 4) إزالة الأعمدة المكررة مرة ثانية بعد الدمج
        # ============================================================
        df = df.loc[:, ~df.columns.duplicated()]

        # ============================================================
        # 5) تنظيف محتوى الخلايا
        # ============================================================
        df = df.applymap(lambda x: x.strip() if isinstance(x, str) else x)

        # ============================================================
        # 6) ضمان وجود عمود "مؤرشف"
        # ============================================================
        if "مؤرشف" not in df.columns:
            df["مؤرشف"] = "0"

        df["مؤرشف"] = df["مؤرشف"].astype(str).apply(
            lambda x: "1" if str(x).strip().lower() in ["1", "نعم", "true", "yes", "y"] else "0"
        )

        return df

    except Exception as e:
        print("load_requests error:", e)
        return pd.DataFrame()

def save_requests(df):
    # 1) إزالة الأعمدة المكررة قبل أي شيء
    df = remove_duplicate_columns(df)

    # 2) تطبيع أسماء الأعمدة
    df = normalize_columns(df)

    # 3) خريطة تصحيح أسماء الأعمدة لمنع تكرارها
    rename_map = {
        "الحاله": "الحالة",
        "اخر تحديث بواسطه": "آخر تحديث بواسطة",
        "اخر تحديث بواسطة": "آخر تحديث بواسطة",
        "بدا التنفيذ بواسطه": "بدأ التنفيذ بواسطة",
        "اغلق بواسطه": "أغلق بواسطة",
        "القسم المرسل ": "القسم المرسل",
        "القسم المستلم ": "القسم المستلم",
    }

    for old, new in rename_map.items():
        if old in df.columns:
            df.rename(columns={old: new}, inplace=True)

    # 4) إزالة الأعمدة المكررة مرة ثانية بعد التوحيد
    df = df.loc[:, ~df.columns.duplicated()]

    # 5) التأكد أن كل الأعمدة الأساسية موجودة
    required_cols = [
        'رقم الطلب', 'التاريخ', 'العنوان', 'الوصف',
        'القسم المرسل', 'القسم المستلم', 'الحالة',
        'الموظف المعين', 'آخر تحديث بواسطة',
        'بدأ التنفيذ بواسطة', 'أغلق بواسطة',
        'الوقت', 'الملف',
        'اسم المرسل', 'اسم المستلم',
        'مؤرشف'
    ]

    for col in required_cols:
        if col not in df.columns:
            df[col] = ""

    # 6) ترتيب الأعمدة بنفس ترتيب required_cols لمنع أي فوضى
    df = df[required_cols]

    # 7) حفظ الملف بدون إعادة إنتاج أعمدة مكررة
    df.to_excel(REQUESTS_PATH, index=False, sheet_name=REQUESTS_SHEET)

def generate_request_id():
    df = load_requests()
    if df.empty or 'رقم الطلب' not in df.columns or df['رقم الطلب'].dropna().empty:
        return f"REQ-{datetime.now().year}-001"
    try:
        last_id = str(df['رقم الطلب'].dropna().iloc[-1])
        number = int(last_id.split('-')[-1]) + 1
        return f"REQ-{datetime.now().year}-{number:03}"
    except:
        return f"REQ-{datetime.now().year}-001"

# ============== الصفحات ==============
@app.route('/')
def index(): return render_template('Login.html')

@app.route('/Login.html')
def login_page(): return render_template('Login.html')

@app.route('/EmployeePage.html')
def emp_page(): return render_template('EmployeePage.html')

@app.route('/DepartmentManagerPage.html')
def mgr_page(): return render_template('DepartmentManagerPage.html')

@app.route('/GeneralManager.html')
def gm_page(): return render_template('GeneralManager.html')
@app.route('/HrPage.html')
def hr_page():
    return render_template('HrPage.html')
@app.route('/ForgotYourPassword.html')
def forgot_page(): return render_template('ForgotYourPassword.html')

@app.route('/Portal.html')
def portal_page():
    return render_template('Portal.html')


@app.route("/api/portal/apps")
def portal_apps():
    user = session.get("user")
    if not user:
        return jsonify({"apps": []})

    return jsonify({
        "apps": user.get("apps", [])
    })


@app.route('/admin.html')
def admin_page():
    return render_template('admin.html')
# ============== API: الدخول ==============
def normalize_role(text):
    if not isinstance(text, str):
        text = str(text)

    t = text.strip()
    t = re.sub(r'[إأآا]', 'ا', t)
    t = t.replace('ة', 'ه')
    t = t.replace('  ', '')
    t = t.replace('–', '-').replace('_', '').replace('/', '')
    t = t.replace('ال', '')
    t = t.replace(' ', '').lower()

    # مدير عام
    if t in ["مديرعام", "مديرعامه", "جنرال", "generalmanager", "gm"]:
        return "general_manager"

    # مدير قسم
    if any(x in t for x in [
        "مديرقسم",
        "مديرالقسم",
        "رئيسقسم",
        "manager",
        "head",
    ]):
        return "manager"

    # الموارد البشرية

    if any(x in t for x in [
        "مواردبشريه",
        "مواردبشرية",
        "المواردالبشريه",
        "المواردالبشرية",
        "hr",
        "humanresource"
    ]):
        return "hr"

    # أدمن
    if t in ["ادمن", "admin", "مشرف", "مديرنظام"]:
        return "admin"

    # موظف
    if t in ["موظف", "عامل", "staff", "employee"]:
        return "employee"

    return t


from flask import session
app.secret_key = "SEVENS-SECRET-2025"

@app.route('/api/auth/session_check', methods=['GET'])
def session_check():
    """يتأكد أن الجلسة صحيحة بعد أي تعديل"""
    user = session.get("user")
    if not user:
        return jsonify({"valid": False})

    df = load_users()

    email_col = next((c for c in df.columns if "بريد" in c or "email" in c.lower()), None)
    role_col  = next((c for c in df.columns if "صلاح" in c or "role" in c.lower()), None)
    dept_col  = next((c for c in df.columns if "قسم" in c), None)
    name_col  = next((c for c in df.columns if "اسم" in c), None)
    status_col= next((c for c in df.columns if "حال" in c), None)

    df[email_col] = df[email_col].astype(str).str.lower().str.strip()
    row = df[df[email_col] == user["email"]]

    if row.empty:
        session.clear()
        return jsonify({"valid": False})

    row = row.iloc[0]

    # تحقق من التغييرات
    new_role = normalize_role(str(row[role_col]))
    new_dept = str(row[dept_col]).strip()
    new_name = str(row[name_col]).strip()
    new_status = str(row.get(status_col, "نشط")).strip()

    # لو تغيّرت الصلاحية أو القسم → سجل خروج
    # Normalize للطرفين قبل المقارنة
    old_role = user["role_raw"]
    old_dept = normalize_arabic(user["department"])

    new_role_norm = normalize_role(new_role)
    new_dept_norm = normalize_arabic(new_dept)

    # إذا اختلفت الصلاحية أو القسم بعد التطبيع
    if old_role != new_role_norm or old_dept != new_dept_norm:
        session.clear()
        return jsonify({"valid": False})

    if new_status != "نشط":
        session.clear()
        return jsonify({"valid": False})

    # ⭐ إعادة جلب force_reset
    force_raw = str(row.get("force_reset", "0"))
    force_raw = force_raw.replace(".0", "").replace(".00", "").strip()

    force_reset_needed = (force_raw not in ["0"])

    return jsonify({
        "valid": True,
        "user": session.get("user")
    })


@app.route('/api/login', methods=['POST'])
def login():
    data = request.get_json() or {}
    email = (data.get('email', '') or '').strip().lower()
    password = (data.get('password', '') or '').strip()

    df = load_users()
    if df.empty:
        return jsonify({"success": False, "message": "قاعدة المستخدمين فارغة"}), 500

    # ==== اكتشاف الأعمدة ====
    email_col = next((c for c in df.columns if 'بريد' in c or 'email' in c.lower()), None)
    pass_col  = next((c for c in df.columns if 'مرور' in c or 'pass' in c.lower()), None)
    role_col  = next((c for c in df.columns if 'صلاح' in c or 'role' in c.lower()), None)
    dept_col  = next((c for c in df.columns if 'قسم' in c), None)
    name_col  = next((c for c in df.columns if 'اسم' in c), None)
    company_col = next((c for c in df.columns if "شركة" in c), None)
    branch_col = next((c for c in df.columns if "فرع" in c), None)

    # ==== تنظيف ====
    df[email_col] = df[email_col].astype(str).str.lower().str.strip()
    df[pass_col]  = df[pass_col].astype(str).str.strip()

    # ==== المطابقة ====
    user = df[(df[email_col] == email) & (df[pass_col] == password)]
    if user.empty:
        return jsonify({"success": False, "message": "البريد أو كلمة المرور غير صحيحة"}), 401

    user = user.iloc[0]

    raw_role = str(user[role_col]).strip()
    dept_raw = str(user.get(dept_col, '')).strip()
    name     = str(user.get(name_col, '')).strip()
    company = str(user.get(company_col, '')).strip() if company_col else ""
    branch = str(user.get(branch_col, '')).strip() if branch_col else ""

    # ==== Normalize ====
    dept_norm = normalize_arabic(dept_raw)
    role_norm = normalize_role(raw_role)

    # ==== توجيه موحد إلى البوابة فقط ====
    role = role_norm
    redirect = "Portal.html"

    # ==== قراءة الأقسام الأخرى Extra Departments ====
    extra_col = next((c for c in df.columns if "أخرى" in c or "اخرى" in c), None)
    extra_depts = []

    if extra_col:
        raw_extra = str(user.get(extra_col, "")).strip()

        if raw_extra:
            # 🔥 إزالة الرموز الخفية
            raw_extra = raw_extra.replace("\u200f", "").replace("\u200e", "")

            # 🔥 توحيد الفواصل العربية والإنجليزية
            raw_extra = raw_extra.replace(" ،", "،").replace("، ", "،")
            raw_extra = raw_extra.replace(" ,", ",").replace(", ", ",")

            # 🔥 إزالة المسافات حول الفواصل
            raw_extra = re.sub(r"\s*,\s*", ",", raw_extra)
            raw_extra = re.sub(r"\s*،\s*", "،", raw_extra)

            # 🔥 استبدال الفاصلة العربية بالإنجليزية (اختياري للتسهيل)
            raw_extra = raw_extra.replace("،", ",")

            # 🔥 تقسيم مضمون 100%
            parts = [p.strip() for p in raw_extra.split(",") if p.strip()]

            extra_depts = [normalize_arabic(p) for p in parts]
        else:
            extra_depts = []

    # ==== بناء session ====
    session["user"] = {
        "email": email,
        "name": name,
        "role": role,
        "role_raw": role_norm,
        "department": dept_norm,
        "company": company,
        "branch": branch,
        "extra_departments": extra_depts
    }

    apps_raw = str(user.get("apps", "")).strip().lower()

    # 🔥 معالجة NaN / None بشكل صريح
    if apps_raw in ["nan", "none", "null"]:
        apps_raw = ""

    apps_list = [a.strip() for a in apps_raw.split(",") if a.strip()]
    session["user"]["apps"] = apps_list

    # =====================================================
    #                FORCE RESET — النسخة السليمة
    # =====================================================
    force_raw = str(user.get("force_reset", "0")).strip()
    force_raw = force_raw.replace(".0", "").replace(".00", "").strip()

    # القيمة الصحيحة الوحيدة لإجبار التغيير هي 1 فقط
    needs_reset = (force_raw == "1")

    session["user"]["force_reset"] = needs_reset

    # =====================================================

    return jsonify({
        "success": True,
        "redirect": redirect,
        "user": {
    "email": email,
    "name": name,
    "role": role,
    "department": dept_norm,
    "company": company,
    "branch": branch,
    "apps": str(user.get("apps", "")),
    "force_reset": needs_reset,
    "extra_departments": extra_depts
}

    })

@app.route("/api/session")
def api_session():
    user = session.get("user")
    if not user:
        return jsonify({"error": "no session"}), 401
    return jsonify(user)

@app.route('/api/admin/check', methods=['POST'])
def admin_check():
    """التحقق هل المستخدم أدمن"""
    data = request.get_json() or {}
    role = data.get('role', '').strip().lower()

    if 'admin' in role or 'ادمن' in normalize_arabic(role):
        return jsonify({"admin": True})

    return jsonify({"admin": False})


@app.route('/api/admin/get_info', methods=['POST'])
def admin_info():
    data = request.get_json() or {}
    email = (data.get('email') or '').strip().lower()

    df = load_users()
    if df.empty:
        return jsonify({"success": False, "error": "فشل تحميل قاعدة المستخدمين"}), 500

    # 🔍 اكتشاف اسم عمود البريد الإلكتروني ديناميكيًا (مثل login)
    email_col = next((c for c in df.columns if any(k in str(c).lower() for k in ['بريد', 'email', 'ايميل'])), None)

    if not email_col:
        return jsonify({"success": False, "error": "عمود البريد الإلكتروني غير موجود في قاعدة البيانات"}), 500

    # ✅ التنظيف والبحث
    df[email_col] = df[email_col].astype(str).str.lower().str.strip()
    user_row = df[df[email_col] == email]

    if user_row.empty:
        return jsonify({"success": False, "message": "لم يُعثر على المستخدم"})

    user = user_row.iloc[0].to_dict()
    # تحويل القيم إلى نص لتجنب مشاكل JSON (مثل NaN)
    user = {str(k): str(v) if pd.notna(v) else '' for k, v in user.items()}

    return jsonify({"success": True, "user": user})

# ============== API: جلب الموظفين لكل قسم ==============
@app.route('/api/get_employees', methods=['POST'])
def get_employees():
    """
    جلب الموظفين بناءً على المدير (كل الموظفين التابعين له بغض النظر عن القسم)
    """
    try:
        data = request.get_json() or {}
        manager_name = (data.get('manager_name', '') or '').strip()
        dept = (data.get('department', '') or '').strip()

        df = load_users()
        if df.empty:
            return jsonify({"success": False, "message": "لا توجد بيانات مستخدمين"})

        # 🔹 اكتشاف الأعمدة الأساسية
        name_col = next((c for c in df.columns if 'اسم' in str(c)), 'الاسم')
        role_col = next((c for c in df.columns if 'صلاح' in str(c)), 'الصلاحية')
        dept_col = next((c for c in df.columns if 'قسم' in str(c)), 'القسم')

        df['الاسم'] = df[name_col].astype(str).str.strip()
        df['الصلاحية'] = df[role_col].astype(str).str.strip()
        df['القسم'] = df[dept_col].astype(str).str.strip()

        # ✅ المنطق الجديد:
        # إذا المستخدم مدير قسم → يشوف كل الموظفين اللي صلاحيتهم "موظف"
        if manager_name:
            df = df[df['الصلاحية'].isin(['موظف', 'عامل'])]

        # ✅ المدير العام يشوف الكل
        employees = df[['الاسم', 'القسم', 'الصلاحية']].dropna().to_dict(orient='records')
        return jsonify({"success": True, "employees": employees})

    except Exception as e:
        print("❌ get_employees error:", e)
        return jsonify({"success": False, "message": str(e)})


# ============== API: الطلبات ==============
@app.route('/api/get_requests', methods=['POST'])
def get_requests():
    try:
        # ================================
        # 1) بيانات المستخدم من الجلسة
        # ================================
        user = session.get("user")
        if not user:
            return jsonify([])

        role = user.get("role", "")
        user_name = normalize_arabic(user.get("name", "")).replace(" ", "")


        # من الجلسة
        main_dept = normalize_arabic(user.get("department", ""))
        extra_depts = [normalize_arabic(x) for x in (user.get("extra_departments") or [])]

        # ================================
        # 2) دمج الأقسام القادمة عبر POST
        # ================================
        data = request.get_json(silent=True) or {}
        posted_depts = data.get("departments", [])
        posted_depts = [normalize_arabic(str(d)).strip() for d in posted_depts if d]


        # تجميع جميع الأقسام بدون تكرار
        user_departments = list({main_dept, *extra_depts, *posted_depts})

        # ================================
        # 3) تحميل الطلبات
        # ================================
        df = load_requests()
        if df.empty:
            return jsonify([])

        df = df.loc[:, ~df.columns.duplicated()]

        # تطبيع الأعمدة
        df["القسم المرسل"] = df["القسم المرسل"].astype(str).apply(normalize_arabic)
        df["القسم المستلم"] = df["القسم المستلم"].astype(str).apply(normalize_arabic)
        df["اسم المرسل_norm"] = df["اسم المرسل"].astype(str).apply(
            lambda x: normalize_arabic(x)
        )

        # ================================
        # 4) دالة مطابقة الأقسام
        # ================================
        def dept_match(req_dept):
            req_dept = normalize_arabic(req_dept)
            for d in user_departments:
                d = normalize_arabic(d)
                if req_dept == d or req_dept in d or d in req_dept:
                    return True
            return False

        # ================================
        # 5) استبعاد المؤرشف + صلاحيات الوصول
        # ================================
        role_norm = normalize_role(role)

        # ✅ استبعاد الطلبات المؤرشفة من جميع الصفحات ما عدا صفحة الأدمن
        if "مؤرشف" in df.columns and role_norm != "admin":
            df["مؤرشف"] = df["مؤرشف"].astype(str).str.strip()
            df = df[df["مؤرشف"] != "1"]

        if role_norm == "employee":
            incoming = df[df["القسم المستلم"].apply(dept_match)]
            outgoing = df[df["اسم المرسل_norm"] == normalize_arabic(user.get("name", ""))]
            result = pd.concat([incoming, outgoing]).drop_duplicates()

        elif role_norm == "manager":
            incoming = df[df["القسم المستلم"].apply(dept_match)]
            outgoing = df[df["القسم المرسل"].apply(dept_match)]
            result = pd.concat([incoming, outgoing]).drop_duplicates()

        elif role_norm in ["general_manager", "admin", "hr"]:
            result = df.copy()

        else:
            result = pd.DataFrame()

        return jsonify(result.fillna('').to_dict(orient='records'))

    except Exception as e:
        print("❌ get_requests error:", e)
        return jsonify([])


@app.route('/api/create_request', methods=['POST'])
def create_request():
    try:
        title  = request.form.get('title', '').strip()
        desc   = request.form.get('description', '').strip()
        target = request.form.get('targetDept', '').strip()
        sender = request.form.get('senderDept', '').strip()
        sender_name = request.form.get('senderName', '').strip()

        file = request.files.get('file')

        if not all([title, desc, target, sender]):
            return jsonify({"success": False, "message": "جميع الحقول مطلوبة"}), 400

        df = load_requests()
        for col in ['رقم الطلب','التاريخ','العنوان','الوصف','القسم المرسل','القسم المستلم',
                    'الحالة','الموظف المعين','آخر تحديث بواسطة','الوقت','بدأ التنفيذ بواسطة','أغلق بواسطة','الملف']:
            if col not in df.columns:
                df[col] = ""

        req_id = generate_request_id()
        file_name = ""
        if file:
            safe_name = f"{req_id}_{file.filename}"
            file_path = os.path.join(UPLOAD_DIR, safe_name)
            file.save(file_path)
            file_name = safe_name

        new_row = {
            'رقم الطلب': req_id,
            'التاريخ': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            'العنوان': title,
            'الوصف': desc,
            'القسم المرسل': sender,
            'اسم المرسل': sender_name,
            'القسم المستلم': target,
            'اسم المستلم': '',
            'الحالة': 'جديد',
            'الموظف المعين': '-',
            'آخر تحديث بواسطة': sender_name or '-',
            'بدأ التنفيذ بواسطة': '',
            'أغلق بواسطة': '',
            'الوقت': '',
            'الملف': file_name
        }

        df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
        save_requests(df)

        # 🔁 إضافة مزامنة ونسخ احتياطي كاملة بعد إنشاء طلب
        try:
            full_sync_and_backup()
        except Exception as _e:
            print("⚠️ post-create_request full_sync skipped:", _e)

        return jsonify({"success": True})
    except Exception as e:
        print("❌ create_request error:", e)
        return jsonify({"success": False, "message": str(e)}), 500

@app.route('/uploads/<path:filename>')
def get_uploaded_file(filename):
    # ✅ يعرض الملف مباشرة داخل المتصفح بدل التحميل
    return send_from_directory(UPLOAD_DIR, filename)

@app.route('/api/update_request_status', methods=['POST'])
def update_request_status():
    data = request.get_json()
    req_id = (data.get('requestId','') or '').strip()
    new_status = (data.get('status','') or '').strip()
    updater = (data.get('updater','') or '').strip()
    duration = data.get('duration')

    df = load_requests()
    if df.empty or 'رقم الطلب' not in df.columns:
        return jsonify({"success": False}), 404

    idx_list = df.index[df['رقم الطلب'] == req_id].tolist()
    if not idx_list:
        return jsonify({"success": False}), 404
    idx = idx_list[0]

    # ✅ ضمان أن الأعمدة النصية هي من نوع str لتفادي تحذير pandas
    text_cols = ['اسم المستلم', 'بدأ التنفيذ بواسطة', 'أغلق بواسطة', 'آخر تحديث بواسطة', 'الوقت']
    for col in text_cols:
        if col in df.columns:
            df[col] = df[col].astype(str)

    # 🔹 تحديث الحالة والاسم
    df.at[idx, 'الحالة'] = new_status
    df.at[idx, 'آخر تحديث بواسطة'] = updater

    # 🔹 تعيين اسم المستلم فقط إذا لم يكن موجود سابقًا
    if not df.at[idx, 'اسم المستلم']:
        df.at[idx, 'اسم المستلم'] = updater

    if new_status == 'جاري التنفيذ':
        df.at[idx, 'بدأ التنفيذ بواسطة'] = updater
        df.at[idx, 'وقت البداية'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

    elif new_status == 'معلق':
        df.at[idx, 'وقت التوقف المؤقت'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

    elif new_status == 'مغلق':
        if 'وقت البداية' in df.columns:
            start_str = df.at[idx, 'وقت البداية']
            if start_str:
                start_time = datetime.strptime(start_str, '%Y-%m-%d %H:%M:%S')
                diff = datetime.now() - start_time
                df.at[idx, 'الوقت'] = str(diff).split('.')[0]
        if duration:
            df.at[idx, 'الوقت'] = duration
        df.at[idx, 'أغلق بواسطة'] = updater

    if new_status == 'معلق':
        # حفظ وقت الإيقاف المؤقت فقط
        if 'وقت البداية' in df.columns:
            start_str = df.at[idx, 'وقت البداية']
            if start_str:
                start_time = datetime.strptime(start_str, '%Y-%m-%d %H:%M:%S')
                diff = datetime.now() - start_time
                df.at[idx, 'الوقت'] = str(diff).split('.')[0]

    save_requests(df)
    # 🔁 إضافة مزامنة ونسخ احتياطي بعد تحديث حالة الطلب
    try:
        full_sync_and_backup()
    except Exception as _e:
        print("⚠️ post-update_request_status full_sync skipped:", _e)
    return jsonify({"success": True})


@app.route('/api/delegate_request', methods=['POST'])
def delegate_request():
    data = request.get_json() or {}

    # ✅ يدعم كل أنواع المفاتيح (camelCase أو snake_case)
    req_id = data.get('requestId') or data.get('request_id')
    delegate = data.get('delegate') or data.get('delegateName')
    delegated_by = data.get('delegatedBy') or data.get('delegated_by')

    if not req_id or not delegate:
        return jsonify({'success': False, 'message': 'بيانات غير مكتملة (رقم الطلب أو اسم الموظف مفقود)'})

    df = load_requests()
    if df.empty or 'رقم الطلب' not in df.columns:
        return jsonify({'success': False, 'message': 'قاعدة الطلبات فارغة'})

    mask = df['رقم الطلب'] == req_id
    if not mask.any():
        return jsonify({'success': False, 'message': f'الطلب {req_id} غير موجود'})

    # ✅ تحديث الحقول
    df.loc[mask, 'اسم المستلم'] = delegate
    df.loc[mask, 'آخر تحديث بواسطة'] = delegated_by
    df.loc[mask, 'الحالة'] = 'موكل'

    save_requests(df)
    print(f"✅ تم توكيل الطلب {req_id} إلى {delegate} بواسطة {delegated_by}")
    # 🔁 إضافة مزامنة ونسخ احتياطي بعد التوكيل
    try:
        full_sync_and_backup()
    except Exception as _e:
        print("⚠️ post-delegate_request full_sync skipped:", _e)
    return jsonify({'success': True})

# ============== API: تصدير الطلبات ==============
@app.route('/api/export_requests', methods=['POST'])
def export_requests():
    """
    📦 تصدير الطلبات إلى ملف Excel يحتوي على عدة أوراق:
    ✅ فقط الطلبات التي استلمها القسم (القسم المستلم)
    كل ورقة تمثل حالة من حالات الطلب (جديد، جاري التنفيذ، مغلق، مرفوض، إلخ)
    """
    try:
        data = request.get_json() or {}
        dept = (data.get('department', '') or '').strip()
        start = (data.get('start_date', '') or '').strip()
        end   = (data.get('end_date', '') or '').strip()

        if not os.path.exists(REQUESTS_PATH):
            return jsonify({"success": False, "message": "ملف الطلبات غير موجود."})

        df = pd.read_excel(REQUESTS_PATH)
        if df.empty:
            return jsonify({"success": False, "message": "لا توجد بيانات لتصديرها."})

        # 🧹 تنظيف الأعمدة المهمة
        for col in ['القسم المستلم', 'الحالة', 'التاريخ']:
            if col in df.columns:
                df[col] = (
                    df[col]
                    .astype(str)
                    .str.strip()
                    .str.replace('\u200f', '', regex=True)
                    .str.replace('\u200e', '', regex=True)
                )

        # ✅ فلترة الطلبات التي استلمها القسم فقط
        dept_norm = normalize_arabic(dept)
        df = df[df['القسم المستلم'].apply(lambda x: dept_norm in normalize_arabic(x) or normalize_arabic(x) in dept_norm)]

        # ✅ فلترة حسب التاريخ إن وجد
        if start:
            df = df[pd.to_datetime(df['التاريخ'], errors='coerce') >= pd.to_datetime(start)]
        if end:
            end_dt = pd.to_datetime(end) + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)
            df = df[pd.to_datetime(df['التاريخ'], errors='coerce') <= end_dt]

        if df.empty:
            return jsonify({"success": False, "message": "لا توجد طلبات استلمها القسم ضمن الشروط المحددة."})

        # 🗂️ تقسيم البيانات حسب الحالة
        grouped = {status: sub_df for status, sub_df in df.groupby('الحالة')}

        # 📘 إنشاء ملف Excel بعدة أوراق (كل ورقة = حالة)
        ts = datetime.now().strftime('%Y%m%d_%H%M%S')
        fname = f"طلبات_الواردة_{dept}_{ts}.xlsx".replace(' ', '_')
        fpath = os.path.join(EXPORT_DIR, fname)

        with pd.ExcelWriter(fpath, engine='openpyxl') as writer:
            for status, sub_df in grouped.items():
                clean_status = str(status).replace('/', '-').strip() or 'غير_محدد'
                sub_df.to_excel(writer, index=False, sheet_name=clean_status[:31])

        return jsonify({"success": True, "file": fname})

    except Exception as e:
        print("❌ export_requests error:", e)
        return jsonify({"success": False, "message": f"حدث خطأ أثناء التصدير: {str(e)}"})

@app.route('/download/<path:filename>')
def download(filename):
    return send_from_directory(EXPORT_DIR, filename, as_attachment=True)

# ============== API: الشات العام ==============
@app.route("/chatbot", methods=["POST"])
def chatbot():
    """رد ذكي باستخدام OpenRouter بسرعة أعلى"""
    user_input = request.json.get("message", "").strip()
    if not user_input:
        return jsonify({"reply": "الرسالة فارغة!"})

    headers = {
        "Authorization": f"Bearer {OPENROUTER_API_KEY}",
        "Content-Type": "application/json",
    }

    payload = {
        "model": "qwen/qwen-2.5-7b-instruct",
        "messages": [
            {"role": "system", "content": "أنت مساعد ذكي تتحدث العربية وتساعد موظفي نظام SEVENS."},
            {"role": "user", "content": user_input}
        ],
        "temperature": 0.6,
        "max_tokens": 200
    }

    try:
        response = requests.post(
            "https://openrouter.ai/api/v1/chat/completions",
            headers=headers,
            json=payload,
            timeout=15,   # ⏱️ أقصى مهلة للرد 15 ثانية فقط
        )

        if response.status_code == 200:
            data = response.json()
            if "choices" in data and len(data["choices"]) > 0:
                reply = data["choices"][0]["message"]["content"].strip()
                return jsonify({"reply": reply})
            else:
                return jsonify({"reply": "لم يصل رد من نموذج الذكاء الاصطناعي."})
        else:
            print("❌ OpenRouter Error:", response.text)
            return jsonify({"reply": "حدث خطأ في الخادم أثناء معالجة الطلب."})

    except requests.Timeout:
        return jsonify({"reply": "الخادم تأخر في الرد، حاول مرة أخرى لاحقاً."})
    except Exception as e:
        print("❌ chatbot error:", e)
        return jsonify({"reply": "تعذر الاتصال بخدمة الذكاء الاصطناعي."})

# ============== API: دردشة بين الموظفين ==============
CHAT_UPLOAD_DIR = os.path.join(BASE_DIR, "chat_uploads")
os.makedirs(CHAT_UPLOAD_DIR, exist_ok=True)

@app.route('/api/chat_send_file', methods=['POST'])
def chat_send_file():
    req_id = request.form.get('request_id')
    sender = request.form.get('sender')
    dept = request.form.get('department')
    msg = request.form.get('message', '')
    file = request.files.get('file')
    filename = ""

    if file:
        safe_name = f"{req_id}_{datetime.now().strftime('%Y%m%d_%H%M%S')}_{file.filename}"
        path = os.path.join(CHAT_UPLOAD_DIR, safe_name)
        file.save(path)
        filename = safe_name

    df = load_chats()
    new = pd.DataFrame([{
        'رقم الطلب': req_id,
        'المرسل': sender,
        'القسم': dept,
        'الرسالة': msg,
        'الملف': filename,
        'الوقت': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    }])
    df = pd.concat([df, new], ignore_index=True)
    df.to_excel(CHAT_PATH, index=False)

    # ✅ تحديث الطلب بآخر مرسل
    req_df = load_requests()
    mask = req_df['رقم الطلب'] == req_id
    if mask.any():
        req_df.loc[mask, 'آخر تحديث بواسطة'] = sender
        save_requests(req_df)

        # 🔁 إضافة مزامنة ونسخ احتياطي بعد إرسال ملف دردشة
    try:
        full_sync_and_backup()
    except Exception as _e:
        print("⚠️ post-chat_send_file full_sync skipped:", _e)

    return jsonify({"success": True})

@app.route('/api/chat_get/<req_id>', methods=['GET'])
def chat_get(req_id):
    """إرجاع جميع الرسائل الخاصة بطلب محدد"""
    try:
        df = load_chats()
        if df.empty:
            return jsonify([])
        msgs = df[df['رقم الطلب'].astype(str) == str(req_id)].fillna('').to_dict(orient='records')
        return jsonify(msgs)
    except Exception as e:
        print("❌ chat_get error:", e)
        return jsonify([])


@app.route('/chat_uploads/<path:filename>')
def chat_uploads(filename):
    return send_from_directory(CHAT_UPLOAD_DIR, filename)

@app.route('/api/force_reset_password', methods=['POST'])
def force_reset_password():
    data = request.get_json() or {}

    email = (data.get("email") or "").strip().lower()
    new_password = (data.get("newPassword") or "").strip()

    if not email or not new_password:
        return jsonify({"success": False, "message": "البيانات ناقصة"}), 400

    df = pd.read_excel(DB_PATH)

    # اكتشاف أعمدة البريد وكلمة المرور
    email_col = next((c for c in df.columns if "بريد" in c or "email" in c.lower()), None)
    pass_col  = next((c for c in df.columns if "مرور" in c or "pass" in c.lower()), None)

    if "force_reset" not in df.columns:
        df["force_reset"] = "1"

    df[email_col] = df[email_col].astype(str).str.lower().str.strip()
    mask = df[email_col] == email

    if not mask.any():
        return jsonify({"success": False, "message": "المستخدم غير موجود"}), 404

    # تحديث كلمة المرور
    df.loc[mask, pass_col] = new_password

    # تصحيح force_reset
    df["force_reset"] = df["force_reset"].astype(str)
    df["force_reset"] = df["force_reset"].str.replace(".0", "", regex=False).str.replace(".00", "", regex=False).str.strip()

    df.loc[mask, "force_reset"] = "0"

    # ⭐⭐ الحفظ الفعلي في Excel
    df.to_excel(DB_PATH, index=False)

    # تسجيل خروج المستخدم
    session.pop("user", None)

    return jsonify({"success": True, "message": "تم تحديث كلمة المرور"})


# ============== API: استعادة / إعادة تعيين كلمة المرور ==============
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import random
from flask import session
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import base64


GMAIL_SCOPES = ["https://www.googleapis.com/auth/gmail.send"]


GMAIL_CREDENTIALS_PATH = os.path.join(DATA_DIR, "gmail_credentials.json")
GMAIL_TOKEN_PATH = os.path.join(DATA_DIR, "gmail_token.json")

def get_gmail_service():

    creds = None


    if os.path.exists(GMAIL_TOKEN_PATH):
        creds = Credentials.from_authorized_user_file(GMAIL_TOKEN_PATH, GMAIL_SCOPES)


    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:

            flow = InstalledAppFlow.from_client_secrets_file(
                GMAIL_CREDENTIALS_PATH,
                GMAIL_SCOPES
            )
            creds = flow.run_local_server(port=0)


        with open(GMAIL_TOKEN_PATH, "w", encoding="utf-8") as token:
            token.write(creds.to_json())

    from googleapiclient.discovery import build
    service = build("gmail", "v1", credentials=creds)
    return service

def send_html_email_via_gmail(to_email: str, subject: str, html_body: str):

    try:
        service = get_gmail_service()

        msg = MIMEMultipart("alternative")
        msg["To"] = to_email
        msg["From"] = "SEVENS System"
        msg["Subject"] = subject

        msg.attach(MIMEText(html_body, "html", "utf-8"))

        raw = base64.urlsafe_b64encode(msg.as_bytes()).decode("utf-8")
        body = {"raw": raw}

        service.users().messages().send(userId="me", body=body).execute()
        print(f"✅ Gmail API: email sent to {to_email}")

    except Exception as e:
        print("❌ Gmail API send error:", e)

# ================================
# API: إرسال رمز التحقق عبر SMTP
# ================================
@app.route('/api/send_reset_code', methods=['POST'])
def send_reset_code():
    import threading, time

    data = request.get_json() or {}
    email = (data.get("email") or "").strip().lower()

    if not email:
        return jsonify({"success": False, "message": "يرجى إدخال البريد"}), 400

    # التحقق من وجود البريد في قاعدة البيانات
    if not os.path.exists(DB_PATH):
        return jsonify({"success": False, "message": "ملف المستخدمين غير موجود"}), 500

    df = pd.read_excel(DB_PATH)
    df_cols = [str(c).replace(" ", "").lower() for c in df.columns]

    email_variants = ["البريدالالكتروني", "البريدالإلكتروني", "الايميل", "email"]
    email_col = None

    for i, col in enumerate(df_cols):
        if any(v in col for v in email_variants):
            email_col = df.columns[i]
            break

    if not email_col:
        return jsonify({"success": False, "message": "لا يوجد عمود بريد إلكتروني"}), 500

    df[email_col] = df[email_col].astype(str).str.lower().str.strip()

    if email not in df[email_col].values:
        return jsonify({"success": False, "message": "البريد غير موجود"}), 404

    # إنشاء رمز التحقق
    code = str(random.randint(100000, 999999))

    session["reset_code"] = code
    session["reset_email"] = email
    session["reset_code_time"] = time.time()
    session["reset_verified"] = False

    # الرسالة HTML
    subject = "رمز التحقق لإعادة تعيين كلمة المرور - SEVENS"
    html = f"""
    <html><body style='direction:rtl;font-family:Tajawal;'>
        <h3>رمز التحقق الخاص بك</h3>
        <p>رمزك هو:</p>
        <div style='font-size:32px;font-weight:bold;color:#1976d2'>{code}</div>
        <p>صالح لمدة 10 دقائق.</p>
    </body></html>
    """

    def send_email_background():
        try:
            send_html_email_via_company(email, subject, html)
        except Exception as e:
            print("❌ SMTP background error:", e)

    threading.Thread(target=send_email_background, daemon=True).start()

    return jsonify({"success": True, "message": "تم إرسال رمز التحقق"})


# ============== API: التحقق من رمز إعادة التعيين ==============
@app.route('/api/verify_reset_code', methods=['POST'])
def verify_reset_code():
    import time

    data = request.get_json() or {}
    code = (data.get("code") or "").strip()

    saved_code = session.get("reset_code")
    saved_time = session.get("reset_code_time")

    if not saved_code or not saved_time:
        return jsonify({"success": False, "message": "لا يوجد رمز مُرسل"}), 400

    if time.time() - float(saved_time) > 600:
        session.pop("reset_code", None)
        session.pop("reset_email", None)
        session.pop("reset_code_time", None)
        return jsonify({"success": False, "message": "انتهت صلاحية الرمز"}), 400

    if code != saved_code:
        return jsonify({"success": False, "message": "رمز غير صحيح"}), 400

    session["reset_verified"] = True
    return jsonify({"success": True, "message": "تم التحقق بنجاح"})


@app.route('/api/forgot_reset_password', methods=['POST'])
def forgot_reset_password():
    """تحديث كلمة المرور عبر البريد الإلكتروني بدون إنشاء عمود جديد"""
    try:
        data = request.get_json() or {}
        email = (data.get('email', '') or '').strip().lower()
        new_password = (data.get('newPassword', '') or '').strip()
        # التحقق من أن المستخدم تحقق من الكود
        if email != session.get("reset_email"):
            return jsonify({"success": False, "message": "يرجى إعادة طلب رمز التحقق"}), 403

        if not email or not new_password:
            return jsonify({"success": False, "message": "يرجى إدخال البريد وكلمة المرور الجديدة"}), 400

        df = pd.read_excel(DB_PATH)

        # 🔹 تنظيف أسماء الأعمدة من الرموز والمسافات والاختلافات الإملائية
        df.columns = (
            df.columns.astype(str)
            .str.replace('\u200f', '', regex=True)
            .str.replace('\u200e', '', regex=True)
            .str.replace(' ', '', regex=True)
            .str.strip()
        )

        # 🧩 تعريف جميع احتمالات اسم العمود
        password_variants = ['كلمهالمرور', 'كلمه المرور', 'كلمةالمرور', 'كلمة المرور', 'كلمةالسر', 'password', 'pass']
        email_variants = ['البريدالإلكتروني', 'البريدالالكتروني', 'الايميل', 'email']

        # 🔍 تحديد أسماء الأعمدة الفعلية
        pass_col = next((col for col in df.columns if any(p.replace(' ', '') in col for p in password_variants)), None)
        email_col = next((col for col in df.columns if any(e.replace(' ', '') in col for e in email_variants)), None)

        if not email_col or not pass_col:
            return jsonify({"success": False, "message": "تعذر العثور على أعمدة البريد أو كلمة المرور في الملف"}), 500

        # 🔹 توحيد البريد الإلكتروني للمقارنة
        df[email_col] = df[email_col].astype(str).str.lower().str.strip()

        # 🔍 البحث عن المستخدم المستهدف
        mask = df[email_col] == email
        if not mask.any():
            return jsonify({"success": False, "message": "البريد الإلكتروني غير موجود"}), 404

        # ✏️ تعديل كلمة المرور داخل نفس العمود الموجود
        df.loc[mask, pass_col] = new_password

        # 🧼 إزالة الأعمدة المكررة التي تحمل نفس الاسم بعد التعديل (لتفادي التكرار)
        df = df.loc[:, ~df.columns.duplicated()]

        df.to_excel(DB_PATH, index=False)

        print(f"🔑 Password updated successfully for {email} (column: {pass_col})")

        # 🔁 مزامنة تلقائية بعد التعديل
        try:
            full_sync_and_backup()
        except Exception as _e:
            print("⚠️ post-forgot_reset_password full_sync skipped:", _e)

        return jsonify({"success": True, "message": "تم تحديث كلمة المرور بنجاح ✅"})

    except Exception as e:
        print("❌ forgot_reset_password error:", e)
        return jsonify({"success": False, "message": "حدث خطأ أثناء تحديث كلمة المرور"})

# ====== ★★★ HR APIs ★★★ ======

@app.route('/api/hr/list_users', methods=['GET'])
def hr_list_users():
    """عرض جميع المستخدمين كما هم في ملف Excel بدون إنشاء أعمدة جديدة"""
    try:
        if not os.path.exists(DB_PATH):
            return jsonify([])

        df = pd.read_excel(DB_PATH)
        df = df.loc[:, ~df.columns.duplicated()]  # إزالة الأعمدة المكررة
        # ✅ توحيد أسماء الأعمدة (حتى لو اختلفت المسافات أو الهمزات)
        rename_map = {
            'كلمهالمرور': 'كلمة المرور',
            'كلمه المرور': 'كلمة المرور',
            'كلمةالمرور': 'كلمة المرور',
            'كلمة السر': 'كلمة المرور',
            'password': 'كلمة المرور',
            'pass': 'كلمة المرور',
            'الايميل': 'البريد الإلكتروني',
            'email': 'البريد الإلكتروني',
            'البريدالالكتروني': 'البريد الإلكتروني',
            'البريدالإلكتروني': 'البريد الإلكتروني',
            'الحاله': 'الحالة',
            'status': 'الحالة',
            'role': 'الصلاحية'
        }
        for old, new in rename_map.items():
            if old in df.columns:
                df.rename(columns={old: new}, inplace=True)

        # التحقق من وجود الأعمدة الأساسية فقط (بدون إنشائها)
        required = ['الاسم','الصلاحية','كلمة المرور','البريد الإلكتروني','القسم','الحالة']
        for col in required:
            if col not in df.columns:
                print(f"⚠️ الملف ناقص العمود: {col}")
                return jsonify([])

        return jsonify(df.fillna('').to_dict(orient='records'))
    except Exception as e:
        print("hr_list_users error:", e)
        return jsonify([]), 500


@app.route('/api/hr/add_user', methods=['POST'])
def hr_add_user():
    """إضافة مستخدم جديد فقط بالأعمدة الموجودة فعليًا"""

    data = request.get_json() or {}
    name  = (data.get('name','') or '').strip()
    role  = (data.get('role','') or '').strip()
    pwd   = (data.get('password','') or '').strip()
    email = (data.get('email','') or '').strip().lower()
    dept  = (data.get('department','') or '').strip()
    status= (data.get('status','نشط') or 'نشط').strip()
    extra = (data.get("extra_departments", "") or "").strip()

    if not all([name, role, pwd, email, dept]):
        return jsonify({"success": False, "message": "الحقول مطلوبة"}), 400

    if not os.path.exists(DB_PATH):
        return jsonify({"success": False, "message": "ملف المستخدمين غير موجود"}), 500

    df = pd.read_excel(DB_PATH)
    df = df.loc[:, ~df.columns.duplicated()]
    # ✅ توحيد أسماء الأعمدة (حتى لو اختلفت المسافات أو الهمزات)
    rename_map = {
        'كلمهالمرور': 'كلمة المرور',
        'كلمه المرور': 'كلمة المرور',
        'كلمةالمرور': 'كلمة المرور',
        'كلمة السر': 'كلمة المرور',
        'password': 'كلمة المرور',
        'pass': 'كلمة المرور',
        'الايميل': 'البريد الإلكتروني',
        'email': 'البريد الإلكتروني',
        'البريدالالكتروني': 'البريد الإلكتروني',
        'البريدالإلكتروني': 'البريد الإلكتروني',
        'الحاله': 'الحالة',
        'status': 'الحالة',
        'role': 'الصلاحية'
    }
    for old, new in rename_map.items():
        if old in df.columns:
            df.rename(columns={old: new}, inplace=True)

    required_cols = ['الاسم','الصلاحية','كلمة المرور','البريد الإلكتروني','القسم','الحالة']
    for col in required_cols:
        if col not in df.columns:
            return jsonify({"success": False, "message": f"الملف ناقص العمود: {col}"}), 500

    # منع التكرار بالبريد
    mask = df['البريد الإلكتروني'].astype(str).str.lower().str.strip() == email
    if mask.any():
        return jsonify({"success": False, "message": "البريد موجود مسبقاً"}), 409

    new_row = {
        'الاسم': name,
        'الصلاحية': role,
        'كلمة المرور': pwd,
        'البريد الإلكتروني': email,
        'القسم': dept,
        'الحالة': status,
        'الأقسام الأخرى': extra

    }

    # الاحتفاظ فقط بالأعمدة الموجودة
    df = pd.concat([df, pd.DataFrame([[new_row.get(c, '') for c in df.columns]], columns=df.columns)], ignore_index=True)
    df.to_excel(DB_PATH, index=False)
    print(f"✅ تمت إضافة مستخدم جديد: {email}")

    return jsonify({"success": True})


@app.route('/api/hr/update_user', methods=['POST'])
def hr_update_user():
    """تحديث بيانات المستخدم + تحديث جميع الطلبات المرتبطة + مزامنة كاملة"""

    data = request.get_json() or {}

    email = (data.get('email','') or '').strip().lower()
    if not email:
        return jsonify({"success": False, "message": "البريد مطلوب"}), 400

    # تحميل المستخدمين من Excel
    df = pd.read_excel(DB_PATH)
    df = df.loc[:, ~df.columns.duplicated()]

    # توحيد الأعمدة
    rename_map = {
        'الايميل': 'البريد الإلكتروني',
        'email': 'البريد الإلكتروني',
        'البريدالالكتروني': 'البريد الإلكتروني',

        'password': 'كلمة المرور',
        'pass': 'كلمة المرور',
        'كلمه المرور': 'كلمة المرور',
        'كلمةالمرور': 'كلمة المرور',

        'الحاله': 'الحالة',
        'status': 'الحالة',

        'role': 'الصلاحية',

        'extra_departments': 'الأقسام الأخرى'

    }
    for old, new in rename_map.items():
        if old in df.columns:
            df.rename(columns={old: new}, inplace=True)

    # التأكد من الأعمدة الأساسية
    required_cols = ['الاسم','الصلاحية','كلمة المرور','البريد الإلكتروني','القسم','الحالة']
    for col in required_cols:
        if col not in df.columns:
            df[col] = ''

    # إيجاد الموظف
    mask = df['البريد الإلكتروني'].astype(str).str.lower().str.strip() == email
    if not mask.any():
        return jsonify({"success": False, "message": "المستخدم غير موجود"}), 404

    # حفظ البيانات القديمة قبل التعديل
    old_email = str(df.loc[mask, 'البريد الإلكتروني'].values[0]).strip().lower()
    old_name  = str(df.loc[mask, 'الاسم'].values[0]).strip()
    old_dept  = str(df.loc[mask, 'القسم'].values[0]).strip()

    # تحديث بيانات الموظف
    fields = {
        'name': 'الاسم',
        'role': 'الصلاحية',
        'password': 'كلمة المرور',
        'department': 'القسم',
        'status': 'الحالة',
        'extra_departments': 'الأقسام الأخرى'
    }

    for key, col in fields.items():
        if key in data and data[key] is not None:
            df.loc[mask, col] = str(data[key]).strip()

    # البيانات الجديدة بعد التعديل
    new_email = str(df.loc[mask, 'البريد الإلكتروني'].values[0]).strip().lower()
    new_name  = str(df.loc[mask, 'الاسم'].values[0]).strip()
    new_dept  = str(df.loc[mask, 'القسم'].values[0]).strip()

    # حفظ Excel
    df.to_excel(DB_PATH, index=False)


    # =============================
    #  تحديث الطلبات المرتبطة
    # =============================
    req_df = load_requests()
    if not req_df.empty:

        # تحديث الأقسام إذا تغيّر القسم
        old_dept_norm = normalize_arabic(old_dept)
        new_dept_norm = normalize_arabic(new_dept)

        if new_dept != old_dept:
            for col in ['القسم المرسل', 'القسم المستلم']:
                if col in req_df.columns:
                    req_df[col] = req_df[col].apply(
                        lambda x: new_dept if normalize_arabic(str(x)) == old_dept_norm else x
                    )

        # تحديث اسم المرسل إذا تغيّر الاسم
        if 'اسم المرسل' in req_df.columns:
            req_df['اسم المرسل'] = req_df['اسم المرسل'].apply(
                lambda x: new_name if normalize_arabic(str(x)) == normalize_arabic(old_name) else x
            )

        # تحديث اسم المستلم إذا تغيّر
        if 'اسم المستلم' in req_df.columns:
            req_df['اسم المستلم'] = req_df['اسم المستلم'].apply(
                lambda x: new_name if normalize_arabic(str(x)) == normalize_arabic(old_name) else x
            )

        save_requests(req_df)

    # =============================
    #  تحديث دردشات الطلبات
    # =============================
    try:
        chats = load_chats()
        if not chats.empty:
            if 'المرسل' in chats.columns:
                chats['المرسل'] = chats['المرسل'].apply(
                    lambda x: new_name if normalize_arabic(str(x)) == normalize_arabic(old_name) else x
                )
            chats.to_excel(CHAT_PATH, index=False)
    except:
        pass

    # =============================
    #  مزامنة SQLite بالكامل
    # =============================
    try:
        sync_excel_to_sqlite()
    except Exception as e:
        print("SQLite error:", e)

    # =============================
    #  مزامنة + نسخ احتياطي
    # =============================
    try:
        full_sync_and_backup()
    except Exception as e:
        print("Backup error:", e)
    # Force logout if updating own account
    try:
        if session.get("user", {}).get("email") == new_email:
            session.clear()
    except:
        pass

    print(f"🔧 Updated user: {old_email} → {new_email}")

    return jsonify({"success": True})

@app.route('/api/hr/archive_user', methods=['POST'])
def hr_archive_user():
    """أرشفة المستخدم (تحديث الحالة فقط إن وجدت)"""
    data = request.get_json() or {}
    email = (data.get('email','') or '').strip().lower()
    if not email:
        return jsonify({"success": False, "message": "البريد مطلوب"}), 400

    if not os.path.exists(DB_PATH):
        return jsonify({"success": False, "message": "ملف المستخدمين غير موجود"}), 500

    df = pd.read_excel(DB_PATH)
    sync_excel_to_sqlite()
    df = df.loc[:, ~df.columns.duplicated()]
    # ✅ توحيد أسماء الأعمدة (حتى لو اختلفت المسافات أو الهمزات)
    rename_map = {
        'كلمهالمرور': 'كلمة المرور',
        'كلمه المرور': 'كلمة المرور',
        'كلمةالمرور': 'كلمة المرور',
        'كلمة السر': 'كلمة المرور',
        'password': 'كلمة المرور',
        'pass': 'كلمة المرور',
        'الايميل': 'البريد الإلكتروني',
        'email': 'البريد الإلكتروني',
        'البريدالالكتروني': 'البريد الإلكتروني',
        'البريدالإلكتروني': 'البريد الإلكتروني',
        'الحاله': 'الحالة',
        'status': 'الحالة',
        'role': 'الصلاحية'
    }
    for old, new in rename_map.items():
        if old in df.columns:
            df.rename(columns={old: new}, inplace=True)

    if 'الحالة' not in df.columns:
        return jsonify({"success": False, "message": "عمود الحالة غير موجود في الملف"}), 500

    mask = df['البريد الإلكتروني'].astype(str).str.lower().str.strip() == email
    if not mask.any():
        return jsonify({"success": False, "message": "المستخدم غير موجود"}), 404

    df.loc[mask, 'الحالة'] = 'مؤرشف'
    df.to_excel(DB_PATH, index=False)
    print(f"📦 تمت أرشفة المستخدم: {email}")

    return jsonify({"success": True})

def sync_excel_to_sqlite():
    """ينسخ محتوى Excel إلى SQLite إذا تم التعديل على Excel"""

    try:
        conn = sqlite3.connect(DB_SQLITE)
        cur = conn.cursor()

        # 🧱 مزامنة المستخدمين
        if os.path.exists(DB_PATH):
            df_users = pd.read_excel(DB_PATH)

            # 🔹 تنظيف أسماء الأعمدة من أي رموز ومسافات
            df_users.columns = (
                df_users.columns
                .astype(str)
                .str.replace('\u200f', '', regex=True)
                .str.replace('\u200e', '', regex=True)
                .str.replace(' ', '', regex=True)  # ← تحذف المسافات بين الحروف
                .str.strip()
            )

            # 🧩 خريطة التطبيع للأعمدة المحتملة
            rename_map = {
                'الاسم': 'الاسم',
                'الا سم': 'الاسم',
                'الإسم': 'الاسم',
                'الاسمالكامل': 'الاسم',

                'email': 'البريد الإلكتروني',
                'الايميل': 'البريد الإلكتروني',
                'البريدالالكتروني': 'البريد الإلكتروني',
                'البريدالإلكتروني': 'البريد الإلكتروني',

                'القسم': 'القسم',
                'ادارة': 'القسم',

                'الصلاحيه': 'الصلاحية',
                'الوظيفة': 'الصلاحية',
                'role': 'الصلاحية',

                # 👇 أضف كل الاحتمالات الممكنة لكلمة المرور
                'كلمهالمرور': 'كلمة المرور',
                'كلمه المرور': 'كلمة المرور',
                'كلمةالمرور': 'كلمة المرور',
                'كلمة المرور': 'كلمة المرور',
                'كلمةالسر': 'كلمة المرور',
                'password': 'كلمة المرور',
                'pass': 'كلمة المرور',
            }

            # ✅ إعادة التسمية بناءً على تطابق جزئي (حتى لو كان اختلاف بسيط)
            for col in list(df_users.columns):
                normalized = re.sub(r'[إأآا]', 'ا', col).replace(' ', '').lower()
                for k, v in rename_map.items():
                    if re.sub(r'[إأآا]', 'ا', k).replace(' ', '').lower() in normalized:
                        df_users.rename(columns={col: v}, inplace=True)

            # ✅ ضمان الأعمدة الأساسية موجودة
            for col in ['الاسم', 'الصلاحية', 'كلمة المرور', 'البريد الإلكتروني', 'القسم', 'الحالة']:
                if col not in df_users.columns:
                    df_users[col] = ''

            # ✅ إدخال المستخدمين إلى SQLite (نسخة محسّنة تتفادى NaN أو أعمدة غير مفهومة)
            for _, row in df_users.iterrows():
                try:
                    # 🧩 استخلاص آمن لكل حقل
                    email_val = str(row.get('البريد الإلكتروني', '')).strip().lower()
                    name_val = str(row.get('الاسم', '')).strip()
                    role_val = str(row.get('الصلاحية', '')).strip()
                    dept_val = str(row.get('القسم', '')).strip()
                    status_val = str(row.get('الحالة', 'نشط')).strip()

                    # 🧩 معالجة كلمة المرور بشكل خاص (لأنها سبب الخطأ)
                    pwd_val = row.get('كلمة المرور', '')
                    if isinstance(pwd_val, (pd.Series, pd.DataFrame)):
                        pwd_val = pwd_val.iloc[0] if not pwd_val.empty else ''
                    pwd_val = str(pwd_val).strip()
                    if pwd_val.lower() in ['nan', 'none']:
                        pwd_val = ''

                    cur.execute("""
                        INSERT OR REPLACE INTO users (email, name, role, password, department, status)
                        VALUES (?, ?, ?, ?, ?, ?)
                    """, (email_val, name_val, role_val, pwd_val, dept_val, status_val))
                except Exception as e:
                    print(f"⚠️ Error inserting user row: {e}")

        # 🧾 مزامنة الطلبات
        if os.path.exists(REQUESTS_PATH):
            df_req = pd.read_excel(REQUESTS_PATH)
            df_req.columns = [c.strip() for c in df_req.columns]
            for _, row in df_req.iterrows():
                cur.execute("""
                    INSERT OR REPLACE INTO requests (req_id, date, title, description, sender_dept, receiver_dept, status, assigned_to, updated_by, duration, file_name)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (
                    str(row.get('رقم الطلب', '')).strip(),
                    str(row.get('التاريخ', '')).strip(),
                    str(row.get('العنوان', '')).strip(),
                    str(row.get('الوصف', '')).strip(),
                    str(row.get('القسم المرسل', '')).strip(),
                    str(row.get('القسم المستلم', '')).strip(),
                    str(row.get('الحالة', '')).strip(),
                    str(row.get('الموظف المعين', '')).strip(),
                    str(row.get('آخر تحديث بواسطة', '')).strip(),
                    str(row.get('الوقت', '')).strip(),
                    str(row.get('الملف', '')).strip(),
                ))

        conn.commit()
        conn.close()
        print("🔁 Excel → SQLite sync done successfully ✅")

    except Exception as e:
        print("❌ sync_excel_to_sqlite error:", e)

def full_sync_and_backup():
    """مزامنة من Excel إلى SQLite فقط + رفع النسخة الاحتياطية"""
    try:
        # ✅ فقط Excel → SQLite
        sync_excel_to_sqlite()

        # ✅ رفع ملفات Excel إلى Google Drive (اختياري)
        upload_to_drive(DB_PATH)
        upload_to_drive(REQUESTS_PATH)
        upload_to_drive(CHAT_PATH)

        print("✅ One-way sync (Excel → SQLite) done successfully.")
    except Exception as e:
        print("⚠️ full_sync_and_backup error:", e)

# ✅ مزامنة قواعد البيانات قبل التشغيل
init_sqlite()
sync_excel_to_sqlite()

import threading
import time

def watch_excel_changes(interval=30):
    """يراقب أي تغييرات في ملفات Excel ويعمل مزامنة تلقائية"""
    last_users_time = os.path.getmtime(DB_PATH)
    last_requests_time = os.path.getmtime(REQUESTS_PATH)

    while True:
        time.sleep(interval)
        try:
            # تحقق من آخر وقت تعديل
            new_users_time = os.path.getmtime(DB_PATH)
            new_requests_time = os.path.getmtime(REQUESTS_PATH)

            # إذا تغير أي ملف → أعد المزامنة
            if new_users_time != last_users_time or new_requests_time != last_requests_time:
                print("🔄 Detected Excel file change, syncing to SQLite...")
                sync_excel_to_sqlite()
                last_users_time = new_users_time
                last_requests_time = new_requests_time

        except Exception as e:
            print("⚠️ watch_excel_changes error:", e)

# 🔁 تشغيل المراقبة في خيط منفصل
threading.Thread(target=watch_excel_changes, daemon=True).start()

def auto_backup(interval_hours=24):
    """نسخ احتياطي تلقائي إلى Google Drive كل فترة محددة"""
    while True:
        try:
            print("🕐 Running scheduled backup...")
            upload_to_drive(DB_PATH)
            upload_to_drive(REQUESTS_PATH)
            # 📎 إضافة نسخة احتياطية لملف دردشة الطلبات أيضًا
            upload_to_drive(CHAT_PATH)
            print("✅ Backup completed successfully.")
        except Exception as e:
            print("❌ auto_backup error:", e)
        time.sleep(interval_hours * 3600)


threading.Thread(target=auto_backup, daemon=True).start()



import pandas as pd
from datetime import datetime
import os
import re


# ====== المسارات ======
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")
os.makedirs(DATA_DIR, exist_ok=True)

CARS_XLSX = os.path.join(DATA_DIR, "cars_data.xlsx")
OIL_XLSX  = os.path.join(DATA_DIR, "oil_history.xlsx")

# أعمدة ملف السيارات كما هي في الجدول المرفوع
AR_COLS = {
    "vin": "رقم الهيكل",
    "plate": "اللوحة",
    "brand": "الشركة",
    "model": "فئة السيارة",
    "color": "اللون",
    "year": "سنة الصناعة",
    # أعمدة إضافية اختيارية نحدِّثها من آخر تغيير زيت
    "last_oil_date": "تاريخ تغيير الزيت",
    "last_odometer": "عداد السيارة",
    "last_oil_run": "ممشى الزيت",
    "updated_at": "آخر تحديث"
}

# جدول سجل الزيت
H_COLS = {
    "plate": "اللوحة",
    "vin": "رقم الهيكل",
    "date": "تاريخ التغيير",
    "odometer": "عداد السيارة",
    "oil_run": "ممشى الزيت",
    "notes": "ملاحظات",
    "created_at": "تاريخ الإدخال"
}

def _now():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

def ensure_files():
    # ملف السيارات يجب أن يكون موجوداً (منك)
    if not os.path.exists(CARS_XLSX):
        # ننشئ ملفاً فارغاً بنفس الأعمدة الأساسية لكي لا يتعطل التشغيل
        df = pd.DataFrame(columns=[AR_COLS[c] for c in ["vin","plate","brand","model","color","year"]])
        df.to_excel(CARS_XLSX, index=False)
    # ملف سجل الزيت ننشئه إن لم يوجد
    if not os.path.exists(OIL_XLSX):
        hdf = pd.DataFrame(columns=[H_COLS[k] for k in ["plate","vin","date","odometer","oil_run","notes","created_at"]])
        hdf.to_excel(OIL_XLSX, index=False)

def read_cars():
    ensure_files()
    df = pd.read_excel(CARS_XLSX, dtype=str).fillna("")

    # ✅ توحيد أسماء الأعمدة
    rename_map = {}
    for c in df.columns:
        clean = c.strip().replace(" ", "")
        if clean in ["فئىةالسيارة", "فئةالسياره", "فئةالسيارة"]:
            rename_map[c] = "فئة السيارة"
        elif clean in ["اللوحه", "الوحه", "اللوحة"]:
            rename_map[c] = "اللوحة"
        elif clean in ["رقمالهيكل", "رقمهيكل"]:
            rename_map[c] = "رقم الهيكل"
        elif clean in ["اللون"]:
            rename_map[c] = "اللون"
        elif clean in ["الشركه"]:
            rename_map[c] = "الشركة"
        elif clean in ["سنهاالصناعه", "سنةالصناعه"]:
            rename_map[c] = "سنة الصناعة"
    if rename_map:
        df = df.rename(columns=rename_map)

    # ✅ تنظيف النصوص من الرموز الخفية والفراغات الغريبة
    def clean_text(x):
        if not isinstance(x, str):
            return str(x)
        x = x.strip()
        x = x.replace("\u200f", "").replace("\u200e", "")  # رموز الاتجاه
        x = x.replace("ـ", "")  # شرطة التمديد
        x = re.sub(r"\s+", " ", x)  # تقليل المسافات
        return x

    for col in df.columns:
        df[col] = df[col].apply(clean_text)

    # ✅ تأكد من الأعمدة الإضافية الخاصة بتواريخ الزيت
    for c in ["last_oil_date", "last_odometer", "last_oil_run", "updated_at"]:
        col = AR_COLS[c]
        if col not in df.columns:
            df[col] = ""

    return df

def write_cars(df):
    df.to_excel(CARS_XLSX, index=False)

def read_oil():
    ensure_files()
    return pd.read_excel(OIL_XLSX, dtype=str).fillna("")

def write_oil(df):
    df.to_excel(OIL_XLSX, index=False)

def delete_history_by_plate_or_vin(old_plate: str, vin: str):
    """يحذف كل سجلات الزيت المرتبطة باللوحة القديمة أو رقم الهيكل"""
    h = read_oil()
    opn = normalize_plate(old_plate or "")
    if h.empty:
        return
    m = pd.Series([False]*len(h))
    if old_plate:
        m = m | (h[H_COLS["plate"]].astype(str).apply(normalize_plate) == opn)
    if vin:
        m = m | (h[H_COLS["vin"]].astype(str) == str(vin))
    if m.any():
        h = h[~m].copy()
        write_oil(h)

def normalize_plate(s: str) -> str:
    s = (s or "").strip()
    # إزالة مسافات وتطبيع بسيط للحروف العربية المتفرقة
    s = s.replace(" ", "").replace("ـ", "")
    s = s.replace("\u200f","").replace("\u200e","")
    return s

@app.route("/maintenance.html")
def maintenance_page():
    return render_template("maintenance.html")

@app.route("/rental.html")
def rental_page():
    return render_template("rental.html")

# ---------- API: قراءة السيارات مع بحث + فلتر ----------
@app.route("/api/cars", methods=["GET"])
def api_cars():
    q = (request.args.get("q") or "").strip()
    limit = (request.args.get("limit") or "all").lower()
    df = read_cars()

    if q:
        qn = normalize_plate(q)
        # نبحث في اللوحة والهيكل وبقية الأعمدة
        mask = (
            df[AR_COLS["plate"]].astype(str).apply(normalize_plate).str.contains(qn, na=False) |
            df[AR_COLS["vin"]].astype(str).str.contains(q, na=False) |
            df[AR_COLS["brand"]].astype(str).str.contains(q, case=False, na=False) |
            df[AR_COLS["model"]].astype(str).str.contains(q, case=False, na=False) |
            df[AR_COLS["color"]].astype(str).str.contains(q, case=False, na=False) |
            df[AR_COLS["year"]].astype(str).str.contains(q, na=False)
        )
        df = df[mask]

    # الترتيب من الأحدث تحديثاً للأقدم
    if AR_COLS["updated_at"] in df.columns:
        df = df.copy()
        # حاول تحويل التاريخ للترتيب
        try:
            df["_u"] = pd.to_datetime(df[AR_COLS["updated_at"]], errors="coerce")
            df = df.sort_values(by="_u", ascending=False)
            df = df.drop(columns=["_u"])
        except Exception:
            pass

    if limit.isdigit():
        df = df.head(int(limit))

    # نعيد صفوف كـ JSON
    records = df.to_dict(orient="records")
    return jsonify(records)

# ---------- API: تفاصيل سيارة واحدة ----------
@app.route("/api/car/<plate>", methods=["GET"])
def api_car_detail(plate):
    df = read_cars()
    pn = normalize_plate(plate)
    plate = requests.utils.unquote(plate)

    # نبحث باللوحة أو رقم الهيكل
    row = df[
        (df[AR_COLS["plate"]].astype(str).apply(normalize_plate) == pn)
        | (df[AR_COLS["vin"]].astype(str) == plate)
    ]

    if row.empty:
        return jsonify({"ok": False, "msg": "السيارة غير موجودة"}), 404
    return jsonify({"ok": True, "data": row.iloc[0].to_dict()})

# ---------- API: إضافة/تحديث سيارة (ورشة فقط) ----------
@app.route("/api/car/save", methods=["POST"])
def api_car_save():
    data = request.json or {}
    vin   = (data.get("vin") or "").strip()
    plate = (data.get("plate") or "").strip()
    brand = (data.get("brand") or "").strip()
    model = (data.get("model") or "").strip()
    color = (data.get("color") or "").strip()
    year  = (data.get("year") or "").strip()

    if not vin:
        return jsonify({"ok": False, "msg": "رقم الهيكل مطلوب"}), 400

    df = read_cars()

    # البحث عن السيارة حسب VIN (VIN لا يتغير)
    exist_idx = df.index[df[AR_COLS["vin"]].astype(str) == vin].tolist()
    if exist_idx:
        i = exist_idx[0]

        # فحص تغيير اللوحة
        old_plate = str(df.at[i, AR_COLS["plate"]] or "").strip()
        new_plate = plate.strip() if plate else old_plate
        old_norm = normalize_plate(old_plate)
        new_norm = normalize_plate(new_plate)

        # تحديث الحقول المسموحة فقط (VIN ثابت)
        if plate:
            df.at[i, AR_COLS["plate"]] = new_plate
        if brand: df.at[i, AR_COLS["brand"]] = brand
        if model: df.at[i, AR_COLS["model"]] = model
        if color: df.at[i, AR_COLS["color"]] = color
        if year:  df.at[i, AR_COLS["year"]]  = year

        # ✅ إذا تغيّرت اللوحة → احذف كل سجل الزيت المرتبط
        if old_plate and (new_norm != old_norm):
            delete_history_by_plate_or_vin(old_plate=old_plate, vin=vin)
            for c in ["last_oil_date", "last_odometer", "last_oil_run"]:
                df.at[i, AR_COLS[c]] = ""

        df.at[i, AR_COLS["updated_at"]] = _now()

    else:
        # إنشاء سجل جديد (يُسمح بإدخال VIN جديد لأن هذا تسجيل جديد)
        new_row = {
            AR_COLS["vin"]: vin,
            AR_COLS["plate"]: plate,
            AR_COLS["brand"]: brand,
            AR_COLS["model"]: model,
            AR_COLS["color"]: color,
            AR_COLS["year"]: year,
            AR_COLS["last_oil_date"]: "",
            AR_COLS["last_odometer"]: "",
            AR_COLS["last_oil_run"]: "",
            AR_COLS["updated_at"]: _now(),
        }
        df = pd.concat([pd.DataFrame([new_row]), df], ignore_index=True)

    write_cars(df)
    return jsonify({"ok": True, "msg": "تم الحفظ بنجاح"})

# ---------- API: سجل الزيت لسيارة ----------
@app.route("/api/oil_history/<plate>", methods=["GET"])
def api_oil_history(plate):
    h = read_oil()
    pn = normalize_plate(plate)
    mask = h[H_COLS["plate"]].astype(str).apply(normalize_plate) == pn
    sub = h[mask].copy()
    # أحدث سجل أولاً
    try:
        sub["_c"] = pd.to_datetime(sub[H_COLS["created_at"]], errors="coerce")
        sub = sub.sort_values(by="_c", ascending=False).drop(columns=["_c"])
    except Exception:
        pass
    return jsonify(sub.to_dict(orient="records"))

# ---------- API: إضافة سجل تغيير زيت (ورشة فقط) ----------
@app.route("/api/oil_history/add", methods=["POST"])
def api_oil_add():
    data = request.json or {}
    plate    = (data.get("plate") or "").strip()
    vin      = (data.get("vin") or "").strip()
    date     = (data.get("date") or "").strip()
    odometer = (data.get("odometer") or "").strip()
    oil_run  = (data.get("oil_run") or "").strip()
    notes    = (data.get("notes") or "").strip()

    if not plate and not vin:
        return jsonify({"ok": False, "msg": "اللوحة أو رقم الهيكل مطلوب"}), 400

    # ✅ إلزامية التاريخ والممشى
    if not date or not odometer:
        return jsonify({"ok": False, "msg": "يجب إدخال التاريخ وعداد السيارة"}), 400

    h = read_oil()
    newh = {
        H_COLS["plate"]: plate,
        H_COLS["vin"]: vin,
        H_COLS["date"]: date,
        H_COLS["odometer"]: odometer,
        H_COLS["oil_run"]: oil_run,
        H_COLS["notes"]: notes,
        H_COLS["created_at"]: _now()
    }
    h = pd.concat([pd.DataFrame([newh]), h], ignore_index=True)
    write_oil(h)

    # تحديث بيانات السيارة
    if vin or plate:
        df = read_cars()
        if vin:
            idx = df.index[df[AR_COLS["vin"]].astype(str) == vin].tolist()
        else:
            pn = normalize_plate(plate)
            idx = df.index[df[AR_COLS["plate"]].astype(str).apply(normalize_plate) == pn].tolist()
        if idx:
            i = idx[0]
            df.at[i, AR_COLS["last_oil_date"]] = date
            df.at[i, AR_COLS["last_odometer"]] = odometer
            if oil_run:
                df.at[i, AR_COLS["last_oil_run"]] = oil_run
            df.at[i, AR_COLS["updated_at"]] = _now()
            write_cars(df)

    return jsonify({"ok": True, "msg": "تم إضافة سجل تغيير الزيت بنجاح"})



@app.route('/api/admin/requests/list', methods=['POST'])
def admin_list_requests():
    try:
        df = load_requests()
        if df.empty:
            return jsonify([])

        # 🔥 لا تحويل إلى Boolean!
        # فقط تأكد أن القيمة "1" أو "0"
        df["مؤرشف"] = df["مؤرشف"].astype(str).str.strip()
        df["مؤرشف"] = df["مؤرشف"].apply(lambda x: "1" if x in ["1","نعم","true","True","yes","y"] else "0")

        return jsonify(df.fillna('').to_dict(orient='records'))

    except Exception as e:
        print("❌ admin_list_requests error:", e)
        return jsonify([]), 500

@app.route("/api/admin/requests/archive", methods=["POST"])
def archive_request():
    data = request.get_json()
    req_id = str(data.get("request_id")).strip()
    archive = data.get("archive", False)
    updated_by = data.get("updated_by", "admin")

    # تحميل الملف بدون أي تطبيع
    df = pd.read_excel(REQUESTS_XLSX, dtype=str)

    # تنظيف القيم
    df.columns = df.columns.str.strip()
    df = df.applymap(lambda x: x.strip() if isinstance(x,str) else x)

    # البحث الصحيح عن رقم الطلب
    idx = df.index[df["رقم الطلب"].astype(str).str.strip() == req_id]

    if len(idx) == 0:
        print("❌ رقم الطلب لم يُعثر عليه داخل Excel:", req_id)
        return jsonify({"success": False, "msg": "Request not found"})

    row = idx[0]

    # تعديل قيمة المؤرشف
    df.loc[row, "مؤرشف"] = "1" if archive else "0"
    df.loc[row, "آخر تحديث بواسطة"] = updated_by

    # حفظ التعديل
    df.to_excel(REQUESTS_XLSX, index=False)

    print("✅ تم تحديث أرشفة الطلب:", req_id)
    return jsonify({"success": True})



@app.route('/api/get_departments')
def get_departments():
    import pandas as pd
    try:
        df = pd.read_excel(USERS_XLSX)

        # الأقسام الأساسية
        main_depts = set()

        if "القسم" in df.columns:
            for d in df["القسم"].dropna().tolist():
                d = str(d).strip()
                if d:
                    main_depts.add(d)

        # الأقسام الإضافية (الأقسام الأخرى)
        extra_depts = set()

        if "الأقسام الأخرى" in df.columns:
            for row in df["الأقسام الأخرى"].dropna().tolist():
                row = str(row).strip()
                if row:
                    # تقسيم بالقيم: "،" أو ","
                    parts = re.split(r"[،,]", row)
                    for p in parts:
                        p = p.strip()
                        if p:
                            extra_depts.add(p)

        # دمج الأقسام
        all_depts = main_depts.union(extra_depts)

        # استثناء الأقسام الخاصة بالأدمن
        blacklist = ["admin", "ادمن", "مشرف", "مدير نظام", "إدارة النظام"]

        clean = []
        for d in all_depts:
            dn = str(d).strip()
            dn_norm = dn.replace("إ", "ا").replace("أ", "ا").replace("آ", "ا").lower()
            if any(b in dn_norm for b in blacklist):
                continue
            clean.append(dn)

        # ترتيب أبجدي
        clean = sorted(clean)

        return jsonify({"departments": clean})

    except Exception as e:
        return jsonify({"departments": [], "error": str(e)})

# ============================================================
# 🔌 CORE SYSTEM (stor7s-backend) INJECTION
# ============================================================

def require_core_access():
    user = session.get("user")
    if not user:
        return False

    apps = user.get("apps", [])
    if not isinstance(apps, list):
        return False

    # السماح بالدخول إذا لديه نظام المستودع
    return ("warehouse" in apps) or ("core" in apps) or ("all" in apps)

@app.route("/core")
def core_entry():
    if not session.get("user"):
        return redirect("/Login.html")

    user = session["user"]

    apps = user.get("apps", [])
    if isinstance(apps, str):
        apps = [a.strip().lower() for a in apps.split(",")]

    role_norm = normalize_role(user.get("role", ""))

    allowed_by_apps = any(a in apps for a in ["warehouse", "core", "all"])
    allowed_by_role = role_norm in ["manager", "general_manager", "admin"]

    if not (allowed_by_apps or allowed_by_role):
        return "🚫 غير مصرح لك بالدخول إلى نظام المستودع والعهد", 403

    role = normalize_arabic(user.get("role", ""))
    dept = normalize_arabic(user.get("department", ""))

    page = None  # ⭐ تعريف صريح

    # ===============================
    # 1️⃣ أولوية مدير القسم
    # ===============================
    if role_norm == "manager":
        page = "manager1.html"

    # ===============================
    # 2️⃣ التوجيه حسب القسم
    # ===============================
    elif "تقنية" in dept or "it" in dept:
        page = "it1.html"

    elif "مالي" in dept:
        page = "finance1.html"

    elif "مشتريات" in dept:
        page = "purchasing1.html"

    elif "موارد" in dept:
        page = "hr1.html"

    elif "ادارة" in dept and "العامة" in dept:
        page = "admin1.html"

    # ===============================
    # 2️⃣ fallback حسب الصلاحية فقط
    # ===============================
    if not page:
        if role == "manager":
            page = "manager1.html"

        elif role in ["general_manager", "admin"]:
            page = "admin1.html"

        elif role == "employee":
            page = "employee1.html"

    # ===============================
    # 3️⃣ حماية نهائية
    # ===============================
    if not page:
        return "❌ لا توجد صفحة مناسبة لهذا المستخدم", 403

    return render_template(f"templates1/{page}")

# 📡 Core APIs (Blueprints)
from modules.employee import api as core_employee_api
from modules.manager import api as core_manager_api
from modules.purchasing import api as core_purchasing_api
from modules.it import api as core_it_api
from modules.hr import api as core_hr_api
from modules.finance import api as core_finance_api
from modules.admin import api as core_admin_api

app.register_blueprint(core_employee_api, url_prefix="/api/core/employee")
app.register_blueprint(core_manager_api, url_prefix="/api/core/manager")
app.register_blueprint(core_purchasing_api, url_prefix="/api/core/purchasing")
app.register_blueprint(core_it_api, url_prefix="/api/core/it")
app.register_blueprint(core_hr_api, url_prefix="/api/core/hr")
app.register_blueprint(core_finance_api, url_prefix="/api/core/finance")
app.register_blueprint(core_admin_api, url_prefix="/api/core/admin")

# ============== التشغيل ==============
if __name__ == "__main__":
    import os
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
