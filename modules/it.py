from flask import Blueprint, request, jsonify
from utils.excel import load, save, append
from werkzeug.utils import secure_filename
from datetime import datetime
import os

api = Blueprint("it", __name__)

# =========================================
# 📥 الطلبات المحالة من المشتريات إلى IT
# =========================================
@api.get("/incoming")
def incoming():
    df = load("requests")
    df = df[df["الحالة"] == "محول لقسم IT"]
    return jsonify(df.to_dict("records"))


# =========================================
# 🛠️ تقييم فني + توصية + إعادة الطلب للمشتريات
# =========================================
@api.post("/evaluate")
def evaluate():
    data = request.form

    req_id = str(data.get("رقم_الطلب", "")).strip()
    evaluation = data.get("نوع_التقييم", "")
    recommendation = data.get("التوصية", "")
    notes = data.get("الوصف_الفني", "")
    tech = data.get("اسم_الفني", "")

    # ===============================
    # 📎 رفع المرفق إن وجد
    # ===============================
    file = request.files.get("file")
    filename = ""

    if file and file.filename:
        filename = secure_filename(file.filename)

        upload_dir = os.path.join(
            os.path.dirname(__file__), "..", "uploads"
        )
        os.makedirs(upload_dir, exist_ok=True)

        file.save(os.path.join(upload_dir, filename))

    # ===============================
    # 📄 حفظ تقرير IT
    # ===============================
    append("it_reports", [
        req_id,
        evaluation,          # نوع التقييم
        recommendation,      # التوصية
        notes,               # الوصف الفني
        tech,                # اسم الفني
        datetime.now().date(),
        datetime.now().time(),
        filename              # المرفق
    ])

    # ===============================
    # 🔄 تحديث حالة الطلب
    # ===============================
    df = load("requests")
    df.loc[
        df["رقم_الطلب"].astype(str) == req_id,
        "الحالة"
    ] = "أعيد من IT"
    save("requests", df)

    # ===============================
    # 📝 تسجيل Log
    # ===============================
    append("logs", [
        req_id,
        "تقييم فني",
        tech,
        "تقنية المعلومات",
        datetime.now().date(),
        datetime.now().time(),
        recommendation
    ])

    return jsonify({
        "ok": True,
        "msg": "✔️ تم إرسال التقييم الفني وإعادة الطلب للمشتريات"
    })


# =========================================
# 🛒 رفع طلب شراء جديد من IT (اعتماد تلقائي)
# =========================================
@api.post("/create-request")
def create_request_from_it():
    data = request.get_json()

    req_id = str(data.get("رقم_الطلب"))
    items  = data.get("items", [])
    user   = data.get("user", {})

    # ===============================
    # 📄 إنشاء الطلب
    # ===============================
    append("requests", [
        req_id,
        "تقنية المعلومات",
        user.get("name",""),
        user.get("department","IT"),
        user.get("company",""),
        user.get("branch",""),
        "شراء",
        "بانتظار المشتريات",
        data.get("الوصف","")
    ])

    # ===============================
    # 📦 تفاصيل الطلب
    # ===============================
    for item in items:
        append("items", [
            req_id,
            item.get("كود",""),
            item.get("اسم",""),
            item.get("كمية",1),
            item.get("ملاحظات","")
        ])

    # ===============================
    # 📝 Log
    # ===============================
    append("logs", [
        req_id,
        "طلب شراء مباشر من IT",
        user.get("name",""),
        "تقنية المعلومات",
        datetime.now().date(),
        datetime.now().time(),
        "اعتماد تلقائي"
    ])

    return jsonify({
        "ok": True,
        "msg": "✔️ تم إرسال طلب الشراء للمشتريات مباشرة"
    })
