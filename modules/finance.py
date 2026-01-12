from flask import Blueprint, jsonify, request

from utils.excel import load

api = Blueprint("finance", __name__)

@api.get("/executed")
def executed():
    df = load("requests")
    df = df[df["الحالة"].str.contains("تم")]
    return jsonify(df.to_dict("records"))

@api.get("/purchases")
def purchases():
    purchases = load("purchase")
    requests  = load("requests")

    # ربط الشراء بالطلب
    merged = purchases.merge(
        requests,
        on="رقم_الطلب",
        how="left"
    )

    return jsonify(merged.to_dict("records"))

@api.get("/attachments/<req_id>")
def attachments(req_id):
    df = load("attachments")
    r = df[df["رقم_الطلب"] == str(req_id)]
    return jsonify(r.to_dict("records"))

from datetime import datetime
from utils.excel import append

@api.post("/create-request")
def create_request():
    data = request.get_json() or {}

    from utils.excel import load, save
    from datetime import datetime

    # تحميل الطلبات الحالية
    df = load("requests")

    # الأعمدة الفعلية لملف الطلبات.xlsx
    required_cols = [
        "رقم_الطلب",
        "الدور",
        "الرافع",
        "القسم",
        "الشركة",
        "الفرع",
        "النوع",
        "الحالة",
        "الوصف",
    ]

    # ضمان وجود الأعمدة (لو ملف قديم)
    for col in required_cols:
        if col not in df.columns:
            df[col] = ""

    # إنشاء رقم الطلب
    req_id = f"FIN-{int(datetime.now().timestamp())}"

    new_row = {
        "رقم_الطلب": req_id,
        "الدور": "المالية",
        "الرافع": "المالية",
        "القسم": "المالية",
        "الشركة": data.get("الشركة", ""),
        "الفرع": data.get("الفرع", ""),
        "النوع": "طلب مالي",
        "الحالة": "بانتظار المشتريات",
        "الوصف": data.get("الوصف", ""),
    }

    # الإضافة الآمنة
    df.loc[len(df)] = new_row
    save("requests", df)

    # سجل الحركات (هذا ملفه مختلف ومضبوط)
    from utils.excel import append
    append(
        "logs",
        [
            req_id,  # رقم_الطلب
            "إنشاء طلب مالي",  # الحدث
            "المالية",  # منفذ
            "المالية",  # الدور
            datetime.now().strftime("%Y-%m-%d"),  # التاريخ
            datetime.now().strftime("%H:%M:%S"),  # الوقت
            "اعتماد تلقائي"  # ملاحظات
        ],
        cols=[
            "رقم_الطلب",
            "الحدث",
            "منفذ",
            "الدور",
            "التاريخ",
            "الوقت",
            "ملاحظات"
        ]
    )

    return jsonify({
        "ok": True,
        "msg": "✔️ تم إنشاء الطلب المالي وتحويله للمشتريات",
        "رقم_الطلب": req_id
    })


