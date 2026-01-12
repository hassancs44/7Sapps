from flask import Blueprint, request, jsonify
from utils.excel import load, save
from utils.workflow import to_purchasing
import re

api = Blueprint("manager", __name__)

COLS = ["رقم_الطلب", "الحالة"]


# ✅ دالة التطبيع (محلية – بدون أي import خطير)
def normalize_arabic(text):
    if not text:
        return ""
    text = str(text)

    # توحيد المسافات
    text = re.sub(r"\s+", " ", text).strip()

    # إزالة محارف RTL الخفية
    text = text.replace("\u200f", "").replace("\u200e", "")

    return text


@api.get("/pending/<dept>")
def pending(dept):
    df = load("requests")

    dept_norm = normalize_arabic(dept)
    df["القسم_norm"] = df["القسم"].astype(str).apply(normalize_arabic)

    df = df[
        (df["القسم_norm"] == dept_norm) &
        (df["الحالة"].isin([
            "جديد",
            "بانتظار مدير القسم"
        ]))
    ]

    return jsonify(df.fillna("").to_dict("records"))



@api.post("/approve")
def approve():
    data = request.get_json() or {}
    num = str(data.get("رقم_الطلب", "")).strip()

    if not num:
        return jsonify({"ok": False, "msg": "رقم الطلب غير صحيح"})

    df = load("requests")

    mask = df["رقم_الطلب"].astype(str) == num
    if not mask.any():
        return jsonify({"ok": False, "msg": "الطلب غير موجود"})

    # ✅ تحديث الحالة
    df.loc[mask, "الحالة"] = "بانتظار المشتريات"
    save("requests", df)

    # ✅ تحويل الطلب
    to_purchasing(num)

    return jsonify({"ok": True, "msg": "✔️ تم اعتماد الطلب وتحويله للمشتريات"})
