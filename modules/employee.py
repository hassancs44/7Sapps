import datetime
from utils.excel import load, save, append
from utils.files import upload
from utils.workflow import initial_state
from flask import Blueprint, request, jsonify
from utils.excel import load, save, append
from utils.files import upload
from utils.workflow import initial_state
import os
from werkzeug.utils import secure_filename

api = Blueprint("employee", __name__)

REQ_COLS = ["رقم_الطلب", "الدور", "الرافع", "القسم", "الشركة", "الفرع", "النوع", "الحالة", "الوصف"]
ITEM_COLS = ["رقم_الطلب", "كود", "اسم", "كمية", "ملاحظات"]
ATT_COLS = [
    "رقم_الطلب","اسم_الملف","رافع","دور","القسم","الشركة","الفرع","تاريخ","وقت"
]

@api.post("/create")
def create_request():
    data = request.form
    df = load("requests")
    new_id = len(df) + 1

    role = data.get("الدور", "")

    # تحديد الحالة حسب الدور (موحد مع النظام)
    الحالة = initial_state(role)

    append("requests", [
        new_id,
        data.get("الدور", ""),
        data.get("الرافع", ""),
        data.get("القسم", ""),
        data.get("الشركة", "غير محدد"),
        data.get("الفرع", ""),
        data.get("النوع", ""),
        الحالة,
        data.get("الوصف", "")
    ], REQ_COLS)

    # ☑️ لو المدير رفع الطلب → نرسله للمشتريات مباشرة
    if role in ["مدير قسم", "الموارد البشرية", "الإدارة العامة"]:
        from utils.workflow import to_purchasing
        to_purchasing(new_id)

    return jsonify({"done": True, "رقم_الطلب": new_id})



@api.post("/upload")
def upload_file():
    f = request.files.get("file")
    req = request.form.get("رقم_الطلب")

    if not f or not req:
        return jsonify({"error": "❌ ملف أو رقم الطلب مفقود"}), 400

    from app import UPLOAD_FOLDER
    os.makedirs(UPLOAD_FOLDER, exist_ok=True)

    filename = f"REQ{req}_{secure_filename(f.filename)}"
    f.save(os.path.join(UPLOAD_FOLDER, filename))

    df = load("attachments")
    df.loc[len(df)] = {
        "رقم_الطلب": req,
        "اسم_الملف": filename,
        "رافع": request.form.get("رافع", "غير معروف"),
        "دور": request.form.get("دور", "موظف"),
        "القسم": request.form.get("القسم", ""),
        "الشركة": request.form.get("الشركة", ""),
        "الفرع": request.form.get("الفرع", ""),
        "تاريخ": str(datetime.date.today()),
        "وقت": datetime.datetime.now().strftime("%H:%M")
    }
    save("attachments", df)

    return jsonify({"ok": True, "msg": "✔️ تم رفع المرفق بنجاح", "file": filename})


@api.get("/my/<name>")
def my(name):
    df = load("requests")
    result = df[df["الرافع"] == name]
    return jsonify(result.to_dict("records"))


@api.get("/attachments/<req_id>")
def attachments(req_id):
    df = load("attachments")
    files = df[df["رقم_الطلب"].astype(str)==str(req_id)]["اسم_الملف"].tolist()
    return jsonify(files)






