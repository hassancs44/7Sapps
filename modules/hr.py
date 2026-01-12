from flask import Blueprint, request, jsonify
from utils.excel import load, save, append
from datetime import datetime
from utils.id import generate_custody_id
from utils.excel import load, save, append

api = Blueprint("hr", __name__)


# ===============================
# ➕ HR | إنشاء طلب شراء (اعتماد تلقائي)
# ===============================
@api.post("/request/create")
def hr_create_request():
    d = request.json
    now = datetime.now()

    df = load("requests")
    new_id = str(len(df) + 1)

    append("requests", [
        new_id,
        "الموارد البشرية",
        d["الرافع"],
        d["القسم"],
        d["الشركة"],
        d["الفرع"],
        d.get("النوع", "شراء"),
        "بانتظار المشتريات",
        d.get("الوصف", "")
    ])

    append("logs", [
        new_id,
        "إنشاء طلب شراء",
        d["الرافع"],
        "HR",
        now.date(),
        now.time(),
        "طلب موارد بشرية معتمد تلقائياً"
    ])

    return jsonify({"ok": True, "رقم_الطلب": new_id})

@api.get("/custody/all")
def custody_all():
    df = load("custody")
    # عرض كل العهد ما عدا المقفلة
    df["الحالة"] = df["الحالة"].astype(str).fillna("").str.strip()
    return jsonify(df[df["الحالة"] != "مقفلة"].to_dict("records"))


# ===============================
# 📌 عرض جميع العهد
# ===============================
@api.get("/custody/<custody_id>")
def custody_details(custody_id):
    custody = load("custody")
    reqs    = load("requests")
    items   = load("items")
    attach  = load("attachments")
    logs    = load("logs")

    c = custody[custody["رقم_العهدة"] == custody_id]
    if c.empty:
        return jsonify({"ok": False}), 404

    req_id = c.iloc[0]["رقم_الطلب"]

    return jsonify({
        "custody": c.to_dict("records")[0],
        "request": reqs[reqs["رقم_الطلب"] == req_id].to_dict("records"),
        "items": items[items["رقم_الطلب"] == req_id].to_dict("records"),
        "attachments": attach[attach["رقم_الطلب"] == req_id].to_dict("records"),
        "logs": logs[logs["رقم_الطلب"] == req_id].to_dict("records")
    })

# ===============================
# ➕ إضافة عهد متعددة لموظف واحد
# ===============================
@api.post("/custody/add-multi")
def add_multi_custody():
    d = request.json
    now = datetime.now()

    # 🔎 جلب بيانات الموظف الحقيقي
    users = load("users")
    u = users[users["اسم_المستخدم"] == d["الموظف"]]

    if u.empty:
        return jsonify({"ok": False, "msg": "❌ الموظف غير موجود"}), 400

    emp_department = u.iloc[0]["القسم"]
    emp_branch     = u.iloc[0]["الفرع"]

    for item in d["items"]:
        append("custody", [
            generate_custody_id(),          # رقم_العهدة
            d.get("رقم_الطلب",""),         # رقم_الطلب
            item.get("كود",""),            # كود_الصنف
            item.get("اسم",""),            # اسم_الصنف
            item.get("نوع","جهاز"),        # نوع_العهدة
            item.get("سيريال",""),         # سيريال
            d["الموظف"],                   # الموظف
            emp_department,                # ✅ القسم الصحيح
            emp_branch,                    # ✅ الفرع الصحيح
            item.get("كمية",1),            # الكمية
            now.strftime("%Y-%m-%d"),       # تاريخ_التسليم
            "",                             # تاريخ_الاسترجاع
            "نشطة",                         # الحالة
            item.get("ملاحظات",""),        # ملاحظات
            item.get("اسم","")             # الجهاز
        ])

    append("logs", [
        d.get("رقم_الطلب",""),
        "إضافة عهد متعددة",
        "HR",
        "إشراف",
        now.date(),
        now.time(),
        f"إضافة عهد متعددة للموظف {d['الموظف']}"
    ])

    return jsonify({"ok": True})


# ===============================
# 🔁 نقل عهدة
# ===============================
@api.post("/custody/transfer")
def transfer():
    df = load("custody")
    users = load("users")

    cid = request.json["رقم_العهدة"]
    new_emp = request.json["الموظف"]

    # 🔎 جلب بيانات الموظف الجديد من users.xlsx
    u = users[users["اسم_المستخدم"] == new_emp]
    if u.empty:
        return jsonify({"ok": False, "msg": "❌ الموظف غير موجود"}), 400

    new_dep = u.iloc[0]["القسم"]
    new_branch = u.iloc[0]["الفرع"]

    append("logs", [
        "",
        "نقل عهدة",
        "HR",
        "إشراف",
        datetime.now().date(),
        datetime.now().time(),
        f"نقل العهدة {cid} إلى {new_emp} ({new_dep})"
    ])

    df.loc[df["رقم_العهدة"] == cid, ["الموظف", "القسم", "الفرع"]] = [
        new_emp,
        new_dep,
        new_branch
    ]

    save("custody", df)
    return jsonify({"ok": True})


# ===============================
# 🚫 إقفال عهدة
# ===============================
@api.post("/custody/close")
def close():
    data = request.json
    cid  = data.get("رقم_العهدة")

    custody = load("custody")
    wh      = load("warehouse")

    row = custody[custody["رقم_العهدة"] == cid]
    if row.empty:
        return jsonify({"ok": False, "msg": "العهدة غير موجودة"}), 404

    row = row.iloc[0]

    # تحديث العهدة
    custody.loc[custody["رقم_العهدة"] == cid, ["الحالة","تاريخ_الاسترجاع"]] = [
        "مقفلة",
        datetime.now().strftime("%Y-%m-%d")
    ]

    # إرجاع الكمية للمستودع
    wh.loc[wh["كود"] == row["كود_الصنف"], "كمية_حالياً"] = (
        wh.loc[wh["كود"] == row["كود_الصنف"], "كمية_حالياً"].astype(int)
        + int(row["الكمية"])
    )

    save("custody", custody)
    save("warehouse", wh)

    # Log
    append("logs", [
        row["رقم_الطلب"],
        "إقفال عهدة",
        "HR",
        "إقفال",
        datetime.now().date(),
        datetime.now().time(),
        f"إقفال العهدة {cid} وإرجاع {row['الكمية']} للمستودع"
    ])

    return jsonify({"ok": True})



# ===============================
# 📄 عرض الطلبات (Read Only)
# ===============================
@api.get("/requests")
def view_requests():
    return jsonify(load("requests").to_dict("records"))

# ===============================
# 📜 سجل الحركات
# ===============================
@api.get("/logs")
def logs():
    return jsonify(load("logs").to_dict("records"))


# ===============================
# ✅ استلام عهدة
# ===============================
@api.post("/custody/receive")
def receive_custody():
    cid = request.json["رقم_العهدة"]
    df = load("custody")

    df.loc[df["رقم_العهدة"]==cid, "الحالة"] = "تم التسليم"
    save("custody", df)

    return jsonify({"ok":True})


# ===============================
# 🔄 تسليم عهدة
# ===============================
@api.post("/custody/dispatch")
def dispatch_custody():
    cid = request.json["رقم_العهدة"]
    df = load("custody")

    df.loc[df["رقم_العهدة"]==cid, "الحالة"] = "قيد التسليم"
    save("custody", df)

    return jsonify({"ok":True})
