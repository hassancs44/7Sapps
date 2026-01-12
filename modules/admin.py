from flask import Blueprint, request, jsonify
from datetime import datetime
from utils.excel import load, save, append
from utils.workflow import initial_state

api = Blueprint("admin", __name__)

# =========================
# ✅ Helpers
# =========================
def now_date_time():
    dt = datetime.now()
    return dt.strftime("%Y-%m-%d"), dt.strftime("%H:%M:%S")

def next_request_id(df_requests):
    # يولّد رقم طلب جديد بشكل آمن
    if df_requests.empty or "رقم_الطلب" not in df_requests.columns:
        return "1"
    s = df_requests["رقم_الطلب"].astype(str).str.strip()
    nums = []
    for x in s:
        try:
            nums.append(int(x))
        except:
            pass
    return str(max(nums) + 1) if nums else "1"


# =========================
# ✅ Dashboard / Lists
# =========================
@api.get("/requests")
def get_requests():
    df = load("requests")
    return jsonify(df.to_dict("records"))

@api.get("/custody")
def get_custody():
    df = load("custody")
    return jsonify(df.to_dict("records"))

@api.get("/logs")
def get_logs():
    df = load("logs")
    return jsonify(df.to_dict("records"))

@api.get("/approvals")
def get_approvals():
    df = load("approvals")
    return jsonify(df.to_dict("records"))

@api.get("/purchase")
def get_purchase():
    df = load("purchase")
    return jsonify(df.to_dict("records"))


# =========================
# ✅ إنشاء طلب شراء من الإدارة (تعميد تلقائي)
# =========================
@api.post("/create_request")
def create_request():
    data = request.get_json() or {}

    # بيانات المستخدم من الفرونت (من localStorage)
    creator_name = str(data.get("name", "")).strip()
    creator_role = str(data.get("role", "الإدارة العامة")).strip()
    department   = str(data.get("department", "الإدارة العامة")).strip()
    company      = str(data.get("company", "")).strip()
    branch       = str(data.get("branch", "")).strip()

    req_type     = str(data.get("type", "شراء")).strip()
    desc         = str(data.get("description", "")).strip()


    if not creator_name:
        return jsonify({"ok": False, "msg": "اسم المستخدم مطلوب"}), 400

    if not desc:
        return jsonify({"ok": False, "msg": "الوصف مطلوب"}), 400

    df_requests = load("requests")
    req_id = next_request_id(df_requests)

    # ✅ حالة الإدارة العامة: بانتظار المشتريات مباشرة
    state = initial_state("الإدارة العامة")

    # 1) إضافة الطلب في requests
    row_req = {
        "رقم_الطلب": req_id,
        "الدور": "الإدارة العامة",
        "الرافع": creator_name,
        "القسم": department,
        "الشركة": company,
        "الفرع": branch,
        "النوع": req_type,
        "الحالة": state,
        "الوصف": desc
    }

    df_requests.loc[len(df_requests)] = row_req
    save("requests", df_requests)

    # 3) ✅ اعتماد تلقائي في approvals (تعميد الإدارة)
    d, t = now_date_time()
    df_appr = load("approvals")
    df_appr.loc[len(df_appr)] = {
        "رقم_الطلب": req_id,
        "الجهة": "الإدارة العامة",
        "المعتمد": creator_name,
        "التاريخ": d,
        "الوقت": t,
        "ملاحظات": "تعميد تلقائي لطلبات الإدارة العامة"
    }
    save("approvals", df_appr)

    # 4) ✅ log
    df_logs = load("logs")
    df_logs.loc[len(df_logs)] = {
        "رقم_الطلب": req_id,
        "الحدث": "إنشاء طلب (إدارة عامة) + تعميد تلقائي",
        "منفذ": creator_name,
        "الدور": "الإدارة العامة",
        "التاريخ": d,
        "الوقت": t,
        "ملاحظات": f"الحالة -> {state}"
    }
    save("logs", df_logs)

    return jsonify({"ok": True, "msg": "تم إنشاء الطلب وتعميده تلقائياً", "رقم_الطلب": req_id})


# =========================
# ✅ تعميد طلبات الأقسام التي بدون مدير قسم
# (بديل نظيف عن استدعاء /api/manager/approve)
# =========================
@api.post("/approve_request")
def approve_request():
    data = request.get_json() or {}
    req_id = str(data.get("رقم_الطلب","")).strip()

    admin_name = str(data.get("admin_name","")).strip()
    note       = str(data.get("note","")).strip() or "تعميد إداري (لا يوجد مدير قسم)"

    if not req_id:
        return jsonify({"ok": False, "msg": "رقم الطلب مطلوب"}), 400
    if not admin_name:
        return jsonify({"ok": False, "msg": "اسم المعتمد مطلوب"}), 400

    df = load("requests")

    if "رقم_الطلب" not in df.columns:
        return jsonify({"ok": False, "msg": "ملف الطلبات لا يحتوي على رقم_الطلب"}), 400

    mask = df["رقم_الطلب"].astype(str).str.strip() == req_id
    if not mask.any():
        return jsonify({"ok": False, "msg": "الطلب غير موجود"}), 404

    current = str(df.loc[mask, "الحالة"].iloc[0]).strip()

    # نعمد فقط إذا كان بانتظار مدير القسم (أو تبغى تسمح لغيرها)
    if current != "بانتظار مدير القسم":
        return jsonify({"ok": False, "msg": f"لا يمكن التعميد لأن الحالة الحالية: {current}"}), 400

    # ✅ بعد التعميد يروح للمشتريات
    df.loc[mask, "الحالة"] = "بانتظار المشتريات"
    save("requests", df)

    d, t = now_date_time()

    # approvals
    df_appr = load("approvals")
    df_appr.loc[len(df_appr)] = {
        "رقم_الطلب": req_id,
        "الجهة": "الإدارة العامة",
        "المعتمد": admin_name,
        "التاريخ": d,
        "الوقت": t,
        "ملاحظات": note
    }
    save("approvals", df_appr)

    # logs
    df_logs = load("logs")
    df_logs.loc[len(df_logs)] = {
        "رقم_الطلب": req_id,
        "الحدث": "تعميد إداري",
        "منفذ": admin_name,
        "الدور": "الإدارة العامة",
        "التاريخ": d,
        "الوقت": t,
        "ملاحظات": f"{note} | الحالة: {current} -> بانتظار المشتريات"
    }
    save("logs", df_logs)

    return jsonify({"ok": True, "msg": "تم التعميد وتحويل الطلب للمشتريات"})
