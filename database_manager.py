# database_manager.py — النسخة المعدّلة الكاملة
"""
database_manager.py — قاعدة البيانات المحلية بديل Excel
مع إضافة: الإجراءات التأديبية + معلومات الموظف الإضافية + التقييم السابق
"""
import json, os
import pandas as pd
import streamlit as st
from datetime import date, datetime
from constants import MONTH_MAP, PERSONAL_KPIS, PERSONAL_WEIGHT

DB_DIR  = "db"
EMP_DB  = os.path.join(DB_DIR, "employees.json")
KPI_DB  = os.path.join(DB_DIR, "kpis.json")
DATA_DB = os.path.join(DB_DIR, "evaluations.json")
META_FILE = os.path.join(DB_DIR, "db_meta.json")

# ═══════════════════════════════════════════════════════════════════
# ✅ جديد: ملفات الإجراءات التأديبية ومعلومات الموظفين الإضافية
# ═══════════════════════════════════════════════════════════════════
DISCIPLINARY_DB = os.path.join(DB_DIR, "disciplinary.json")
EMP_EXTRA_DB = os.path.join(DB_DIR, "employees_extra.json")

os.makedirs(DB_DIR, exist_ok=True)


def db_exists():
    return all(os.path.exists(p) for p in [EMP_DB, KPI_DB, DATA_DB])


def get_db_meta():
    if os.path.exists(META_FILE):
        with open(META_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}


# ═══════════════════════════════════════════════════════════════════
# تحميل البيانات الأساسية
# ═══════════════════════════════════════════════════════════════════
@st.cache_data(ttl=30)
def load_data_from_db():
    try:
        with open(EMP_DB,  "r", encoding="utf-8") as f: emp_records  = json.load(f)
        with open(KPI_DB,  "r", encoding="utf-8") as f: kpi_records  = json.load(f)
        with open(DATA_DB, "r", encoding="utf-8") as f: data_records = json.load(f)

        df_emp  = pd.DataFrame(emp_records)
        df_kpi  = pd.DataFrame(kpi_records)
        df_data = pd.DataFrame(data_records)

        # ── تنظيف جميع الأعمدة النصية ─────────────────────────────
        for df in [df_emp, df_kpi, df_data]:
            df.columns = [str(c).strip() for c in df.columns]
            for col in df.columns:
                if str(df[col].dtype) in ('object', 'str', 'string'):
                    df[col] = df[col].astype(str).str.strip()

        # ── تصحيح الأنواع الرقمية ──────────────────────────────────
        df_kpi["Weight"] = pd.to_numeric(df_kpi["Weight"], errors="coerce").fillna(0)
        for col in ["KPI_%", "Weight"]:
            if col in df_data.columns:
                df_data[col] = pd.to_numeric(df_data[col], errors="coerce").fillna(0)
        if "Year" in df_data.columns:
            df_data["Year"] = pd.to_numeric(df_data["Year"], errors="coerce").fillna(2025).astype(int)

        # ── توحيد اسم عمود المقيم ─────────────────────────────────
        if "اسم المقيم " not in df_emp.columns and "اسم المقيم" in df_emp.columns:
            df_emp.rename(columns={"اسم المقيم": "اسم المقيم "}, inplace=True)

        return df_emp, df_kpi, df_data

    except FileNotFoundError:
        st.error("⚠️ قاعدة البيانات غير موجودة.")
        return None, None, None
    except Exception as e:
        st.error(f"❌ خطأ في تحميل قاعدة البيانات: {e}")
        return None, None, None


# ═══════════════════════════════════════════════════════════════════
# ✅ جديد: معلومات الموظفين الإضافية (تاريخ التعيين، الراتب، إلخ)
# ═══════════════════════════════════════════════════════════════════
def load_employees_extra():
    """تحميل المعلومات الإضافية للموظفين"""
    if os.path.exists(EMP_EXTRA_DB):
        with open(EMP_EXTRA_DB, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}


def save_employees_extra(data):
    """حفظ المعلومات الإضافية للموظفين"""
    with open(EMP_EXTRA_DB, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def get_employee_extra(emp_name):
    """جلب المعلومات الإضافية لموظف معين"""
    extras = load_employees_extra()
    return extras.get(emp_name, {
        "hire_date": "",
        "current_salary": "",
        "phone": "",
        "email": "",
        "nationality": "",
        "notes": ""
    })


def update_employee_extra(emp_name, data):
    """تحديث المعلومات الإضافية لموظف"""
    extras = load_employees_extra()
    if emp_name not in extras:
        extras[emp_name] = {}
    extras[emp_name].update(data)
    save_employees_extra(extras)
    return True


# ═══════════════════════════════════════════════════════════════════
# ✅ جديد: الإجراءات التأديبية
# ═══════════════════════════════════════════════════════════════════
def load_disciplinary_actions():
    """تحميل جميع الإجراءات التأديبية"""
    if os.path.exists(DISCIPLINARY_DB):
        with open(DISCIPLINARY_DB, "r", encoding="utf-8") as f:
            return json.load(f)
    return []


def save_disciplinary_actions(actions):
    """حفظ الإجراءات التأديبية"""
    with open(DISCIPLINARY_DB, "w", encoding="utf-8") as f:
        json.dump(actions, f, ensure_ascii=False, indent=2)


def add_disciplinary_action(emp_name, action_date, action_type, description, created_by):
    """إضافة إجراء تأديبي جديد"""
    actions = load_disciplinary_actions()
    # حساب id جديد
    new_id = max([a.get("id", 0) for a in actions]) + 1 if actions else 1
    new_action = {
        "id": new_id,
        "employee_name": emp_name,
        "action_date": action_date,
        "action_month": datetime.strptime(action_date, "%Y-%m-%d").strftime("%Y-%m") if action_date else "",
        "action_type": action_type,
        "description": description,
        "created_by": created_by,
        "created_at": datetime.now().strftime("%Y-%m-%d %H:%M")
    }
    actions.append(new_action)
    save_disciplinary_actions(actions)
    return True, new_action["id"]


def get_employee_disciplinary(emp_name, year=None, month=None):
    """
    جلب الإجراءات التأديبية لموظف معين
    يمكن تصفيتها حسب السنة والشهر
    """
    actions = load_disciplinary_actions()
    emp_actions = [a for a in actions if a["employee_name"] == emp_name]
    
    if year:
        emp_actions = [a for a in emp_actions 
                       if a.get("action_date", "").startswith(str(year))]
    
    if month and year:
        month_str = f"{year}-{str(month).zfill(2)}"
        emp_actions = [a for a in emp_actions 
                       if a.get("action_month", "") == month_str]
    
    return emp_actions


def get_disciplinary_summary(emp_name):
    """
    ملخص الإجراءات التأديبية للموظف
    يُستخدم في التقارير
    """
    actions = get_employee_disciplinary(emp_name)
    if not actions:
        return "لا يوجد إجراءات تأديبية"
    
    summary_parts = []
    for a in actions:
        summary_parts.append(f"• {a['action_date']}: {a['action_type']} - {a['description']}")
    
    return "\n".join(summary_parts)


def delete_disciplinary_action(action_id):
    """حذف إجراء تأديبي"""
    actions = load_disciplinary_actions()
    actions = [a for a in actions if a.get("id") != action_id]
    save_disciplinary_actions(actions)
    return True


def update_disciplinary_action(action_id, updated_data):
    """تعديل إجراء تأديبي"""
    actions = load_disciplinary_actions()
    for a in actions:
        if a.get("id") == action_id:
            a.update(updated_data)
            break
    save_disciplinary_actions(actions)
    return True


# ═══════════════════════════════════════════════════════════════════
# ✅ جديد: جلب التقييم السابق للموظف
# ═══════════════════════════════════════════════════════════════════
def get_previous_evaluation(emp_name, current_year, current_month_en):
    """
    جلب التقييم السابق للموظف (الشهر الذي يسبق الشهر الحالي)
    """
    # تحميل البيانات
    result = load_data_from_db()
    if result is None:
        return None
    df_emp, df_kpi, df_data = result
    
    if df_data is None or df_data.empty:
        return None
    
    # ترتيب الأشهر
    months_order = ["January", "February", "March", "April", "May", "June",
                    "July", "August", "September", "October", "November", "December"]
    
    if current_month_en not in months_order:
        return None
    
    current_idx = months_order.index(current_month_en)
    
    # حساب الشهر والسنة السابقة
    if current_idx == 0:
        # يناير → ديسمبر من السنة الماضية
        prev_month_en = "December"
        prev_year = int(current_year) - 1
    else:
        prev_month_en = months_order[current_idx - 1]
        prev_year = int(current_year)
    
    # البحث عن التقييم السابق
    prev_data = df_data[
        (df_data["EmployeeName"] == emp_name) &
        (df_data["Month"] == prev_month_en) &
        (df_data["Year"] == prev_year)
    ]
    
    if prev_data.empty:
        return None
    
    # حساب النتيجة الإجمالية للتقييم السابق
    total_score = prev_data["KPI_%"].sum()
    
    # حساب عدد المؤشرات المقيمة
    kpi_count = len(prev_data)
    
    # حساب النسبة المئوية (إذا كان هناك أوزان)
    total_weight = prev_data["Weight"].sum() if "Weight" in prev_data.columns else 0
    percentage = round((total_score / total_weight) * 100, 1) if total_weight > 0 else 0
    
    return {
        "year": prev_year,
        "month": prev_month_en,
        "total_score": total_score,
        "percentage": percentage,
        "kpi_count": kpi_count,
        "employee_name": emp_name
    }


def get_all_previous_evaluations(emp_name, current_year, current_month_en):
    """
    جلب آخر 3 تقييمات سابقة للموظف (للمقارنة)
    """
    result = load_data_from_db()
    if result is None:
        return []
    df_emp, df_kpi, df_data = result
    
    if df_data is None or df_data.empty:
        return []
    
    months_order = ["January", "February", "March", "April", "May", "June",
                    "July", "August", "September", "October", "November", "December"]
    
    if current_month_en not in months_order:
        return []
    
    current_idx = months_order.index(current_month_en)
    
    previous_evaluations = []
    year = int(current_year)
    
    # نجيب آخر 3 أشهر سابقة
    for back in range(1, 4):
        prev_idx = current_idx - back
        prev_year = year
        
        if prev_idx < 0:
            prev_idx += 12
            prev_year -= 1
        
        if prev_idx >= 0 and prev_idx < 12:
            prev_month_en = months_order[prev_idx]
            
            prev_data = df_data[
                (df_data["EmployeeName"] == emp_name) &
                (df_data["Month"] == prev_month_en) &
                (df_data["Year"] == prev_year)
            ]
            
            if not prev_data.empty:
                total_score = prev_data["KPI_%"].sum()
                previous_evaluations.append({
                    "year": prev_year,
                    "month": prev_month_en,
                    "total_score": total_score
                })
    
    return previous_evaluations


# ═══════════════════════════════════════════════════════════════════
# حفظ وتحديث وحذف التقييمات (للقاعدة)
# ═══════════════════════════════════════════════════════════════════
def save_evaluation_to_db(emp_name, month_ar, year, manager, dept,
                           kpi_rows, notes="", training=""):
    try:
        with open(DATA_DB, "r", encoding="utf-8") as f:
            records = json.load(f)
        month_en  = MONTH_MAP.get(month_ar, month_ar)
        eval_date = date.today().strftime("%d/%m/%Y")
        for item in kpi_rows:
            if len(item) == 4:
                kpi_name, weight, grade, rating_lbl = item
            else:
                kpi_name, weight, grade = item[:3]
                rating_lbl = ""
            records.append({
                "EmployeeName": emp_name,
                "Month":        month_en,
                "KPI_Name":     kpi_name,
                "Weight":       float(weight),
                "KPI_%":        round(float(grade), 2),
                "RatingLabel":  rating_lbl,
                "Evaluator":    manager,
                "Nots":         notes,
                "Year":         int(year),
                "EvalDate":     eval_date,
                "Training":     training,
            })
        with open(DATA_DB, "w", encoding="utf-8") as f:
            json.dump(records, f, ensure_ascii=False, indent=2)
        load_data_from_db.clear()
        return True, None
    except Exception as e:
        return False, str(e)


def update_evaluation_in_db(act_emp, act_month_en, act_year, new_grades: dict):
    with open(DATA_DB, "r", encoding="utf-8") as f:
        records = json.load(f)
    updated = 0
    for r in records:
        if (r["EmployeeName"] == act_emp and
                r["Month"] == act_month_en and
                int(r["Year"]) == int(act_year) and
                r["KPI_Name"] in new_grades):
            r["KPI_%"] = float(new_grades[r["KPI_Name"]])
            updated += 1
    with open(DATA_DB, "w", encoding="utf-8") as f:
        json.dump(records, f, ensure_ascii=False, indent=2)
    load_data_from_db.clear()
    return updated


def delete_evaluation_from_db(act_emp, act_month_en, act_year, kpi_to_del=None):
    with open(DATA_DB, "r", encoding="utf-8") as f:
        records = json.load(f)
    before = len(records)
    if kpi_to_del:
        records = [r for r in records if not (
            r["EmployeeName"] == act_emp and r["Month"] == act_month_en and
            int(r["Year"]) == int(act_year) and r["KPI_Name"] == kpi_to_del
        )]
    else:
        records = [r for r in records if not (
            r["EmployeeName"] == act_emp and r["Month"] == act_month_en and
            int(r["Year"]) == int(act_year)
        )]
    with open(DATA_DB, "w", encoding="utf-8") as f:
        json.dump(records, f, ensure_ascii=False, indent=2)
    load_data_from_db.clear()
    return before - len(records)


def import_from_excel(file_path: str) -> dict:
    result = {"success": False, "message": "", "counts": {}}
    try:
        df_emp  = pd.read_excel(file_path, sheet_name="EMPLOYEES", engine="openpyxl")
        df_kpi  = pd.read_excel(file_path, sheet_name="KPIs",      engine="openpyxl")
        df_data = pd.read_excel(file_path, sheet_name="DATA",       engine="openpyxl")

        for df in [df_emp, df_kpi, df_data]:
            df.columns = [str(c).strip() for c in df.columns]
            for col in df.select_dtypes(include=['object']).columns:
                df[col] = df[col].astype(str).str.strip()

        df_data = df_data.dropna(subset=["EmployeeName","Month","KPI_Name"], how="all")
        df_data = df_data[df_data["EmployeeName"].astype(str).str.strip().ne("nan")]

        if "Year" not in df_data.columns:
            df_data["Year"] = 2025
        else:
            df_data["Year"] = pd.to_numeric(df_data["Year"], errors="coerce").fillna(2025).astype(int)
        for col in ["EvalDate","Training","Nots"]:
            if col not in df_data.columns: df_data[col] = ""
            else: df_data[col] = df_data[col].fillna("").astype(str).replace("nan","")

        df_kpi["Weight"] = pd.to_numeric(df_kpi["Weight"], errors="coerce").fillna(0)
        df_data["Weight"] = pd.to_numeric(df_data["Weight"], errors="coerce").fillna(0)
        df_data["KPI_%"]  = pd.to_numeric(df_data["KPI_%"],  errors="coerce").fillna(0)

        df_emp = df_emp.dropna(subset=["EmployeeName"])
        df_emp = df_emp[df_emp["EmployeeName"].astype(str).str.strip().ne("nan")]

        # حفظ EMPLOYEES
        emp_records = []
        for _, r in df_emp.iterrows():
            emp_records.append({
                "EmployeeName": str(r.get("EmployeeName","")).strip(),
                "JobTitle":     str(r.get("JobTitle","")).strip(),
                "القسم":        str(r.get("القسم","")).strip(),
                "اسم المقيم ":  str(r.get("اسم المقيم ", r.get("اسم المقيم",""))).strip(),
            })
        with open(EMP_DB, "w", encoding="utf-8") as f:
            json.dump(emp_records, f, ensure_ascii=False, indent=2)

        # حفظ KPIs مع تعديل الأوزان 80% + 20% صفات شخصية
        jobs_total = df_kpi.groupby("JobTitle")["Weight"].sum().to_dict()
        kpi_records = []
        for _, r in df_kpi.iterrows():
            job   = str(r["JobTitle"]).strip()
            total = jobs_total.get(job, 100)
            w     = float(r["Weight"])
            if total == 100:
                w = round(w * 0.8, 1)
            kpi_records.append({"JobTitle": job, "KPI_Name": str(r["KPI_Name"]).strip(), "Weight": w})
        for job in df_kpi["JobTitle"].unique():
            for p in PERSONAL_KPIS:
                kpi_records.append({"JobTitle": str(job).strip(), "KPI_Name": p, "Weight": float(PERSONAL_WEIGHT)})
        with open(KPI_DB, "w", encoding="utf-8") as f:
            json.dump(kpi_records, f, ensure_ascii=False, indent=2)

        # حفظ DATA
        data_records = []
        for _, r in df_data.iterrows():
            data_records.append({
                "EmployeeName": str(r.get("EmployeeName","")).strip(),
                "Month":        str(r.get("Month","")).strip(),
                "KPI_Name":     str(r.get("KPI_Name","")).strip(),
                "Weight":       float(r.get("Weight",0)),
                "KPI_%":        float(r.get("KPI_%",0)),
                "Evaluator":    str(r.get("Evaluator","")).strip(),
                "Nots":         str(r.get("Nots","")).replace("nan","").strip(),
                "Year":         int(r.get("Year",2025)),
                "EvalDate":     str(r.get("EvalDate","")).replace("nan","").strip(),
                "Training":     str(r.get("Training","")).replace("nan","").strip(),
            })
        with open(DATA_DB, "w", encoding="utf-8") as f:
            json.dump(data_records, f, ensure_ascii=False, indent=2)

        meta = {
            "imported_at": datetime.now().strftime("%Y-%m-%d %H:%M"),
            "source_file": os.path.basename(file_path),
            "employees_count": len(emp_records),
            "kpis_count": len(kpi_records),
            "evaluations_count": len(data_records),
        }
        with open(META_FILE, "w", encoding="utf-8") as f:
            json.dump(meta, f, ensure_ascii=False, indent=2)

        load_data_from_db.clear()
        result.update({"success": True, "message": "✅ تم الاستيراد بنجاح",
                        "counts": {"employees": len(emp_records), "kpis": len(kpi_records),
                                   "evaluations": len(data_records)}})
    except Exception as e:
        result["message"] = f"❌ خطأ: {e}"
    return result


def export_db_to_excel():
    import io
    df_emp, df_kpi, df_data = load_data_from_db()
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        if df_emp  is not None: df_emp.to_excel(writer,  sheet_name="EMPLOYEES", index=False)
        if df_kpi  is not None: df_kpi.to_excel(writer,  sheet_name="KPIs",      index=False)
        if df_data is not None: df_data.to_excel(writer, sheet_name="DATA",       index=False)
    buf.seek(0)
    return buf.read()


def sync_from_excel_if_updated(file_path: str = "final Apprisal.xlsm"):
    """
    يتحقق إذا تم تعديل ملف Excel بعد آخر استيراد،
    وإذا نعم يُحدّث قاعدة البيانات تلقائياً.
    يُستدعى في بداية كل تشغيل للبرنامج.
    """
    import os
    if not os.path.exists(file_path):
        return False, "ملف Excel غير موجود"

    meta = get_db_meta()
    last_import = meta.get("imported_at", "")

    try:
        excel_mtime = os.path.getmtime(file_path)
        from datetime import datetime
        excel_dt = datetime.fromtimestamp(excel_mtime)

        if last_import:
            last_dt = datetime.strptime(last_import, "%Y-%m-%d %H:%M")
            # إذا ملف Excel أحدث من آخر استيراد بأكثر من دقيقة
            if (excel_dt - last_dt).total_seconds() > 60:
                result = import_from_excel(file_path)
                if result["success"]:
                    return True, f"✅ تم تحديث قاعدة البيانات تلقائياً من Excel ({excel_dt.strftime('%H:%M')})"
                return False, result["message"]
    except Exception as e:
        return False, str(e)

    return False, ""
