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
    new_action = {
        "id": len(actions) + 1,
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
    جلب 
