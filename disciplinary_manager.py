"""
disciplinary_manager.py — إدارة الإجراءات التأديبية الكاملة
"""
import json
import os
import pandas as pd
import streamlit as st
from datetime import datetime

DB_DIR = "db"
DISCIPLINARY_DB = os.path.join(DB_DIR, "disciplinary.json")

# ═══════════════════════════════════════════════════════════════════
# التأكد من وجود المجلد والملف
# ═══════════════════════════════════════════════════════════════════
os.makedirs(DB_DIR, exist_ok=True)

def _ensure_file():
    """التأكد من وجود ملف الإجراءات"""
    if not os.path.exists(DISCIPLINARY_DB):
        with open(DISCIPLINARY_DB, "w", encoding="utf-8") as f:
            json.dump([], f, ensure_ascii=False, indent=2)

# ═══════════════════════════════════════════════════════════════════
# الدوال الأساسية (التحميل والحفظ)
# ═══════════════════════════════════════════════════════════════════
def load_actions():
    """تحميل جميع الإجراءات التأديبية"""
    _ensure_file()
    with open(DISCIPLINARY_DB, "r", encoding="utf-8") as f:
        try:
            return json.load(f)
        except:
            return []

def save_actions(actions):
    """حفظ الإجراءات التأديبية"""
    with open(DISCIPLINARY_DB, "w", encoding="utf-8") as f:
        json.dump(actions, f, ensure_ascii=False, indent=2)

# ═══════════════════════════════════════════════════════════════════
# دوال البحث والتصفية
# ═══════════════════════════════════════════════════════════════════
def get_actions_by_employee(emp_name, year=None, month=None):
    """جلب إجراءات موظف معين"""
    actions = load_actions()
    result = [a for a in actions if a.get("employee_name", "") == emp_name]
    
    if year:
        result = [a for a in result if a.get("year") == year]
    
    if month and year:
        result = [a for a in result if a.get("month") == month]
    
    return sorted(result, key=lambda x: x.get("action_date", ""), reverse=True)

def get_actions_summary(emp_name, year=None):
    """ملخص الإجراءات لنص التقرير"""
    actions = get_actions_by_employee(emp_name, year)
    if not actions:
        return "لا توجد إجراءات تأديبية"
    
    lines = []
    for a in actions:
        lines.append(f"• {a.get('action_date', '')}: {a.get('warning_type', '')} - {a.get('reason', '')}")
    return "\n".join(lines)

# ═══════════════════════════════════════════════════════════════════
# دوال الإضافة والتعديل والحذف
# ═══════════════════════════════════════════════════════════════════
def add_action(emp_name, emp_id, action_date, warning_type, reason, deduction_days):
    """إضافة إجراء تأديبي جديد"""
    actions = load_actions()
    
    # حساب ID جديد
    new_id = max([a.get("id", 0) for a in actions]) + 1 if actions else 1
    
    # استخراج السنة والشهر من التاريخ
    try:
        date_obj = datetime.strptime(action_date, "%Y-%m-%d")
        year = date_obj.year
        month = date_obj.month
    except:
        year = datetime.now().year
        month = datetime.now().month
    
    new_action = {
        "id": new_id,
        "employee_name": emp_name,
        "employee_id": emp_id,
        "action_date": action_date,
        "year": year,
        "month": month,
        "warning_type": warning_type,
        "reason": reason,
        "deduction_days": deduction_days,
        "created_at": datetime.now().strftime("%Y-%m-%d %H:%M")
    }
    
    actions.append(new_action)
    save_actions(actions)
    return new_id

def update_action(action_id, updated_data):
    """تعديل إجراء تأديبي"""
    actions = load_actions()
    for i, a in enumerate(actions):
        if a.get("id") == action_id:
            actions[i].update(updated_data)
            save_actions(actions)
            return True
    return False

def delete_action(action_id):
    """حذف إجراء تأديبي"""
    actions = load_actions()
    actions = [a for a in actions if a.get("id") != action_id]
    save_actions(actions)
    return True

# ═══════════════════════════════════════════════════════════════════
# استيراد من Excel
# ═══════════════════════════════════════════════════════════════════
def import_from_excel(uploaded_file):
    """استيراد الإجراءات من ملف Excel"""
    try:
        df = pd.read_excel(uploaded_file, engine="openpyxl")
        df.columns = [str(c).strip() for c in df.columns]
        
        # توحيد أسماء الأعمدة
        expected_columns = {
            "عدد ايام الخصم": "deduction_days",
            "نوع الإنذار": "warning_type",
            "تاريخ الإنذار": "action_date",
            "سبب الإنذار": "reason",
            "اسم الموظف": "employee_name",
            "رقم الموظف": "employee_id"
        }
        
        for old, new in expected_columns.items():
            if old in df.columns:
                df.rename(columns={old: new}, inplace=True)
        
        # تحويل التاريخ
        df["action_date"] = pd.to_datetime(df["action_date"]).dt.strftime("%Y-%m-%d")
        
        # تنظيف الأسماء
        df["employee_name"] = df["employee_name"].astype(str).str.strip()
        df["warning_type"] = df["warning_type"].astype(str).str.strip()
        df["reason"] = df["reason"].astype(str).str.strip()
        df["deduction_days"] = pd.to_numeric(df["deduction_days"], errors="coerce").fillna(0).astype(int)
        df["employee_id"] = df["employee_id"].astype(str).str.strip()
        
        # حفظ البيانات
        imported_count = 0
        for _, row in df.iterrows():
            add_action(
                emp_name=row["employee_name"],
                emp_id=row.get("employee_id", ""),
                action_date=row["action_date"],
                warning_type=row["warning_type"],
                reason=row["reason"],
                deduction_days=int(row["deduction_days"])
            )
            imported_count += 1
        
        return True, f"✅ تم استيراد {imported_count} إجراء تأديبي"
    
    except Exception as e:
        return False, f"❌ خطأ: {str(e)}"

# ═══════════════════════════════════════════════════════════════════
# تصدير إلى Excel
# ═══════════════════════════════════════════════════════════════════
def export_to_excel():
    """تصدير جميع الإجراءات إلى Excel"""
    import io
    actions = load_actions()
    df = pd.DataFrame(actions)
    
    # اختيار الأعمدة المطلوبة
    columns = ["employee_name", "employee_id", "action_date", "warning_type", "reason", "deduction_days"]
    available = [c for c in columns if c in df.columns]
    
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df[available].to_excel(writer, sheet_name="الإجراءات التأديبية", index=False)
    
    buf.seek(0)
    return buf

# ═══════════════════════════════════════════════════════════════════
# دوال الحصول على الأسماء للموجهات
# ═══════════════════════════════════════════════════════════════════
def get_all_employee_names(df_emp):
    """جلب جميع أسماء الموظفين من قاعدة البيانات الرئيسية"""
    if df_emp is not None and not df_emp.empty:
        return sorted(df_emp["EmployeeName"].dropna().astype(str).str.strip().tolist())
    return []
