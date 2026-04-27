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
    _ensure_file()
    with open(DISCIPLINARY_DB, "w", encoding="utf-8") as f:
        json.dump(actions, f, ensure_ascii=False, indent=2)

# ═══════════════════════════════════════════════════════════════════
# دوال البحث والتصفية (المُصلحة)
# ═══════════════════════════════════════════════════════════════════
def get_actions_by_employee(emp_name, year=None, month=None):
    """
    جلب إجراءات موظف معين
    إذا تم تمرير month فقط مع year، يتم التصفية حسب الشهر أيضاً
    """
    actions = load_actions()
    result = [a for a in actions if a.get("employee_name", "") == emp_name]
    
    if year is not None:
        result = [a for a in result if a.get("year") == year]
    
    if month is not None and year is not None:
        # تأكد من أن الشهر من النوع int وقارنه
        month_int = int(month) if not isinstance(month, int) else month
        result = [a for a in result if a.get("month") == month_int]
    elif month is not None and year is None:
        # لو تم إرسال شهر بدون سنة، لا نفلتر (لتجنب النتائج الخاطئة)
        pass
    
    # ترتيب تنازلي حسب التاريخ
    return sorted(result, key=lambda x: x.get("action_date", ""), reverse=True)

def get_actions_by_month(year, month):
    """جلب جميع الإجراءات لشهر وسنة محددين"""
    actions = load_actions()
    return [a for a in actions if a.get("year") == year and a.get("month") == month]

def get_actions_summary(emp_name, year=None):
    """ملخص الإجراءات لنص التقرير"""
    actions = get_actions_by_employee(emp_name, year)
    if not actions:
        return "لا توجد إجراءات تأديبية"
    
    lines = []
    for a in actions:
        lines.append(f"• {a.get('action_date', '')}: {a.get('warning_type', '')} - {a.get('reason', '')}")
    return "\n".join(lines)

def get_statistics():
    """إحصائيات سريعة للإجراءات"""
    actions = load_actions()
    total = len(actions)
    unique_emps = len(set(a.get("employee_name", "") for a in actions))
    
    now = datetime.now()
    current_month_actions = [a for a in actions if a.get("year") == now.year and a.get("month") == now.month]
    
    return {
        "total": total,
        "unique_employees": unique_emps,
        "current_month": len(current_month_actions)
    }

# ═══════════════════════════════════════════════════════════════════
# دوال الإضافة والتعديل والحذف
# ═══════════════════════════════════════════════════════════════════
def add_action(emp_name, emp_id, action_date, warning_type, reason, deduction_days):
    """إضافة إجراء تأديبي جديد (مع منع التكرار)"""
    actions = load_actions()
    
    # منع التكرار
    for existing in actions:
        if (existing.get("employee_name") == emp_name and
            existing.get("action_date") == action_date and
            existing.get("warning_type") == warning_type):
            return None
    
    # حساب ID جديد
    new_id = max([a.get("id", 0) for a in actions]) + 1 if actions else 1
    
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

def clear_all_actions():
    """مسح جميع الإجراءات (استخدام بحذر)"""
    save_actions([])
    return True

# ═══════════════════════════════════════════════════════════════════
# استيراد من Excel (ذكي: لا يكرر، يضيف فقط الجديد)
# ═══════════════════════════════════════════════════════════════════
def import_from_excel(uploaded_file, clear_old=False):
    try:
        df = pd.read_excel(uploaded_file, engine="openpyxl")
        df.columns = [str(c).strip() for c in df.columns]
        
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
        
        if "employee_name" not in df.columns or "action_date" not in df.columns or "warning_type" not in df.columns:
            return False, "⚠️ الأعمدة المطلوبة غير موجودة في ملف Excel"
        
        df["action_date"] = pd.to_datetime(df["action_date"]).dt.strftime("%Y-%m-%d")
        df["employee_name"] = df["employee_name"].astype(str).str.strip()
        df["warning_type"] = df["warning_type"].astype(str).str.strip()
        df["reason"] = df["reason"].astype(str).str.strip()
        df["deduction_days"] = pd.to_numeric(df["deduction_days"], errors="coerce").fillna(0).astype(int)
        
        if "employee_id" in df.columns:
            df["employee_id"] = df["employee_id"].astype(str).str.strip()
        else:
            df["employee_id"] = ""
        
        if clear_old:
            save_actions([])
            existing_actions = []
        else:
            existing_actions = load_actions()
        
        existing_keys = {(a.get("employee_name", ""), a.get("action_date", ""), a.get("warning_type", "")) 
                         for a in existing_actions}
        
        imported_count = 0
        skipped_count = 0
        
        for _, row in df.iterrows():
            emp_name = row["employee_name"]
            action_date = row["action_date"]
            warning_type = row["warning_type"]
            key = (emp_name, action_date, warning_type)
            
            if key not in existing_keys:
                add_action(
                    emp_name=emp_name,
                    emp_id=row.get("employee_id", ""),
                    action_date=action_date,
                    warning_type=warning_type,
                    reason=row.get("reason", ""),
                    deduction_days=int(row.get("deduction_days", 0))
                )
                imported_count += 1
                existing_keys.add(key)
            else:
                skipped_count += 1
        
        if imported_count == 0 and skipped_count > 0:
            return True, f"✅ لا توجد إجراءات جديدة للاستيراد (تخطي {skipped_count} إجراء مكرر)"
        elif imported_count > 0 and skipped_count > 0:
            return True, f"✅ تم استيراد {imported_count} إجراء جديد (تخطي {skipped_count} مكرر)"
        else:
            return True, f"✅ تم استيراد {imported_count} إجراء تأديبي"
    
    except Exception as e:
        return False, f"❌ خطأ: {str(e)}"

# ═══════════════════════════════════════════════════════════════════
# تصدير إلى Excel
# ═══════════════════════════════════════════════════════════════════
def export_to_excel(year=None, month=None):
    import io
    actions = load_actions()
    
    if year and month:
        actions = [a for a in actions if a.get("year") == year and a.get("month") == month]
    elif year:
        actions = [a for a in actions if a.get("year") == year]
    
    df = pd.DataFrame(actions)
    columns = ["employee_name", "employee_id", "action_date", "warning_type", "reason", "deduction_days"]
    
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        if not df.empty:
            available_cols = [c for c in columns if c in df.columns]
            df[available_cols].to_excel(writer, sheet_name="الإجراءات التأديبية", index=False)
        else:
            pd.DataFrame(columns=columns).to_excel(writer, sheet_name="الإجراءات التأديبية", index=False)
    
    buf.seek(0)
    return buf

# ═══════════════════════════════════════════════════════════════════
# دوال الحصول على الأسماء للواجهات
# ═══════════════════════════════════════════════════════════════════
def get_all_employee_names(df_emp):
    if df_emp is not None and not df_emp.empty:
        return sorted(df_emp["EmployeeName"].dropna().astype(str).str.strip().tolist())
    return []

def get_unique_years():
    actions = load_actions()
    years = sorted(set(a.get("year") for a in actions if a.get("year")))
    return years

def get_months_with_actions(year):
    actions = load_actions()
    months = sorted(set(a.get("month") for a in actions if a.get("year") == year))
    return months
