"""
attendance_manager.py — إدارة بيانات الالتزام بالدوام (التأخير)
"""
import json
import os
import pandas as pd
import streamlit as st
from datetime import datetime

DB_DIR = "db"
ATTENDANCE_DB = os.path.join(DB_DIR, "attendance.json")

# ═══════════════════════════════════════════════════════════════════
# التأكد من وجود المجلد والملف
# ═══════════════════════════════════════════════════════════════════
os.makedirs(DB_DIR, exist_ok=True)

def _ensure_file():
    """التأكد من وجود ملف attendance"""
    if not os.path.exists(ATTENDANCE_DB):
        with open(ATTENDANCE_DB, "w", encoding="utf-8") as f:
            json.dump([], f, ensure_ascii=False, indent=2)

# ═══════════════════════════════════════════════════════════════════
# الدوال الأساسية (التحميل والحفظ)
# ═══════════════════════════════════════════════════════════════════
def load_attendance():
    """تحميل جميع بيانات الالتزام بالدوام"""
    _ensure_file()
    with open(ATTENDANCE_DB, "r", encoding="utf-8") as f:
        try:
            return json.load(f)
        except:
            return []

def save_attendance(data):
    """حفظ بيانات الالتزام بالدوام"""
    _ensure_file()
    with open(ATTENDANCE_DB, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

# ═══════════════════════════════════════════════════════════════════
# دوال البحث والتصفية
# ═══════════════════════════════════════════════════════════════════
def get_attendance_by_employee(emp_name, emp_id, year=None, month=None):
    """
    جلب بيانات الالتزام بالدوام لموظف معين
    يمكن تصفيتها حسب السنة والشهر
    """
    records = load_attendance()
    result = [r for r in records if r.get("employee_name") == emp_name and r.get("employee_id") == emp_id]
    
    if year is not None:
        result = [r for r in result if r.get("year") == year]
    
    if month is not None and year is not None:
        result = [r for r in result if r.get("year") == year and r.get("month") == month]
    
    return sorted(result, key=lambda x: x.get("date", ""), reverse=True)

def get_attendance_summary(emp_name, emp_id, year=None, month=None):
    """
    الحصول على ملخص الالتزام بالدوام للموظف:
    - عدد مرات التأخير
    - مجموع ساعات التأخير
    """
    records = get_attendance_by_employee(emp_name, emp_id, year, month)
    total_late_hours = sum(r.get("late_hours", 0) for r in records)
    total_late_count = len(records)
    
    return {
        "count": total_late_count,
        "hours": total_late_hours,
        "records": records
    }

def get_statistics():
    """إحصائيات سريعة لبيانات الالتزام بالدوام"""
    records = load_attendance()
    total = len(records)
    unique_emps = len(set(r.get("employee_name", "") for r in records))
    
    return {
        "total": total,
        "unique_employees": unique_emps
    }

# ═══════════════════════════════════════════════════════════════════
# دوال الإضافة والتعديل والحذف
# ═══════════════════════════════════════════════════════════════════
def add_attendance_record(emp_name, emp_id, date_str, late_hours, days_count, amount=None):
    """
    إضافة سجل تأخير جديد
    - emp_name: اسم الموظف
    - emp_id: الرقم الوظيفي
    - date_str: تاريخ التأخير (YYYY-MM-DD)
    - late_hours: عدد ساعات التأخير
    - days_count: عدد الأيام (عادة 1)
    - amount: القيمة (اختياري)
    """
    records = load_attendance()
    
    # منع التكرار (نفس الموظف ونفس التاريخ)
    for existing in records:
        if (existing.get("employee_name") == emp_name and
            existing.get("employee_id") == emp_id and
            existing.get("date") == date_str):
            return None  # موجود مسبقاً
    
    # حساب ID جديد
    new_id = max([r.get("id", 0) for r in records]) + 1 if records else 1
    
    # استخراج السنة والشهر من التاريخ
    try:
        date_obj = datetime.strptime(date_str, "%Y-%m-%d")
        year = date_obj.year
        month = date_obj.month
    except:
        year = datetime.now().year
        month = datetime.now().month
    
    new_record = {
        "id": new_id,
        "employee_name": emp_name,
        "employee_id": emp_id,
        "date": date_str,
        "year": year,
        "month": month,
        "late_hours": float(late_hours),
        "days_count": int(days_count),
        "amount": float(amount) if amount is not None else 0.0,
        "created_at": datetime.now().strftime("%Y-%m-%d %H:%M")
    }
    
    records.append(new_record)
    save_attendance(records)
    return new_id

def delete_attendance_record(record_id):
    """حذف سجل تأخير"""
    records = load_attendance()
    records = [r for r in records if r.get("id") != record_id]
    save_attendance(records)
    return True

def clear_all_attendance():
    """مسح جميع بيانات الالتزام بالدوام"""
    save_attendance([])
    return True

# ═══════════════════════════════════════════════════════════════════
# استيراد من Excel
# ═══════════════════════════════════════════════════════════════════
def import_from_excel(uploaded_file, clear_old=False):
    """
    استيراد بيانات الالتزام بالدوام من ملف Excel
    الأعمدة المتوقعة:
    - الرقم الوظيفي / رقم الموظف / employee_id
    - اسم الموظف / employee_name
    - التاريخ / date
    - ساعات التاخير / late_hours
    - عدد الايام / days_count
    - القيمة / amount (اختياري)
    """
    try:
        df = pd.read_excel(uploaded_file, engine="openpyxl")
        df.columns = [str(c).strip() for c in df.columns]
        
        # توحيد أسماء الأعمدة
        column_mapping = {}
        for col in df.columns:
            col_lower = col.lower()
            if "رقم الموظف" in col_lower or "employee_id" in col_lower or "الرقم الوظيفي" in col_lower:
                column_mapping[col] = "employee_id"
            elif "اسم الموظف" in col_lower or "employee_name" in col_lower:
                column_mapping[col] = "employee_name"
            elif "تاريخ" in col_lower or "date" in col_lower:
                column_mapping[col] = "date"
            elif "ساعات التاخير" in col_lower or "late_hours" in col_lower:
                column_mapping[col] = "late_hours"
            elif "عدد الايام" in col_lower or "days_count" in col_lower:
                column_mapping[col] = "days_count"
            elif "قيمة" in col_lower or "amount" in col_lower:
                column_mapping[col] = "amount"
        
        df = df.rename(columns=column_mapping)
        
        # التحقق من الأعمدة المطلوبة
        required = ["employee_name", "employee_id", "date", "late_hours", "days_count"]
        missing = [r for r in required if r not in df.columns]
        if missing:
            return False, f"⚠️ الأعمدة المطلوبة غير موجودة: {missing}"
        
        # تنظيف البيانات
        df["employee_name"] = df["employee_name"].astype(str).str.strip()
        df["employee_id"] = df["employee_id"].astype(str).str.strip()
        df["date"] = pd.to_datetime(df["date"]).dt.strftime("%Y-%m-%d")
        df["late_hours"] = pd.to_numeric(df["late_hours"], errors="coerce").fillna(0).astype(float)
        df["days_count"] = pd.to_numeric(df["days_count"], errors="coerce").fillna(1).astype(int)
        
        if "amount" in df.columns:
            df["amount"] = pd.to_numeric(df["amount"], errors="coerce").fillna(0).astype(float)
        else:
            df["amount"] = 0.0
        
        # مسح القديم إذا طلب
        if clear_old:
            save_attendance([])
            existing_records = []
        else:
            existing_records = load_attendance()
        
        # بناء مجموعة المفاتيح الموجودة لتجنب التكرار
        existing_keys = {(r.get("employee_name", ""), r.get("employee_id", ""), r.get("date", "")) 
                         for r in existing_records}
        
        imported_count = 0
        skipped_count = 0
        
        for _, row in df.iterrows():
            emp_name = row["employee_name"]
            emp_id = row["employee_id"]
            date_str = row["date"]
            key = (emp_name, emp_id, date_str)
            
            if key not in existing_keys:
                add_attendance_record(
                    emp_name=emp_name,
                    emp_id=emp_id,
                    date_str=date_str,
                    late_hours=row["late_hours"],
                    days_count=row["days_count"],
                    amount=row.get("amount", 0.0)
                )
                imported_count += 1
                existing_keys.add(key)
            else:
                skipped_count += 1
        
        if imported_count == 0 and skipped_count > 0:
            return True, f"✅ لا توجد سجلات جديدة للاستيراد (تخطي {skipped_count} سجل مكرر)"
        elif imported_count > 0 and skipped_count > 0:
            return True, f"✅ تم استيراد {imported_count} سجل جديد (تخطي {skipped_count} مكرر)"
        else:
            return True, f"✅ تم استيراد {imported_count} سجل تأخير"
    
    except Exception as e:
        return False, f"❌ خطأ: {str(e)}"

# ═══════════════════════════════════════════════════════════════════
# تصدير إلى Excel
# ═══════════════════════════════════════════════════════════════════
def export_to_excel(year=None, month=None):
    """تصدير بيانات الالتزام بالدوام إلى Excel"""
    import io
    records = load_attendance()
    
    if year and month:
        records = [r for r in records if r.get("year") == year and r.get("month") == month]
    elif year:
        records = [r for r in records if r.get("year") == year]
    
    df = pd.DataFrame(records)
    columns = ["employee_name", "employee_id", "date", "late_hours", "days_count", "amount"]
    
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        if not df.empty:
            available_cols = [c for c in columns if c in df.columns]
            df[available_cols].to_excel(writer, sheet_name="الالتزام بالدوام", index=False)
        else:
            pd.DataFrame(columns=columns).to_excel(writer, sheet_name="الالتزام بالدوام", index=False)
    
    buf.seek(0)
    return buf

# ═══════════════════════════════════════════════════════════════════
# دوال مساعدة للواجهات
# ═══════════════════════════════════════════════════════════════════
def get_employee_attendance_summary(emp_name, emp_id, year, month):
    """
    دالة مبسطة لجلب ملخص الالتزام بالدوام للموظف لشهر وسنة محددين
    تُستخدم في entry.py و employee_report.py
    """
    result = get_attendance_summary(emp_name, emp_id, year, month)
    return {
        "count": result["count"],
        "hours": result["hours"]
    }

def get_unique_years():
    """جلب السنوات الموجودة في البيانات"""
    records = load_attendance()
    years = sorted(set(r.get("year") for r in records if r.get("year")))
    return years
