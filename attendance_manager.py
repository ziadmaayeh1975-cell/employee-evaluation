"""
attendance_manager.py — إدارة بيانات الالتزام بالدوام (التأخير)
"""
import json
import os
import zipfile
import pandas as pd
import streamlit as st
from datetime import datetime

DB_DIR = "db"
ATTENDANCE_DB = os.path.join(DB_DIR, "attendance.json")

os.makedirs(DB_DIR, exist_ok=True)

def _ensure_file():
    if not os.path.exists(ATTENDANCE_DB):
        with open(ATTENDANCE_DB, "w", encoding="utf-8") as f:
            json.dump([], f, ensure_ascii=False, indent=2)

def load_attendance():
    _ensure_file()
    with open(ATTENDANCE_DB, "r", encoding="utf-8") as f:
        try:
            return json.load(f)
        except:
            return []

def save_attendance(data):
    _ensure_file()
    with open(ATTENDANCE_DB, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

def get_attendance_by_employee(emp_name, emp_id, year=None, month=None):
    records = load_attendance()
    result = [r for r in records if r.get("employee_name") == emp_name and r.get("employee_id") == emp_id]
    if year is not None:
        result = [r for r in result if r.get("year") == year]
    if month is not None and year is not None:
        result = [r for r in result if r.get("year") == year and r.get("month") == month]
    return sorted(result, key=lambda x: x.get("year", 0), reverse=True)

def get_attendance_summary(emp_name, emp_id, year=None, month=None):
    records = get_attendance_by_employee(emp_name, emp_id, year, month)
    total_late_hours = sum(r.get("total_late_hours", 0) for r in records)
    total_late_count = sum(r.get("late_count", 0) for r in records)
    return {"count": total_late_count, "hours": total_late_hours, "records": records}

def get_statistics():
    records = load_attendance()
    total_records = len(records)
    unique_emps = len(set(r.get("employee_name", "") for r in records))
    return {"total": total_records, "unique_employees": unique_emps}

def _time_str_to_hours(time_str):
    if time_str is None or time_str == "" or pd.isna(time_str):
        return 0.0
    time_str = str(time_str).strip()
    try:
        return float(time_str)
    except:
        pass
    parts = time_str.split(":")
    if len(parts) >= 2:
        try:
            hours = int(parts[0])
            minutes = int(parts[1])
            return hours + minutes / 60.0
        except:
            pass
    return 0.0

def import_from_excel(uploaded_file, clear_old=False):
    try:
        if uploaded_file.name.endswith('.zip'):
            return False, "⚠️ الملف مضغوط (ZIP). يرجى رفع ملف Excel فقط (.xlsx أو .xls)"

        try:
            df = pd.read_excel(uploaded_file, engine='openpyxl')
        except:
            df = pd.read_excel(uploaded_file, engine='xlrd')

        df.columns = [str(c).strip() for c in df.columns]
        
        # تعيين أسماء الأعمدة بناءً على ملفك
        column_mapping = {}
        for col in df.columns:
            col_lower = col.lower()
            if "رقم الموظف" in col_lower:
                column_mapping[col] = "employee_id"
            elif "اسم الموظف" in col_lower:
                column_mapping[col] = "employee_name"
            elif "تاريخ" in col_lower:
                column_mapping[col] = "date"
            elif "ساعات التاخير" in col_lower or "ساعات التأخير" in col_lower:
                column_mapping[col] = "late_hours"
        
        df = df.rename(columns=column_mapping)
        
        required = ["employee_name", "employee_id", "date", "late_hours"]
        missing = [r for r in required if r not in df.columns]
        if missing:
            return False, f"⚠️ الأعمدة المطلوبة غير موجودة: {missing}"
        
        df["employee_name"] = df["employee_name"].astype(str).str.strip()
        df["employee_id"] = df["employee_id"].astype(str).str.strip()
        df["date"] = pd.to_datetime(df["date"], errors="coerce")
        df["year"] = df["date"].dt.year
        df["month"] = df["date"].dt.month
        df["late_hours"] = df["late_hours"].apply(_time_str_to_hours)
        
        # إزالة الصفوف الفارغة
        df = df.dropna(subset=["employee_name", "employee_id", "year", "month", "late_hours"])
        
        if df.empty:
            return False, "⚠️ لا توجد بيانات صالحة للاستيراد"
        
        # تجميع البيانات لكل موظف في كل شهر
        grouped = df.groupby(["employee_name", "employee_id", "year", "month"]).agg(
            late_count=("late_hours", "count"),
            total_late_hours=("late_hours", "sum")
        ).reset_index()
        
        # تاريخ تمثيلي لأول يوم في الشهر
        grouped["date"] = pd.to_datetime(grouped["year"].astype(str) + "-" + grouped["month"].astype(str) + "-01").dt.strftime("%Y-%m-%d")
        
        if clear_old:
            save_attendance([])
            existing_records = []
        else:
            existing_records = load_attendance()
        
        existing_keys = {(r.get("employee_name", ""), r.get("employee_id", ""), r.get("date", "")) for r in existing_records}
        
        imported_count = 0
        skipped_count = 0
        
        for _, row in grouped.iterrows():
            key = (row["employee_name"], row["employee_id"], row["date"])
            if key not in existing_keys:
                new_id = max([r.get("id", 0) for r in existing_records]) + 1 if existing_records else 1
                new_record = {
                    "id": new_id,
                    "employee_name": row["employee_name"],
                    "employee_id": row["employee_id"],
                    "date": row["date"],
                    "year": int(row["year"]),
                    "month": int(row["month"]),
                    "late_count": int(row["late_count"]),
                    "total_late_hours": round(float(row["total_late_hours"]), 2),
                    "created_at": datetime.now().strftime("%Y-%m-%d %H:%M")
                }
                existing_records.append(new_record)
                imported_count += 1
            else:
                skipped_count += 1
        
        save_attendance(existing_records)
        
        if imported_count == 0 and skipped_count > 0:
            return True, f"✅ لا توجد سجلات جديدة (تخطي {skipped_count})"
        elif imported_count > 0 and skipped_count > 0:
            return True, f"✅ تم استيراد {imported_count} سجل جديد (تخطي {skipped_count})"
        else:
            return True, f"✅ تم استيراد {imported_count} سجل"
    
    except zipfile.BadZipFile:
        return False, "⚠️ ملف Excel تالف أو غير صالح"
    except Exception as e:
        return False, f"❌ خطأ: {str(e)}"

def add_attendance_manual(emp_name, emp_id, year, month, late_count, total_late_hours):
    records = load_attendance()
    date_str = f"{year}-{month:02d}-01"
    
    for existing in records:
        if existing.get("employee_name") == emp_name and existing.get("employee_id") == emp_id and existing.get("date") == date_str:
            return False
    
    new_id = max([r.get("id", 0) for r in records]) + 1 if records else 1
    new_record = {
        "id": new_id,
        "employee_name": emp_name,
        "employee_id": emp_id,
        "date": date_str,
        "year": int(year),
        "month": int(month),
        "late_count": int(late_count),
        "total_late_hours": float(total_late_hours),
        "created_at": datetime.now().strftime("%Y-%m-%d %H:%M")
    }
    records.append(new_record)
    save_attendance(records)
    return True

def delete_attendance_record(record_id):
    records = load_attendance()
    records = [r for r in records if r.get("id") != record_id]
    save_attendance(records)
    return True

def clear_all_attendance():
    save_attendance([])
    return True

def export_to_excel(year=None, month=None):
    import io
    records = load_attendance()
    if year and month:
        records = [r for r in records if r.get("year") == year and r.get("month") == month]
    elif year:
        records = [r for r in records if r.get("year") == year]
    df = pd.DataFrame(records)
    columns = ["employee_name", "employee_id", "year", "month", "late_count", "total_late_hours"]
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        if not df.empty:
            available_cols = [c for c in columns if c in df.columns]
            df[available_cols].to_excel(writer, sheet_name="الالتزام بالدوام", index=False)
        else:
            pd.DataFrame(columns=columns).to_excel(writer, sheet_name="الالتزام بالدوام", index=False)
    buf.seek(0)
    return buf

def get_employee_attendance_summary(emp_name, emp_id, year, month):
    result = get_attendance_summary(emp_name, emp_id, year, month)
    return {"count": result["count"], "hours": result["hours"]}

def get_unique_years():
    records = load_attendance()
    years = sorted(set(r.get("year") for r in records if r.get("year")))
    return years
