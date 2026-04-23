import pandas as pd
import json
import os

# إنشاء مجلد db إذا لم يكن موجوداً
os.makedirs("db", exist_ok=True)

# قراءة ملف الاكسل
print("جاري قراءة ملف الاكسل...")
excel_file = "final Appraisal.xlsm"

# نقل الموظفين
df_emp = pd.read_excel(excel_file, sheet_name="Employees")
employees = []
for _, row in df_emp.iterrows():
    employees.append({
        "id": str(row.get("ID", row.get("الرقم", ""))),
        "name": str(row.get("Name", row.get("الاسم", ""))),
        "job_title": str(row.get("Job Title", row.get("المسمى الوظيفي", ""))),
        "department": str(row.get("Department", row.get("القسم", "")))
    })

with open("emp_profiles.json", "w", encoding="utf-8") as f:
    json.dump({"employees": employees}, f, ensure_ascii=False, indent=4)
print(f"تم نقل {len(employees)} موظف")

# نقل المؤشرات
df_kpi = pd.read_excel(excel_file, sheet_name="KPIs")
kpis = []
for i, row in df_kpi.iterrows():
    kpis.append({
        "id": f"KPI_{i+1:03d}",
        "name": str(row.get("Name", row.get("اسم المؤشر", ""))),
        "category": str(row.get("Category", row.get("الفئة", "أداء وظيفي"))),
        "default_weight": float(row.get("Weight", row.get("الوزن", 0)))
    })

os.makedirs("db", exist_ok=True)
with open("db/kpis.json", "w", encoding="utf-8") as f:
    json.dump({"kpis": kpis}, f, ensure_ascii=False, indent=4)
print(f"تم نقل {len(kpis)} مؤشر")

print("\n✅ تم نقل جميع البيانات بنجاح!")
print("يمكنك الآن حذف ملف final Appraisal.xlsm إذا أردت")
