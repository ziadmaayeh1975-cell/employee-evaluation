"""
أداة نقل البيانات من Excel إلى قاعدة البيانات JSON
يتم تشغيل هذا الملف مرة واحدة فقط لنقل البيانات ثم يتم حذفه أو تجاهله
"""

import pandas as pd
import json
import os
from datetime import datetime

def migrate_all_data():
    """نقل جميع البيانات من Excel إلى ملفات JSON"""
    
    excel_file = "final Appraisal.xlsm"
    
    if not os.path.exists(excel_file):
        print(f"❌ ملف Excel غير موجود: {excel_file}")
        return False
    
    print("🚀 بدء عملية نقل البيانات من Excel إلى قاعدة البيانات...")
    
    try:
        # 1️⃣ قراءة بيانات الموظفين
        print("\n📋 جاري قراءة بيانات الموظفين...")
        df_employees = pd.read_excel(excel_file, sheet_name="Employees")
        
        # تحويل بيانات الموظفين إلى التنسيق المطلوب
        employees_data = []
        for _, row in df_employees.iterrows():
            employee = {
                "id": str(row.get("ID", row.get("الرقم", ""))),
                "name": str(row.get("Name", row.get("الاسم", ""))),
                "job_title": str(row.get("Job Title", row.get("المسمى الوظيفي", ""))),
                "department": str(row.get("Department", row.get("القسم", ""))),
                "email": str(row.get("Email", row.get("البريد", ""))),
                "hire_date": str(row.get("Hire Date", row.get("تاريخ التعيين", ""))),
                "status": "active"
            }
            employees_data.append(employee)
        
        # حفظ الموظفين
        with open("emp_profiles.json", "w", encoding="utf-8") as f:
            json.dump({"employees": employees_data}, f, ensure_ascii=False, indent=4)
        print(f"✅ تم حفظ {len(employees_data)} موظف")
        
        # 2️⃣ قراءة مؤشرات الأداء (KPIs)
        print("\n📊 جاري قراءة مؤشرات الأداء...")
        df_kpis = pd.read_excel(excel_file, sheet_name="KPIs")
        
        # تحويل المؤشرات
        kpis_data = []
        for _, row in df_kpis.iterrows():
            kpi = {
                "id": f"KPI_{len(kpis_data)+1:03d}",
                "code": str(row.get("Code", row.get("الكود", f"KPI{len(kpis_data)+1}"))),
                "name": str(row.get("Name", row.get("اسم المؤشر", ""))),
                "description": str(row.get("Description", row.get("الوصف", ""))),
                "category": str(row.get("Category", row.get("الفئة", "أداء وظيفي"))),
                "measurement_unit": str(row.get("Unit", row.get("وحدة القياس", "نسبة مئوية"))),
                "target_type": str(row.get("Target Type", row.get("نوع الهدف", "max"))),
                "default_weight": float(row.get("Weight", row.get("الوزن الافتراضي", 0))),
                "is_active": True,
                "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            }
            kpis_data.append(kpi)
        
        # حفظ المؤشرات
        with open("db/kpis.json", "w", encoding="utf-8") as f:
            json.dump({"kpis": kpis_data}, f, ensure_ascii=False, indent=4)
        print(f"✅ تم حفظ {len(kpis_data)} مؤشر أداء")
        
        # 3️⃣ ربط المؤشرات بالوظائف
        print("\n🔗 جاري ربط المؤشرات بالوظائف...")
        df_job_kpis = pd.read_excel(excel_file, sheet_name="Job_KPIs")
        
        job_kpis_mapping = {}
        for _, row in df_job_kpis.iterrows():
            job_title = str(row.get("Job Title", row.get("المسمى الوظيفي", "")))
            kpi_code = str(row.get("KPI Code", row.get("كود المؤشر", "")))
            weight = float(row.get("Weight", row.get("الوزن", 0)))
            
            if job_title not in job_kpis_mapping:
                job_kpis_mapping[job_title] = []
            
            job_kpis_mapping[job_title].append({
                "kpi_id": kpi_code,
                "weight": weight
            })
        
        # حفظ ربط المؤشرات
        os.makedirs("db", exist_ok=True)
        with open("db/job_kpis_mapping.json", "w", encoding="utf-8") as f:
            json.dump(job_kpis_mapping, f, ensure_ascii=False, indent=4)
        print(f"✅ تم حفظ ربط المؤشرات لـ {len(job_kpis_mapping)} وظيفة")
        
        # 4️⃣ إنشاء ملف الإعدادات
        settings = {
            "data_source": "json",
            "last_migration": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "excel_file": excel_file,
            "total_employees": len(employees_data),
            "total_kpis": len(kpis_data),
            "kpi_categories": ["أداء وظيفي", "صفات شخصية"],
            "total_weight": 100
        }
        
        with open("db/migration_settings.json", "w", encoding="utf-8") as f:
            json.dump(settings, f, ensure_ascii=False, indent=4)
        
        print("\n✨ تم نقل جميع البيانات بنجاح!")
        print(f"📊 إجمالي الموظفين: {len(employees_data)}")
        print(f"📊 إجمالي المؤشرات: {len(kpis_data)}")
        print(f"📊 الوظائف المرتبطة: {len(job_kpis_mapping)}")
        
        return True
        
    except Exception as e:
        print(f"❌ خطأ أثناء نقل البيانات: {e}")
        import traceback
        traceback.print_exc()
        return False

def verify_migration():
    """التحقق من صحة البيانات المنقولة"""
    print("\n🔍 التحقق من البيانات المنقولة...")
    
    # التأكد من وجود الملفات
    files_to_check = [
        "emp_profiles.json",
        "db/kpis.json",
        "db/job_kpis_mapping.json"
    ]
    
    all_exist = True
    for file in files_to_check:
        if os.path.exists(file):
            with open(file, "r", encoding="utf-8") as f:
                data = json.load(f)
            print(f"✅ {file}: موجود ويحتوي على بيانات")
        else:
            print(f"❌ {file}: غير موجود")
            all_exist = False
    
    # التحقق من مجموع الأوزان
    if os.path.exists("db/job_kpis_mapping.json"):
        with open("db/job_kpis_mapping.json", "r", encoding="utf-8") as f:
            job_kpis = json.load(f)
        
        for job, kpis in job_kpis.items():
            total_weight = sum(kpi["weight"] for kpi in kpis)
            if abs(total_weight - 100) > 0.01:
                print(f"⚠️ تنبيه: مجموع أوزان المؤشرات لوظيفة {job} = {total_weight}%")
    
    return all_exist

if __name__ == "__main__":
    print("=" * 50)
    print("🔄 أداة نقل البيانات من Excel إلى قاعدة البيانات")
    print("=" * 50)
    
    # تنفيذ النقل
    success = migrate_all_data()
    
    if success:
        # التحقق من النقل
        verify_migration()
        print("\n✅ العملية اكتملت. يمكنك الآن تشغيل التطبيق.")
        print("💡 ملاحظة: يمكن حذف هذا الملف أو إضافة 'migrate_from_excel.py' إلى .gitignore")
    else:
        print("\n❌ فشلت عملية النقل. راجع الأخطاء أعلاه.")
