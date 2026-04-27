@st.cache_data(ttl=30)
def load_data():
    if _DB_AVAILABLE and db_exists():
        return load_data_from_db()
    try:
        import openpyxl
        df_emp = pd.read_excel(FILE_PATH, sheet_name="EMPLOYEES")
        df_kpi = pd.read_excel(FILE_PATH, sheet_name="KPIs")
        df_data = pd.read_excel(FILE_PATH, sheet_name="DATA")

        for df in [df_emp, df_kpi, df_data]:
            df.columns = [str(c).strip() for c in df.columns]
            for col in df.select_dtypes("object").columns:
                df[col] = df[col].astype(str).str.strip()

        # تأكد من وجود الأعمدة المطلوبة في EMPLOYEES
        required_emp_cols = ["رقم الموظف", "EmployeeName", "JobTitle", "القسم", "اسم المقيم"]
        for col in required_emp_cols:
            if col not in df_emp.columns:
                st.error(f"⚠️ العمود '{col}' غير موجود في شيت EMPLOYEES. الرجاء إضافته.")
                return None, None, None

        # توحيد الأعمدة (اختياري)
        df_emp = df_emp[required_emp_cols]

        # باقي الكود كما هو (معالجة DATA و KPIs...)
        ...
        return df_emp, df_kpi, df_data
    except Exception as e:
        st.error(f"❌ تعذّر تحميل البيانات: {e}")
        return None, None, None
