"""
db_settings_panel.py — لوحة إدارة قاعدة البيانات
"""
import streamlit as st
import os, json
from database_manager import (
    import_from_excel, db_exists, get_db_meta,
    load_data_from_db, export_db_to_excel, DATA_DB, DB_DIR
)
from constants import FILE_PATH

def render_db_panel():
    st.markdown("""
    <div style="background:linear-gradient(135deg,#1F3864,#2E75B6);
                border-radius:16px;padding:20px 28px;margin-bottom:20px;color:white;">
      <h2 style="margin:0;font-size:22px;">🗄️ قاعدة البيانات المحلية</h2>
      <p style="margin:6px 0 0;opacity:0.85;font-size:13px;">
        إدارة بيانات النظام بدون الاعتماد على ملف Excel
      </p>
    </div>""", unsafe_allow_html=True)

    if db_exists():
        meta = get_db_meta()
        c1,c2,c3,c4 = st.columns(4)
        c1.metric("👥 الموظفون",        meta.get("employees_count","—"))
        c2.metric("📊 مؤشرات الأداء",   meta.get("kpis_count","—"))
        c3.metric("📋 التقييمات",        meta.get("evaluations_count","—"))
        c4.metric("🕒 آخر تحديث",        meta.get("imported_at","—"))
        st.success("✅ قاعدة البيانات نشطة — النظام يعمل بدون Excel")
    else:
        st.warning("⚠️ قاعدة البيانات غير مهيأة — استخدم الاستيراد أدناه")

    tab1, tab2 = st.tabs(["📥 استيراد من Excel", "💾 تصدير / نسخ احتياطي"])

    with tab1:
        st.markdown("#### استيراد البيانات من ملف Excel")
        if os.path.exists(FILE_PATH):
            st.info(f"📁 الملف الموجود: `{FILE_PATH}`")
            if st.button("📥 استيراد من الملف الافتراضي", type="primary", use_container_width=True):
                with st.spinner("جاري الاستيراد..."):
                    res = import_from_excel(FILE_PATH)
                if res["success"]:
                    st.success(res["message"])
                    st.json(res["counts"])
                    st.rerun()
                else:
                    st.error(res["message"])
        else:
            st.warning(f"⚠️ الملف `{FILE_PATH}` غير موجود")

        st.markdown("---")
        uploaded = st.file_uploader("أو ارفع ملف Excel جديد", type=["xlsx","xlsm"])
        if uploaded:
            import tempfile
            with tempfile.NamedTemporaryFile(suffix=".xlsm", delete=False) as tmp:
                tmp.write(uploaded.read()); tmp_path = tmp.name
            if st.button("📥 استيراد من الملف المرفوع", type="primary", use_container_width=True):
                with st.spinner("جاري الاستيراد..."): res = import_from_excel(tmp_path)
                os.unlink(tmp_path)
                if res["success"]:
                    st.success(res["message"]); st.json(res["counts"]); st.rerun()
                else:
                    st.error(res["message"])

    with tab2:
        col1, col2 = st.columns(2)
        with col1:
            st.markdown("##### 📊 تصدير إلى Excel")
            if st.button("إنشاء ملف Excel", use_container_width=True):
                xlsx = export_db_to_excel()
                st.download_button("⬇️ تحميل Excel", data=xlsx,
                    file_name=f"database_export.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True)
        with col2:
            st.markdown("##### 📦 نسخة احتياطية JSON")
            if st.button("إنشاء نسخة احتياطية", use_container_width=True):
                import zipfile, io
                buf = io.BytesIO()
                with zipfile.ZipFile(buf,"w") as zf:
                    for f in ["db/employees.json","db/kpis.json","db/evaluations.json"]:
                        if os.path.exists(f): zf.write(f)
                buf.seek(0)
                st.download_button("⬇️ تحميل ZIP", data=buf,
                    file_name="db_backup.zip", mime="application/zip",
                    use_container_width=True)
