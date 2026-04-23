"""
employees_kpis_panel.py
تبويبان: إدارة الموظفين | قائمة مؤشرات الأداء
"""
import streamlit as st
import pandas as pd
import json, os
from datetime import date
from constants import PERSONAL_KPIS, PERSONAL_WEIGHT

# ─── مسارات الملفات ──────────────────────────────────────
DB_DIR = "db"
EMP_DB = "emp_profiles.json"
KPI_DB = os.path.join(DB_DIR, "kpis.json")
DATA_DB = os.path.join(DB_DIR, "evaluations.json")


# ─── دوال الحفظ والتحميل ─────────────────────────────────
def _load_emps():
    """تحميل الموظفين من JSON"""
    if not os.path.exists(EMP_DB):
        return []
    try:
        with open(EMP_DB, "r", encoding="utf-8") as f:
            data = json.load(f)
            return data.get("employees", [])
    except:
        return []

def _save_emps(records):
    """حفظ الموظفين في JSON"""
    os.makedirs(os.path.dirname(EMP_DB) if os.path.dirname(EMP_DB) else ".", exist_ok=True)
    with open(EMP_DB, "w", encoding="utf-8") as f:
        json.dump({"employees": records}, f, ensure_ascii=False, indent=2)

def _load_kpis():
    """تحميل المؤشرات من JSON"""
    if not os.path.exists(KPI_DB):
        return []
    try:
        with open(KPI_DB, "r", encoding="utf-8") as f:
            data = json.load(f)
            # المؤشرات قد تكون محفوظة كـ {"kpis": [...]} أو كـ {job: [{name, weight}]}
            if isinstance(data, dict):
                if "kpis" in data:
                    kpis = data["kpis"]
                else:
                    # تنسيق {JobTitle: [{name, weight}]}
                    kpis = []
                    for job, items in data.items():
                        if isinstance(items, list):
                            for item in items:
                                kpis.append({
                                    "JobTitle": job,
                                    "KPI_Name": item.get("name", item.get("KPI_Name", "")),
                                    "Weight": item.get("weight", item.get("Weight", 0))
                                })
                return kpis
            return []
    except:
        return []

def _save_kpis(records):
    """حفظ المؤشرات في JSON بصيغة kpis"""
    os.makedirs(DB_DIR, exist_ok=True)
    with open(KPI_DB, "w", encoding="utf-8") as f:
        json.dump({"kpis": records}, f, ensure_ascii=False, indent=2)


# ══════════════════════════════════════════════════════════
# تبويب إدارة الموظفين
# ══════════════════════════════════════════════════════════
def render_employees_panel():
    st.markdown("#### 👨‍💼 إدارة الموظفين")

    emps = _load_emps()
    
    if not emps:
        st.warning("⚠️ لا يوجد موظفون. قم بنقل البيانات من Excel أولاً من أعلى صفحة الإعدادات.")
        return
    
    df = pd.DataFrame(emps)

    # ── تأكد من وجود الأعمدة المطلوبة ─────────────────
    for col in ["EmployeeName", "JobTitle", "Department", "Manager"]:
        if col not in df.columns:
            df[col] = ""
    
    # ترجمة أسماء الأعمدة للعرض
    rename_map = {}
    if "EmployeeName" in df.columns:
        rename_map["EmployeeName"] = "اسم الموظف"
    if "JobTitle" in df.columns:
        rename_map["JobTitle"] = "المسمى الوظيفي"
    if "Department" in df.columns:
        rename_map["Department"] = "القسم"
    if "Manager" in df.columns:
        rename_map["Manager"] = "المقيم"
    if "القسم" in df.columns:
        rename_map["القسم"] = "القسم"
    if "اسم المقيم " in df.columns:
        rename_map["اسم المقيم "] = "المقيم"
    if "اسم المقيم" in df.columns:
        rename_map["اسم المقيم"] = "المقيم"

    # ── بحث ───────────────────────────────────────────────
    search = st.text_input("🔍 بحث باسم الموظف أو القسم أو الوظيفة")
    df_view = df.copy()
    if search:
        mask = df_view.apply(lambda r: search.lower() in " ".join(str(v).lower() for v in r.values), axis=1)
        df_view = df_view[mask]

    st.markdown(f"**إجمالي الموظفين: {len(df)}** | نتائج البحث: {len(df_view)}")
    
    # عرض الجدول
    if not df_view.empty and rename_map:
        st.dataframe(df_view.rename(columns=rename_map), use_container_width=True, height=250)
    else:
        st.dataframe(df_view, use_container_width=True, height=250)

    tab_add, tab_edit, tab_del = st.tabs(["➕ إضافة موظف", "✏️ تعديل موظف", "🗑️ حذف موظف"])

    # ── إضافة ──────────────────────────────────────────────
    with tab_add:
        with st.form("add_emp_form"):
            c1, c2 = st.columns(2)
            with c1:
                n_name    = st.text_input("اسم الموظف *")
                n_id      = st.text_input("الرقم الوظيفي")
                n_hire    = st.date_input("تاريخ التعيين", value=date.today())
                n_status  = st.selectbox("الحالة", ["فعّال", "غير فعّال"])
            with c2:
                kpis      = _load_kpis()
                # استخراج الوظائف من المؤشرات
                all_jobs  = sorted(set(k.get("JobTitle", "") for k in kpis
                                       if k.get("KPI_Name", "") not in PERSONAL_KPIS))
                if not all_jobs:
                    all_jobs = ["--- لا توجد وظائف ---"]
                n_job     = st.selectbox("المسمى الوظيفي *", all_jobs)
                
                # استخراج الأقسام من الموظفين الحاليين
                depts = sorted(set(e.get("Department", e.get("القسم", "")) for e in emps if e.get("Department", e.get("القسم", ""))))
                n_dept    = st.selectbox("القسم *", depts + ["--- أخرى ---"] if depts else ["--- أخرى ---"])
                if n_dept == "--- أخرى ---":
                    n_dept = st.text_input("اسم القسم الجديد")
                
                # استخراج المقيمين
                reviewers = sorted(set(e.get("Manager", e.get("اسم المقيم ", e.get("اسم المقيم", ""))) for e in emps if e.get("Manager", e.get("اسم المقيم ", e.get("اسم المقيم", "")))))
                if not reviewers:
                    reviewers = [""]
                n_rev     = st.selectbox("المقيم *", reviewers)

            if st.form_submit_button("➕ إضافة الموظف", use_container_width=True, type="primary"):
                if not n_name.strip():
                    st.error("اسم الموظف إلزامي")
                elif n_job == "--- لا توجد وظائف ---":
                    st.error("يجب نقل بيانات المؤشرات أولاً من صفحة الإعدادات")
                else:
                    emps.append({
                        "EmployeeName": n_name.strip(),
                        "JobTitle":     n_job,
                        "Department":   n_dept,
                        "Manager":      n_rev,
                        "emp_id":       n_id.strip(),
                        "hire_date":    str(n_hire),
                        "status":       n_status,
                    })
                    _save_emps(emps)
                    st.success(f"✅ تمت إضافة {n_name}")
                    st.rerun()

    # ── تعديل ──────────────────────────────────────────────
    with tab_edit:
        emp_names = [e.get("EmployeeName", "") for e in emps]
        if not emp_names:
            st.info("لا يوجد موظفون للتعديل")
        else:
            sel       = st.selectbox("اختر الموظف للتعديل", emp_names, key="edit_sel")
            emp_data  = next((e for e in emps if e.get("EmployeeName","") == sel), {})

            with st.form("edit_emp_form"):
                c1, c2 = st.columns(2)
                with c1:
                    e_name   = st.text_input("اسم الموظف", value=emp_data.get("EmployeeName", ""))
                    e_id     = st.text_input("الرقم الوظيفي", value=emp_data.get("emp_id", ""))
                    e_hire   = st.text_input("تاريخ التعيين", value=emp_data.get("hire_date", ""))
                    statuses = ["فعّال", "غير فعّال"]
                    cur_st   = emp_data.get("status", "فعّال")
                    e_status = st.selectbox("الحالة", statuses,
                        index=statuses.index(cur_st) if cur_st in statuses else 0)
                with c2:
                    kpis     = _load_kpis()
                    all_jobs = sorted(set(k.get("JobTitle", "") for k in kpis
                                          if k.get("KPI_Name", "") not in PERSONAL_KPIS))
                    if not all_jobs:
                        all_jobs = [emp_data.get("JobTitle", "")]
                    cur_job  = emp_data.get("JobTitle", "")
                    e_job    = st.selectbox("المسمى الوظيفي", all_jobs,
                        index=all_jobs.index(cur_job) if cur_job in all_jobs else 0)
                    e_dept   = st.text_input("القسم", value=emp_data.get("Department", emp_data.get("القسم", "")))
                    reviewers = sorted(set(e.get("Manager", e.get("اسم المقيم ", e.get("اسم المقيم", ""))) for e in emps if e.get("Manager", e.get("اسم المقيم ", e.get("اسم المقيم", "")))))
                    if not reviewers:
                        reviewers = [emp_data.get("Manager", "")]
                    cur_rev  = emp_data.get("Manager", emp_data.get("اسم المقيم ", emp_data.get("اسم المقيم", "")))
                    e_rev    = st.selectbox("المقيم", reviewers,
                        index=reviewers.index(cur_rev) if cur_rev in reviewers else 0)

                if st.form_submit_button("💾 حفظ التعديلات", use_container_width=True, type="primary"):
                    for e in emps:
                        if e.get("EmployeeName", "") == sel:
                            e["EmployeeName"] = e_name.strip()
                            e["JobTitle"]     = e_job
                            e["Department"]   = e_dept.strip()
                            e["Manager"]      = e_rev
                            e["emp_id"]       = e_id.strip()
                            e["hire_date"]    = e_hire.strip()
                            e["status"]       = e_status
                            break
                    _save_emps(emps)
                    st.success("✅ تم حفظ التعديلات")
                    st.rerun()

    # ── حذف ────────────────────────────────────────────────
    with tab_del:
        emp_names = [e.get("EmployeeName", "") for e in emps]
        if not emp_names:
            st.info("لا يوجد موظفون للحذف")
        else:
            del_sel   = st.selectbox("اختر الموظف للحذف", emp_names, key="del_sel")
            st.warning(f"⚠️ سيتم حذف **{del_sel}** من قائمة الموظفين فقط — لن تُحذف بياناته التقييمية.")
            if st.checkbox("✅ أؤكد رغبتي في الحذف", key="del_emp_confirm"):
                if st.button("🗑️ تنفيذ الحذف", type="primary", use_container_width=True):
                    emps = [e for e in emps if e.get("EmployeeName", "") != del_sel]
                    _save_emps(emps)
                    st.success(f"✅ تم حذف {del_sel}")
                    st.rerun()


# ══════════════════════════════════════════════════════════
# تبويب قائمة مؤشرات الأداء
# ══════════════════════════════════════════════════════════
def render_kpis_panel():
    st.markdown("#### 📊 قائمة مؤشرات الأداء")
    st.info("""
    **ملاحظة مهمة:** التعديل أو الحذف يسري فقط على التقييمات الجديدة من تاريخ التعديل.
    التقييمات القديمة المحفوظة لن تتأثر.
    """)

    kpis = _load_kpis()
    
    if not kpis:
        st.warning("⚠️ لا توجد مؤشرات. قم بنقل البيانات من Excel أولاً من أعلى صفحة الإعدادات.")
        return
    
    all_jobs = sorted(set(k.get("JobTitle", "") for k in kpis if k.get("KPI_Name", "") not in PERSONAL_KPIS))

    # ── بحث بالوظيفة ───────────────────────────────────────
    sel_job = st.selectbox("🔍 اختر الوظيفة", ["-- اختر --"] + all_jobs)
    if sel_job == "-- اختر --":
        st.info("اختر وظيفة لعرض مؤشراتها.")
        return

    job_kpis  = [k for k in kpis if k.get("JobTitle","") == sel_job and k.get("KPI_Name","") not in PERSONAL_KPIS]
    pers_kpis = [k for k in kpis if k.get("KPI_Name","") in PERSONAL_KPIS]

    job_w  = round(sum(k.get("Weight", 0) for k in job_kpis), 1)
    pers_w = round(sum(k.get("Weight", 0) for k in pers_kpis[:5]), 1) if pers_kpis else PERSONAL_WEIGHT * len(PERSONAL_KPIS)

    c1, c2, c3 = st.columns(3)
    c1.metric("مؤشرات وظيفية", f"{len(job_kpis)} مؤشر")
    c2.metric("وزن وظيفي", f"{job_w}%")
    c3.metric("وزن شخصي", f"{pers_w}%")

    # ── جدول المؤشرات الوظيفية ────────────────────────────
    st.markdown("##### 🎯 المؤشرات الوظيفية")
    if job_kpis:
        df_job = pd.DataFrame(job_kpis)
        df_job.index = range(1, len(df_job)+1)
        st.dataframe(df_job.rename(columns={"JobTitle":"الوظيفة","KPI_Name":"المؤشر","Weight":"الوزن%" if "Weight" in df_job.columns else "الوزن"}),
                     use_container_width=True)
    else:
        st.info("لا توجد مؤشرات وظيفية لهذه الوظيفة.")

    # ── الصفات الشخصية (ثابتة لكل الوظائف) ───────────────
    st.markdown("##### 🌟 مؤشرات الصفات الشخصية (ثابتة لجميع الوظائف)")
    df_pers = pd.DataFrame([{"المؤشر": p, "الوزن%": PERSONAL_WEIGHT} for p in PERSONAL_KPIS])
    df_pers.index = range(1, len(df_pers)+1)
    st.dataframe(df_pers, use_container_width=True)

    st.markdown("---")
    tab_add, tab_edit, tab_del = st.tabs(["➕ إضافة مؤشر", "✏️ تعديل مؤشر", "🗑️ حذف مؤشر"])

    # ── إضافة مؤشر ────────────────────────────────────────
    with tab_add:
        with st.form("add_kpi_form"):
            new_kpi_name = st.text_input("اسم المؤشر الجديد *")
            new_kpi_w    = st.number_input("الوزن% *", min_value=1, max_value=50, value=8)
            st.caption(f"مجموع الأوزان الحالية: {job_w}% | بعد الإضافة: {round(job_w+new_kpi_w,1)}%")
            if st.form_submit_button("➕ إضافة", use_container_width=True, type="primary"):
                if not new_kpi_name.strip():
                    st.error("اسم المؤشر إلزامي")
                else:
                    kpis.append({"JobTitle": sel_job, "KPI_Name": new_kpi_name.strip(), "Weight": float(new_kpi_w)})
                    _save_kpis(kpis)
                    st.success(f"✅ تمت إضافة: {new_kpi_name}")
                    st.rerun()

    # ── تعديل مؤشر ────────────────────────────────────────
    with tab_edit:
        kpi_names = [k.get("KPI_Name", "") for k in job_kpis]
        if not kpi_names:
            st.info("لا توجد مؤشرات.")
        else:
            sel_kpi = st.selectbox("اختر المؤشر للتعديل", kpi_names, key="edit_kpi_sel")
            cur_kpi = next((k for k in job_kpis if k.get("KPI_Name","") == sel_kpi), {})
            with st.form("edit_kpi_form"):
                e_kname = st.text_input("اسم المؤشر", value=cur_kpi.get("KPI_Name", ""))
                e_kw    = st.number_input("الوزن%", min_value=1, max_value=50,
                                           value=int(cur_kpi.get("Weight", 8)))
                st.warning("⚠️ التعديل يسري على التقييمات الجديدة فقط من تاريخ اليوم.")
                if st.form_submit_button("💾 حفظ التعديل", use_container_width=True, type="primary"):
                    for k in kpis:
                        if k.get("JobTitle","") == sel_job and k.get("KPI_Name","") == sel_kpi:
                            k["KPI_Name"] = e_kname.strip()
                            k["Weight"]   = float(e_kw)
                            break
                    _save_kpis(kpis)
                    st.success("✅ تم التعديل — يسري على التقييمات الجديدة")
                    st.rerun()

    # ── حذف مؤشر ──────────────────────────────────────────
    with tab_del:
        kpi_names = [k.get("KPI_Name", "") for k in job_kpis]
        if not kpi_names:
            st.info("لا توجد مؤشرات.")
        else:
            del_kpi = st.selectbox("اختر المؤشر للحذف", kpi_names, key="del_kpi_sel")
            st.warning("⚠️ الحذف يسري على التقييمات الجديدة فقط. التقييمات القديمة محفوظة.")
            if st.checkbox("✅ أؤكد الحذف", key="del_kpi_confirm"):
                if st.button("🗑️ تنفيذ الحذف", type="primary", use_container_width=True):
                    kpis = [k for k in kpis if not (k.get("JobTitle","") == sel_job and k.get("KPI_Name","") == del_kpi)]
                    _save_kpis(kpis)
                    st.success(f"✅ تم حذف: {del_kpi}")
                    st.rerun()
