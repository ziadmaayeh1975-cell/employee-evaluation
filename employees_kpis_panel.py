"""
employees_kpis_panel.py
تبويبان: إدارة الموظفين | قائمة مؤشرات الأداء
"""
import streamlit as st
import pandas as pd
import json, os
from datetime import date
from database_manager import (
    load_data_from_db, DATA_DB, EMP_DB, KPI_DB
)
from constants import PERSONAL_KPIS, PERSONAL_WEIGHT


# ─── حفظ مباشر ────────────────────────────────────────────
def _save_emps(records):
    with open(EMP_DB,"w",encoding="utf-8") as f:
        json.dump(records, f, ensure_ascii=False, indent=2)
    load_data_from_db.clear()

def _save_kpis(records):
    with open(KPI_DB,"w",encoding="utf-8") as f:
        json.dump(records, f, ensure_ascii=False, indent=2)
    load_data_from_db.clear()

def _load_emps():
    with open(EMP_DB,"r",encoding="utf-8") as f:
        return json.load(f)

def _load_kpis():
    with open(KPI_DB,"r",encoding="utf-8") as f:
        return json.load(f)


# ══════════════════════════════════════════════════════════
# تبويب إدارة الموظفين
# ══════════════════════════════════════════════════════════
def render_employees_panel():
    st.markdown("#### 👨‍💼 إدارة الموظفين")

    emps = _load_emps()
    df   = pd.DataFrame(emps)

    # ── بحث ───────────────────────────────────────────────
    search = st.text_input("🔍 بحث باسم الموظف أو القسم أو الوظيفة")
    df_view = df.copy()
    if search:
        mask = df_view.apply(lambda r: search in " ".join(r.astype(str)), axis=1)
        df_view = df_view[mask]

    st.markdown(f"**إجمالي الموظفين: {len(df)}** | نتائج البحث: {len(df_view)}")
    st.dataframe(df_view.rename(columns={
        "EmployeeName":"اسم الموظف","JobTitle":"المسمى الوظيفي",
        "القسم":"القسم","اسم المقيم ":"المقيم",
        "emp_id":"الرقم الوظيفي","hire_date":"تاريخ التعيين","status":"الحالة"
    }), use_container_width=True, height=250)

    tab_add, tab_edit, tab_del = st.tabs(["➕ إضافة موظف","✏️ تعديل موظف","🗑️ حذف موظف"])

    # ── إضافة ──────────────────────────────────────────────
    with tab_add:
        with st.form("add_emp_form"):
            c1, c2 = st.columns(2)
            with c1:
                n_name    = st.text_input("اسم الموظف *")
                n_id      = st.text_input("الرقم الوظيفي")
                n_hire    = st.date_input("تاريخ التعيين", value=date.today())
                n_status  = st.selectbox("الحالة", ["فعّال","غير فعّال"])
            with c2:
                kpis      = _load_kpis()
                all_jobs  = sorted(set(k["JobTitle"] for k in kpis
                                       if k["KPI_Name"] not in PERSONAL_KPIS))
                n_job     = st.selectbox("المسمى الوظيفي *", all_jobs)
                depts     = sorted(set(e["القسم"] for e in emps if e.get("القسم","")))
                n_dept    = st.selectbox("القسم *", depts + ["--- أخرى ---"])
                if n_dept == "--- أخرى ---":
                    n_dept = st.text_input("اسم القسم الجديد")
                reviewers = sorted(set(e["اسم المقيم "] for e in emps if e.get("اسم المقيم ","")))
                n_rev     = st.selectbox("المقيم *", reviewers)

            if st.form_submit_button("➕ إضافة الموظف", use_container_width=True, type="primary"):
                if not n_name.strip():
                    st.error("اسم الموظف إلزامي")
                else:
                    emps.append({
                        "EmployeeName": n_name.strip(),
                        "JobTitle":     n_job,
                        "القسم":        n_dept,
                        "اسم المقيم ":  n_rev,
                        "emp_id":       n_id.strip(),
                        "hire_date":    str(n_hire),
                        "status":       n_status,
                    })
                    _save_emps(emps)
                    st.success(f"✅ تمت إضافة {n_name}")
                    st.rerun()

    # ── تعديل ──────────────────────────────────────────────
    with tab_edit:
        emp_names = [e["EmployeeName"] for e in emps]
        sel       = st.selectbox("اختر الموظف للتعديل", emp_names, key="edit_sel")
        emp_data  = next((e for e in emps if e["EmployeeName"]==sel), {})

        with st.form("edit_emp_form"):
            c1, c2 = st.columns(2)
            with c1:
                e_name   = st.text_input("اسم الموظف", value=emp_data.get("EmployeeName",""))
                e_id     = st.text_input("الرقم الوظيفي", value=emp_data.get("emp_id",""))
                e_hire   = st.text_input("تاريخ التعيين", value=emp_data.get("hire_date",""))
                statuses = ["فعّال","غير فعّال"]
                cur_st   = emp_data.get("status","فعّال")
                e_status = st.selectbox("الحالة", statuses,
                    index=statuses.index(cur_st) if cur_st in statuses else 0)
            with c2:
                kpis     = _load_kpis()
                all_jobs = sorted(set(k["JobTitle"] for k in kpis
                                      if k["KPI_Name"] not in PERSONAL_KPIS))
                cur_job  = emp_data.get("JobTitle","")
                e_job    = st.selectbox("المسمى الوظيفي", all_jobs,
                    index=all_jobs.index(cur_job) if cur_job in all_jobs else 0)
                e_dept   = st.text_input("القسم", value=emp_data.get("القسم",""))
                reviewers = sorted(set(e["اسم المقيم "] for e in emps if e.get("اسم المقيم ","")))
                cur_rev  = emp_data.get("اسم المقيم ","")
                e_rev    = st.selectbox("المقيم", reviewers,
                    index=reviewers.index(cur_rev) if cur_rev in reviewers else 0)

            if st.form_submit_button("💾 حفظ التعديلات", use_container_width=True, type="primary"):
                for e in emps:
                    if e["EmployeeName"] == sel:
                        e["EmployeeName"] = e_name.strip()
                        e["JobTitle"]     = e_job
                        e["القسم"]        = e_dept.strip()
                        e["اسم المقيم "]  = e_rev
                        e["emp_id"]       = e_id.strip()
                        e["hire_date"]    = e_hire.strip()
                        e["status"]       = e_status
                        break
                _save_emps(emps)
                st.success("✅ تم حفظ التعديلات")
                st.rerun()

    # ── حذف ────────────────────────────────────────────────
    with tab_del:
        emp_names = [e["EmployeeName"] for e in emps]
        del_sel   = st.selectbox("اختر الموظف للحذف", emp_names, key="del_sel")
        st.warning(f"⚠️ سيتم حذف **{del_sel}** من قائمة الموظفين فقط — لن تُحذف بياناته التقييمية.")
        if st.checkbox("✅ أؤكد رغبتي في الحذف", key="del_emp_confirm"):
            if st.button("🗑️ تنفيذ الحذف", type="primary", use_container_width=True):
                emps = [e for e in emps if e["EmployeeName"] != del_sel]
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

    kpis     = _load_kpis()
    all_jobs = sorted(set(k["JobTitle"] for k in kpis if k["KPI_Name"] not in PERSONAL_KPIS))

    # ── بحث بالوظيفة ───────────────────────────────────────
    sel_job  = st.selectbox("🔍 اختر الوظيفة", ["-- اختر --"] + all_jobs)
    if sel_job == "-- اختر --":
        st.info("اختر وظيفة لعرض مؤشراتها.")
        return

    job_kpis  = [k for k in kpis if k["JobTitle"]==sel_job and k["KPI_Name"] not in PERSONAL_KPIS]
    pers_kpis = [k for k in kpis if k["KPI_Name"] in PERSONAL_KPIS]

    job_w  = round(sum(k["Weight"] for k in job_kpis), 1)
    pers_w = round(sum(k["Weight"] for k in pers_kpis[:5]), 1)

    c1, c2, c3 = st.columns(3)
    c1.metric("مؤشرات وظيفية", f"{len(job_kpis)} مؤشر")
    c2.metric("وزن وظيفي", f"{job_w}%")
    c3.metric("وزن شخصي", f"{pers_w}%")

    # ── جدول المؤشرات الوظيفية ────────────────────────────
    st.markdown("##### 🎯 المؤشرات الوظيفية")
    df_job = pd.DataFrame(job_kpis)
    df_job.index = range(1, len(df_job)+1)
    st.dataframe(df_job.rename(columns={"JobTitle":"الوظيفة","KPI_Name":"المؤشر","Weight":"الوزن%"}),
                 use_container_width=True)

    # ── الصفات الشخصية (ثابتة لكل الوظائف) ───────────────
    st.markdown("##### 🌟 مؤشرات الصفات الشخصية (ثابتة لجميع الوظائف)")
    df_pers = pd.DataFrame([{"المؤشر": p, "الوزن%": PERSONAL_WEIGHT} for p in PERSONAL_KPIS])
    df_pers.index = range(1, len(df_pers)+1)
    st.dataframe(df_pers, use_container_width=True)

    st.markdown("---")
    tab_add, tab_edit, tab_del = st.tabs(["➕ إضافة مؤشر","✏️ تعديل مؤشر","🗑️ حذف مؤشر"])

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
        kpi_names = [k["KPI_Name"] for k in job_kpis]
        if not kpi_names:
            st.info("لا توجد مؤشرات.")
        else:
            sel_kpi = st.selectbox("اختر المؤشر للتعديل", kpi_names, key="edit_kpi_sel")
            cur_kpi = next((k for k in job_kpis if k["KPI_Name"]==sel_kpi), {})
            with st.form("edit_kpi_form"):
                e_kname = st.text_input("اسم المؤشر", value=cur_kpi.get("KPI_Name",""))
                e_kw    = st.number_input("الوزن%", min_value=1, max_value=50,
                                           value=int(cur_kpi.get("Weight",8)))
                st.warning("⚠️ التعديل يسري على التقييمات الجديدة فقط من تاريخ اليوم.")
                if st.form_submit_button("💾 حفظ التعديل", use_container_width=True, type="primary"):
                    for k in kpis:
                        if k["JobTitle"]==sel_job and k["KPI_Name"]==sel_kpi:
                            k["KPI_Name"] = e_kname.strip()
                            k["Weight"]   = float(e_kw)
                            break
                    _save_kpis(kpis)
                    st.success("✅ تم التعديل — يسري على التقييمات الجديدة")
                    st.rerun()

    # ── حذف مؤشر ──────────────────────────────────────────
    with tab_del:
        kpi_names = [k["KPI_Name"] for k in job_kpis]
        if not kpi_names:
            st.info("لا توجد مؤشرات.")
        else:
            del_kpi = st.selectbox("اختر المؤشر للحذف", kpi_names, key="del_kpi_sel")
            st.warning("⚠️ الحذف يسري على التقييمات الجديدة فقط. التقييمات القديمة محفوظة.")
            if st.checkbox("✅ أؤكد الحذف", key="del_kpi_confirm"):
                if st.button("🗑️ تنفيذ الحذف", type="primary", use_container_width=True):
                    kpis = [k for k in kpis if not (k["JobTitle"]==sel_job and k["KPI_Name"]==del_kpi)]
                    _save_kpis(kpis)
                    st.success(f"✅ تم حذف: {del_kpi}")
                    st.rerun()
