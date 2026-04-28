import io
from datetime import date
import openpyxl
import pandas as pd
import streamlit as st
try:
    import plotly.graph_objects as go
    PLOTLY_OK = True
except ImportError:
    PLOTLY_OK = False
from constants import MONTHS_AR, MONTHS_EN, MONTHS_SHORT, MONTH_MAP, PERSONAL_KPIS, PERSONAL_WEIGHT
from calculations import calc_monthly, get_kpi_avgs, verbal_grade, grade_color_hex, kpi_score_to_pct, rating_label
from data_loader import get_emp_notes
from auth import get_current_reviewer, get_current_role
from report_export import build_employee_sheet, print_preview_html

def _reviewer_emp_set(df_emp):
    role = get_current_role()
    current_reviewer = get_current_reviewer()
    if role == "super_admin":
        return None
    if role == "admin" and not current_reviewer:
        return None
    reviewer_col = "اسم المقيم"
    return set(df_emp[df_emp[reviewer_col].astype(str).str.strip() == current_reviewer]["EmployeeName"].dropna().tolist())

def _reviewer_emp_list(df_emp):
    allowed = _reviewer_emp_set(df_emp)
    if allowed is None:
        return df_emp["EmployeeName"].dropna().astype(str).str.strip().tolist()
    return list(allowed)

def _safe_df(df):
    if df is None or not isinstance(df, pd.DataFrame):
        return pd.DataFrame(columns=["EmployeeName","Month","KPI_Name","Weight","KPI_%","Evaluator","Notes","Year","EvalDate","Training"])
    df = df.copy()
    for col in ["EmployeeName","Month","KPI_Name","Weight","KPI_%","Year","EvalDate","Notes","Training"]:
        if col not in df.columns:
            df[col] = pd.Series(dtype="object")
    return df

def _get_month_meta(df_data, emp, month_en, year):
    mask = (df_data["EmployeeName"] == emp) & (df_data["Month"] == month_en) & (df_data["Year"] == int(year))
    sub = df_data[mask]
    if sub.empty:
        return "", "", ""
    row = sub.iloc[0]
    notes = row.get("Notes", "") if "Notes" in sub.columns else ""
    eval_d = row.get("EvalDate", "") if "EvalDate" in sub.columns else ""
    training = row.get("Training", "") if "Training" in sub.columns else ""
    return str(eval_d), str(notes), str(training)

def render_employee_report(df_emp, df_kpi, df_data):
    st.subheader("📄 نموذج التقييم النهائي للموظف")
    df_data = _safe_df(df_data)
    allowed_emps = set(_reviewer_emp_list(df_emp))
    all_evaluated = [str(e) for e in df_data["EmployeeName"].dropna().unique().tolist() if str(e).strip() in allowed_emps]
    if not all_evaluated:
        st.info("ℹ️ لا توجد تقييمات محفوظة.")
        return

    ca, cb, cc = st.columns([2,2,1])
    with ca:
        sel2 = st.selectbox("اختر الموظف", all_evaluated, key="rep_emp")
    with cc:
        sel2_year = st.selectbox("السنة", [2025,2026,2027], key="rep_year")
    emp_eval_months_en = df_data[(df_data["EmployeeName"] == sel2) & (df_data["Year"] == int(sel2_year))]["Month"].dropna().unique().tolist()
    emp_eval_months_ar = [MONTHS_AR[MONTHS_EN.index(m)] for m in emp_eval_months_en if m in MONTHS_EN]
    with cb:
        if emp_eval_months_ar:
            sel2_months = st.multiselect("تصفية بأشهر", emp_eval_months_ar, key="rep_months")
        else:
            st.warning("لا توجد تقييمات.")
            sel2_months = []
    months_en_f = [MONTH_MAP[m] for m in sel2_months] if sel2_months else None

    ei = df_emp[df_emp["EmployeeName"] == sel2]
    if ei.empty:
        st.error("بيانات الموظف غير موجودة في EMPLOYEES")
        return
    ei = ei.iloc[0]
    emp_id = str(ei.get("رقم الموظف", ""))
    job2 = str(ei.get("JobTitle", ""))
    dept2 = str(ei.get("القسم", ""))
    mgr2 = str(ei.get("اسم المقيم", ""))

    monthly_rep = []
    for idx, (en, short) in enumerate(zip(MONTHS_EN, MONTHS_SHORT)):
        if months_en_f and en not in months_en_f:
            monthly_rep.append((idx+1, short, 0.0, "", "", ""))
        else:
            score = calc_monthly(df_data, sel2, en, sel2_year)
            ev, nm, tr = _get_month_meta(df_data, sel2, en, sel2_year)
            monthly_rep.append((idx+1, short, score, ev, nm, tr))

    done2 = [(n,m,s) for n,m,s,*_ in monthly_rep if s > 0]
    
    job_kpis2 = []
    job_kpis_df = df_kpi[df_kpi["JobTitle"] == job2]
    for _, row in job_kpis_df.iterrows():
        kpi_name = row["KPI_Name"]
        if kpi_name in PERSONAL_KPIS:
            continue
        weight = float(row["Weight"])
        scores = []
        for en in MONTHS_EN:
            if months_en_f and en not in months_en_f:
                continue
            mask = (df_data["EmployeeName"] == sel2) & (df_data["Month"] == en) & (df_data["Year"] == int(sel2_year)) & (df_data["KPI_Name"] == kpi_name)
            sub = df_data[mask]
            if not sub.empty:
                scores.append(sub["KPI_%"].sum())
        avg_score = sum(scores) / len(scores) if scores else 0.0
        job_kpis2.append((kpi_name, weight, avg_score))
    
    pers_kpis2 = []
    personal_kpis_df = df_kpi[(df_kpi["JobTitle"] == job2) & (df_kpi["KPI_Name"].isin(PERSONAL_KPIS))]
    if personal_kpis_df.empty:
        for kpi_name in PERSONAL_KPIS:
            weight = PERSONAL_WEIGHT
            scores = []
            for en in MONTHS_EN:
                if months_en_f and en not in months_en_f:
                    continue
                mask = (df_data["EmployeeName"] == sel2) & (df_data["Month"] == en) & (df_data["Year"] == int(sel2_year)) & (df_data["KPI_Name"] == kpi_name)
                sub = df_data[mask]
                if not sub.empty:
                    scores.append(sub["KPI_%"].sum())
            avg_score = sum(scores) / len(scores) if scores else 0.0
            pers_kpis2.append((kpi_name, weight, avg_score))
    else:
        for _, row in personal_kpis_df.iterrows():
            kpi_name = row["KPI_Name"]
            weight = float(row["Weight"])
            scores = []
            for en in MONTHS_EN:
                if months_en_f and en not in months_en_f:
                    continue
                mask = (df_data["EmployeeName"] == sel2) & (df_data["Month"] == en) & (df_data["Year"] == int(sel2_year)) & (df_data["KPI_Name"] == kpi_name)
                sub = df_data[mask]
                if not sub.empty:
                    scores.append(sub["KPI_%"].sum())
            avg_score = sum(scores) / len(scores) if scores else 0.0
            pers_kpis2.append((kpi_name, weight, avg_score))

    avg2 = sum(s for _,_,s in done2)/len(done2) if done2 else 0.0
    pct2 = avg2 * 100
    verb2 = verbal_grade(pct2)
    clr2 = grade_color_hex(pct2)

    job_scores_monthly = []
    pers_scores_monthly = []
    for en in MONTHS_EN:
        if months_en_f and en not in months_en_f:
            continue
        s_all = df_data[(df_data["EmployeeName"] == sel2) & (df_data["Month"] == en) & (df_data["Year"] == int(sel2_year))]
        s_job = s_all[~s_all["KPI_Name"].isin(PERSONAL_KPIS)]
        s_per = s_all[s_all["KPI_Name"].isin(PERSONAL_KPIS)]
        if not s_job.empty:
            job_scores_monthly.append(s_job["KPI_%"].sum())
        if not s_per.empty:
            pers_scores_monthly.append(s_per["KPI_%"].sum())

    job_avg2 = round(sum(job_scores_monthly)/len(job_scores_monthly),1) if job_scores_monthly else 0.0
    pers_avg2 = round(sum(pers_scores_monthly)/len(pers_scores_monthly),1) if pers_scores_monthly else 0.0

    notes2, training2 = "", ""
    for _,_,sc_,ev,nm,tr in monthly_rep:
        if sc_>0:
            notes2, training2 = nm, tr
            break
    if not notes2 and not training2:
        _fb = get_emp_notes(sel2)
        notes2, training2 = _fb[0] if len(_fb)>0 else "", _fb[1] if len(_fb)>1 else ""

    disciplinary_df = None
    try:
        from disciplinary_manager import get_actions_by_employee
        disc_actions_list = get_actions_by_employee(sel2, sel2_year)
        if disc_actions_list:
            disciplinary_df = pd.DataFrame(disc_actions_list)
    except Exception as e:
        st.warning(f"⚠️ خطأ في الإجراءات التأديبية: {e}")

    st.markdown(f"""
    <div style="background:#F8FAFC;border:1px solid #CBD5E1;border-radius:12px;padding:16px;margin-bottom:10px;direction:rtl;">
        <h2 style="margin:0 0 4px;color:#1E3A8A;">{sel2}</h2>
        <p style="margin:3px 0;color:#475569;">
            🆔 {emp_id} &nbsp;|&nbsp; 💼 {job2} &nbsp;|&nbsp; 🏢 {dept2} &nbsp;|&nbsp; 👨‍💼 {mgr2} &nbsp;|&nbsp; 📅 {date.today().strftime('%d/%m/%Y')} &nbsp;|&nbsp; 📆 {len(done2)}/12
        </p>
    </div>
    """, unsafe_allow_html=True)

    st.markdown(f"""
    <div style="background:white;border:2px solid #1E3A8A;border-radius:12px;padding:18px;text-align:center;margin-bottom:12px;">
        <div style="font-size:13px;color:#64748B;">✅ النتيجة النهائية السنوية</div>
        <div style="font-size:3rem;font-weight:bold;color:{clr2};">{int(round(pct2))}%</div>
        <div style="font-size:1.1rem;color:{clr2};">{verb2}</div>
    </div>
    """, unsafe_allow_html=True)

    col1, col2 = st.columns(2)
    with col1:
        st.markdown(f"""
        <div style="background:white;border:1px solid #1E3A8A;border-radius:12px;padding:14px;text-align:center;">
            <div style="font-size:12px;color:#64748B;">🎯 متوسط الأداء الوظيفي</div>
            <div style="font-size:2rem;font-weight:bold;color:#1E3A8A;">{job_avg2}%</div>
        </div>
        """, unsafe_allow_html=True)
    with col2:
        st.markdown(f"""
        <div style="background:white;border:1px solid #ED7D31;border-radius:12px;padding:14px;text-align:center;">
            <div style="font-size:12px;color:#64748B;">🌟 متوسط الصفات الشخصية</div>
            <div style="font-size:2rem;font-weight:bold;color:#ED7D31;">{pers_avg2}%</div>
        </div>
        """, unsafe_allow_html=True)

    if done2:
        st.markdown("---")
        st.markdown("### 📅 نتيجة التقييم الشهري")
        
        monthly_table_data = []
        for n, short, score, ev_date, note, train in monthly_rep:
            month_name = MONTHS_AR[n-1]
            if score > 0:
                monthly_table_data.append({
                    "الشهر": month_name,
                    "الدرجة (%)": f"{round(score*100, 1)}%",
                    "التقييم اللفظي": verbal_grade(score*100),
                    "تاريخ التقييم": ev_date if ev_date else "—",
                    "ملاحظات المقيم": note if note else "—"
                })
            else:
                monthly_table_data.append({
                    "الشهر": month_name,
                    "الدرجة (%)": "—",
                    "التقييم اللفظي": "—",
                    "تاريخ التقييم": "—",
                    "ملاحظات المقيم": "—"
                })
        
        st.dataframe(
            pd.DataFrame(monthly_table_data), 
            use_container_width=True, 
            hide_index=True,
            column_config={
                "الشهر": st.column_config.TextColumn("الشهر", width="small"),
                "الدرجة (%)": st.column_config.TextColumn("الدرجة", width="small"),
                "التقييم اللفظي": st.column_config.TextColumn("التقييم", width="medium"),
                "تاريخ التقييم": st.column_config.TextColumn("تاريخ التقييم", width="medium"),
                "ملاحظات المقيم": st.column_config.TextColumn("ملاحظات المقيم", width="large"),
            }
        )
        
        st.markdown("---")
        st.subheader("🎯 مؤشرات الأداء الوظيفي")
        if job_kpis2:
            job_df = pd.DataFrame([{
                "المؤشر": k,
                "الوزن (%)": w,
                "الدرجة (0-100)": round(kpi_score_to_pct(g, w), 1),
                "التقييم": rating_label(kpi_score_to_pct(g, w))
            } for k, w, g in job_kpis2])
            st.dataframe(job_df, use_container_width=True, hide_index=True)
        else:
            st.info("لا توجد مؤشرات أداء وظيفي")

        st.subheader("🌟 مؤشرات الصفات الشخصية")
        if pers_kpis2:
            pers_df = pd.DataFrame([{
                "المؤشر": k,
                "الوزن (%)": w,
                "الدرجة (0-100)": round(kpi_score_to_pct(g, w), 1),
                "التقييم": rating_label(kpi_score_to_pct(g, w))
            } for k, w, g in pers_kpis2])
            st.dataframe(pers_df, use_container_width=True, hide_index=True)
        else:
            st.info("لا توجد مؤشرات صفات شخصية")

        if disciplinary_df is not None and not disciplinary_df.empty:
            st.subheader("⚠️ الإجراءات التأديبية المسجلة")
            disc_display = disciplinary_df.copy()
            disc_display = disc_display.rename(columns={
                "action_date": "التاريخ",
                "warning_type": "نوع الإنذار",
                "reason": "السبب",
                "deduction_days": "خصم (أيام)"
            })
            cols_to_show = ["التاريخ", "نوع الإنذار", "السبب", "خصم (أيام)"]
            available_cols = [c for c in cols_to_show if c in disc_display.columns]
            if available_cols:
                st.dataframe(disc_display[available_cols], use_container_width=True, hide_index=True)

        cn, ct = st.columns(2)
        with cn:
            st.info(f"📝 **ملاحظات المقيم:** {notes2 or '—'}")
        with ct:
            st.info(f"🎓 **الاحتياجات التدريبية:** {training2 or '—'}")

        months_done_list = [(MONTHS_AR[n-1], round(s*100, 1)) for n, _, s, *_ in monthly_rep if s > 0]
        if months_done_list and PLOTLY_OK:
            st.markdown("---")
            fig = go.Figure()
            colors = ["#4472C4","#ED7D31","#A5A5A5","#FFC000","#5B9BD5","#70AD47","#264478","#9E480E","#636363","#997300","#255E91","#43682B"]
            for i,(mon, sc) in enumerate(months_done_list):
                fig.add_trace(go.Bar(name=mon, x=[mon], y=[sc], marker_color=colors[i%len(colors)], text=f"{sc}%", textposition="outside"))
            fig.update_layout(barmode="group", title=f"التقييم السنوي — {sel2} — {sel2_year}", xaxis_title="الأشهر", yaxis_title="الدرجة %", yaxis_range=[0, 120], height=420)
            st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("لا توجد تقييمات")

    st.subheader("⬇️ تحميل التقرير")
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    kpis_export = [{"KPI_Name":k,"Weight":w,"avg_score":g} for k,w,g in job_kpis2+pers_kpis2]
    build_employee_sheet(
        wb, sel2, job2, dept2, mgr2, sel2_year,
        kpis_export, monthly_rep, notes2, training2,
        employee_id=emp_id,
        disciplinary_actions=disciplinary_df
    )
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    st.download_button("📥 تحميل Excel", data=buf, file_name=f"تقييم_{sel2}_{sel2_year}.xlsx", use_container_width=True)
    st.markdown("---")
    st.markdown("#### 🖨️ معاينة الطباعة")
    html_preview = print_preview_html(io.BytesIO(buf.getvalue()), f"تقييم {sel2}")
    st.components.v1.html(html_preview, height=1100, scrolling=True)
