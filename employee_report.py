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
from constants import MONTHS_AR, MONTHS_EN, MONTHS_SHORT, MONTH_MAP, PERSONAL_KPIS
from calculations import calc_monthly, get_kpi_avgs, verbal_grade, grade_color_hex, kpi_score_to_pct, rating_label
from data_loader import get_emp_notes
from auth import get_current_reviewer, get_current_role
from report_export import build_employee_sheet, print_preview_html

# استيراد دوال الإجراءات التأديبية
try:
    from disciplinary_loader import load_disciplinary_actions, get_employee_disciplinary, format_disciplinary_text
    DISCIPLINARY_AVAILABLE = True
except ImportError:
    DISCIPLINARY_AVAILABLE = False

def _reviewer_emp_set(df_emp):
    """
    None  = كل الموظفين (super_admin)
    set() = موظفو المقيم فقط
    """
    from auth import get_current_reviewer, get_current_role
    role             = get_current_role()
    current_reviewer = get_current_reviewer()

    # الأدمن الرئيسي → كل الموظفين دائماً
    if role == "super_admin":
        return None

    # أدمن عادي بدون reviewer → كل الموظفين
    if role == "admin" and not current_reviewer:
        return None

    # أدمن عادي أو user مع reviewer → موظفوه فقط
    reviewer_col = df_emp.columns[3] if len(df_emp.columns) > 3 else df_emp.columns[-1]
    return set(
        str(e).strip() for e in
        df_emp[df_emp[reviewer_col].astype(str).str.strip() == current_reviewer
               ]["EmployeeName"].dropna().tolist()
        if str(e).strip() not in ("","nan")
    )


def _reviewer_emp_list(df_emp):
    """يُعيد قائمة الموظفين المسموح برؤيتهم."""
    allowed = _reviewer_emp_set(df_emp)
    if allowed is None:
        return df_emp["EmployeeName"].dropna().astype(str).str.strip().tolist()
    return list(allowed)



def _safe_df(df):
    if df is None or not isinstance(df, pd.DataFrame):
        return pd.DataFrame(columns=[
            "EmployeeName","Month","KPI_Name","Weight","KPI_%",
            "Evaluator","Notes","Year","EvalDate","Training"
        ])
    df = df.copy()
    for col in ["EmployeeName","Month","KPI_Name","Weight","KPI_%","Year","EvalDate","Notes","Training"]:
        if col not in df.columns:
            df[col] = pd.Series(dtype="object")
    return df


def _get_month_meta(df_data, emp, month_en, year):
    mask = (
        (df_data["EmployeeName"] == emp) &
        (df_data["Month"]        == month_en) &
        (df_data["Year"]         == int(year))
    )
    sub = df_data[mask]
    if sub.empty:
        return "", "", ""
    row = sub.iloc[0]

    def _by_name(*names):
        for n in names:
            if n in sub.columns:
                v = str(row[n] or "").strip()
                if v and v not in ("nan","None",""):
                    return v
        return ""

    def _by_idx(idx):
        cols = list(sub.columns)
        if idx < len(cols):
            v = str(row[cols[idx]] or "").strip()
            if v and v not in ("nan","None",""):
                return v
        return ""

    notes    = _by_name("Notes","notes")                    or _by_idx(6)
    eval_d   = _by_name("EvalDate","eval_date","EntryDate") or _by_idx(8)
    training = _by_name("Training","training")              or _by_idx(9)
    return eval_d, notes, training


def render_employee_report(df_emp, df_kpi, df_data):
    st.subheader("📄 نموذج التقييم النهائي للموظف")

    df_data = _safe_df(df_data)

    # ── فلترة الموظفين حسب المقيم ───────────────────────────────
    allowed_emps = set(_reviewer_emp_list(df_emp))

    all_evaluated = [
        str(e) for e in df_data["EmployeeName"].dropna().unique().tolist()
        if str(e).strip() not in ("","nan") and str(e).strip() in allowed_emps
    ]

    if not all_evaluated:
        st.info("ℹ️ لا توجد تقييمات محفوظة لموظفيك حتى الآن.")
        return

    ca, cb, cc = st.columns([2, 2, 1])
    with ca:
        sel2 = st.selectbox("اختر الموظف", all_evaluated, key="rep_emp")
    with cc:
        sel2_year = st.selectbox("السنة", [2025, 2026, 2027], key="rep_year")

    emp_eval_months_en = df_data[
        (df_data["EmployeeName"] == sel2) &
        (df_data["Year"] == int(sel2_year))
    ]["Month"].dropna().unique().tolist()
    emp_eval_months_ar = [MONTHS_AR[MONTHS_EN.index(m)] for m in emp_eval_months_en if m in MONTHS_EN]

    with cb:
        if emp_eval_months_ar:
            sel2_months = st.multiselect(
                "تصفية بأشهر (الأشهر المُقيَّمة فقط)",
                emp_eval_months_ar, key="rep_months"
            )
        else:
            st.warning("لا توجد تقييمات لهذا الموظف في السنة المختارة.")
            sel2_months = []

    months_en_f = [MONTH_MAP[m] for m in sel2_months] if sel2_months else None

    if sel2_months:
        missing = [m for m in sel2_months if MONTH_MAP.get(m) not in emp_eval_months_en]
        if missing:
            st.warning(f"⚠️ لا يوجد تقييم للأشهر التالية: {', '.join(missing)}")

    ei = df_emp[df_emp["EmployeeName"] == sel2]
    ei    = ei.iloc[0] if not ei.empty else df_emp.iloc[0]
    job2  = str(ei.iloc[1]).strip()
    dept2 = str(ei.iloc[2]).strip()
    mgr2  = str(ei.iloc[3]).strip()

    monthly_rep = []
    for idx, (en, short) in enumerate(zip(MONTHS_EN, MONTHS_SHORT)):
        if months_en_f and en not in months_en_f:
            monthly_rep.append((idx+1, short, 0.0, "", "", ""))
        else:
            score      = calc_monthly(df_data, sel2, en, sel2_year)
            ev, nm, tr = _get_month_meta(df_data, sel2, en, sel2_year)
            monthly_rep.append((idx+1, short, score, ev, nm, tr))

    done2  = [(n,m,s) for n,m,s,*_ in monthly_rep if s > 0]
    kpis2  = get_kpi_avgs(df_data, df_kpi, sel2, job2, months_en_f, sel2_year)

    _P = PERSONAL_KPIS
    job_kpis2  = [(k,w,g) for k,w,g in kpis2 if k not in _P]
    pers_kpis2 = [(k,w,g) for k,w,g in kpis2 if k in _P]

    avg2   = sum(s for _,_,s in done2)/len(done2) if done2 else 0.0
    pct2   = avg2 * 100
    verb2  = verbal_grade(pct2)
    clr2   = grade_color_hex(pct2)

    job_scores_monthly = []
    pers_scores_monthly = []
    for en in MONTHS_EN:
        if months_en_f and en not in months_en_f:
            continue
        mask_base = (
            (df_data["EmployeeName"] == sel2) &
            (df_data["Month"] == en) &
            (df_data["Year"] == int(sel2_year))
        )
        s_all = df_data[mask_base]
        s_job = s_all[~s_all["KPI_Name"].isin(PERSONAL_KPIS)]
        s_per = s_all[s_all["KPI_Name"].isin(PERSONAL_KPIS)]
        if not s_job.empty:
            job_scores_monthly.append(s_job["KPI_%"].sum())
        if not s_per.empty:
            pers_scores_monthly.append(s_per["KPI_%"].sum())

    job_avg2  = round(sum(job_scores_monthly)/len(job_scores_monthly), 1) if job_scores_monthly else 0.0
    pers_avg2 = round(sum(pers_scores_monthly)/len(pers_scores_monthly), 1) if pers_scores_monthly else 0.0

    notes2 = ""; training2 = ""
    for _, _, sc_, ev, nm, tr in monthly_rep:
        if sc_ > 0:
            notes2    = nm
            training2 = tr
            break
    if not notes2 and not training2:
        _fb = get_emp_notes(sel2)
        notes2    = _fb[0] if len(_fb) > 0 else ""
        training2 = _fb[1] if len(_fb) > 1 else ""

    # ═══════════════════════════════════════════════════════════════════
    # 🔍 تحميل الإجراءات التأديبية للموظف
    # ═══════════════════════════════════════════════════════════════════
    disciplinary_df = None
    if DISCIPLINARY_AVAILABLE:
        try:
            df_disc = load_disciplinary_actions()
            if not df_disc.empty:
                disciplinary_df = get_employee_disciplinary(df_disc, sel2, sel2_year)
        except Exception as e:
            st.warning(f"⚠️ خطأ في تحميل الإجراءات التأديبية: {e}")

    st.markdown(f"""
    <div style="background:#F8FAFC;border:1px solid #CBD5E1;border-radius:12px;
                padding:16px;margin-bottom:10px;direction:rtl;">
        <h2 style="margin:0 0 4px;color:#1E3A8A;">{sel2}</h2>
        <p style="margin:3px 0;color:#475569;">💼 {job2} &nbsp;|&nbsp; 🏢 {dept2}
           &nbsp;|&nbsp; 👨‍💼 {mgr2} &nbsp;|&nbsp; 📅 {date.today().strftime('%d/%m/%Y')}
           &nbsp;|&nbsp; 📆 أشهر مُقيَّمة: {len(done2)}</p>
    </div>""", unsafe_allow_html=True)

    st.markdown(f"""
    <div style="background:white;border:2px solid #1E3A8A;border-radius:12px;
                padding:18px;text-align:center;direction:rtl;margin-bottom:12px;">
        <div style="font-size:13px;color:#64748B;font-weight:600;margin-bottom:6px;">
            ✅ النتيجة النهائية السنوية
        </div>
        <div style="font-size:3rem;font-weight:bold;color:{clr2};">{int(round(pct2))}%</div>
        <div style="font-size:1.1rem;color:{clr2};font-weight:600;">{verb2}</div>
    </div>""", unsafe_allow_html=True)

    sc_c1, sc_c2 = st.columns(2)
    with sc_c1:
        st.markdown(f"""
        <div style="background:white;border:1px solid #1E3A8A;border-radius:12px;
                    padding:14px;text-align:center;direction:rtl;">
            <div style="font-size:12px;color:#64748B;font-weight:600;margin-bottom:6px;">
                🎯 متوسط مؤشرات الأداء الوظيفي
            </div>
            <div style="font-size:2rem;font-weight:bold;color:#1E3A8A;">{job_avg2}%</div>
        </div>""", unsafe_allow_html=True)
    with sc_c2:
        st.markdown(f"""
        <div style="background:white;border:1px solid #ED7D31;border-radius:12px;
                    padding:14px;text-align:center;direction:rtl;">
            <div style="font-size:12px;color:#64748B;font-weight:600;margin-bottom:6px;">
                🌟 متوسط مؤشرات الصفات الشخصية
            </div>
            <div style="font-size:2rem;font-weight:bold;color:#ED7D31;">{pers_avg2}%</div>
        </div>""", unsafe_allow_html=True)

    st.markdown("<div style='margin:10px 0'></div>", unsafe_allow_html=True)

    if done2:
        ca2, cb2 = st.columns(2)
        with ca2:
            st.subheader("📅 التقييم الشهري")
            st.dataframe(pd.DataFrame([
                {
                    "الشهر":          MONTHS_AR[n-1],
                    "الدرجة (%)":     round(s*100, 1),
                    "التقييم اللفظي": verbal_grade(s*100),
                }
                for n,_,s,*_ in monthly_rep if s > 0
            ]), hide_index=True, use_container_width=True)

        with cb2:
            job_kpis_show  = [(k,w,g) for k,w,g in kpis2 if k not in PERSONAL_KPIS and g > 0]
            pers_kpis_show = [(k,w,g) for k,w,g in kpis2 if k in PERSONAL_KPIS and g > 0]

            st.subheader("🎯 مؤشرات الأداء الوظيفي")
            if job_kpis_show:
                st.dataframe(pd.DataFrame([
                    {
                        "المؤشر":           k,
                        "الوزن النسبي (%)": w,
                        "الدرجة (0-100)":   round(kpi_score_to_pct(g, w), 1),
                        "التقييم":          rating_label(kpi_score_to_pct(g, w)),
                    }
                    for k,w,g in job_kpis_show
                ]), hide_index=True, use_container_width=True)

            if pers_kpis_show:
                st.subheader("🌟 مؤشرات الصفات الشخصية")
                st.dataframe(pd.DataFrame([
                    {
                        "المؤشر":           k,
                        "الوزن النسبي (%)": w,
                        "الدرجة (0-100)":   round(kpi_score_to_pct(g, w), 1),
                        "التقييم":          rating_label(kpi_score_to_pct(g, w)),
                    }
                    for k,w,g in pers_kpis_show
                ]), hide_index=True, use_container_width=True)

        # ═══════════════════════════════════════════════════════════════════
        # ⚠️ عرض الإجراءات التأديبية في واجهة Streamlit
        # ═══════════════════════════════════════════════════════════════════
        if disciplinary_df is not None and not disciplinary_df.empty:
            st.markdown("---")
            st.markdown("#### ⚠️ سجل الإجراءات التأديبية")
            
            # عرض جدول الإجراءات
            display_df = disciplinary_df.copy()
            display_df = display_df.rename(columns={
                "warning_date": "التاريخ",
                "warning_type": "نوع الإنذار",
                "reason": "السبب",
                "deduction_days": "خصم (أيام)"
            })
            
            cols_to_show = ["التاريخ", "نوع الإنذار", "السبب", "خصم (أيام)"]
            available_cols = [c for c in cols_to_show if c in display_df.columns]
            
            st.dataframe(display_df[available_cols], hide_index=True, use_container_width=True)

        if notes2 or training2:
            cn, ct = st.columns(2)
            with cn:
                st.info(f"📝 **ملاحظات المقيم:** {notes2 or '—'}")
            with ct:
                st.info(f"🎓 **الاحتياجات التدريبية:** {training2 or '—'}")

        months_done_list = [(MONTHS_AR[n-1], round(s*100,1))
                             for n,_,s,*_ in monthly_rep if s > 0]

        if months_done_list and PLOTLY_OK:
            st.markdown("---")
            MONTH_COLORS = ["#4472C4","#ED7D31","#A5A5A5","#FFC000","#5B9BD5",
                            "#70AD47","#264478","#9E480E","#636363","#997300","#255E91","#43682B"]
            fig = go.Figure()
            for i, (mon, sc) in enumerate(zip([m for m,_ in months_done_list],
                                               [s for _,s in months_done_list])):
                fig.add_trace(go.Bar(
                    name=mon, x=[mon], y=[sc],
                    marker_color=MONTH_COLORS[i % len(MONTH_COLORS)],
                    text=f"{sc}%", textposition="outside",
                    textfont=dict(size=10), showlegend=True,
                ))
            fig.update_layout(
                barmode="group",
                title=dict(text=f"التقييم السنوي — {sel2} — {sel2_year}",
                           font=dict(size=13, color="#1E3A8A"), x=0.5),
                xaxis=dict(title="الأشهر", tickfont=dict(size=10), gridcolor="#F0F0F0"),
                yaxis=dict(title="الدرجة %", range=[0,120], tickfont=dict(size=9),
                           gridcolor="#F0F0F0", tickvals=list(range(0,121,10))),
                plot_bgcolor="white", paper_bgcolor="white", bargap=0.25,
                legend=dict(orientation="h", yanchor="bottom", y=-0.28,
                            xanchor="center", x=0.5, font=dict(size=9)),
                margin=dict(l=50, r=40, t=50, b=80), height=420,
            )
            st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("لا توجد تقييمات في الفترة المحددة.")

    st.subheader("⬇️ تحميل نموذج التقييم النهائي")
    wb2 = openpyxl.Workbook()
    wb2.remove(wb2.active)
    
    # تمرير الإجراءات التأديبية إلى دالة build_employee_sheet
    build_employee_sheet(
        wb2, sel2, job2, dept2, mgr2, sel2_year,
        kpis2, monthly_rep, notes2, training2,
        disciplinary_actions=disciplinary_df  # إضافة الإجراءات التأديبية
    )
    
    buf2 = io.BytesIO()
    wb2.save(buf2)
    buf2.seek(0)
    st.download_button(
        label="📥 تحميل Excel",
        data=buf2,
        file_name=f"تقييم_{sel2}_{date.today()}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
    st.markdown("---")
    st.markdown("#### 🖨️ معاينة وطباعة التقرير")
    html_prev2 = print_preview_html(io.BytesIO(buf2.getvalue()), f"تقييم {sel2}")
    st.components.v1.html(html_prev2, height=1100, scrolling=True)
