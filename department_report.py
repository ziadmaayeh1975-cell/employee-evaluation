import io
from datetime import date
import openpyxl
import pandas as pd
import streamlit as st
from constants import MONTHS_AR, MONTHS_EN, MONTHS_SHORT, MONTH_MAP, PERSONAL_KPIS, PERSONAL_WEIGHT
from calculations import calc_monthly, get_kpi_avgs, verbal_grade, kpi_score_to_pct, rating_label
from data_loader import get_emp_notes
from auth import get_current_reviewer, get_current_role
from report_export import build_employee_sheet, build_summary_sheet, print_preview_html

# ── استيراد الإجراءات التأديبية ────────────────────────────────────────────
try:
    from disciplinary_manager import get_actions_by_employee
    DISCIPLINARY_AVAILABLE = True
except ImportError:
    DISCIPLINARY_AVAILABLE = False

# ── استيراد نظام الالتزام بالدوام ──────────────────────────────────────────
try:
    from attendance_manager import get_employee_attendance_summary
    ATTENDANCE_AVAILABLE = True
except ImportError:
    ATTENDANCE_AVAILABLE = False


# ══════════════════════════════════════════════════════════════════════════════
# Helper functions
# ══════════════════════════════════════════════════════════════════════════════

def _reviewer_emp_set(df_emp):
    role = get_current_role()
    current_reviewer = get_current_reviewer()
    if role == "super_admin":
        return None
    if role == "admin" and not current_reviewer:
        return None
    reviewer_col = df_emp.columns[3] if len(df_emp.columns) > 3 else df_emp.columns[-1]
    return set(
        str(e).strip() for e in
        df_emp[df_emp[reviewer_col].astype(str).str.strip() == current_reviewer
               ]["EmployeeName"].dropna().tolist()
        if str(e).strip() not in ("", "nan")
    )


def _reviewer_emp_list(df_emp):
    allowed = _reviewer_emp_set(df_emp)
    if allowed is None:
        return df_emp["EmployeeName"].dropna().astype(str).str.strip().tolist()
    return list(allowed)


def _safe_df(df):
    if df is None or not isinstance(df, pd.DataFrame):
        return pd.DataFrame(columns=[
            "EmployeeName", "Month", "KPI_Name", "Weight", "KPI_%",
            "Evaluator", "Notes", "Year", "EvalDate", "Training"
        ])
    df = df.copy()
    for col in ["EmployeeName", "Month", "Year", "KPI_%", "KPI_Name",
                "Weight", "EvalDate", "Notes", "Training"]:
        if col not in df.columns:
            df[col] = pd.Series(dtype="object")
    return df


def _get_month_details(df_data, emp_name, month_en, year):
    mask = (
        (df_data["EmployeeName"] == emp_name) &
        (df_data["Month"] == month_en) &
        (df_data["Year"] == int(year))
    )
    subset = df_data[mask]
    if subset.empty:
        return "", "", ""
    row = subset.iloc[0]

    eval_date = ""
    for col in ["EvalDate", "eval_date", "EntryDate"]:
        if col in subset.columns and pd.notna(row.get(col, None)):
            val = str(row[col]).strip()
            if val and val not in ("", "nan", "None"):
                try:
                    eval_date = pd.to_datetime(val).strftime("%d/%m/%Y")
                except:
                    eval_date = val
                break

    notes = ""
    for col in ["Notes", "notes"]:
        if col in subset.columns and pd.notna(row.get(col, None)):
            val = str(row[col]).strip()
            if val and val not in ("", "nan", "None"):
                notes = val
                break

    training = ""
    for col in ["Training", "training"]:
        if col in subset.columns and pd.notna(row.get(col, None)):
            val = str(row[col]).strip()
            if val and val not in ("", "nan", "None"):
                training = val
                break

    return eval_date, notes, training


# ══════════════════════════════════════════════════════════════════════════════
# Main render function
# ══════════════════════════════════════════════════════════════════════════════

def render_department_report(df_emp, df_kpi, df_data):
    st.subheader("📊 التقرير المفصّل بالأقسام")
    df_data = _safe_df(df_data)
    allowed_emps = _reviewer_emp_set(df_emp)

    # ── تحديد عمود القسم بالاسم الصحيح ────────────────────────────────────
    # البحث عن عمود القسم بالاسم أولاً، وإلا الرجوع للعمود الثالث
    if "القسم" in df_emp.columns:
        dept_col = "القسم"
    elif "Department" in df_emp.columns:
        dept_col = "Department"
    else:
        dept_col = df_emp.columns[2]

    # تحديد عمود المسمى الوظيفي بالاسم الصحيح
    if "JobTitle" in df_emp.columns:
        job_col = "JobTitle"
    else:
        job_col = df_emp.columns[1]

    # تحديد عمود المقيم
    if "اسم المقيم" in df_emp.columns:
        mgr_col = "اسم المقيم"
    elif len(df_emp.columns) > 3:
        mgr_col = df_emp.columns[3]
    else:
        mgr_col = None

    # ── فلاتر الاختيار ──────────────────────────────────────────────────────
    c3a, c3b, c3c = st.columns([2, 2, 1])
    with c3a:
        if allowed_emps is None:
            dept_emps_all = df_emp
        else:
            dept_emps_all = df_emp[df_emp["EmployeeName"].isin(allowed_emps)]
        dept_list = dept_emps_all[dept_col].dropna().astype(str).str.strip().unique().tolist()
        sel3_dept = st.selectbox("🏢 القسم", ["-- الكل --"] + dept_list, key="dept_sel")
    with c3b:
        sel3_months = st.multiselect("📅 الأشهر", MONTHS_AR, key="dept_months")
    with c3c:
        sel3_year = st.selectbox("🗓️ السنة", [2025, 2026, 2027], key="dept_year")

    if sel3_dept == "-- الكل --":
        dept_emps = dept_emps_all["EmployeeName"].dropna().tolist()
    else:
        dept_emps = dept_emps_all[
            dept_emps_all[dept_col].astype(str).str.strip() == sel3_dept
        ]["EmployeeName"].tolist()

    if df_data.empty or "EmployeeName" not in df_data.columns:
        st.info("ℹ️ لا توجد تقييمات محفوظة حتى الآن.")
        return

    dept_with_data = [e for e in dept_emps if e in df_data["EmployeeName"].values]
    if not dept_with_data:
        st.info("لا توجد بيانات تقييم لهذا القسم بعد.")
        return

    st.markdown(f"**{len(dept_with_data)} موظف** لديهم بيانات تقييم.")
    months_en_f3 = [MONTH_MAP[m] for m in sel3_months] if sel3_months else None

    # ══════════════════════════════════════════════════════════════════════════
    # جمع بيانات كل موظف
    # ══════════════════════════════════════════════════════════════════════════
    summary3 = []

    for emp in dept_with_data:
        ei3 = df_emp[df_emp["EmployeeName"] == emp]
        if ei3.empty:
            continue
        ei3_row = ei3.iloc[0]

        dept_val = str(ei3_row[dept_col]).strip() if dept_col in ei3.columns else str(ei3_row.iloc[2]).strip()
        job_val  = str(ei3_row[job_col]).strip()  if job_col  in ei3.columns else str(ei3_row.iloc[1]).strip()
        mgr_val  = str(ei3_row[mgr_col]).strip()  if mgr_col and mgr_col in ei3.columns else (str(ei3_row.iloc[3]).strip() if len(ei3_row) > 3 else "")
        emp_id   = str(ei3_row.get("رقم الموظف", "")) if "رقم الموظف" in ei3.columns else ""

        # ── الدرجات الشهرية ────────────────────────────────────────────────
        s_list  = [calc_monthly(df_data, emp, m, sel3_year) for m in (months_en_f3 or MONTHS_EN)]
        active3 = [s for s in s_list if s > 0]
        avg3    = sum(active3) / len(active3) if active3 else 0.0

        # ── monthly_rep مثل employee_report ────────────────────────────────
        monthly_rep = []
        for idx, (en, short) in enumerate(zip(MONTHS_EN, MONTHS_SHORT)):
            if months_en_f3 and en not in months_en_f3:
                monthly_rep.append((idx + 1, short, 0.0, "", "", ""))
            else:
                score      = calc_monthly(df_data, emp, en, sel3_year)
                ev, nm, tr = _get_month_details(df_data, emp, en, sel3_year)
                monthly_rep.append((idx + 1, short, score, ev, nm, tr))

        # ── مؤشرات الأداء الوظيفي ──────────────────────────────────────────
        job_kpis = []
        job_kpis_df = df_kpi[df_kpi["JobTitle"] == job_val]
        for _, row in job_kpis_df.iterrows():
            kpi_name = row["KPI_Name"]
            if kpi_name in PERSONAL_KPIS:
                continue
            weight = float(row["Weight"])
            scores = []
            for en in (months_en_f3 or MONTHS_EN):
                mask = (
                    (df_data["EmployeeName"] == emp) &
                    (df_data["Month"]        == en) &
                    (df_data["Year"]         == int(sel3_year)) &
                    (df_data["KPI_Name"]     == kpi_name)
                )
                sub = df_data[mask]
                if not sub.empty:
                    scores.append(sub["KPI_%"].sum())
            job_kpis.append((kpi_name, weight, sum(scores) / len(scores) if scores else 0.0))

        # ── مؤشرات الصفات الشخصية ─────────────────────────────────────────
        pers_kpis = []
        personal_kpis_df = df_kpi[
            (df_kpi["JobTitle"] == job_val) & (df_kpi["KPI_Name"].isin(PERSONAL_KPIS))
        ]
        source_kpis = personal_kpis_df.iterrows() if not personal_kpis_df.empty else [
            (None, {"KPI_Name": k, "Weight": PERSONAL_WEIGHT}) for k in PERSONAL_KPIS
        ]
        for _, row in source_kpis:
            kpi_name = row["KPI_Name"]
            weight   = float(row["Weight"])
            scores   = []
            for en in (months_en_f3 or MONTHS_EN):
                mask = (
                    (df_data["EmployeeName"] == emp) &
                    (df_data["Month"]        == en) &
                    (df_data["Year"]         == int(sel3_year)) &
                    (df_data["KPI_Name"]     == kpi_name)
                )
                sub = df_data[mask]
                if not sub.empty:
                    scores.append(sub["KPI_%"].sum())
            pers_kpis.append((kpi_name, weight, sum(scores) / len(scores) if scores else 0.0))

        # ── ملاحظات وتدريب ─────────────────────────────────────────────────
        emp_notes, emp_train = "", ""
        for _, _, sc_, ev, nm, tr in monthly_rep:
            if sc_ > 0:
                emp_notes, emp_train = nm, tr
                break
        if not emp_notes and not emp_train:
            _fb = get_emp_notes(emp)
            emp_notes  = _fb[0] if len(_fb) > 0 else ""
            emp_train  = _fb[1] if len(_fb) > 1 else ""

        # ── الإجراءات التأديبية ────────────────────────────────────────────
        disciplinary_df = None
        if DISCIPLINARY_AVAILABLE:
            try:
                disc_actions_list = get_actions_by_employee(emp, sel3_year)
                if disc_actions_list:
                    disciplinary_df = pd.DataFrame(disc_actions_list)
            except:
                pass

        # ── بيانات الالتزام بالدوام (شهرياً) ─────────────────────────────
        attendance_monthly_rows = []
        att_cnt_by_month        = {}
        att_hrs_by_month        = {}
        attendance_count        = 0
        attendance_hours        = 0.0

        if ATTENDANCE_AVAILABLE:
            try:
                for month_num in range(1, 13):
                    att_summary = get_employee_attendance_summary(emp, emp_id, sel3_year, month_num)
                    lc = att_summary.get("count", 0) or 0
                    lh = att_summary.get("hours", 0) or 0.0
                    att_cnt_by_month[month_num] = lc
                    att_hrs_by_month[month_num] = lh
                    if lc > 0 or lh > 0:
                        attendance_monthly_rows.append({
                            "month":      month_num,
                            "late_count": lc,
                            "late_hours": lh,
                        })
                    attendance_count += lc
                    attendance_hours += lh
            except:
                pass

        # بناء DataFrame للتصدير
        if attendance_monthly_rows:
            attendance_export = pd.DataFrame(attendance_monthly_rows)
        elif attendance_count > 0 or attendance_hours > 0:
            attendance_export = pd.DataFrame([{
                "month":      0,
                "late_count": attendance_count,
                "late_hours": attendance_hours,
            }])
        else:
            attendance_export = None

        # ── kpis_export للتصدير ────────────────────────────────────────────
        kpis_export = [
            {"KPI_Name": k, "Weight": w, "avg_score": g}
            for k, w, g in job_kpis + pers_kpis
        ]

        summary3.append({
            "emp":               emp,
            "dept":              dept_val,
            "job":               job_val,
            "mgr":               mgr_val,
            "emp_id":            emp_id,
            "months":            len(active3),
            "pct":               round(avg3 * 100, 1),
            "verb":              verbal_grade(avg3 * 100) if active3 else "—",
            "monthly_rep":       monthly_rep,
            "job_kpis":          job_kpis,
            "pers_kpis":         pers_kpis,
            "kpis_export":       kpis_export,
            "emp_notes":         emp_notes,
            "emp_train":         emp_train,
            "disciplinary_df":   disciplinary_df,
            "attendance_count":  attendance_count,
            "attendance_hours":  attendance_hours,
            "att_cnt_by_month":  att_cnt_by_month,
            "att_hrs_by_month":  att_hrs_by_month,
            "attendance_export": attendance_export,
        })

    summary3.sort(key=lambda x: x["pct"], reverse=True)

    # ══════════════════════════════════════════════════════════════════════════
    # جدول الملخص العام
    # ══════════════════════════════════════════════════════════════════════════
    display_df = pd.DataFrame([{
        "الموظف":              s["emp"],
        "القسم":               s["dept"],
        "السنة":               sel3_year,
        "الأشهر":              s["months"],
        "المعدل (%)":          s["pct"],
        "التقييم":             s["verb"],
        "الإجراءات التأديبية": len(s["disciplinary_df"]) if s["disciplinary_df"] is not None and not s["disciplinary_df"].empty else 0,
        "عدد مرات التأخير":    s["attendance_count"],
        "ساعات التأخير":       f"{s['attendance_hours']:.2f}",
    } for s in summary3])
    st.dataframe(display_df, hide_index=True, use_container_width=True)

    # ══════════════════════════════════════════════════════════════════════════
    # عرض تفصيلي لكل موظف (مثل employee_report)
    # ══════════════════════════════════════════════════════════════════════════
    st.markdown("---")
    st.markdown("## 📋 التفاصيل الكاملة لكل موظف")

    for s in summary3:
        with st.expander(f"👤 {s['emp']} — {s['pct']}% — {s['verb']}", expanded=False):

            # بطاقة معلومات الموظف
            st.markdown(f"""
            <div style="background:#F8FAFC;border:1px solid #CBD5E1;border-radius:12px;
                        padding:14px;margin-bottom:10px;direction:rtl;">
                <h3 style="margin:0 0 4px;color:#1E3A8A;">{s['emp']}</h3>
                <p style="margin:3px 0;color:#475569;">
                    🆔 {s['emp_id']} &nbsp;|&nbsp; 💼 {s['job']} &nbsp;|&nbsp;
                    🏢 {s['dept']} &nbsp;|&nbsp; 👨‍💼 {s['mgr']}
                </p>
            </div>
            """, unsafe_allow_html=True)

            # النتيجة النهائية
            st.markdown(f"""
            <div style="background:white;border:2px solid #1E3A8A;border-radius:12px;
                        padding:14px;text-align:center;margin-bottom:10px;">
                <div style="font-size:12px;color:#64748B;">✅ المعدل السنوي</div>
                <div style="font-size:2.5rem;font-weight:bold;color:#1E3A8A;">{s['pct']}%</div>
                <div style="font-size:1rem;color:#1E3A8A;">{s['verb']}</div>
            </div>
            """, unsafe_allow_html=True)

            # ── جدول التقييم الشهري ─────────────────────────────────────────
            if any(sc > 0 for _, _, sc, *_ in s["monthly_rep"]):
                st.markdown("#### 📅 نتيجة التقييم الشهري")
                monthly_table_data = []
                for n, short, score, ev_date, note, train in s["monthly_rep"]:
                    month_name = MONTHS_AR[n - 1]
                    late_cnt   = s["att_cnt_by_month"].get(n, 0)
                    late_hrs   = s["att_hrs_by_month"].get(n, 0.0)
                    monthly_table_data.append({
                        "الشهر":             month_name,
                        "الدرجة (%)":        f"{round(score * 100, 1)}%" if score > 0 else "—",
                        "التقييم اللفظي":    verbal_grade(score * 100) if score > 0 else "—",
                        "تاريخ التقييم":     ev_date if ev_date else "—",
                        "ملاحظات المقيم":    note    if note    else "—",
                        "عدد مرات التأخير":  str(int(late_cnt)) if late_cnt > 0 else "0",
                        "ساعات التأخير":     f"{late_hrs:.2f}",
                    })
                st.dataframe(
                    pd.DataFrame(monthly_table_data),
                    use_container_width=True,
                    hide_index=True,
                    column_config={
                        "الشهر":            st.column_config.TextColumn("الشهر",            width="small"),
                        "الدرجة (%)":       st.column_config.TextColumn("الدرجة",           width="small"),
                        "التقييم اللفظي":   st.column_config.TextColumn("التقييم",          width="medium"),
                        "تاريخ التقييم":    st.column_config.TextColumn("تاريخ التقييم",    width="medium"),
                        "ملاحظات المقيم":   st.column_config.TextColumn("ملاحظات المقيم",   width="large"),
                        "عدد مرات التأخير": st.column_config.TextColumn("عدد مرات التأخير", width="small"),
                        "ساعات التأخير":    st.column_config.TextColumn("ساعات التأخير",    width="small"),
                    }
                )
                st.markdown("---")

            # ── مؤشرات الأداء والصفات الشخصية ─────────────────────────────
            col_kpi, col_pers = st.columns(2)

            with col_kpi:
                st.subheader("🎯 مؤشرات الأداء الوظيفي")
                if s["job_kpis"]:
                    job_df = pd.DataFrame([{
                        "المؤشر":         k,
                        "الوزن (%)":      w,
                        "الدرجة (0-100)": round(kpi_score_to_pct(g, w), 1),
                        "التقييم":        rating_label(kpi_score_to_pct(g, w))
                    } for k, w, g in s["job_kpis"]])
                    st.dataframe(job_df, use_container_width=True, hide_index=True)
                else:
                    st.info("لا توجد مؤشرات أداء وظيفي")

            with col_pers:
                st.subheader("🌟 مؤشرات الصفات الشخصية")
                if s["pers_kpis"]:
                    pers_df = pd.DataFrame([{
                        "المؤشر":         k,
                        "الوزن (%)":      w,
                        "الدرجة (0-100)": round(kpi_score_to_pct(g, w), 1),
                        "التقييم":        rating_label(kpi_score_to_pct(g, w))
                    } for k, w, g in s["pers_kpis"]])
                    st.dataframe(pers_df, use_container_width=True, hide_index=True)
                else:
                    st.info("لا توجد مؤشرات صفات شخصية")

            st.markdown("---")

            # ── الإجراءات التأديبية ─────────────────────────────────────────
            if s["disciplinary_df"] is not None and not s["disciplinary_df"].empty:
                st.subheader("⚠️ الإجراءات التأديبية المسجلة")
                disc_display = s["disciplinary_df"].copy().rename(columns={
                    "action_date":    "التاريخ",
                    "warning_type":   "نوع الإنذار",
                    "reason":         "السبب",
                    "deduction_days": "خصم (أيام)"
                })
                cols_to_show   = ["التاريخ", "نوع الإنذار", "السبب", "خصم (أيام)"]
                available_cols = [c for c in cols_to_show if c in disc_display.columns]
                if available_cols:
                    st.dataframe(disc_display[available_cols], use_container_width=True, hide_index=True)
            else:
                st.info("✅ لا توجد إجراءات تأديبية مسجلة")

            # ── ملخص الالتزام بالدوام ───────────────────────────────────────
            st.subheader("⏰ الالتزام بالدوام")
            att_col1, att_col2 = st.columns(2)
            with att_col1:
                st.metric("📋 عدد مرات التأخير",    s["attendance_count"])
            with att_col2:
                st.metric("⏱️ إجمالي ساعات التأخير", f"{s['attendance_hours']:.2f}")

            # ── ملاحظات واحتياجات تدريبية ───────────────────────────────────
            cn, ct = st.columns(2)
            with cn:
                st.info(f"📝 **ملاحظات المقيم:** {s['emp_notes'] or '—'}")
            with ct:
                st.info(f"🎓 **الاحتياجات التدريبية:** {s['emp_train'] or '—'}")

    # ══════════════════════════════════════════════════════════════════════════
    # تصدير Excel
    # ══════════════════════════════════════════════════════════════════════════
    st.markdown("---")
    wb3 = openpyxl.Workbook()
    wb3.remove(wb3.active)

    for s in summary3:
        build_employee_sheet(
            wb3,
            s["emp"], s["job"], s["dept"], s["mgr"], sel3_year,
            s["kpis_export"],
            s["monthly_rep"],
            s["emp_notes"],
            s["emp_train"],
            employee_id          = s["emp_id"],
            disciplinary_actions = s["disciplinary_df"],
            attendance_data      = s["attendance_export"],
        )

    period_label = "، ".join(sel3_months) if sel3_months else "كل الأشهر"
    sum_title = (
        f"ملخص – {sel3_dept if sel3_dept != '-- الكل --' else 'الكل'}"
        f" – {sel3_year} – {period_label}"
    )
    build_summary_sheet(
        wb3,
        [
            (
                s["emp"],
                s["dept"],
                s["months"],
                s["pct"],
                s["verb"],
                len(s["disciplinary_df"]) if s["disciplinary_df"] is not None and not s["disciplinary_df"].empty else 0,
                s["attendance_count"],
                round(s["attendance_hours"], 2),
            )
            for s in summary3
        ],
        sum_title,
        year=sel3_year
    )
    wb3.move_sheet(wb3.worksheets[-1], offset=-(len(wb3.worksheets) - 1))

    buf3 = io.BytesIO()
    wb3.save(buf3)
    buf3.seek(0)

    d_label = sel3_dept.replace(" ", "_") if sel3_dept != "-- الكل --" else "الكل"
    st.download_button(
        label=f"📥 تحميل Excel ({len(summary3)} موظف)",
        data=buf3,
        file_name=f"تقارير_{d_label}_{date.today()}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

    st.markdown("---")
    st.markdown("#### 🖨️ معاينة وطباعة")
    html_prev3 = print_preview_html(io.BytesIO(buf3.getvalue()))
    st.components.v1.html(html_prev3, height=1400, scrolling=True)
