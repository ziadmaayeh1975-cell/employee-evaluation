import io
from datetime import date
import openpyxl
import pandas as pd
import streamlit as st
from constants import MONTHS_AR, MONTHS_EN, MONTHS_SHORT, MONTH_MAP, PERSONAL_KPIS
from calculations import calc_monthly, get_kpi_avgs, verbal_grade
from data_loader import get_emp_notes
from auth import get_current_reviewer, get_current_role
from report_export import build_employee_sheet, build_summary_sheet, print_preview_html

# استيراد نظام الإجراءات التأديبية
try:
    from disciplinary_manager import get_actions_by_employee
    DISCIPLINARY_AVAILABLE = True
except ImportError:
    DISCIPLINARY_AVAILABLE = False

# استيراد نظام الالتزام بالدوام
try:
    from attendance_manager import get_employee_attendance_summary
    ATTENDANCE_AVAILABLE = True
except ImportError:
    ATTENDANCE_AVAILABLE = False


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
        if str(e).strip() not in ("","nan")
    )


def _reviewer_emp_list(df_emp):
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
    for col in ["EmployeeName","Month","Year","KPI_%","KPI_Name"]:
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
    for col in ["EvalDate","eval_date","EntryDate"]:
        if col in subset.columns and pd.notna(row.get(col, None)):
            val = str(row[col]).strip()
            if val and val not in ("","nan","None"):
                try: eval_date = pd.to_datetime(val).strftime("%d/%m/%Y")
                except: eval_date = val
                break
    notes = ""
    for col in ["Notes","notes"]:
        if col in subset.columns and pd.notna(row.get(col, None)):
            val = str(row[col]).strip()
            if val and val not in ("","nan","None"):
                notes = val; break
    training = ""
    for col in ["Training","training"]:
        if col in subset.columns and pd.notna(row.get(col, None)):
            val = str(row[col]).strip()
            if val and val not in ("","nan","None"):
                training = val; break
    return eval_date, notes, training


def render_department_report(df_emp, df_kpi, df_data):
    st.subheader("📊 التقرير المفصّل بالأقسام")
    df_data = _safe_df(df_data)
    allowed_emps = _reviewer_emp_set(df_emp)

    c3a, c3b, c3c = st.columns([2, 2, 1])
    with c3a:
        dept_col = df_emp.columns[2]
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
        dept_emps = dept_emps_all[dept_emps_all[dept_col].astype(str).str.strip() == sel3_dept]["EmployeeName"].tolist()

    if df_data.empty or "EmployeeName" not in df_data.columns:
        st.info("ℹ️ لا توجد تقييمات محفوظة حتى الآن.")
        return

    dept_with_data = [e for e in dept_emps if e in df_data["EmployeeName"].values]
    if not dept_with_data:
        st.info("لا توجد بيانات تقييم لهذا القسم بعد.")
        return

    st.markdown(f"**{len(dept_with_data)} موظف** لديهم بيانات تقييم.")
    months_en_f3 = [MONTH_MAP[m] for m in sel3_months] if sel3_months else None

    summary3 = []
    for emp in dept_with_data:
        ei3 = df_emp[df_emp["EmployeeName"] == emp]
        d3 = str(ei3.iloc[0, 2]).strip() if not ei3.empty else "—"
        emp_id = str(ei3.iloc[0].get("رقم الموظف", "")) if "رقم الموظف" in ei3.columns else ""
        s_list = [calc_monthly(df_data, emp, m, sel3_year) for m in (months_en_f3 or MONTHS_EN)]
        active3 = [s for s in s_list if s > 0]
        avg3 = sum(active3) / len(active3) if active3 else 0.0
        
        # جلب بيانات الالتزام بالدوام للموظف
        attendance_monthly = {}
        attendance_count = 0
        attendance_hours = 0.0
        if ATTENDANCE_AVAILABLE:
            try:
                for month_num in range(1, 13):
                    att_summary = get_employee_attendance_summary(emp, emp_id, sel3_year, month_num)
                    att_count = att_summary.get("count", 0) or 0
                    att_hrs = att_summary.get("hours", 0) or 0.0
                    attendance_monthly[month_num] = {"count": att_count, "hours": att_hrs}
                    attendance_count += att_count
                    attendance_hours += att_hrs
            except:
                pass
        
        # جلب الإجراءات التأديبية للموظف
        disciplinary_df = None
        if DISCIPLINARY_AVAILABLE:
            try:
                disc_actions_list = get_actions_by_employee(emp, sel3_year)
                if disc_actions_list:
                    disciplinary_df = pd.DataFrame(disc_actions_list)
            except Exception as e:
                pass
        
        summary3.append({
            "emp": emp, 
            "dept": d3,
            "months": len(active3),
            "pct": round(avg3 * 100, 1),
            "verb": verbal_grade(avg3 * 100) if active3 else "—",
            "emp_id": emp_id,
            "attendance_count": attendance_count,
            "attendance_hours": attendance_hours,
            "attendance_monthly": attendance_monthly,
            "disciplinary_df": disciplinary_df
        })
    summary3.sort(key=lambda x: x["pct"], reverse=True)

    # عرض جدول الملخص
    display_df = pd.DataFrame([{
        "الموظف": s["emp"], 
        "القسم": s["dept"], 
        "السنة": sel3_year,
        "الأشهر": s["months"], 
        "المعدل (%)": s["pct"], 
        "التقييم": s["verb"],
        "عدد مرات التأخير": s["attendance_count"],
        "ساعات التأخير": f"{s['attendance_hours']:.2f}"
    } for s in summary3])
    st.dataframe(display_df, hide_index=True, use_container_width=True)

    wb3 = openpyxl.Workbook()
    wb3.remove(wb3.active)
    
    for s in summary3:
        ei3 = df_emp[df_emp["EmployeeName"] == s["emp"]]
        if ei3.empty: 
            continue
        ei3 = ei3.iloc[0]
        job3 = str(ei3.iloc[1]).strip()
        d3 = str(ei3.iloc[2]).strip()
        m3 = str(ei3.iloc[3]).strip()
        
        # ✅ جلب مؤشرات الأداء (مع التأكد من تمرير months_en_f3)
        kpis3 = get_kpi_avgs(df_data, df_kpi, s["emp"], job3, months_en_f3, sel3_year)
        
        # ✅ للتصحيح: طباعة عدد المؤشرات التي تم جلبها
        if kpis3:
            print(f"✅ تم جلب {len(kpis3)} مؤشر للموظف {s['emp']}")
        else:
            print(f"⚠️ لم يتم جلب أي مؤشر للموظف {s['emp']}")
        
        # إعداد البيانات الشهرية
        ms3 = []
        for idx, (en, short) in enumerate(zip(MONTHS_EN, MONTHS_SHORT)):
            if months_en_f3 and en not in months_en_f3:
                ms3.append((idx+1, short, 0.0, "", "", ""))
            else:
                score = calc_monthly(df_data, s["emp"], en, sel3_year)
                ev, nm, tr = _get_month_details(df_data, s["emp"], en, sel3_year)
                ms3.append((idx+1, short, score, ev, nm, tr))
        
        # جلب ملاحظات وتدريب الموظف
        emp_notes = ""; emp_train = ""
        for item in ms3:
            if item[2] > 0:
                emp_notes = item[4] or ""; emp_train = item[5] or ""; break
        if not emp_notes and not emp_train:
            emp_notes, emp_train = get_emp_notes(s["emp"])
        
        # تحضير بيانات الالتزام بالدوام للتصدير
        attendance_export = None
        if s["attendance_count"] > 0 or s["attendance_hours"] > 0:
            monthly_list = []
            for month_num, att_data in s.get("attendance_monthly", {}).items():
                if att_data["count"] > 0 or att_data["hours"] > 0:
                    monthly_list.append({
                        "month": month_num,
                        "late_count": att_data["count"],
                        "late_hours": att_data["hours"]
                    })
            if monthly_list:
                attendance_export = pd.DataFrame(monthly_list)
            else:
                attendance_export = pd.DataFrame([{
                    "employee_name": s["emp"],
                    "employee_id": s["emp_id"],
                    "year": sel3_year,
                    "late_count": s["attendance_count"],
                    "total_late_hours": s["attendance_hours"]
                }])
        
        # ✅ بناء شيت الموظف مع تمرير جميع البيانات
        build_employee_sheet(
            wb3, 
            s["emp"], 
            job3, 
            d3, 
            m3, 
            sel3_year,
            kpis3,  # ✅ مؤشرات الأداء (الوظيفي + الشخصي)
            ms3,    # ✅ البيانات الشهرية
            emp_notes, 
            emp_train,
            employee_id=s["emp_id"],
            attendance_data=attendance_export,
            disciplinary_actions=s.get("disciplinary_df")
        )

    period_label = "، ".join(sel3_months) if sel3_months else "كل الأشهر"
    sum_title = f"ملخص – {sel3_dept if sel3_dept != '-- الكل --' else 'الكل'} – {sel3_year} – {period_label}"
    build_summary_sheet(wb3, [(s["emp"], s["dept"], s["months"], s["pct"], s["verb"]) for s in summary3], sum_title, year=sel3_year)
    wb3.move_sheet(wb3.worksheets[-1], offset=-(len(wb3.worksheets)-1))

    buf3 = io.BytesIO()
    wb3.save(buf3); buf3.seek(0)
    d_label = sel3_dept.replace(" ","_") if sel3_dept != "-- الكل --" else "الكل"
    st.download_button(
        label=f"📥 تحميل Excel ({len(summary3)} موظف)",
        data=buf3, file_name=f"تقارير_{d_label}_{date.today()}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", 
        use_container_width=True,
    )
    st.markdown("---")
    st.markdown("#### 🖨️ معاينة وطباعة")
    html_prev3 = print_preview_html(io.BytesIO(buf3.getvalue()))
    st.components.v1.html(html_prev3, height=1400, scrolling=True)
