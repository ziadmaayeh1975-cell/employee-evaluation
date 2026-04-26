import io
from datetime import date
import openpyxl
import pandas as pd
import streamlit as st
from constants import MONTHS_EN
from calculations import calc_monthly, calc_yearly, verbal_grade
from auth import get_current_reviewer, get_current_role
from report_export import build_summary_sheet, print_preview_html

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
    for col in ["EmployeeName","Month","Year","KPI_%"]:
        if col not in df.columns:
            df[col] = pd.Series(dtype="object")
    return df


def render_yearly_summary(df_emp, df_kpi, df_data):
    st.subheader("🏆 الملخص السنوي – جميع الموظفين")

    df_data = _safe_df(df_data)

    allowed_emps = _reviewer_emp_set(df_emp)  # None = الكل

    all_evaled = [
        str(e) for e in df_data["EmployeeName"].dropna().unique().tolist()
        if str(e).strip() not in ("","nan")
        and (allowed_emps is None or str(e).strip() in allowed_emps)
    ]

    if not all_evaled:
        st.info("ℹ️ لا توجد تقييمات محفوظة لموظفيك حتى الآن.")
        return

    sel4_year = st.selectbox("🗓️ السنة", [2025, 2026, 2027], key="sum_year")

    rows4 = []
    for emp in all_evaled:
        ei4   = df_emp[df_emp["EmployeeName"] == emp]
        d4    = str(ei4.iloc[0, 2]).strip() if not ei4.empty else "—"
        avg4  = calc_yearly(df_data, emp, sel4_year)
        done4 = len([m for m in MONTHS_EN if calc_monthly(df_data, emp, m, sel4_year) > 0])
        rows4.append({
            "الموظف":     emp,
            "القسم":      d4,
            "الأشهر":     done4,
            "المعدل (%)": round(avg4*100, 1),
            "التقييم":    verbal_grade(avg4*100) if avg4 > 0 else "—",
        })
    rows4.sort(key=lambda x: x["المعدل (%)"], reverse=True)

    st.dataframe(pd.DataFrame(rows4), hide_index=True, use_container_width=True)

    wb4 = openpyxl.Workbook()
    wb4.remove(wb4.active)
    build_summary_sheet(
        wb4,
        [(r["الموظف"], r["القسم"], r["الأشهر"], r["المعدل (%)"], r["التقييم"]) for r in rows4],
        f"الملخص السنوي {sel4_year}",
        year=sel4_year
    )
    buf4 = io.BytesIO(); wb4.save(buf4); buf4.seek(0)
    st.download_button(
        label="📥 تحميل Excel",
        data=buf4,
        file_name=f"الملخص_السنوي_{date.today()}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
    st.markdown("---")
    st.markdown("#### 🖨️ معاينة وطباعة")
    html_prev4 = print_preview_html(io.BytesIO(buf4.getvalue()))
    st.components.v1.html(html_prev4, height=800, scrolling=True)
