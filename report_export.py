"""
report_export.py
================
بناء تقارير Excel لنظام تقييم الأداء
"""

import io
import os
import base64
from datetime import date

import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.drawing.image import Image as XLImage
from openpyxl.worksheet.page import PageMargins
from openpyxl.utils import get_column_letter, column_index_from_string

from constants import (
    LOGO_PATH, MONTHS_AR, MONTHS_EN, MONTHS_SHORT, PERSONAL_KPIS,
    DARK, MID, LBLUE, ORANGE, YELLOW, LGRAY, GREEN_BG, RED_BG, WHITE, CREAM,
    OUTER_B, INNER_B,
)
from calculations import verbal_grade, kpi_score_to_pct, rating_label


# ═══════════════════════════════════════════════════════════════════
# الثوابت المشتركة
# ═══════════════════════════════════════════════════════════════════
_WARM     = "FFF3E0"
_NOTE_BG  = "FFFDE7"
_TRAIN_BG = "F3E5F5"
_INFO_BG  = "EBF3FB"
_DISC_BG  = "FEE2E2"
_ATT_BG   = "E0F2FE"

_med  = Side(style="medium", color="000000")
_thn  = Side(style="thin",   color="AAAAAA")
_BK   = Border(left=_med, right=_med, top=_med, bottom=_med)
_TN   = Border(left=_thn, right=_thn, top=_thn, bottom=_thn)

_MAR  = {
    "Jan":"يناير","Feb":"فبراير","Mar":"مارس","Apr":"أبريل",
    "May":"مايو","Jun":"يونيو","Jul":"يوليو","Aug":"أغسطس",
    "Sep":"سبتمبر","Oct":"أكتوبر","Nov":"نوفمبر","Dec":"ديسمبر",
}
_MONTHS_LIST = [
    "يناير","فبراير","مارس","أبريل","مايو","يونيو",
    "يوليو","أغسطس","سبتمبر","أكتوبر","نوفمبر","ديسمبر",
]


def _fill(hex_c):
    return PatternFill("solid", fgColor=str(hex_c).lstrip("#"))


def _font(bold=False, color="000000", size=9):
    return Font(name="Arial", bold=bold, color=str(color).lstrip("#"), size=size)


def _align(h="right", v="center", wrap=False):
    return Alignment(horizontal=h, vertical=v, wrapText=wrap, readingOrder=2)


def _sc(cell, val=None, bold=False, sz=9, color="000000",
        bg=None, ah="right", av="center", wrap=False, brd=None):
    if val is not None:
        cell.value = val
    cell.font      = _font(bold=bold, color=color, size=sz)
    cell.alignment = _align(h=ah, v=av, wrap=wrap)
    if bg:
        cell.fill = _fill(bg)
    if brd:
        cell.border = brd


def _mc(ws, r1, c1, r2, c2, val=None, **kw):
    ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)
    _sc(ws.cell(r1, c1, val), **kw)


def _company_header():
    _company, _branch = "مجموعة شركات فنون", ""
    try:
        from auth import load_app_settings as _las
        cfg = _las()
        _company = cfg.get("company_name", _company)
        _branch  = cfg.get("branch_name",  "")
    except Exception:
        pass
    return (f"نموذج تقييم الأداء السنوي — {_company}"
            + (f" — {_branch}" if _branch else ""))


def _add_logo(ws, anchor="A1", h=45, w=36):
    for _lp in [LOGO_PATH, "logo.png"]:
        if os.path.exists(_lp):
            try:
                img = XLImage(_lp)
                img.height, img.width = h, w
                img.anchor = anchor
                ws.add_image(img)
            except Exception:
                pass
            break


def _auto_fit_columns_range(ws, start_col, end_col, start_row=1, end_row=None, min_width=5, max_width=35):
    """
    ضبط عرض الأعمدة تلقائياً في نطاق محدد
    start_col, end_col: رقم العمود (1-indexed)
    """
    if end_row is None:
        end_row = ws.max_row
    
    for col in range(start_col, end_col + 1):
        max_len = 0
        col_letter = get_column_letter(col)
        
        # حساب أقصى طول للنص في هذا العمود
        for row in range(start_row, end_row + 1):
            cell = ws.cell(row, col)
            if cell.value:
                try:
                    text = str(cell.value)
                    text_len = len(text)
                    # النصوص العربية تحتاج عرض أكبر قليلاً
                    if any('\u0600' <= c <= '\u06FF' for c in text):
                        text_len = int(text_len * 1.2)
                    # النصوص الطويلة جداً (مثل الملاحظات) نحدها
                    if text_len > 50:
                        text_len = 35
                    max_len = max(max_len, text_len)
                except:
                    pass
        
        # تعيين العرض المناسب
        new_width = min(max(max_len + 2, min_width), max_width)
        if new_width > 0:
            ws.column_dimensions[col_letter].width = new_width


def _auto_fit_column_a(ws, start_row=1, end_row=None):
    """ضبط عرض العمود A فقط"""
    max_len = 0
    if end_row is None:
        end_row = ws.max_row
    for row in range(start_row, end_row + 1):
        cell = ws.cell(row, 1)
        if cell.value:
            try:
                max_len = max(max_len, len(str(cell.value)))
            except:
                pass
    ws.column_dimensions["A"].width = min(max(max_len * 0.7 + 2, 8), 35)


def _disc_by_month(disciplinary_actions):
    result = {}
    if disciplinary_actions is None:
        return result
    try:
        if getattr(disciplinary_actions, "empty", True):
            return result
        for _, row in disciplinary_actions.iterrows():
            dd = str(row.get("action_date", "") or "")
            if not dd:
                continue
            try:
                mn = int(dd.split("-")[1])
                result.setdefault(mn, []).append(
                    str(row.get("warning_type", "") or ""))
            except Exception:
                pass
    except Exception:
        pass
    return result


def _att_by_month(attendance_data):
    """
    تحويل بيانات التأخير إلى قاموس لكل شهر
    """
    result = {}
    result_hours = {}
    if attendance_data is None:
        return result, result_hours
    try:
        import pandas as _pd
        if isinstance(attendance_data, _pd.DataFrame):
            for _, row in attendance_data.iterrows():
                mn = row.get("month")
                if mn:
                    result[int(mn)] = int(row.get("late_count", 0) or 0)
                    result_hours[int(mn)] = float(row.get("late_hours", 0) or 0.0)
        elif isinstance(attendance_data, dict):
            mn = attendance_data.get("month")
            if mn:
                result[int(mn)] = int(attendance_data.get("late_count", 0) or 0)
                result_hours[int(mn)] = float(attendance_data.get("late_hours", 0) or 0.0)
        elif isinstance(attendance_data, list):
            for item in attendance_data:
                mn = item.get("month")
                if mn:
                    result[int(mn)] = int(item.get("late_count", 0) or 0)
                    result_hours[int(mn)] = float(item.get("late_hours", 0) or 0.0)
    except Exception:
        pass
    return result, result_hours


# ═══════════════════════════════════════════════════════════════════
# build_employee_sheet
# ═══════════════════════════════════════════════════════════════════
def build_employee_sheet(
    wb,
    emp_name, job_title, dept, manager, year,
    kpis,
    monthly_scores,
    notes="", training="",
    chart_img=None,
    disciplinary_actions=None,
    employee_id="",
    attendance_data=None,
):
    safe = emp_name[:28]
    if safe in [s.title for s in wb.worksheets]:
        safe = safe[:25] + "_2"
    ws = wb.create_sheet(safe)
    ws.sheet_view.rightToLeft = True
    ws.sheet_view.showGridLines = False

    # معالجة البيانات
    m_train = {}
    for item in monthly_scores:
        ms = item[1]
        def _v(x): return str(x).strip() if x not in (None, "nan", "None", "—") else ""
        m_train[ms] = _v(item[5]) if len(item) > 5 else ""

    done = [(n, m, s) for n, m, s, *_ in monthly_scores if s > 0]
    pct = sum(s for _, _, s in done) / len(done) * 100 if done else 0
    verb = verbal_grade(pct)
    sc_c = "375623" if pct >= 80 else ("C00000" if pct < 60 else "7F6000")
    sbg = GREEN_BG if pct >= 80 else (YELLOW if pct >= 60 else RED_BG)

    job_kpis = [(k["KPI_Name"], k["Weight"], k.get("avg_score", 0))
                for k in kpis if k["KPI_Name"] not in PERSONAL_KPIS]
    per_kpis = [(k["KPI_Name"], k["Weight"], k.get("avg_score", 0))
                for k in kpis if k["KPI_Name"] in PERSONAL_KPIS]

    disc_map = _disc_by_month(disciplinary_actions)
    att_count_map, att_hours_map = _att_by_month(attendance_data)
    
    total_late_count = sum(att_count_map.values())
    total_late_hours = sum(att_hours_map.values())

    # ════════════════════════════════════════════════════════════
    # ROW 1 — ترويسة
    # ════════════════════════════════════════════════════════════
    ws.row_dimensions[1].height = 32
    _mc(ws, 1, 1, 1, 15, _company_header(),
        bold=True, sz=11, color="FFFFFF", bg=DARK, ah="center")
    _add_logo(ws, anchor="A1", h=45, w=36)

    # ════════════════════════════════════════════════════════════
    # ROWS 2-8 — معلومات الموظف (A-D)
    # ════════════════════════════════════════════════════════════
    INFO = [
        ("اسم الموظف", emp_name),
        ("رقم الموظف", employee_id),
        ("الوظيفة", job_title),
        ("القسم", dept),
        ("السنة", str(year)),
        ("اسم المقيم", manager),
        ("تاريخ التقييم", date.today().strftime("%d/%m/%Y")),
    ]
    for i, (lbl, val) in enumerate(INFO):
        rr = 2 + i
        ws.row_dimensions[rr].height = 16
        _sc(ws.cell(rr, 1, lbl), bold=True, color="FFFFFF", bg=DARK, ah="center", brd=_TN)
        ws.merge_cells(start_row=rr, start_column=2, end_row=rr, end_column=4)
        _sc(ws.cell(rr, 2, val), color="000000", bg=_INFO_BG, ah="right", brd=_TN)

    # ════════════════════════════════════════════════════════════
    # ROW 3 — عنوان الجدول الشهري
    # ════════════════════════════════════════════════════════════
    ws.row_dimensions[3].height = 18
    _mc(ws, 3, 7, 3, 14, "نتيجة التقييم الشهري",
        bold=True, sz=10, color="FFFFFF", bg=DARK, ah="center", brd=_TN)

    # ════════════════════════════════════════════════════════════
    # ROW 4 — رؤوس أعمدة الجدول الشهري
    # ════════════════════════════════════════════════════════════
    ws.row_dimensions[4].height = 15
    mth_hdrs = [
        "الشهر", "الدرجة (%)", "التقييم اللفظي",
        "تاريخ التقييم", "ملاحظات المقيم",
        "الإجراءات التأديبية", "عدد مرات التأخير", "ساعات التأخير"
    ]
    for ci, h in enumerate(mth_hdrs, 7):
        _sc(ws.cell(4, ci, h), bold=True, sz=8, color="FFFFFF", bg=MID,
            ah="center", brd=_TN)

    # ════════════════════════════════════════════════════════════
    # ROWS 5-16 — بيانات الأشهر الـ12 (تبدأ من الصف 5)
    # ════════════════════════════════════════════════════════════
    for month_idx, month_name in enumerate(_MONTHS_LIST, 1):
        mr = 4 + month_idx
        ws.row_dimensions[mr].height = 15
        rbg = LGRAY if month_idx % 2 == 0 else WHITE

        month_data = None
        for item in monthly_scores:
            short = item[1]
            mn_num = (list(_MAR.keys()).index(short) + 1) if short in _MAR else 0
            if mn_num == month_idx:
                month_data = item
                break

        if month_data and month_data[2] > 0:
            score = month_data[2]
            eval_date = str(month_data[3]) if len(month_data) > 3 else ""
            note = str(month_data[4]) if len(month_data) > 4 else ""
            score_pct = f"{round(score * 100, 1)}%"
            verbal_val = verbal_grade(score * 100)
            if eval_date in ("None", "nan", ""):
                eval_date = "—"
            if note in ("None", "nan", ""):
                note = "—"
        else:
            score_pct = verbal_val = eval_date = note = "—"

        disc_text = ("، ".join(set(disc_map[month_idx]))
                     if month_idx in disc_map else "—")
        late_count = att_count_map.get(month_idx, 0)
        late_hours = att_hours_map.get(month_idx, 0.0)
        late_count_txt = str(late_count) if late_count > 0 else "—"
        late_hours_txt = f"{late_hours:.2f}" if late_hours > 0 else "—"

        if score_pct != "—":
            sv = float(score_pct.replace("%", ""))
            sbg2 = GREEN_BG if sv >= 80 else (YELLOW if sv >= 60 else RED_BG)
        else:
            sbg2 = rbg

        _sc(ws.cell(mr, 7, month_name), bg=rbg, ah="center", brd=_TN)
        _sc(ws.cell(mr, 8, score_pct), bg=sbg2, ah="center", bold=(score_pct != "—"), brd=_TN)
        _sc(ws.cell(mr, 9, verbal_val), bg=sbg2, ah="center", brd=_TN)
        _sc(ws.cell(mr, 10, eval_date), bg=rbg, ah="center", sz=8, brd=_TN)
        _sc(ws.cell(mr, 11, note), bg=rbg, ah="right", wrap=True, sz=8, brd=_TN)
        _sc(ws.cell(mr, 12, disc_text), bg=(_DISC_BG if disc_text != "—" else rbg),
            ah="center", sz=8, brd=_TN)
        _sc(ws.cell(mr, 13, late_count_txt), bg=(_ATT_BG if late_count > 0 else rbg),
            ah="center", sz=8, brd=_TN)
        _sc(ws.cell(mr, 14, late_hours_txt), bg=(_ATT_BG if late_hours > 0 else rbg),
            ah="center", sz=8, brd=_TN)

    # ════════════════════════════════════════════════════════════
    # ROW 17 — إجمالي الإجراءات والتأخير
    # ════════════════════════════════════════════════════════════
    ws.row_dimensions[17].height = 15
    total_disc = sum(len(v) for v in disc_map.values())
    _mc(ws, 17, 7, 17, 11, "الإجماليات السنوية",
        bold=True, color="FFFFFF", bg=DARK, ah="center", brd=_TN)
    _sc(ws.cell(17, 12,
                f"إجمالي الإجراءات: {total_disc}" if total_disc else "لا توجد إجراءات"),
        bold=True, bg=_DISC_BG, ah="center", sz=8, brd=_TN)
    _sc(ws.cell(17, 13,
                f"إجمالي مرات التأخير: {total_late_count}" if total_late_count > 0 else "لا تأخير"),
        bold=True, bg=_ATT_BG, ah="center", sz=8, brd=_TN)
    _sc(ws.cell(17, 14,
                f"إجمالي ساعات التأخير: {total_late_hours:.2f}" if total_late_hours > 0 else "لا تأخير"),
        bold=True, bg=_ATT_BG, ah="center", sz=8, brd=_TN)

    # ════════════════════════════════════════════════════════════
    # ROW 10 — نتيجة التقييم السنوي
    # ════════════════════════════════════════════════════════════
    ws.row_dimensions[10].height = 17
    _mc(ws, 10, 1, 10, 2, "نتيجة التقييم السنوي",
        bold=True, color="FFFFFF", bg=ORANGE, ah="center", brd=_TN)
    _mc(ws, 10, 3, 10, 4, f"{int(round(pct))}% — {verb}",
        bold=True, sz=10, color=sc_c, bg=sbg, ah="center", brd=_TN)

    # ════════════════════════════════════════════════════════════
    # KPI SECTION (يبدأ من ROW 11)
    # ════════════════════════════════════════════════════════════
    r = 11

    # ── مؤشرات الأداء الوظيفي ──
    ws.row_dimensions[r].height = 15
    _sc(ws.cell(r, 1, "مؤشرات الأداء الوظيفي"), bold=True, color="FFFFFF", bg=DARK, ah="right", brd=_TN)
    _sc(ws.cell(r, 2, "الوزن %"), bold=True, color="FFFFFF", bg=DARK, ah="center", brd=_TN)
    _sc(ws.cell(r, 3, "الدرجة (0-100)"), bold=True, color="FFFFFF", bg=DARK, ah="center", brd=_TN)
    _sc(ws.cell(r, 4, "التقييم"), bold=True, color="FFFFFF", bg=DARK, ah="center", brd=_TN)
    r += 1

    job_total_score, job_total_weight = 0.0, 0.0
    for i, (kname, weight, grade) in enumerate(job_kpis):
        rbg = LGRAY if i % 2 == 0 else WHITE
        w, g = float(weight), float(grade)
        pct_val = round(kpi_score_to_pct(g, w), 1)
        lbl = rating_label(pct_val)
        job_total_score += g
        job_total_weight += w
        kbg = (GREEN_BG if pct_val >= 80
               else (YELLOW if pct_val >= 60
               else (RED_BG if pct_val > 0 else rbg)))
        ws.row_dimensions[r].height = 14
        _sc(ws.cell(r, 1, kname), bg=rbg, wrap=True, sz=8, brd=_TN)
        _sc(ws.cell(r, 2, f"{w:.1f}%"), bg=rbg, ah="center", sz=8, brd=_TN)
        _sc(ws.cell(r, 3, pct_val), bold=True, bg=kbg, ah="center", sz=8, brd=_TN)
        _sc(ws.cell(r, 4, lbl), bold=True, bg=kbg, ah="center", sz=8, brd=_TN)
        r += 1

    ws.row_dimensions[r].height = 14
    jp = round(kpi_score_to_pct(job_total_score, job_total_weight), 1) if job_total_weight > 0 else 0
    _sc(ws.cell(r, 1, "مجموع الأداء الوظيفي"), bold=True, color="FFFFFF", bg=MID, ah="right", brd=_TN)
    _sc(ws.cell(r, 2, f"{job_total_weight:.1f}%"), bold=True, color="FFFFFF", bg=MID, ah="center", brd=_TN)
    _sc(ws.cell(r, 3, f"{jp}%"), bold=True, color="FFFFFF", bg=MID, ah="center", brd=_TN)
    _sc(ws.cell(r, 4, rating_label(jp)), bold=True, color="FFFFFF", bg=MID, ah="center", brd=_TN)
    r += 2

    # ── مؤشرات الصفات الشخصية ──
    ws.row_dimensions[r].height = 14
    _mc(ws, r, 1, r, 4, "مؤشرات الصفات الشخصية",
        bold=True, color="FFFFFF", bg=ORANGE, ah="center", brd=_TN)
    r += 1
    ws.row_dimensions[r].height = 14
    _sc(ws.cell(r, 1, "المؤشر"), bold=True, color="FFFFFF", bg=MID, ah="right", brd=_TN)
    _sc(ws.cell(r, 2, "الوزن %"), bold=True, color="FFFFFF", bg=MID, ah="center", brd=_TN)
    _sc(ws.cell(r, 3, "الدرجة (0-100)"), bold=True, color="FFFFFF", bg=MID, ah="center", brd=_TN)
    _sc(ws.cell(r, 4, "التقييم"), bold=True, color="FFFFFF", bg=MID, ah="center", brd=_TN)
    r += 1

    per_total_score, per_total_weight = 0.0, 0.0
    for i, (kname, weight, grade) in enumerate(per_kpis):
        rbg = _WARM if i % 2 == 0 else WHITE
        w, g = float(weight), float(grade)
        pct_val = round(kpi_score_to_pct(g, w), 1)
        lbl = rating_label(pct_val)
        per_total_score += g
        per_total_weight += w
        kbg = (GREEN_BG if pct_val >= 80
               else (YELLOW if pct_val >= 60
               else (RED_BG if pct_val > 0 else rbg)))
        ws.row_dimensions[r].height = 14
        _sc(ws.cell(r, 1, kname), bg=rbg, wrap=True, sz=8, brd=_TN)
        _sc(ws.cell(r, 2, f"{w:.1f}%"), bg=rbg, ah="center", sz=8, brd=_TN)
        _sc(ws.cell(r, 3, pct_val), bold=True, bg=kbg, ah="center", sz=8, brd=_TN)
        _sc(ws.cell(r, 4, lbl), bold=True, bg=kbg, ah="center", sz=8, brd=_TN)
        r += 1

    ws.row_dimensions[r].height = 14
    pp = round(kpi_score_to_pct(per_total_score, per_total_weight), 1) if per_total_weight > 0 else 0
    _sc(ws.cell(r, 1, "مجموع الصفات الشخصية"), bold=True, color="FFFFFF", bg=ORANGE, ah="right", brd=_TN)
    _sc(ws.cell(r, 2, f"{per_total_weight:.1f}%"), bold=True, color="FFFFFF", bg=ORANGE, ah="center", brd=_TN)
    _sc(ws.cell(r, 3, f"{pp}%"), bold=True, color="FFFFFF", bg=ORANGE, ah="center", brd=_TN)
    _sc(ws.cell(r, 4, rating_label(pp)), bold=True, color="FFFFFF", bg=ORANGE, ah="center", brd=_TN)
    r += 2

    # ════════════════════════════════════════════════════════════
    # الإجراءات التأديبية المسجلة
    # ════════════════════════════════════════════════════════════
    if disc_map:
        ws.row_dimensions[r].height = 14
        _mc(ws, r, 1, r, 4, "⚠️ الإجراءات التأديبية المسجلة",
            bold=True, color="FFFFFF", bg="C00000", ah="right", brd=_TN)
        r += 1
        for mn, actions in sorted(disc_map.items()):
            month_name_d = _MONTHS_LIST[mn - 1] if 1 <= mn <= 12 else str(mn)
            for act in actions:
                ws.row_dimensions[r].height = 13
                _sc(ws.cell(r, 1, month_name_d), bg=_DISC_BG, ah="center", sz=8, brd=_TN)
                ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=4)
                _sc(ws.cell(r, 2, act or "إجراء تأديبي"),
                    bg=_DISC_BG, ah="right", sz=8, brd=_TN)
                r += 1
        r += 1
    else:
        ws.row_dimensions[r].height = 14
        _mc(ws, r, 1, r, 4, "⚠️ الإجراءات التأديبية المسجلة",
            bold=True, color="FFFFFF", bg="C00000", ah="right", brd=_TN)
        r += 1
        ws.row_dimensions[r].height = 13
        _mc(ws, r, 1, r, 4, "لا توجد إجراءات تأديبية مسجلة",
            bg=_DISC_BG, ah="center", sz=8, brd=_TN)
        r += 1
        r += 1

    # ════════════════════════════════════════════════════════════
    # الالتزام بالدوام
    # ════════════════════════════════════════════════════════════
    ws.row_dimensions[r].height = 14
    _mc(ws, r, 1, r, 4, "⏰ الالتزام بالدوام",
        bold=True, color="FFFFFF", bg="1F3864", ah="right", brd=_TN)
    r += 1
    
    ws.row_dimensions[r].height = 13
    _sc(ws.cell(r, 1, "إجمالي عدد مرات التأخير"), bold=True, bg=_ATT_BG, ah="right", sz=8, brd=_TN)
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=4)
    _sc(ws.cell(r, 2, str(total_late_count)), bg=_ATT_BG, ah="right", sz=8, brd=_TN)
    r += 1
    
    ws.row_dimensions[r].height = 13
    _sc(ws.cell(r, 1, "إجمالي ساعات التأخير"), bold=True, bg=_ATT_BG, ah="right", sz=8, brd=_TN)
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=4)
    _sc(ws.cell(r, 2, f"{total_late_hours:.2f}"), bg=_ATT_BG, ah="right", sz=8, brd=_TN)
    r += 2

    # ════════════════════════════════════════════════════════════
    # ملاحظات المقيم والاحتياجات التدريبية
    # ════════════════════════════════════════════════════════════
    ws.row_dimensions[r].height = 20
    _train_vals = [v for v in m_train.values()
                   if v and str(v).strip() not in ("", "nan", "None", "—")]
    _train = _train_vals[0] if _train_vals else (training or "")
    _mc(ws, r, 1, r, 4, f"ملاحظات المقيم: {notes or '—'}",
        bg=_NOTE_BG, wrap=True, brd=_TN)
    r += 1
    _mc(ws, r, 1, r, 4,
        f"الاحتياجات التدريبية: {_train}" if _train else "الاحتياجات التدريبية: —",
        bg=_TRAIN_BG, wrap=True, brd=_TN)
    r += 2

    # ════════════════════════════════════════════════════════════
    # التوقيع
    # ════════════════════════════════════════════════════════════
    ws.row_dimensions[r].height = 16
    _sc(ws.cell(r, 1, f"المسؤول المباشر: {manager}"),
        bold=True, ah="center", brd=_BK)
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=4)
    _sc(ws.cell(r, 2, f"اسم الموظف: {emp_name}"),
        bold=True, ah="right", brd=_BK)
    r += 1
    ws.row_dimensions[r].height = 16
    _sc(ws.cell(r, 1, "التوقيع: _______________"),
        bold=True, bg=LGRAY, ah="center", brd=_BK)
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=4)
    _sc(ws.cell(r, 2, "التوقيع: _______________"),
        bold=True, bg=LGRAY, ah="center", brd=_BK)

    # ════════════════════════════════════════════════════════════
    # ضبط عرض الأعمدة تلقائياً (Auto Width) - فقط للأعمدة G إلى O (7 إلى 15)
    # بدون المساس ببقية الأعمدة أو الهيكل
    # ════════════════════════════════════════════════════════════
    
    # 1. تثبيت عرض العمود A (لأنه يحتوي على عناوين أطول)
    _auto_fit_column_a(ws, start_row=1, end_row=ws.max_row)
    
    # 2. ضبط عرض الأعمدة من G (7) إلى O (15) تلقائياً حسب المحتوى
    # هذه هي الأعمدة التي نريد ضبطها:
    # G = الشهر (عمود 7)
    # H = الدرجة (%) (عمود 8)
    # I = التقييم اللفظي (عمود 9)
    # J = تاريخ التقييم (عمود 10)
    # K = ملاحظات المقيم (عمود 11)
    # L = الإجراءات التأديبية (عمود 12)
    # M = عدد مرات التأخير (عمود 13)
    # N = ساعات التأخير (عمود 14)
    # O = عمود فارغ أو إضافي (عمود 15)
    
    _auto_fit_columns_range(ws, start_col=7, end_col=15, 
                            start_row=1, end_row=ws.max_row,
                            min_width=6, max_width=30)
    
    # 3. معالجة خاصة لعمود الملاحظات (K - عمود 11) لأنه قد يحتوي على نصوص طويلة
    # نعطيه عرض أكبر قليلاً
    if ws.column_dimensions["K"].width < 22:
        ws.column_dimensions["K"].width = 22

    # إعداد الطباعة
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = 9
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUp
