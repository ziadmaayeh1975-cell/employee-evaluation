"""
report_export.py
================
بناء تقارير Excel لنظام تقييم الأداء
- build_employee_sheet  : تقرير الموظف الفردي
- build_summary_sheet   : ملخص الأقسام
- print_preview_html    : معاينة HTML للطباعة
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


def _add_logo(ws, anchor="A1", h=60, w=48):
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


def _disc_by_month(disciplinary_actions):
    """إرجاع dict {month_num: [warning_type, ...]}"""
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
    """إرجاع dict {month_num: late_count}"""
    result = {}
    if attendance_data is None:
        return result
    try:
        import pandas as _pd
        if isinstance(attendance_data, _pd.DataFrame):
            for _, row in attendance_data.iterrows():
                mn = row.get("month")
                if mn:
                    result[int(mn)] = int(row.get("late_count", 0) or 0)
        elif isinstance(attendance_data, dict):
            mn = attendance_data.get("month")
            if mn:
                result[int(mn)] = int(attendance_data.get("late_count", 0) or 0)
    except Exception:
        pass
    return result


# ═══════════════════════════════════════════════════════════════════
# build_employee_sheet
# ═══════════════════════════════════════════════════════════════════
def build_employee_sheet(
    wb,
    emp_name, job_title, dept, manager, year,
    kpis,           # list of dict {KPI_Name, Weight, avg_score}
    monthly_scores, # list of (emp, short, score [,eval_date, notes, training])
    notes="", training="",
    chart_img=None,
    disciplinary_actions=None,
    employee_id="",
    attendance_data=None,
):
    """
    التخطيط (Landscape A4 — صفحة واحدة مضمونة):
    ┌──────────────────────────────────────────────────────────────┐
    │  R1  : ترويسة كاملة A-N                                      │
    ├───────────────────┬──────────────────────────────────────────┤
    │  R2-8 : معلومات   │  R2: عنوان جدول الشهري  (G-N)           │
    │  الموظف (A-D)     │  R3: رؤوس أعمدة الجدول                   │
    │  R9: نتيجة سنوية  │  R4-15: 12 شهراً                         │
    ├───────────────────┤  R16-17: مجموع شهري + مرات التأخير       │
    │  R10+: KPIs       ├──────────────────────────────────────────┤
    │  الوظيفي          │  (فارغ — يتوافق مع ارتفاع KPIs)          │
    │  ثم الشخصي        │                                           │
    │  ثم ملاحظات       │                                           │
    │  ثم توقيع         │                                           │
    └───────────────────┴──────────────────────────────────────────┘
    """
    safe = emp_name[:28]
    if safe in [s.title for s in wb.worksheets]:
        safe = safe[:25] + "_2"
    ws = wb.create_sheet(safe)
    ws.sheet_view.rightToLeft  = True
    ws.sheet_view.showGridLines = False

    # ── عرض الأعمدة ─────────────────────────────────────────────
    # A-D  : بلوك اليسار (معلومات + KPIs)
    # E-F  : فاصل ضيق
    # G-N  : بلوك اليمين (جدول شهري)
    col_cfg = {
        "A": 5,  "B": 26, "C": 16, "D": 13,
        "E": 2,  "F": 2,
        "G": 10, "H": 11, "I": 14, "J": 14,
        "K": 22, "L": 14, "M": 14, "N": 4,
    }
    for col, w in col_cfg.items():
        ws.column_dimensions[col].width = w

    # ── pre-process ──────────────────────────────────────────────
    m_train = {}
    for item in monthly_scores:
        ms = item[1]
        def _v(x): return str(x).strip() if x not in (None,"nan","None","—") else ""
        m_train[ms] = _v(item[5]) if len(item) > 5 else ""

    done = [(n, m, s) for n, m, s, *_ in monthly_scores if s > 0]
    pct  = sum(s for _, _, s in done) / len(done) * 100 if done else 0
    verb = verbal_grade(pct)
    sc_c = "375623" if pct >= 80 else ("C00000" if pct < 60 else "7F6000")
    sbg  = GREEN_BG if pct >= 80 else (YELLOW if pct >= 60 else RED_BG)

    job_kpis = [(k["KPI_Name"], k["Weight"], k.get("avg_score", 0))
                for k in kpis if k["KPI_Name"] not in PERSONAL_KPIS]
    per_kpis = [(k["KPI_Name"], k["Weight"], k.get("avg_score", 0))
                for k in kpis if k["KPI_Name"] in PERSONAL_KPIS]

    disc_map = _disc_by_month(disciplinary_actions)
    att_map  = _att_by_month(attendance_data)

    # ════════════════════════════════════════════════════════════
    # ROW 1 — ترويسة كاملة
    # ════════════════════════════════════════════════════════════
    ws.row_dimensions[1].height = 28
    _mc(ws, 1, 1, 1, 14, _company_header(),
        bold=True, sz=11, color="FFFFFF", bg=DARK, ah="center")
    _add_logo(ws, anchor="N1", h=55, w=44)

    # ════════════════════════════════════════════════════════════
    # ROWS 2-8 — معلومات الموظف (A-D)  +  عنوان ورؤوس الجدول (G-N)
    # ════════════════════════════════════════════════════════════
    INFO = [
        ("اسم الموظف",    emp_name),
        ("رقم الموظف",    employee_id),
        ("الوظيفة",       job_title),
        ("القسم",         dept),
        ("السنة",         str(year)),
        ("اسم المقيم",    manager),
        ("تاريخ التقييم", date.today().strftime("%d/%m/%Y")),
    ]
    for i, (lbl, val) in enumerate(INFO):
        rr = 2 + i
        ws.row_dimensions[rr].height = 16
        _sc(ws.cell(rr, 1, lbl), bold=True, color="FFFFFF", bg=DARK, ah="center", brd=_TN)
        ws.merge_cells(start_row=rr, start_column=2, end_row=rr, end_column=4)
        _sc(ws.cell(rr, 2, val), color="000000", bg=_INFO_BG, ah="right", brd=_TN)

    # عنوان الجدول الشهري — يمتد صفين (2-3) على G-N
    ws.row_dimensions[2].height = 16
    _mc(ws, 2, 7, 3, 13, "نتيجة التقييم الشهري",
        bold=True, sz=10, color="FFFFFF", bg=DARK, ah="center")
    ws.row_dimensions[3].height = 15

    # ════════════════════════════════════════════════════════════
    # ROW 9 — نتيجة التقييم السنوي (A-D)
    # ════════════════════════════════════════════════════════════
    ws.row_dimensions[9].height = 17
    _mc(ws, 9, 1, 9, 2, "نتيجة التقييم السنوي",
        bold=True, color="FFFFFF", bg=ORANGE, ah="center", brd=_TN)
    _mc(ws, 9, 3, 9, 4, f"{int(round(pct))}% — {verb}",
        bold=True, sz=10, color=sc_c, bg=sbg, ah="center", brd=_TN)

    # ════════════════════════════════════════════════════════════
    # ROW 4 — رؤوس أعمدة الجدول الشهري (G-N)
    # ════════════════════════════════════════════════════════════
    ws.row_dimensions[4].height = 15
    mth_hdrs = [
        "الشهر", "الدرجة (%)", "التقييم",
        "تاريخ التقييم", "ملاحظات المقيم",
        "الإجراءات التأديبية", "مرات التأخير",
    ]
    for ci, h in enumerate(mth_hdrs, 7):
        _sc(ws.cell(4, ci, h), bold=True, sz=8, color="FFFFFF", bg=MID,
            ah="center", brd=_TN)

    # ════════════════════════════════════════════════════════════
    # ROWS 5-16 — بيانات الأشهر الـ12 (G-N)
    # ════════════════════════════════════════════════════════════
    for month_idx, month_name in enumerate(_MONTHS_LIST, 1):
        mr  = 4 + month_idx   # rows 5-16
        ws.row_dimensions[mr].height = 15
        rbg = LGRAY if month_idx % 2 == 0 else WHITE

        month_data = None
        for item in monthly_scores:
            short  = item[1]
            mn_num = (list(_MAR.keys()).index(short) + 1) if short in _MAR else 0
            if mn_num == month_idx:
                month_data = item
                break

        if month_data and month_data[2] > 0:
            score      = month_data[2]
            eval_date  = str(month_data[3]) if len(month_data) > 3 else ""
            note       = str(month_data[4]) if len(month_data) > 4 else ""
            score_pct  = f"{round(score * 100, 1)}%"
            verbal_val = verbal_grade(score * 100)
            if eval_date in ("None","nan",""): eval_date = "—"
            if note      in ("None","nan",""): note      = "—"
        else:
            score_pct = verbal_val = eval_date = note = "—"

        disc_text  = ("، ".join(set(disc_map[month_idx]))
                      if month_idx in disc_map else "—")
        late_count = att_map.get(month_idx, 0)
        late_txt   = str(late_count) if late_count > 0 else "—"

        # لون خلفية درجة التقييم
        if score_pct != "—":
            sv = float(score_pct.replace("%",""))
            sbg2 = GREEN_BG if sv >= 80 else (YELLOW if sv >= 60 else RED_BG)
        else:
            sbg2 = rbg

        _sc(ws.cell(mr,  7, month_name),  bg=rbg,  ah="center", brd=_TN)
        _sc(ws.cell(mr,  8, score_pct),   bg=sbg2, ah="center", bold=(score_pct!="—"), brd=_TN)
        _sc(ws.cell(mr,  9, verbal_val),  bg=sbg2, ah="center", brd=_TN)
        _sc(ws.cell(mr, 10, eval_date),   bg=rbg,  ah="center", sz=8, brd=_TN)
        _sc(ws.cell(mr, 11, note),        bg=rbg,  ah="right",  wrap=True, sz=8, brd=_TN)
        _sc(ws.cell(mr, 12, disc_text),   bg=(_DISC_BG if disc_text != "—" else rbg),
            ah="center", sz=8, brd=_TN)
        _sc(ws.cell(mr, 13, late_txt),    bg=(_ATT_BG if late_txt != "—" else rbg),
            ah="center", sz=8, brd=_TN)

    # ════════════════════════════════════════════════════════════
    # ROW 17 — إجمالي الإجراءات التأديبية والتأخير (G-N)
    # ════════════════════════════════════════════════════════════
    ws.row_dimensions[17].height = 15
    total_disc  = sum(len(v) for v in disc_map.values())
    total_late  = sum(att_map.values())
    _mc(ws, 17, 7, 17, 11, "الإجماليات السنوية",
        bold=True, color="FFFFFF", bg=DARK, ah="center", brd=_TN)
    _sc(ws.cell(17, 12,
                f"إجمالي الإجراءات: {total_disc}" if total_disc else "لا توجد إجراءات"),
        bold=True, bg=_DISC_BG, ah="center", sz=8, brd=_TN)
    _sc(ws.cell(17, 13,
                f"إجمالي التأخير: {total_late}" if total_late else "لا تأخير"),
        bold=True, bg=_ATT_BG, ah="center", sz=8, brd=_TN)

    # ════════════════════════════════════════════════════════════
    # KPI SECTION — يبدأ من ROW 10 (A-D)
    # ════════════════════════════════════════════════════════════
    r = 10

    # ── مؤشرات الأداء الوظيفي ──
    ws.row_dimensions[r].height = 15
    _sc(ws.cell(r, 1, "مؤشرات الأداء الوظيفي"), bold=True, color="FFFFFF", bg=DARK,   ah="right", brd=_TN)
    _sc(ws.cell(r, 2, "الوزن %"),               bold=True, color="FFFFFF", bg=DARK,   ah="center", brd=_TN)
    _sc(ws.cell(r, 3, "الدرجة (0-100)"),        bold=True, color="FFFFFF", bg=DARK,   ah="center", brd=_TN)
    _sc(ws.cell(r, 4, "التقييم"),               bold=True, color="FFFFFF", bg=DARK,   ah="center", brd=_TN)
    r += 1

    job_total_score, job_total_weight = 0.0, 0.0
    for i, (kname, weight, grade) in enumerate(job_kpis):
        rbg = LGRAY if i % 2 == 0 else WHITE
        w, g = float(weight), float(grade)
        pct_val = round(kpi_score_to_pct(g, w), 1)
        lbl = rating_label(pct_val)
        job_total_score  += g
        job_total_weight += w
        kbg = (GREEN_BG if pct_val >= 80
               else (YELLOW if pct_val >= 60
               else (RED_BG if pct_val > 0 else rbg)))
        ws.row_dimensions[r].height = 14
        _sc(ws.cell(r, 1, kname),       bg=rbg, wrap=True, sz=8, brd=_TN)
        _sc(ws.cell(r, 2, f"{w:.1f}%"), bg=rbg, ah="center", sz=8, brd=_TN)
        _sc(ws.cell(r, 3, pct_val),     bold=True, bg=kbg, ah="center", sz=8, brd=_TN)
        _sc(ws.cell(r, 4, lbl),         bold=True, bg=kbg, ah="center", sz=8, brd=_TN)
        r += 1

    ws.row_dimensions[r].height = 14
    jp = round(kpi_score_to_pct(job_total_score, job_total_weight), 1) \
         if job_total_weight > 0 else 0
    _sc(ws.cell(r, 1, "مجموع الأداء الوظيفي"), bold=True, color="FFFFFF", bg=MID, ah="right", brd=_TN)
    _sc(ws.cell(r, 2, f"{job_total_weight:.1f}%"), bold=True, color="FFFFFF", bg=MID, ah="center", brd=_TN)
    _sc(ws.cell(r, 3, f"{jp}%"),          bold=True, color="FFFFFF", bg=MID, ah="center", brd=_TN)
    _sc(ws.cell(r, 4, rating_label(jp)),   bold=True, color="FFFFFF", bg=MID, ah="center", brd=_TN)
    r += 2

    # ── مؤشرات الصفات الشخصية ──
    ws.row_dimensions[r].height = 14
    _mc(ws, r, 1, r, 4, "مؤشرات الصفات الشخصية",
        bold=True, color="FFFFFF", bg=ORANGE, ah="center", brd=_TN)
    r += 1
    ws.row_dimensions[r].height = 14
    _sc(ws.cell(r, 1, "المؤشر"),         bold=True, color="FFFFFF", bg=MID, ah="right",  brd=_TN)
    _sc(ws.cell(r, 2, "الوزن %"),        bold=True, color="FFFFFF", bg=MID, ah="center", brd=_TN)
    _sc(ws.cell(r, 3, "الدرجة (0-100)"), bold=True, color="FFFFFF", bg=MID, ah="center", brd=_TN)
    _sc(ws.cell(r, 4, "التقييم"),        bold=True, color="FFFFFF", bg=MID, ah="center", brd=_TN)
    r += 1

    per_total_score, per_total_weight = 0.0, 0.0
    for i, (kname, weight, grade) in enumerate(per_kpis):
        rbg = _WARM if i % 2 == 0 else WHITE
        w, g = float(weight), float(grade)
        pct_val = round(kpi_score_to_pct(g, w), 1)
        lbl = rating_label(pct_val)
        per_total_score  += g
        per_total_weight += w
        kbg = (GREEN_BG if pct_val >= 80
               else (YELLOW if pct_val >= 60
               else (RED_BG if pct_val > 0 else rbg)))
        ws.row_dimensions[r].height = 14
        _sc(ws.cell(r, 1, kname),       bg=rbg, wrap=True, sz=8, brd=_TN)
        _sc(ws.cell(r, 2, f"{w:.1f}%"), bg=rbg, ah="center", sz=8, brd=_TN)
        _sc(ws.cell(r, 3, pct_val),     bold=True, bg=kbg, ah="center", sz=8, brd=_TN)
        _sc(ws.cell(r, 4, lbl),         bold=True, bg=kbg, ah="center", sz=8, brd=_TN)
        r += 1

    ws.row_dimensions[r].height = 14
    pp = round(kpi_score_to_pct(per_total_score, per_total_weight), 1) \
         if per_total_weight > 0 else 0
    _sc(ws.cell(r, 1, "مجموع الصفات الشخصية"), bold=True, color="FFFFFF", bg=ORANGE, ah="right",  brd=_TN)
    _sc(ws.cell(r, 2, f"{per_total_weight:.1f}%"),  bold=True, color="FFFFFF", bg=ORANGE, ah="center", brd=_TN)
    _sc(ws.cell(r, 3, f"{pp}%"),          bold=True, color="FFFFFF", bg=ORANGE, ah="center", brd=_TN)
    _sc(ws.cell(r, 4, rating_label(pp)),   bold=True, color="FFFFFF", bg=ORANGE, ah="center", brd=_TN)
    r += 2

    # ── الإجراءات التأديبية المجمعة (A-D) ──
    if disc_map:
        ws.row_dimensions[r].height = 14
        _mc(ws, r, 1, r, 4, "⚠️ الإجراءات التأديبية المسجلة",
            bold=True, color="FFFFFF", bg=_DISC_BG.replace("FEE2E2","C00000"),
            ah="right", brd=_TN)
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

    # ── ملاحظات المقيم ──
    ws.row_dimensions[r].height = 20
    _train_vals = [v for v in m_train.values()
                   if v and str(v).strip() not in ("","nan","None","—")]
    _train = _train_vals[0] if _train_vals else (training or "")
    _mc(ws, r, 1, r, 4, f"ملاحظات المقيم: {notes or '—'}",
        bg=_NOTE_BG, wrap=True, brd=_TN)
    r += 1
    _mc(ws, r, 1, r, 4,
        f"الاحتياجات التدريبية: {_train}" if _train else "الاحتياجات التدريبية: —",
        bg=_TRAIN_BG, wrap=True, brd=_TN)
    r += 2

    # ── التوقيع ──
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

    # ── إعداد الطباعة ──
    ws.page_setup.orientation  = "landscape"
    ws.page_setup.paperSize    = 9          # A4
    ws.page_setup.fitToPage    = True
    ws.page_setup.fitToWidth   = 1
    ws.page_setup.fitToHeight  = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins = PageMargins(left=0.4, right=0.4, top=0.4, bottom=0.4)
    ws.print_options.horizontalCentered = True

    return ws


# ═══════════════════════════════════════════════════════════════════
# build_summary_sheet   — ملخص الأقسام مع الإجراءات التأديبية
# ═══════════════════════════════════════════════════════════════════
def build_summary_sheet(
    wb,
    rows,                       # list of (name, dept, months, pct, verb)
    title="ملخص التقييم",
    year=None,
    chart_img=None,
    disciplinary_summary=None,  # dict {emp_name: {count, types}}  — اختياري
    attendance_summary=None,    # dict {emp_name: late_count}       — اختياري
):
    ws = wb.create_sheet(title[:28])
    ws.sheet_view.rightToLeft  = True
    ws.sheet_view.showGridLines = False

    col_cfg = {
        "A": 4, "B": 30, "C": 14, "D": 8,
        "E": 10, "F": 12, "G": 12,
        "H": 14, "I": 14,
    }
    for col, w in col_cfg.items():
        ws.column_dimensions[col].width = w

    # Row 1 — ترويسة
    ws.row_dimensions[1].height = 32
    _mc(ws, 1, 1, 1, 9, title,
        bold=True, sz=12, color="FFFFFF", bg=DARK, ah="center", brd=_BK)
    _add_logo(ws, anchor="I1", h=30, w=24)

    # Row 2 — فراغ
    ws.row_dimensions[2].height = 4

    # Row 3 — رؤوس الأعمدة
    ws.row_dimensions[3].height = 16
    hdrs = ["#", "اسم الموظف", "القسم", "السنة",
            "الأشهر", "المعدل %", "التقييم",
            "الإجراءات التأديبية", "مرات التأخير"]
    for ci, h in enumerate(hdrs, 1):
        _sc(ws.cell(3, ci, h), bold=True, sz=9, color="FFFFFF",
            bg=DARK, ah="center", brd=_BK)

    disc_s = disciplinary_summary or {}
    att_s  = attendance_summary  or {}

    for i, row_data in enumerate(rows, 4):
        # دعم (name, dept, months, pct, verb)
        name, dept_, months, pct_val, verb_ = row_data[:5]
        ws.row_dimensions[i].height = 15
        rbg  = LGRAY if i % 2 == 0 else WHITE
        sc_c = "375623" if pct_val >= 80 else ("C00000" if pct_val < 60 else "000000")
        vbg  = (GREEN_BG if pct_val >= 80
                else (YELLOW if pct_val >= 70
                else (RED_BG if pct_val < 60 else LGRAY)))

        disc_info = disc_s.get(name, {})
        disc_cnt  = disc_info.get("count", 0) if isinstance(disc_info, dict) else int(disc_info or 0)
        late_cnt  = int(att_s.get(name, 0) or 0)

        disc_bg = _DISC_BG if disc_cnt > 0 else rbg
        att_bg  = _ATT_BG  if late_cnt  > 0 else rbg

        _sc(ws.cell(i, 1, i - 3),           sz=8, ah="center", bg=rbg, brd=INNER_B)
        _sc(ws.cell(i, 2, name),             sz=9, ah="right",  bg=rbg, brd=INNER_B)
        _sc(ws.cell(i, 3, dept_),            sz=8, ah="center", bg=rbg, brd=INNER_B)
        _sc(ws.cell(i, 4, year or ""),       sz=9, ah="center", bg=rbg, brd=INNER_B)
        _sc(ws.cell(i, 5, months),           sz=9, ah="center", bg=rbg, brd=INNER_B)
        _sc(ws.cell(i, 6, f"{pct_val:.1f}%"),
            sz=10, bold=True, color=sc_c, ah="center", bg=vbg, brd=INNER_B)
        _sc(ws.cell(i, 7, verb_),
            sz=9, bold=True, color=sc_c, ah="center", bg=vbg, brd=INNER_B)
        _sc(ws.cell(i, 8,
                    f"{disc_cnt} إجراء" if disc_cnt > 0 else "—"),
            sz=8, ah="center", bg=disc_bg, brd=INNER_B)
        _sc(ws.cell(i, 9,
                    f"{late_cnt} مرة" if late_cnt > 0 else "—"),
            sz=8, ah="center", bg=att_bg, brd=INNER_B)

    # حدود خارجية
    last = 3 + len(rows)
    thick_s = Side(style="medium", color="000000")
    for rr in range(3, last + 1):
        for c_idx in [1, 9]:
            b = ws.cell(rr, c_idx).border
            if c_idx == 1:
                ws.cell(rr, c_idx).border = Border(
                    left=thick_s, right=b.right, top=b.top, bottom=b.bottom)
            else:
                ws.cell(rr, c_idx).border = Border(
                    left=b.left, right=thick_s, top=b.top, bottom=b.bottom)
    for c_idx in range(1, 10):
        b = ws.cell(last, c_idx).border
        ws.cell(last, c_idx).border = Border(
            left=b.left, right=b.right, top=b.top, bottom=thick_s)

    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize   = 9
    ws.page_setup.fitToPage   = True
    ws.page_setup.fitToWidth  = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins = PageMargins(left=0.4, right=0.4, top=0.4, bottom=0.4)
    ws.print_options.horizontalCentered = True

    return ws


# ═══════════════════════════════════════════════════════════════════
# print_preview_html
# ═══════════════════════════════════════════════════════════════════
def _rgb_to_hex(color_obj):
    try:
        if color_obj and color_obj.type == "rgb":
            rgb = color_obj.rgb
            return "#" + rgb[2:]
    except Exception:
        pass
    return ""


def print_preview_html(xlsx_buf, title="تقرير", chart_b64=""):
    """
    HTML للطباعة — صفحة واحدة A4 Landscape مضمونة.
    CSS ثابت 100% — بدون JavaScript.
    transform:scale(0.62) ثابت مع transform-origin:top right.
    """
    xlsx_buf.seek(0)
    wb = openpyxl.load_workbook(xlsx_buf, data_only=True)

    _logo_b64 = ""
    for _lp in [LOGO_PATH, "logo.png"]:
        if os.path.exists(_lp):
            try:
                with open(_lp, "rb") as _lf:
                    _logo_b64 = base64.b64encode(_lf.read()).decode()
            except Exception:
                pass
            break

    CSS = """
@import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700&display=swap');
*, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

body {
    font-family: 'Cairo', Arial, sans-serif;
    direction: rtl;
    background: #dde1ea;
    padding: 14px;
}

.no-print {
    text-align: center;
    margin-bottom: 12px;
}
.no-print button {
    padding: 10px 36px;
    background: #1F3864;
    color: #fff;
    border: none;
    border-radius: 8px;
    cursor: pointer;
    font-size: 14px;
    font-family: 'Cairo', Arial, sans-serif;
    font-weight: 700;
    box-shadow: 0 2px 8px rgba(0,0,0,0.3);
}
.no-print button:hover { background: #2E75B6; }

.page {
    background: #fff;
    width: 277mm;
    margin: 0 auto 16px;
    padding: 5mm 6mm;
    box-shadow: 0 3px 14px rgba(0,0,0,0.22);
    overflow: hidden;
}

table {
    border-collapse: collapse;
    width: 100%;
    table-layout: fixed;
    font-size: 8pt;
    direction: rtl;
    font-family: 'Cairo', Arial, sans-serif;
}
td {
    padding: 1px 3px;
    vertical-align: middle;
    overflow: hidden;
    word-break: break-word;
    line-height: 1.22;
}

@media print {
    html, body {
        background: #fff !important;
        margin: 0 !important;
        padding: 0 !important;
        width: 297mm !important;
        height: 210mm !important;
    }
    .no-print { display: none !important; }
    .page {
        transform: scale(0.62) !important;
        transform-origin: top right !important;
        margin-top: 0 !important;
        margin-bottom: -105mm !important;
        margin-right: 0 !important;
        margin-left: auto !important;
        padding: 4mm 5mm !important;
        box-shadow: none !important;
        width: 277mm !important;
        overflow: visible !important;
        page-break-after: avoid !important;
        break-after: avoid !important;
        page-break-inside: avoid !important;
        break-inside: avoid !important;
    }
    table { width: 100% !important; table-layout: fixed !important; font-size: 8pt !important; }
    td    { font-size: 8pt !important; padding: 1px 3px !important; line-height: 1.22 !important; }
    tr    { page-break-inside: avoid !important; break-inside: avoid !important; }
    img   { max-height: 44px !important; object-fit: contain; }
    * {
        -webkit-print-color-adjust: exact !important;
        print-color-adjust: exact !important;
        color-adjust: exact !important;
    }
    @page { size: A4 landscape; margin: 6mm; }
}
"""

    parts = [
        "<!DOCTYPE html>",
        '<html dir="rtl" lang="ar">',
        "<head>",
        '<meta charset="utf-8">',
        f"<title>{title}</title>",
        f"<style>{CSS}</style>",
        "</head><body>",
        '<div class="no-print">',
        f'<button onclick="window.print()">&#128438;&nbsp; {title} &mdash; طباعة صفحة واحدة</button>',
        "</div>",
    ]

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        if ws.max_row == 0:
            continue

        logo_tag = ""
        if _logo_b64:
            logo_tag = (
                '<div style="display:flex;align-items:center;justify-content:space-between;'
                'background:#1F3864;padding:4px 10px;margin-bottom:3px;border-radius:3px;">'
                '<span style="color:#fff;font-size:10pt;font-weight:700;">'
                + _company_header() +
                "</span>"
                f'<img src="data:image/png;base64,{_logo_b64}"'
                ' style="height:40px;width:40px;object-fit:contain;" />'
                "</div>"
            )

        merged = {}
        skip   = set()
        for m in ws.merged_cells.ranges:
            merged[(m.min_row, m.min_col)] = (
                m.max_row - m.min_row + 1,
                m.max_col - m.min_col + 1,
            )
            for r2 in range(m.min_row, m.max_row + 1):
                for c2 in range(m.min_col, m.max_col + 1):
                    if (r2, c2) != (m.min_row, m.min_col):
                        skip.add((r2, c2))

        col_widths = {}
        for col_letter, cd in ws.column_dimensions.items():
            idx = column_index_from_string(col_letter)
            col_widths[idx] = 0 if cd.hidden else max(int((cd.width or 8) * 6.5), 0)

        colgroup = "<colgroup>"
        for c in range(1, ws.max_column + 1):
            colgroup += f'<col style="width:{col_widths.get(c, 50)}px;">'
        colgroup += "</colgroup>"

        parts.append(f'<div class="page">{logo_tag}<table>{colgroup}')

        for r in range(1, ws.max_row + 1):
            rh_raw = ws.row_dimensions[r].height if r in ws.row_dimensions else 13
            rh = 11 if rh_raw is None else max(int(rh_raw * 0.95), 11)
            parts.append(f'<tr style="height:{rh}px;">')

            for c in range(1, ws.max_column + 1):
                if col_widths.get(c, 1) == 0:
                    continue
                if (r, c) in skip:
                    continue

                cell = ws.cell(r, c)
                val  = cell.value
                text = "" if val is None else str(val).replace("\n", "<br>")
                style = "overflow:hidden;word-break:break-word;"

                f_obj = cell.font
                if f_obj:
                    sz = min(f_obj.size or 9, 10)
                    style += f"font-size:{sz}pt;"
                    if f_obj.bold:
                        style += "font-weight:bold;"
                    fc = _rgb_to_hex(f_obj.color)
                    if fc and fc != "#000000":
                        style += f"color:{fc};"

                p = cell.fill
                if p and p.fill_type == "solid":
                    bg = _rgb_to_hex(p.fgColor)
                    if bg and bg.lower() not in ("#000000", "#ffffff", ""):
                        style += f"background:{bg};"

                a = cell.alignment
                if a:
                    ha = {"right":"right","center":"center","left":"left"}.get(
                        a.horizontal or "right", "right")
                    va = {"top":"top","center":"middle","bottom":"bottom"}.get(
                        a.vertical or "center", "middle")
                    style += f"text-align:{ha};vertical-align:{va};"
                else:
                    style += "text-align:right;vertical-align:middle;"
                style += "padding:1px 3px;"

                b = cell.border
                if b:
                    def bs(s):
                        return ("1px solid #000" if s and s.style in ("medium","thick")
                                else ("0.4px solid #AAA" if s and s.style else "none"))
                    style += (f"border-top:{bs(b.top)};border-bottom:{bs(b.bottom)};"
                              f"border-right:{bs(b.right)};border-left:{bs(b.left)};")

                span = ""
                if (r, c) in merged:
                    rs2, cs2 = merged[(r, c)]
                    if rs2 > 1: span += f' rowspan="{rs2}"'
                    if cs2 > 1: span += f' colspan="{cs2}"'

                parts.append(f'<td style="{style}"{span}>{text}</td>')

            parts.append("</tr>")

        chart_section = ""
        if chart_b64:
            chart_section = (
                '<div style="margin-top:6px;text-align:left;padding-left:6px;">'
                f'<img src="data:image/png;base64,{chart_b64}"'
                ' style="width:44%;max-width:360px;height:auto;'
                'border:1px solid #E2E8F0;border-radius:4px;" />'
                "</div>"
            )
        parts.append(f"</table>{chart_section}</div>")

    parts.append("</body></html>")
    return "".join(parts)
