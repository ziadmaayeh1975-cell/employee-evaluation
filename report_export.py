import os
from datetime import date
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.drawing.image import Image as XLImage
from openpyxl.worksheet.page import PageMargins
from openpyxl.utils import get_column_letter
from constants import *
from calculations import verbal_grade, kpi_score_to_pct, rating_label
from excel_reports import print_preview_html


def build_employee_sheet(wb, emp_name, job_title, dept, manager, year,
                         kpis, monthly_scores, notes="", training="", chart_img=None,
                         disciplinary_actions=None, employee_id=""):
    """
    بناء شيت Excel لتقرير الموظف
    kpis: list of dicts with keys: KPI_Name, Weight, avg_score
    monthly_scores: list of (idx, short, score, date, notes, training)
    """
    import os as _os
    from openpyxl.drawing.image import Image as XLImg

    safe = emp_name[:28]
    used = [s.title for s in wb.worksheets]
    if safe in used:
        safe = safe[:25] + "_2"

    ws = wb.create_sheet(safe)
    ws.sheet_view.rightToLeft = True
    ws.sheet_view.showGridLines = False

    # ── ألوان ───────────────────────────────────────────
    DARK = "1F3864"
    MID = "2E75B6"
    ORANGE = "ED7D31"
    LGRAY = "F2F2F2"
    WHITE = "FFFFFF"
    YELLOW = "FFF2CC"
    GREEN_BG = "E2EFDA"
    RED_BG = "FFDAD9"
    WARM = "FFF3E0"
    NOTE_BG = "FFFDE7"
    TRAIN_BG = "F3E5F5"
    DATE_BG = "E3F2FD"
    INFO_BG = "EBF3FB"
    DISC_BG = "FEE2E2"

    # ── حدود ──────────────────────────────────────────
    _med = Side(style="medium", color="000000")
    _thn = Side(style="thin", color="000000")
    BK = Border(left=_med, right=_med, top=_med, bottom=_med)
    TN = Border(left=_thn, right=_thn, top=_thn, bottom=_thn)

    SZ = 9

    def sc(cell, val=None, bold=False, sz=SZ, color="000000",
           bg=None, ah="right", av="center", brd="tn", wrap=False):
        try:
            if val is not None:
                cell.value = val
            cell.font = Font(name="Arial", bold=bold, size=sz, color=color)
            cell.alignment = Alignment(horizontal=ah, vertical=av,
                                       wrapText=wrap, readingOrder=2)
            if bg:
                cell.fill = PatternFill("solid", fgColor=bg)
            if brd == "bk":
                cell.border = BK
            else:
                cell.border = TN
        except:
            pass

    def mc(r1, c1, r2, c2, val=None, **kw):
        ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)
        sc(ws.cell(r1, c1, val), **kw)

    # ── استخراج بيانات الأشهر ─────────────────────────────────────────
    m_score = {}
    m_date = {}
    m_note = {}
    m_train = {}
    for item in monthly_scores:
        ms = item[1]
        def _v(x):
            s = str(x).strip() if x is not None else ""
            return "" if s in ("nan", "None", "") else s
        m_score[ms] = item[2]
        m_date[ms] = _v(item[3]) if len(item) > 3 else ""
        m_note[ms] = _v(item[4]) if len(item) > 4 else ""
        m_train[ms] = _v(item[5]) if len(item) > 5 else ""

    done = [(n, m, s_) for n, m, s_, *_ in monthly_scores if s_ > 0]
    pct = (sum(s_ for _, _, s_ in done) / len(done) * 100) if done else 0
    verb = verbal_grade(pct)
    sc_c = "375623" if pct >= 80 else ("C00000" if pct < 60 else "7F6000")
    sbg = GREEN_BG if pct >= 80 else (YELLOW if pct >= 60 else RED_BG)

    # معالجة الـ KPIs
    job_kpis = [(k.get("KPI_Name", ""), k.get("Weight", 0), k.get("avg_score", 0)) for k in kpis if k.get("KPI_Name", "") not in PERSONAL_KPIS]
    per_kpis = [(k.get("KPI_Name", ""), k.get("Weight", 0), k.get("avg_score", 0)) for k in kpis if k.get("KPI_Name", "") in PERSONAL_KPIS]

    _MAR = {"Jan": "يناير", "Feb": "فبراير", "Mar": "مارس", "Apr": "أبريل",
            "May": "مايو", "Jun": "يونيو", "Jul": "يوليو", "Aug": "أغسطس",
            "Sep": "سبتمبر", "Oct": "أكتوبر", "Nov": "نوفمبر", "Dec": "ديسمبر"}

    # اسم الشركة
    _company = ""
    _branch = ""
    try:
        from auth import load_app_settings as _las
        _cfg = _las()
        _company = _cfg.get("company_name", "مجموعة شركات فنون")
        _branch = _cfg.get("branch_name", "")
    except Exception:
        _company = globals().get("COMPANY_NAME", "مجموعة شركات فنون")
        _branch = globals().get("BRANCH_NAME", "")
    _header = f"نموذج تقييم الأداء السنوي — {_company}"
    if _branch:
        _header += f" — {_branch}"

    # إعداد الأعمدة
    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 14
    ws.column_dimensions["D"].width = 14
    ws.column_dimensions["E"].width = 14
    ws.column_dimensions["F"].width = 14
    ws.column_dimensions["G"].width = 14
    ws.column_dimensions["H"].width = 20
    ws.column_dimensions["I"].width = 20

    r = 1

    # صف 1: ترويسة
    ws.row_dimensions[1].height = 32
    mc(1, 1, 1, 9, _header,
       bold=True, sz=11, color="FFFFFF", bg=DARK, ah="center", brd="bk")

    # لوغو
    _logo = globals().get("LOGO_PATH", "logo.png")
    if _logo and _os.path.exists(_logo):
        try:
            img = XLImg(_logo)
            img.height = 70
            img.width = 56
            img.anchor = "A1"
            ws.add_image(img)
        except:
            pass

    r = 2

    # معلومات الموظف
    INFO = [
        ("اسم الموظف", emp_name),
        ("رقم الموظف", employee_id),
        ("الوظيفة", job_title),
        ("القسم", dept),
        ("السنة", str(year)),
        ("اسم المقيم", manager),
        ("تاريخ التقييم", date.today().strftime("%d/%m/%Y")),
    ]

    ws.row_dimensions[2].height = 16
    for i, (lbl, val) in enumerate(INFO):
        row = r + i
        ws.row_dimensions[row].height = 18
        sc(ws.cell(row, 1, lbl), bold=True, sz=SZ, color="FFFFFF", bg=DARK, ah="center", brd="bk")
        sc(ws.cell(row, 2, val), bold=True, sz=SZ, color="000000", bg=INFO_BG, ah="right", brd="tn")
        # دمج الأعمدة 3-9 للقيمة
        if val:
            mc(row, 3, row, 9, val, bold=True, sz=SZ, color="000000", bg=INFO_BG, ah="right", brd="tn")

    r += len(INFO)

    # نتيجة التقييم السنوي
    ws.row_dimensions[r].height = 18
    sc(ws.cell(r, 1, "نتيجة التقييم السنوي"),
       bold=True, sz=SZ, color="FFFFFF", bg=ORANGE, ah="center", brd="bk")
    mc(r, 2, r, 3, f"{int(round(pct))}%  —  {verb}",
       bold=True, sz=SZ, color=sc_c, bg=sbg, ah="center", brd="bk")
    r += 1

    # جدول مؤشرات الأداء الوظيفي
    ws.row_dimensions[r].height = 16
    sc(ws.cell(r, 1, "مؤشرات الأداء الوظيفي"),
       bold=True, sz=SZ, color="FFFFFF", bg=DARK, ah="right", brd="bk")
    sc(ws.cell(r, 2, "الوزن النسبي %"), bold=True, sz=SZ, color="FFFFFF", bg=DARK, ah="center", brd="bk")
    sc(ws.cell(r, 3, "الدرجة (0-100)"), bold=True, sz=SZ, color="FFFFFF", bg=DARK, ah="center", brd="bk")
    sc(ws.cell(r, 4, "التقييم"), bold=True, sz=SZ, color="FFFFFF", bg=DARK, ah="center", brd="bk")
    r += 1

    _job_total_score = 0.0
    _job_total_weight = 0.0
    for i, (kname, weight, grade) in enumerate(job_kpis):
        rbg = LGRAY if i % 2 == 0 else WHITE
        g = float(grade)
        w = float(weight)
        pct_val = round(kpi_score_to_pct(g, w), 1)
        lbl = rating_label(pct_val)
        _job_total_score += g
        _job_total_weight += w
        kbg = GREEN_BG if pct_val >= 80 else (YELLOW if pct_val >= 60 else (RED_BG if pct_val > 0 else rbg))
        ws.row_dimensions[r].height = 16
        sc(ws.cell(r, 1, kname), sz=10, color="000000", bg=rbg, ah="right", wrap=True)
        sc(ws.cell(r, 2, f"{w:.1f}%"), sz=10, bg=rbg, ah="center")
        sc(ws.cell(r, 3, pct_val), bold=True, sz=10, bg=kbg, ah="center")
        sc(ws.cell(r, 4, lbl), bold=True, sz=10, bg=kbg, ah="center")
        r += 1

    # مجموع الأداء الوظيفي
    ws.row_dimensions[r].height = 16
    sc(ws.cell(r, 1, "مجموع الأداء الوظيفي"), bold=True, sz=SZ, color="FFFFFF", bg=MID, ah="right", brd="bk")
    sc(ws.cell(r, 2, f"{_job_total_weight:.1f}%"), bold=True, sz=SZ, color="FFFFFF", bg=MID, ah="center", brd="bk")
    _job_pct_total = round(kpi_score_to_pct(_job_total_score, _job_total_weight), 1) if _job_total_weight > 0 else 0
    sc(ws.cell(r, 3, f"{_job_pct_total}%"), bold=True, sz=SZ, color="FFFFFF", bg=MID, ah="center", brd="bk")
    sc(ws.cell(r, 4, rating_label(_job_pct_total)), bold=True, sz=SZ, color="FFFFFF", bg=MID, ah="center", brd="bk")
    r += 1

    ws.row_dimensions[r].height = 3
    r += 1

    # مؤشرات الصفات الشخصية
    ws.row_dimensions[r].height = 16
    mc(r, 1, r, 3, "مؤشرات الصفات الشخصية",
       bold=True, sz=SZ, color="FFFFFF", bg=ORANGE, ah="center", brd="bk")
    r += 1
    ws.row_dimensions[r].height = 16
    sc(ws.cell(r, 1, "المؤشر"), bold=True, sz=SZ, color="FFFFFF", bg=MID, ah="right", brd="bk")
    sc(ws.cell(r, 2, "الوزن النسبي %"), bold=True, sz=SZ, color="FFFFFF", bg=MID, ah="center", brd="bk")
    sc(ws.cell(r, 3, "الدرجة (0-100)"), bold=True, sz=SZ, color="FFFFFF", bg=MID, ah="center", brd="bk")
    sc(ws.cell(r, 4, "التقييم"), bold=True, sz=SZ, color="FFFFFF", bg=MID, ah="center", brd="bk")
    r += 1

    _per_total_score = 0.0
    _per_total_weight = 0.0
    for i, (kname, weight, grade) in enumerate(per_kpis):
        rbg = WARM if i % 2 == 0 else WHITE
        g = float(grade)
        w = float(weight)
        pct_val = round(kpi_score_to_pct(g, w), 1)
        lbl = rating_label(pct_val)
        _per_total_score += g
        _per_total_weight += w
        kbg = GREEN_BG if pct_val >= 80 else (YELLOW if pct_val >= 60 else (RED_BG if pct_val > 0 else rbg))
        ws.row_dimensions[r].height = 16
        sc(ws.cell(r, 1, kname), sz=10, color="000000", bg=rbg, ah="right", wrap=True)
        sc(ws.cell(r, 2, f"{w:.1f}%"), sz=10, bg=rbg, ah="center")
        sc(ws.cell(r, 3, pct_val), bold=True, sz=10, bg=kbg, ah="center")
        sc(ws.cell(r, 4, lbl), bold=True, sz=10, bg=kbg, ah="center")
        r += 1

    # مجموع الصفات الشخصية
    ws.row_dimensions[r].height = 16
    sc(ws.cell(r, 1, "مجموع الصفات الشخصיות"), bold=True, sz=SZ, color="FFFFFF", bg=ORANGE, ah="right", br
