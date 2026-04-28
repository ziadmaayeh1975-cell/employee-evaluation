import os
from datetime import date
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.drawing.image import Image as XLImage
from openpyxl.worksheet.page import PageMargins
from openpyxl.utils import get_column_letter
from constants import *
from calculations import verbal_grade, kpi_score_to_pct, rating_label
from excel_reports import print_preview_html


def build_employee_sheet(wb, emp_name, job_title, dept, manager, year, kpis, monthly_scores,
                         notes="", training="", chart_img=None, disciplinary_actions=None,
                         employee_id="", attendance_data=None):
    safe = emp_name[:28]
    if safe in [s.title for s in wb.worksheets]:
        safe = safe[:25] + "_2"
    ws = wb.create_sheet(safe)
    ws.sheet_view.rightToLeft = True
    ws.sheet_view.showGridLines = False

    # ======================== الألوان والحدود ========================
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
    ATT_BG = "E0F2FE"

    _med = Side(style="medium", color="000000")
    _thn = Side(style="thin", color="000000")
    BK = Border(left=_med, right=_med, top=_med, bottom=_med)
    TN = Border(left=_thn, right=_thn, top=_thn, bottom=_thn)

    def sc(cell, val=None, bold=False, sz=9, color="000000", bg=None, ah="right", av="center", wrap=False):
        if val is not None:
            cell.value = val
        cell.font = Font(name="Arial", bold=bold, size=sz, color=color)
        cell.alignment = Alignment(horizontal=ah, vertical=av, wrapText=wrap, readingOrder=2)
        if bg:
            cell.fill = PatternFill("solid", fgColor=bg)

    def mc(r1, c1, r2, c2, val=None, **kw):
        ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)
        sc(ws.cell(r1, c1, val), **kw)

    # تجهيز بيانات الأشهر
    m_score, m_date, m_note, m_train = {}, {}, {}, {}
    for item in monthly_scores:
        ms = item[1]
        def _v(x): return str(x).strip() if x not in (None, "nan", "None") else ""
        m_score[ms] = item[2]
        m_date[ms] = _v(item[3]) if len(item) > 3 else ""
        m_note[ms] = _v(item[4]) if len(item) > 4 else ""
        m_train[ms] = _v(item[5]) if len(item) > 5 else ""

    done = [(n, m, s) for n, m, s, *_ in monthly_scores if s > 0]
    pct = sum(s for _, _, s in done) / len(done) * 100 if done else 0
    verb = verbal_grade(pct)
    sc_c = "375623" if pct >= 80 else ("C00000" if pct < 60 else "7F6000")
    sbg = GREEN_BG if pct >= 80 else (YELLOW if pct >= 60 else RED_BG)

    # معالجة الـ KPIs
    job_kpis = [(k["KPI_Name"], k["Weight"], k.get("avg_score", 0)) for k in kpis
                if k["KPI_Name"] not in PERSONAL_KPIS]
    per_kpis = [(k["KPI_Name"], k["Weight"], k.get("avg_score", 0)) for k in kpis
                if k["KPI_Name"] in PERSONAL_KPIS]

    _MAR = {"Jan": "يناير", "Feb": "فبراير", "Mar": "مارس", "Apr": "أبريل",
            "May": "مايو", "Jun": "يونيو", "Jul": "يوليو", "Aug": "أغسطس",
            "Sep": "سبتمبر", "Oct": "أكتوبر", "Nov": "نوفمبر", "Dec": "ديسمبر"}

    _company, _branch = "", ""
    try:
        from auth import load_app_settings as _las
        _cfg = _las()
        _company = _cfg.get("company_name", "مجموعة شركات فنون")
        _branch = _cfg.get("branch_name", "")
    except:
        pass
    _header = f"نموذج تقييم الأداء السنوي — {_company}" + (f" — {_branch}" if _branch else "")

    # إعداد عرض الأعمدة
    ws.column_dimensions["A"].width = 5
    ws.column_dimensions["B"].width = 25
    ws.column_dimensions["C"].width = 18
    ws.column_dimensions["D"].width = 15
    ws.column_dimensions["E"].width = 3
    ws.column_dimensions["F"].width = 3
    ws.column_dimensions["G"].width = 12
    ws.column_dimensions["H"].width = 12
    ws.column_dimensions["I"].width = 12
    ws.column_dimensions["J"].width = 15
    ws.column_dimensions["K"].width = 25
    ws.column_dimensions["L"].width = 15
    ws.column_dimensions["M"].width = 15
    ws.column_dimensions["N"].width = 5

    r = 1

    # ترويسة (الصف 1)
    ws.row_dimensions[1].height = 32
    mc(1, 1, 1, 14, _header, bold=True, sz=12, color="FFFFFF", bg=DARK, ah="center")
    
    _logo = globals().get("LOGO_PATH", "logo.png")
    if os.path.exists(_logo):
        try:
            img = XLImage(_logo)
            img.height, img.width = 70, 56
            img.anchor = "A1"
            ws.add_image(img)
        except:
            pass

    # معلومات الموظف (الصفوف 2-8)
    r = 2
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
        row = r + i
        ws.row_dimensions[row].height = 18
        sc(ws.cell(row, 1, lbl), bold=True, color="FFFFFF", bg=DARK, ah="center")
        sc(ws.cell(row, 2, val), bold=False, color="000000", bg=INFO_BG, ah="right")
    
    # الصف 9: نتيجة التقييم السنوي
    r = 9
    ws.row_dimensions[r].height = 18
    sc(ws.cell(r, 1, "نتيجة التقييم السنوي"), bold=True, color="FFFFFF", bg=ORANGE, ah="center")
    sc(ws.cell(r, 2, f"{int(round(pct))}% — {verb}"), bold=True, sz=11, color=sc_c, bg=sbg, ah="center")

    # ======================== جدول "نتيجة التقييم الشهري" يبدأ من الصف 4؟ لا، من الصف 11
    # ولكن بناءً على طلبك "يبدأ من صف G4" - لا يمكن أن يكون في الصف 4 لأن معلومات الموظف تأخذ صفوف 2-8
    # الصف 4 مشغول بمعلومات الموظف. سأبدأ الجدول من الصف 11 (بعد معلومات الموظف ونتيجة التقييم)
    
    r = 11  # بداية الجدول (يمكن تعديلها حسب الحاجة)
    
    # عنوان الجدول من G إلى M
    ws.row_dimensions[r].height = 20
    mc(r, 7, r, 13, "نتيجة التقييم الشهري", bold=True, sz=11, color="FFFFFF", bg=DARK, ah="center")
    r += 1

    # رأس الجدول من G إلى M
    headers = ["الشهر", "الدرجة (%)", "التقييم اللفظي", "تاريخ التقييم", "ملاحظات المقيم", "الإجراءات", "عدد مرات التأخير"]
    for col_idx, header in enumerate(headers, 7):
        sc(ws.cell(r, col_idx, header), bold=True, sz=9, color="FFFFFF", bg=MID, ah="center")
    r += 1

    MONTHS_LIST = ["يناير", "فبراير", "مارس", "أبريل", "مايو", "يونيو",
                   "يوليو", "أغسطس", "سبتمبر", "أكتوبر", "نوفمبر", "ديسمبر"]

    # تجهيز بيانات الإجراءات التأديبية لكل شهر
    disc_by_month = {}
    if disciplinary_actions is not None and not disciplinary_actions.empty:
        for _, row_disc in disciplinary_actions.iterrows():
            disc_date = row_disc.get("action_date", "")
            if disc_date:
                try:
                    month_num = int(disc_date.split("-")[1])
                    if month_num not in disc_by_month:
                        disc_by_month[month_num] = []
                    disc_by_month[month_num].append(row_disc.get("warning_type", ""))
                except:
                    pass

    # تجهيز بيانات الالتزام بالدوام لكل شهر
    att_by_month = {}
    if attendance_data is not None:
        if hasattr(attendance_data, 'iterrows'):
            for _, row_att in attendance_data.iterrows():
                month_num = row_att.get("month")
                if month_num:
                    att_by_month[month_num] = row_att.get("late_count", 0)
        elif isinstance(attendance_data, dict):
            month_num = attendance_data.get("month")
            if month_num:
                att_by_month[month_num] = attendance_data.get("late_count", 0)

    # عرض بيانات الأشهر (12 شهرًا)
    for month_idx, month_name in enumerate(MONTHS_LIST, 1):
        ws.row_dimensions[r].height = 16
        rbg = LGRAY if month_idx % 2 == 0 else WHITE

        # بيانات التقييم الشهري
        month_data = None
        for item in monthly_scores:
            short = item[1]
            month_num = list(_MAR.keys()).index(short) + 1 if short in _MAR else 0
            if month_num == month_idx:
                month_data = item
                break

        if month_data and month_data[2] > 0:
            score = month_data[2]
            eval_date = month_data[3] if len(month_data) > 3 else ""
            note = month_data[4] if len(month_data) > 4 else ""
            score_pct = f"{round(score * 100, 1)}%"
            verbal_val = verbal_grade(score * 100)
        else:
            score_pct = "—"
            verbal_val = "—"
            eval_date = "—"
            note = "—"

        sc(ws.cell(r, 7, month_name), bg=rbg, ah="center")
        sc(ws.cell(r, 8, score_pct), bg=rbg, ah="center")
        sc(ws.cell(r, 9, verbal_val), bg=rbg, ah="center")
        sc(ws.cell(r, 10, eval_date), bg=rbg, ah="center")
        sc(ws.cell(r, 11, note), bg=rbg, ah="right", wrap=True)

        # الإجراءات التأديبية
        disc_text = "—"
        if month_idx in disc_by_month:
            disc_text = "، ".join(set(disc_by_month[month_idx]))
        sc(ws.cell(r, 12, disc_text), bg=rbg, ah="center")

        # عدد مرات التأخير
        late_count = att_by_month.get(month_idx, 0)
        sc(ws.cell(r, 13, str(late_count) if late_count > 0 else "—"), bg=rbg, ah="center")

        r += 1

    r += 2

    # ======================== مؤشرات الأداء الوظيفي ========================
    ws.row_dimensions[r].height = 16
    sc(ws.cell(r, 1, "مؤشرات الأداء الوظيفي"), bold=True, color="FFFFFF", bg=DARK, ah="right")
    sc(ws.cell(r, 2, "الوزن النسبي %"), bold=True, color="FFFFFF", bg=DARK, ah="center")
    sc(ws.cell(r, 3, "الدرجة (0-100)"), bold=True, color="FFFFFF", bg=DARK, ah="center")
    sc(ws.cell(r, 4, "التقييم"), bold=True, color="FFFFFF", bg=DARK, ah="center")
    r += 1

    job_total_score, job_total_weight = 0.0, 0.0
    for i, (kname, weight, grade) in enumerate(job_kpis):
        rbg = LGRAY if i % 2 == 0 else WHITE
        w, g = float(weight), float(grade)
        pct_val = round(kpi_score_to_pct(g, w), 1)
        lbl = rating_label(pct_val)
        job_total_score += g
        job_total_weight += w
        kbg = GREEN_BG if pct_val >= 80 else (YELLOW if pct_val >= 60 else (RED_BG if pct_val > 0 else rbg))
        ws.row_dimensions[r].height = 16
        sc(ws.cell(r, 1, kname), bg=rbg, wrap=True)
        sc(ws.cell(r, 2, f"{w:.1f}%"), bg=rbg, ah="center")
        sc(ws.cell(r, 3, pct_val), bold=True, bg=kbg, ah="center")
        sc(ws.cell(r, 4, lbl), bold=True, bg=kbg, ah="center")
        r += 1

    ws.row_dimensions[r].height = 16
    sc(ws.cell(r, 1, "مجموع الأداء الوظيفي"), bold=True, color="FFFFFF", bg=MID, ah="right")
    sc(ws.cell(r, 2, f"{job_total_weight:.1f}%"), bold=True, color="FFFFFF", bg=MID, ah="center")
    job_pct_total = round(kpi_score_to_pct(job_total_score, job_total_weight), 1) if job_total_weight > 0 else 0
    sc(ws.cell(r, 3, f"{job_pct_total}%"), bold=True, color="FFFFFF", bg=MID, ah="center")
    sc(ws.cell(r, 4, rating_label(job_pct_total)), bold=True, color="FFFFFF", bg=MID, ah="center")
    r += 1
    r += 1

    # ======================== مؤشرات الصفات الشخصية ========================
    ws.row_dimensions[r].height = 16
    mc(r, 1, r, 3, "مؤشرات الصفات الشخصية", bold=True, color="FFFFFF", bg=ORANGE, ah="center")
    r += 1
    ws.row_dimensions[r].height = 16
    sc(ws.cell(r, 1, "المؤشر"), bold=True, color="FFFFFF", bg=MID, ah="right")
    sc(ws.cell(r, 2, "الوزن النسبي %"), bold=True, color="FFFFFF", bg=MID, ah="center")
    sc(ws.cell(r, 3, "الدرجة (0-100)"), bold=True, color="FFFFFF", bg=MID, ah="center")
    sc(ws.cell(r, 4, "التقييم"), bold=True, color="FFFFFF", bg=MID, ah="center")
    r += 1

    per_total_score, per_total_weight = 0.0, 0.0
    for i, (kname, weight, grade) in enumerate(per_kpis):
        rbg = WARM if i % 2 == 0 else WHITE
        w, g = float(weight), float(grade)
        pct_val = round(kpi_score_to_pct(g, w), 1)
        lbl = rating_label(pct_val)
        per_total_score += g
        per_total_weight += w
        kbg = GREEN_BG if pct_val >= 80 else (YELLOW if pct_val >= 60 else (RED_BG if pct_val > 0 else rbg))
        ws.row_dimensions[r].height = 16
        sc(ws.cell(r, 1, kname), bg=rbg, wrap=True)
        sc(ws.cell(r, 2, f"{w:.1f}%"), bg=rbg, ah="center")
        sc(ws.cell(r, 3, pct_val), bold=True, bg=kbg, ah="center")
        sc(ws.cell(r, 4, lbl), bold=True, bg=kbg, ah="center")
        r += 1

    ws.row_dimensions[r].height = 16
    sc(ws.cell(r, 1, "مجموع الصفات الشخصية"), bold=True, color="FFFFFF", bg=ORANGE, ah="right")
    sc(ws.cell(r, 2, f"{per_total_weight:.1f}%"), bold=True, color="FFFFFF", bg=ORANGE, ah="center")
    per_pct_total = round(kpi_score_to_pct(per_total_score, per_total_weight), 1) if per_total_weight > 0 else 0
    sc(ws.cell(r, 3, f"{per_pct_total}%"), bold=True, color="FFFFFF", bg=ORANGE, ah="center")
    sc(ws.cell(r, 4, rating_label(per_pct_total)), bold=True, color="FFFFFF", bg=ORANGE, ah="center")
    r += 1
    r += 1

    # ملاحظات المقيم
    ws.row_dimensions[r].height = 22
    mc(r, 1, r, 4, f"ملاحظات المقيم: {notes or ''}", bg=NOTE_BG, wrap=True)
    r += 1

    # الاحتياجات التدريبية
    _train_vals = [v for v in m_train.values() if v and str(v).strip() not in ("", "nan", "None", "—")]
    _train = _train_vals[0] if _train_vals else (training if training else "")
    mc(r, 1, r, 4, f"الاحتياجات التدريبية: {_train}" if _train else "الاحتياجات التدريبية:", bg=TRAIN_BG, wrap=True)
    r += 1
    r += 1

    # التوقيع
    ws.row_dimensions[r].height = 18
    sc(ws.cell(r, 1, f"المسؤول المباشر: {manager}"), bold=True, ah="center")
    sc(ws.cell(r, 2, "اسم الموظف"), bold=True, ah="center")
    _ec = ws.cell(r, 3, emp_name)
    sc(_ec, bold=True, ah="right")
    _ec.border = Border(left=_med, right=_med, top=_med, bottom=_med)
    r += 1
    ws.row_dimensions[r].height = 18
    _t1 = ws.cell(r, 1, "التوقيع: _______________")
    sc(_t1, bold=True, bg=LGRAY, ah="center")
    _t1.border = Border(left=_med, right=_med, top=_med, bottom=_med)
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
    _t2 = ws.cell(r, 2, "التوقيع: _______________")
    sc(_t2, bold=True, bg=LGRAY, ah="center")
    _t2.border = Border(left=_med, right=_med, top=_med, bottom=_med)

    # إعداد الطباعة
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = 9
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins = PageMargins(left=0.5, right=0.5, top=0.5, bottom=0.5)
    ws.print_options.horizontalCentered = True
    ws.print_options.verticalCentered = True

    return ws


def build_summary_sheet(wb, rows, title="ملخص التقييم", year=None, chart_img=None):
    ws = wb.create_sheet(title[:28])
    ws.sheet_view.rightToLeft = True
    ws.sheet_view.showGridLines = False

    DARK = "1F3864"
    LGRAY = "F2F2F2"
    WHITE = "FFFFFF"
    YELLOW = "FFF2CC"
    GREEN_BG = "E2EFDA"
    RED_BG = "FFDAD9"

    for col, w in [("A", 4), ("B", 32), ("C", 14), ("D", 8), ("E", 10), ("F", 13), ("G", 13)]:
        ws.column_dimensions[col].width = w

    ws.row_dimensions[1].height = 36

    def _sc(cell, val=None, bold=False, sz=9, color="000000", bg=None, ah="right", av="center", brd="inner"):
        if val is not None:
            cell.value = val
        cell.font = Font(name="Arial", bold=bold, size=sz, color=color)
        cell.alignment = Alignment(horizontal=ah, vertical=av, readingOrder=2)
        if bg:
            cell.fill = PatternFill("solid", fgColor=bg)
        if brd == "outer":
            cell.border = OUTER_B
        else:
            cell.border = INNER_B

    def _mc(r1, c1, r2, c2, val=None, **kw):
        ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)
        _sc(ws.cell(r1, c1, val), **kw)

    _mc(1, 1, 1, 7, title, bold=True, sz=12, color="FFFFFF", bg=DARK, ah="center", av="center", brd="outer")

    import os as _os2
    from openpyxl.drawing.image import Image as XLImg2
    _logo2 = globals().get("LOGO_PATH", "logo.png")
    if _logo2 and _os2.path.exists(_logo2):
        try:
            _img = XLImg2(_logo2)
            _img.height = 32
            _img.width = 26
            _img.anchor = "G1"
            ws.add_image(_img)
        except:
            pass

    ws.row_dimensions[2].height = 4
    ws.row_dimensions[3].height = 16
    for c, t in [(1, "#"), (2, "اسم الموظف"), (3, "القسم"), (4, "السنة"), (5, "الأشهر"), (6, "المعدل %"), (7, "التقييم")]:
        _sc(ws.cell(3, c, t), bold=True, sz=9, color="FFFFFF", bg=DARK, ah="center", brd="outer")

    for i, (name, dept, months, pct_val, verb) in enumerate(rows, 4):
        ws.row_dimensions[i].height = 16
        rbg = LGRAY if i % 2 == 0 else WHITE
        sc_c = "375623" if pct_val >= 80 else ("C00000" if pct_val < 60 else "000000")
        vbg = GREEN_BG if pct_val >= 80 else (YELLOW if pct_val >= 70 else (RED_BG if pct_val < 60 else LGRAY))
        _sc(ws.cell(i, 1, i - 3), sz=8, ah="center", bg=rbg, brd="inner")
        _sc(ws.cell(i, 2, name), sz=9, ah="right", bg=rbg, brd="inner")
        _sc(ws.cell(i, 3, dept), sz=8, ah="center", bg=rbg, brd="inner")
        _sc(ws.cell(i, 4, year or ""), sz=9, ah="center", bg=rbg, brd="inner")
        _sc(ws.cell(i, 5, months), sz=9, ah="center", bg=rbg, brd="inner")
        _sc(ws.cell(i, 6, f"{pct_val:.1f}%"), sz=10, bold=True, color=sc_c, ah="center", bg=vbg, brd="inner")
        _sc(ws.cell(i, 7, verb), sz=9, bold=True, color=sc_c, ah="center", bg=vbg, brd="inner")

    last = 3 + len(rows)
    for r in range(3, last + 1):
        lft = ws.cell(r, 1).border
        ws.cell(r, 1).border = Border(left=thick_s, right=lft.right, top=lft.top, bottom=lft.bottom)
        rgt = ws.cell(r, 7).border
        ws.cell(r, 7).border = Border(left=rgt.left, right=thick_s, top=rgt.top, bottom=rgt.bottom)
    for c in range(1, 8):
        b = ws.cell(last, c).border
        ws.cell(last, c).border = Border(left=b.left, right=b.right, top=b.top, bottom=thick_s)

    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = 9
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins = PageMargins(left=0.5, right=0.5, top=0.5, bottom=0.5)
    ws.print_options.horizontalCentered = True
    ws.print_options.verticalCentered = True

    return ws
