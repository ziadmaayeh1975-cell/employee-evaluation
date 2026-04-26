import os
from datetime import date
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.drawing.image import Image as XLImage
from openpyxl.worksheet.page import PageMargins
from openpyxl.utils import get_column_letter
from constants import *
from calculations import verbal_grade, kpi_score_to_pct, rating_label
from excel_reports import print_preview_html  # noqa


def auto_col_widths(ws, min_w=6, max_w=80):
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            if cell.value is not None:
                try:
                    if len(str(cell.value)) > max_len:
                        max_len = len(str(cell.value))
                except:
                    pass
        ws.column_dimensions[col_letter].width = min(max(max_len * 1.15 + 2, min_w), max_w)


def apply_data_border(ws, thin_side, thick_side):
    t = Side(style="thin", color="AAAAAA")
    b = Border(left=t, right=t, top=t, bottom=t)
    for row in ws.iter_rows():
        for cell in row:
            if cell.value is not None and str(cell.value).strip() != "":
                cell.border = b


def _sc(cell, val=None, bold=False, sz=9, color="000000", bg=None,
        ah="right", av="center", wrap=False, brd="inner"):
    try:
        if val is not None:
            cell.value = val
        cell.font = Font(name="Arial", bold=bold, size=sz, color=color)
        cell.alignment = Alignment(horizontal=ah, vertical=av,
                                   wrapText=wrap, readingOrder=2)
        if bg:
            cell.fill = PatternFill("solid", fgColor=bg)
        if brd == "outer":  cell.border = OUTER_B
        elif brd == "inner": cell.border = INNER_B
        elif brd == "row":   cell.border = ROW_B
    except:
        pass


def _mc(ws, r1, c1, r2, c2, val=None, **kw):
    ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)
    cell = ws.cell(r1, c1, val)
    _sc(cell, **kw)
    return cell


# ══════════════════════════════════════════════════════════════════════
# تقرير الموظف الفردي — مع إضافة الإجراءات التأديبية
# ══════════════════════════════════════════════════════════════════════
def build_employee_sheet(wb, emp_name, job_title, dept, manager, year,
                         kpis, monthly_scores, notes="", training="", chart_img=None,
                         disciplinary_actions=None):
    """
    monthly_scores: (idx, short, score) أو (idx, short, score, date, notes, train)
    disciplinary_actions: DataFrame أو قائمة تحتوي على الإجراءات التأديبية للموظف
    """
    import os as _os
    from openpyxl.drawing.image import Image as XLImg

    safe = emp_name[:28]
    used = [s.title for s in wb.worksheets]
    if safe in used:
        safe = safe[:25] + "_2"

    ws = wb.create_sheet(safe)
    ws.sheet_view.rightToLeft  = True
    ws.sheet_view.showGridLines = False

    # ── ألوان مطابقة للصورة ───────────────────────────────────────────
    DARK    = "1F3864"   # أزرق داكن — ترويسة وعناوين
    MID     = "2E75B6"   # أزرق متوسط — جدول الأشهر
    ORANGE  = "ED7D31"   # برتقالي — نتيجة التقييم / الصفات الشخصية
    LGRAY   = "F2F2F2"   # رمادي فاتح — صفوف زوجية
    WHITE   = "FFFFFF"
    YELLOW  = "FFF2CC"   # أصفر — نتيجة متوسطة
    GREEN_BG= "E2EFDA"   # أخضر فاتح — درجة جيدة
    RED_BG  = "FFDAD9"   # أحمر فاتح — درجة ضعيفة
    WARM    = "FFF3E0"   # برتقالي فاتح — صفات شخصية زوجية
    NOTE_BG = "FFFDE7"   # أصفر فاتح — ملاحظات
    TRAIN_BG= "F3E5F5"   # بنفسجي فاتح — تدريب
    DATE_BG = "E3F2FD"   # أزرق فاتح جداً — تاريخ
    INFO_BG = "EBF3FB"   # أزرق فاتح — قيم المعلومات
    DISC_BG = "FEE2E2"   # أحمر فاتح جداً — الإجراءات التأديبية

    # ── حدود ──────────────────────────────────────────────────────────
    _med = Side(style="medium", color="000000")
    _thn = Side(style="thin",   color="000000")
    _gry = Side(style="thin",   color="AAAAAA")
    BK = Border(left=_med, right=_med, top=_med, bottom=_med)  # حد سميك
    TN = Border(left=_thn, right=_thn, top=_thn, bottom=_thn)  # حد رفيع
    GR = Border(left=_gry, right=_gry, top=_gry, bottom=_gry)  # حد رمادي

    SZ = 9   # حجم الخط الأساسي

    def sc(cell, val=None, bold=True, sz=SZ, color="000000",
           bg=None, ah="right", av="center", brd="tn", wrap=False):
        try:
            if val is not None: cell.value = val
            cell.font      = Font(name="Arial", bold=bold, size=sz, color=color)
            cell.alignment = Alignment(horizontal=ah, vertical=av,
                                       wrapText=wrap, readingOrder=2)
            if bg: cell.fill = PatternFill("solid", fgColor=bg)
            if   brd == "bk": cell.border = BK
            elif brd == "tn": cell.border = TN
            elif brd == "gr": cell.border = GR
            elif brd == "none": pass
        except:
            pass

    def mc(r1, c1, r2, c2, val=None, **kw):
        ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)
        sc(ws.cell(r1, c1, val), **kw)

    # ── استخراج بيانات الأشهر ─────────────────────────────────────────
    m_score={}; m_date={}; m_note={}; m_train={}
    for item in monthly_scores:
        ms = item[1]
        def _v(x):
            s = str(x).strip() if x is not None else ""
            return "" if s in ("nan","None","") else s
        m_score[ms] = item[2]
        m_date[ms]  = _v(item[3]) if len(item) > 3 else ""
        m_note[ms]  = _v(item[4]) if len(item) > 4 else ""
        m_train[ms] = _v(item[5]) if len(item) > 5 else ""

    done = [(n,m,s_) for n,m,s_,*_ in monthly_scores if s_ > 0]
    pct  = (sum(s_ for _,_,s_ in done)/len(done)*100) if done else 0
    verb = verbal_grade(pct)
    sc_c = "375623" if pct>=80 else ("C00000" if pct<60 else "7F6000")
    sbg  = GREEN_BG if pct>=80 else (YELLOW if pct>=60 else RED_BG)

    job_kpis = [(k,w,g) for k,w,g in kpis if k not in PERSONAL_KPIS]
    per_kpis = [(k,w,g) for k,w,g in kpis if k in     PERSONAL_KPIS]

    _MAR = {"Jan":"يناير","Feb":"فبراير","Mar":"مارس","Apr":"أبريل",
            "May":"مايو","Jun":"يونيو","Jul":"يوليو","Aug":"أغسطس",
            "Sep":"سبتمبر","Oct":"أكتوبر","Nov":"نوفمبر","Dec":"ديسمبر"}

    # ── اسم الشركة ────────────────────────────────────────────────────
    _company = ""; _branch = ""
    try:
        from auth import load_app_settings as _las
        _cfg     = _las()
        _company = _cfg.get("company_name", "مجموعة شركات فنون")
        _branch  = _cfg.get("branch_name",  "")
    except Exception:
        _company = globals().get("COMPANY_NAME", "مجموعة شركات فنون")
        _branch  = globals().get("BRANCH_NAME",  "")
    _header = f"نموذج تقييم الأداء السنوي — {_company}"
    if _branch: _header += f" — {_branch}"

    # ════════════════════════════════════════════════════════════════════
    # تخطيط الأعمدة
    # ════════════════════════════════════════════════════════════════════
    ws.column_dimensions["B"].width = 7
    ws.column_dimensions["C"].width = 9
    ws.column_dimensions["D"].width = 10
    ws.column_dimensions["E"].width = 3
    ws.column_dimensions["F"].width = 9
    ws.column_dimensions["G"].width = 8
    ws.column_dimensions["H"].width = 12
    ws.column_dimensions["I"].width = 12
    ws.column_dimensions["J"].width = 20
    ws.column_dimensions["K"].width = 0
    ws.column_dimensions["K"].hidden = True

    r = 1

    # ════════════════════════════════════════════════════════════════════
    # صف 1: ترويسة
    # ════════════════════════════════════════════════════════════════════
    ws.row_dimensions[1].height = 32
    mc(1,1,1,9, _header,
       bold=True, sz=11, color="FFFFFF", bg=DARK, ah="center", brd="bk")

    _logo = globals().get("LOGO_PATH","logo.png")
    if _logo and _os.path.exists(_logo):
        try:
            img = XLImg(_logo)
            img.height = 70; img.width = 56; img.anchor = "A1"
            ws.add_image(img)
        except:
            pass
    r = 2

    # ════════════════════════════════════════════════════════════════════
    # معلومات الموظف + رأس جدول الأشهر
    # ════════════════════════════════════════════════════════════════════
    INFO = [
        ("اسم الموظف",    emp_name),
        ("الوظيفة",       job_title),
        ("القسم",         dept),
        ("السنة",         str(year)),
        ("اسم المقيم",   manager),
        ("تاريخ التقييم", date.today().strftime("%d/%m/%Y")),
    ]

    ws.row_dimensions[2].height = 16
    mc(2,6,2,9,"نتيجة التقييم الشهري",
       bold=True, sz=SZ, color="FFFFFF", bg=MID, ah="center", brd="bk")

    ws.row_dimensions[3].height = 16
    sc(ws.cell(3,6,"الشهر"),               bold=True,sz=SZ,color="FFFFFF",bg=DARK,ah="center",brd="bk")
    sc(ws.cell(3,7,"الدرجة"),              bold=True,sz=SZ,color="FFFFFF",bg=DARK,ah="center",brd="bk")
    sc(ws.cell(3,8,"التقييم اللفظي"),      bold=True,sz=8, color="FFFFFF",bg=DARK,ah="center",brd="bk")
    sc(ws.cell(3,9,"تاريخ التقييم"),       bold=True,sz=8, color="FFFFFF",bg=DARK,ah="center",brd="bk")
    sc(ws.cell(3,10,"ملاحظات المقيم"),     bold=True,sz=8, color="FFFFFF",bg=DARK,ah="center",brd="bk")

    _info_max_len = max((len(str(v)) for _, v in INFO), default=10)
    ws.column_dimensions["B"].width = min(max(_info_max_len * 0.6 + 3, 14), 35)

    for i, (lbl, val) in enumerate(INFO):
        row = r + i
        ws.row_dimensions[row].height = 18
        sc(ws.cell(row,1,lbl), bold=True,  sz=SZ, color="FFFFFF", bg=DARK, ah="center", brd="bk")
        sc(ws.cell(row,2,val), bold=True,  sz=SZ, color="000000", bg=INFO_BG, ah="right", brd="tn")

    r += len(INFO)

    # ════════════════════════════════════════════════════════════════════
    # بيانات الأشهر الـ12
    # ════════════════════════════════════════════════════════════════════
    MONTHS_LIST = ["Jan","Feb","Mar","Apr","May","Jun",
                   "Jul","Aug","Sep","Oct","Nov","Dec"]
    for i, short in enumerate(MONTHS_LIST, start=4):
        ws.row_dimensions[i].height = 20

        score  = float(m_score.get(short,0) or 0)*100
        edate  = m_date.get(short,"")
        mnote  = m_note.get(short,"")
        mbg    = LGRAY if i%2==0 else WHITE
        sbg_m  = (GREEN_BG if score>=80 else
                  (YELLOW  if score>=60 else
                   (RED_BG if score>0  else mbg)))
        sclr   = ("375623" if score>=80 else
                  ("C00000" if 0<score<60 else "000000"))

        _verbal_m = verbal_grade(score) if score > 0 else "—"
        sc(ws.cell(i,6, _MAR.get(short,short)),
           bold=True, sz=SZ, bg=mbg, ah="center", brd="tn")
        sc(ws.cell(i,7, f"{int(round(score))}%" if score>0 else "—"),
           bold=True, sz=SZ, color=sclr, bg=sbg_m, ah="center", brd="tn")
        sc(ws.cell(i,8, _verbal_m),
           bold=True, sz=8, color=sclr, bg=sbg_m, ah="center", brd="tn")
        sc(ws.cell(i,9, edate if edate else "—"),
           bold=False, sz=8, color="1F3864",
           bg=DATE_BG if edate else mbg, ah="center", brd="tn")
        _mnote_text = mnote if mnote else "—"
        _note_cell  = ws.cell(i,10, _mnote_text)
        sc(_note_cell, bold=False, sz=8, bg=NOTE_BG if mnote else mbg,
           ah="right", av="top", brd="tn", wrap=True)
        if mnote:
            _H_width = ws.column_dimensions["I"].width or 20
            _chars_per_line = max(int(_H_width / 0.55), 10)
            _lines = max(1, -(-len(mnote) // _chars_per_line))
            ws.row_dimensions[i].height = max(20, _lines * 14)

    # ════════════════════════════════════════════════════════════════════
    # نتيجة التقييم السنوي
    # ════════════════════════════════════════════════════════════════════
    ws.row_dimensions[r].height = 18
    sc(ws.cell(r,1,"نتيجة التقييم السنوي"),
       bold=True, sz=SZ, color="FFFFFF", bg=ORANGE, ah="center", brd="bk")
    mc(r,2,r,3, f"{int(round(pct))}%  —  {verb}",
       bold=True, sz=SZ, color=sc_c, bg=sbg, ah="center", brd="bk")
    r += 1

    # ════════════════════════════════════════════════════════════════════
    # مؤشرات الأداء الوظيفي
    # ════════════════════════════════════════════════════════════════════
    ws.row_dimensions[r].height = 16
    sc(ws.cell(r,1,"مؤشرات الأداء الوظيفي"),
       bold=True, sz=SZ, color="FFFFFF", bg=DARK, ah="right", brd="bk")
    sc(ws.cell(r,2,"الوزن النسبي %"), bold=True, sz=SZ, color="FFFFFF", bg=DARK, ah="center", brd="bk")
    sc(ws.cell(r,3,"الدرجة (0-100)"), bold=True, sz=SZ, color="FFFFFF", bg=DARK, ah="center", brd="bk")
    sc(ws.cell(r,4,"التقييم"),        bold=True, sz=SZ, color="FFFFFF", bg=DARK, ah="center", brd="bk")
    r += 1

    _job_total_score = 0.0
    for i,(kname,weight,grade) in enumerate(job_kpis):
        rbg  = LGRAY if i%2==0 else WHITE
        g    = float(grade)
        w    = float(weight)
        pct  = round(kpi_score_to_pct(g, w), 1)
        lbl  = rating_label(pct)
        actual = round(g, 2)
        _job_total_score += actual
        kbg  = GREEN_BG if pct>=80 else (YELLOW if pct>=60 else (RED_BG if pct>0 else rbg))
        ws.row_dimensions[r].height = 16
        sc(ws.cell(r,1,kname), bold=True, sz=10, color="000000",
           bg=rbg, ah="right", av="center", brd="tn", wrap=False)
        sc(ws.cell(r,2,f"{w:.1f}%"),
           bold=True, sz=10, bg=rbg, ah="center", brd="tn")
        sc(ws.cell(r,3,pct),
           bold=True, sz=10, bg=kbg, ah="center", brd="tn")
        sc(ws.cell(r,4,lbl),
           bold=True, sz=10, bg=kbg, ah="center", brd="tn")
        r += 1

    # صف مجموع الأداء الوظيفي
    ws.row_dimensions[r].height = 16
    sc(ws.cell(r,1,"مجموع الأداء الوظيفي"),
       bold=True, sz=SZ, color="FFFFFF", bg=MID, ah="right", brd="bk")
    sc(ws.cell(r,2,"80%"),
       bold=True, sz=SZ, color="FFFFFF", bg=MID, ah="center", brd="bk")
    _job_pct_total = round(kpi_score_to_pct(_job_total_score, 80.0), 1)
    sc(ws.cell(r,3,f"{_job_pct_total}%"),
       bold=True, sz=SZ, color="FFFFFF", bg=MID, ah="center", brd="bk")
    sc(ws.cell(r,4,rating_label(_job_pct_total)),
       bold=True, sz=SZ, color="FFFFFF", bg=MID, ah="center", brd="bk")
    r += 1

    ws.row_dimensions[r].height = 3; r += 1

    # ════════════════════════════════════════════════════════════════════
    # مؤشرات الصفات الشخصية
    # ════════════════════════════════════════════════════════════════════
    ws.row_dimensions[r].height = 16
    mc(r,1,r,3,"مؤشرات الصفات الشخصية",
       bold=True, sz=SZ, color="FFFFFF", bg=ORANGE, ah="center", brd="bk")
    r += 1
    ws.row_dimensions[r].height = 16
    sc(ws.cell(r,1,"المؤشر"),           bold=True, sz=SZ, color="FFFFFF", bg=MID, ah="right",  brd="bk")
    sc(ws.cell(r,2,"الوزن النسبي %"),  bold=True, sz=SZ, color="FFFFFF", bg=MID, ah="center", brd="bk")
    sc(ws.cell(r,3,"الدرجة (0-100)"),  bold=True, sz=SZ, color="FFFFFF", bg=MID, ah="center", brd="bk")
    sc(ws.cell(r,4,"التقييم"),          bold=True, sz=SZ, color="FFFFFF", bg=MID, ah="center", brd="bk")
    r += 1

    _per_total_score = 0.0
    for i,(kname,weight,grade) in enumerate(per_kpis):
        rbg  = WARM if i%2==0 else WHITE
        g    = float(grade)
        w    = float(weight)
        pct  = round(kpi_score_to_pct(g, w), 1)
        lbl  = rating_label(pct)
        actual = round(g, 2)
        _per_total_score += actual
        kbg  = GREEN_BG if pct>=80 else (YELLOW if pct>=60 else (RED_BG if pct>0 else rbg))
        ws.row_dimensions[r].height = 16
        sc(ws.cell(r,1,kname), bold=True, sz=10, color="000000",
           bg=rbg, ah="right", av="center", brd="tn", wrap=False)
        sc(ws.cell(r,2,f"{w:.1f}%"),
           bold=True, sz=10, bg=rbg, ah="center", brd="tn")
        sc(ws.cell(r,3,pct),
           bold=True, sz=10, bg=kbg, ah="center", brd="tn")
        sc(ws.cell(r,4,lbl),
           bold=True, sz=10, bg=kbg, ah="center", brd="tn")
        r += 1

    # صف مجموع الصفات الشخصية
    ws.row_dimensions[r].height = 16
    sc(ws.cell(r,1,"مجموع الصفات الشخصية"),
       bold=True, sz=SZ, color="FFFFFF", bg=ORANGE, ah="right", brd="bk")
    sc(ws.cell(r,2,"20%"),
       bold=True, sz=SZ, color="FFFFFF", bg=ORANGE, ah="center", brd="bk")
    _per_pct_total = round(kpi_score_to_pct(_per_total_score, 20.0), 1)
    sc(ws.cell(r,3,f"{_per_pct_total}%"),
       bold=True, sz=SZ, color="FFFFFF", bg=ORANGE, ah="center", brd="bk")
    sc(ws.cell(r,4,rating_label(_per_pct_total)),
       bold=True, sz=SZ, color="FFFFFF", bg=ORANGE, ah="center", brd="bk")
    r += 1

    ws.row_dimensions[r].height = 3; r += 1

    # ════════════════════════════════════════════════════════════════════
    # ملاحظات المقيم
    # ════════════════════════════════════════════════════════════════════
    ws.row_dimensions[r].height = 22
    mc(r,1,r,3, f"ملاحظات المقيم: {notes or ''}",
       bold=False, sz=SZ, color="000000", bg=NOTE_BG,
       ah="right", av="center", brd="bk", wrap=True)
    r += 1

    # الاحتياجات التدريبية
    _train_vals = [v for v in m_train.values() if v and str(v).strip() not in ("","nan","None","—")]
    _train = (
        _train_vals[0] if _train_vals
        else (training.strip() if training and training.strip() not in ("","nan","None")
              else "")
    )
    _train_label = f"الاحتياجات التدريبية: {_train}" if _train else "الاحتياجات التدريبية:"
    ws.row_dimensions[r].height = 30
    mc(r,1,r,3, _train_label,
       bold=False, sz=SZ, color="000000", bg=TRAIN_BG,
       ah="right", av="center", brd="bk", wrap=True)
    r += 1
    ws.row_dimensions[r].height = 3; r += 1

    # ════════════════════════════════════════════════════════════════════
    # ⚠️ الإجراءات التأديبية (جديدة)
    # ════════════════════════════════════════════════════════════════════
    if disciplinary_actions is not None and not disciplinary_actions.empty:
        ws.row_dimensions[r].height = 18
        mc(r, 1, r, 4, "⚠️ الإجراءات التأديبية المسجلة",
           bold=True, sz=SZ, color="FFFFFF", bg="B91C1C", ah="center", brd="bk")
        r += 1
        
        # رأس جدول الإجراءات
        ws.row_dimensions[r].height = 14
        for ci, hdr in enumerate(["التاريخ", "نوع الإنذار", "السبب", "خصم (أيام)"], 1):
            sc(ws.cell(r, ci, hdr), bold=True, sz=8, color="FFFFFF",
               bg="7F1D1D", ah="center", brd="inner")
        r += 1
        
        # بيانات الإجراءات
        for idx, (_, row_disc) in enumerate(disciplinary_actions.iterrows()):
            ws.row_dimensions[r].height = 14
            rbg = DISC_BG if idx % 2 == 0 else "FFFFFF"
            
            warning_date = row_disc.get("warning_date", "")
            warning_type = row_disc.get("warning_type", "")
            reason = row_disc.get("reason", "")
            deduction = row_disc.get("deduction_days", 0)
            
            # تحويل التاريخ إلى نص
            if hasattr(warning_date, 'strftime'):
                date_str = warning_date.strftime("%Y-%m-%d")
            else:
                date_str = str(warning_date)
            
            sc(ws.cell(r, 1, date_str),
               bold=False, sz=8, bg=rbg, ah="center", brd="inner")
            sc(ws.cell(r, 2, str(warning_type)),
               bold=False, sz=8, bg=rbg, ah="center", brd="inner")
            sc(ws.cell(r, 3, str(reason)),
               bold=False, sz=8, bg=rbg, ah="right", brd="inner", wrap=True)
            sc(ws.cell(r, 4, str(deduction)),
               bold=False, sz=8, bg=rbg, ah="center", brd="inner")
            r += 1
        
        ws.row_dimensions[r].height = 3
        r += 1

    # ════════════════════════════════════════════════════════════════════
    # التوقيع
    # ════════════════════════════════════════════════════════════════════
    ws.row_dimensions[r].height = 18
    sc(ws.cell(r,1, f"المسؤول المباشر: {manager}"),
       bold=True, sz=SZ, ah="center", brd="bk")
    sc(ws.cell(r,2,"اسم الموظف"), bold=True, sz=SZ, ah="center", brd="bk")
    _ec = ws.cell(r,3,emp_name)
    sc(_ec, bold=True, sz=SZ, ah="right", brd="bk")
    _ec.border = Border(left=_med, right=_med, top=_med, bottom=_med)
    r += 1
    
    ws.row_dimensions[r].height = 18
    _t1 = ws.cell(r,1,"التوقيع: _______________")
    sc(_t1, bold=True, sz=SZ, bg=LGRAY, ah="center", brd="bk")
    _t1.border = Border(left=_med, right=_med, top=_med, bottom=_med)
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
    _t2 = ws.cell(r,2,"التوقيع: _______________")
    sc(_t2, bold=True, sz=SZ, bg=LGRAY, ah="center", brd="bk")
    _t2.border = Border(left=_med, right=_med, top=_med, bottom=_med)

    # ════════════════════════════════════════════════════════════════════
    # إعداد الصفحة
    # ════════════════════════════════════════════════════════════════════
    if chart_img:
        try:
            import io as _io
            from openpyxl.drawing.image import Image as _ChartImg
            _cbuf = _io.BytesIO(chart_img)
            _cimg = _ChartImg(_cbuf)
            _W_CM = 9.8
            _H_CM = 4.9
            _PX   = 37.795
            _cimg.width  = int(_W_CM * _PX)
            _cimg.height = int(_H_CM * _PX)
            _cimg.anchor = "E16"
            ws.add_image(_cimg)
        except Exception as _ce:
            pass

    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize   = 9
    ws.page_setup.fitToPage   = True
    ws.page_setup.fitToWidth  = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins = PageMargins(
        left=0.5, right=0.5, top=0.5, bottom=0.5,
        header=0.0, footer=0.0
    )
    ws.print_options.horizontalCentered = True
    ws.print_options.verticalCentered   = True

    # ── ضبط إجباري للصفوف ──────────────────────────────────────────────
    for _rn in range(2, 50):
        _cur = ws.row_dimensions[_rn].height
        if _cur is None or _cur < 20:
            ws.row_dimensions[_rn].height = 20

    # ── موائمة ذكية لعمود A ───────────────────────────────────────────
    all_kpi = [str(k) for (k,_,_) in job_kpis+per_kpis if str(k).strip()]
    mk = max((len(t) for t in all_kpi), default=20)
    _col_a_w = mk * 0.62 + 4
    ws.column_dimensions["A"].width = min(max(_col_a_w, 28), 80)

    # ── حدود شاملة ────────────────────────────────────────────────────
    blk_t = Side(style="thin", color="000000")
    fb    = Border(left=blk_t, right=blk_t, top=blk_t, bottom=blk_t)
    for row in ws.iter_rows():
        for cell in row:
            if cell.value is not None and str(cell.value).strip() != "":
                b = cell.border
                if not any([b.left  and b.left.style,
                             b.right and b.right.style,
                             b.top   and b.top.style,
                             b.bottom and b.bottom.style]):
                    cell.border = fb

    return ws


# ══════════════════════════════════════════════════════════════════════
# ملخص القسم / الكل
# ══════════════════════════════════════════════════════════════════════
def build_summary_sheet(wb, rows, title="ملخص التقييم", year=None, chart_img=None):
    ws = wb.create_sheet(title[:28])
    ws.sheet_view.rightToLeft  = True
    ws.sheet_view.showGridLines = False

    DARK="1F3864"; LGRAY="F2F2F2"; WHITE="FFFFFF"; YELLOW="FFF2CC"
    GREEN_BG="E2EFDA"; RED_BG="FFDAD9"

    for col,w in [("A",4),("B",32),("C",14),("D",8),("E",10),("F",13),("G",13)]:
        ws.column_dimensions[col].width = w

    ws.row_dimensions[1].height = 36
    _mc(ws,1,1,1,7,title,bold=True,sz=12,color="FFFFFF",
        bg=DARK,ah="center",av="center",brd="outer")

    import os as _os2
    from openpyxl.drawing.image import Image as XLImg2
    _logo2=globals().get("LOGO_PATH","logo.png")
    if _logo2 and _os2.path.exists(_logo2):
        try:
            _img=XLImg2(_logo2); _img.height=32; _img.width=26; _img.anchor="G1"
            ws.add_image(_img)
        except: pass

    ws.row_dimensions[2].height = 4
    ws.row_dimensions[3].height = 16
    for c,t in [(1,"#"),(2,"اسم الموظف"),(3,"القسم"),(4,"السنة"),(5,"الأشهر"),(6,"المعدل %"),(7,"التقييم")]:
        _sc(ws.cell(3,c,t),bold=True,sz=9,color="FFFFFF",bg=DARK,ah="center",brd="outer")

    for i,(name,dept,months,pct,verb) in enumerate(rows,4):
        ws.row_dimensions[i].height = 16
        rbg  = LGRAY if i%2==0 else WHITE
        sc_c = "375623" if pct>=80 else ("C00000" if pct<60 else "000000")
        vbg  = GREEN_BG if pct>=80 else (YELLOW if pct>=70 else (RED_BG if pct<60 else LGRAY))
        _sc(ws.cell(i,1,i-3), sz=8,  ah="center",bg=rbg,brd="inner")
        _sc(ws.cell(i,2,name),sz=9,  ah="right", bg=rbg,brd="inner")
        _sc(ws.cell(i,3,dept),sz=8,  ah="center",bg=rbg,brd="inner")
        _sc(ws.cell(i,4,year or ""),sz=9,ah="center",bg=rbg,brd="inner")
        _sc(ws.cell(i,5,months),    sz=9,ah="center",bg=rbg,brd="inner")
        _sc(ws.cell(i,6,f"{pct:.1f}%"),sz=10,bold=True,color=sc_c,ah="center",bg=vbg,brd="inner")
        _sc(ws.cell(i,7,verb),sz=9,bold=True,color=sc_c,ah="center",bg=vbg,brd="inner")

    last=3+len(rows)
    for r in range(3,last+1):
        lft=ws.cell(r,1).border
        ws.cell(r,1).border=Border(left=thick_s,right=lft.right,top=lft.top,bottom=lft.bottom)
        rgt=ws.cell(r,7).border
        ws.cell(r,7).border=Border(left=rgt.left,right=thick_s,top=rgt.top,bottom=rgt.bottom)
    for c in range(1,8):
        b=ws.cell(last,c).border
        ws.cell(last,c).border=Border(left=b.left,right=b.right,top=b.top,bottom=thick_s)

    # حدود شاملة
    _blk=Side(style="thin",color="000000")
    _fb=Border(left=_blk,right=_blk,top=_blk,bottom=_blk)
    for row in ws.iter_rows():
        for cell in row:
            if cell.value is not None and str(cell.value).strip()!="":
                b=cell.border
                if not any([b.left and b.left.style,b.right and b.right.style,
                             b.top and b.top.style,b.bottom and b.bottom.style]):
                    cell.border=_fb

    if chart_img:
        try:
            import io as _io
            from openpyxl.drawing.image import Image as _ChartImg
            _cbuf = _io.BytesIO(chart_img)
            _cimg = _ChartImg(_cbuf)
            _W_CM = 9.8
            _H_CM = 4.9
            _PX   = 37.795
            _cimg.width  = int(_W_CM * _PX)
            _cimg.height = int(_H_CM * _PX)
            _cimg.anchor = "E16"
            ws.add_image(_cimg)
        except Exception as _ce:
            pass

    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize   = 9
    ws.page_setup.fitToPage   = True
    ws.page_setup.fitToWidth  = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins = PageMargins(
        left=0.5, right=0.5, top=0.5, bottom=0.5,
        header=0.0, footer=0.0
    )
    ws.print_options.horizontalCentered = True
    ws.print_options.verticalCentered   = True
