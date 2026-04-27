import os
from datetime import date
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.drawing.image import Image as XLImage
from openpyxl.worksheet.page import PageMargins
from openpyxl.utils import get_column_letter
from constants import *
from calculations import verbal_grade, kpi_score_to_pct, rating_label
from excel_reports import print_preview_html

def build_employee_sheet(wb, emp_name, job_title, dept, manager, year, kpis, monthly_scores, notes="", training="", chart_img=None, disciplinary_actions=None, employee_id=""):
    safe = emp_name[:28]
    if safe in [s.title for s in wb.worksheets]:
        safe = safe[:25]+"_2"
    ws = wb.create_sheet(safe)
    ws.sheet_view.rightToLeft = True
    ws.sheet_view.showGridLines = False

    # ألوان
    DARK, MID, ORANGE = "1F3864","2E75B6","ED7D31"
    LGRAY, WHITE, YELLOW, GREEN_BG, RED_BG = "F2F2F2","FFFFFF","FFF2CC","E2EFDA","FFDAD9"
    WARM, NOTE_BG, TRAIN_BG, DATE_BG, INFO_BG, DISC_BG = "FFF3E0","FFFDE7","F3E5F5","E3F2FD","EBF3FB","FEE2E2"

    _med, _thn = Side(style="medium"), Side(style="thin")
    BK, TN = Border(left=_med,right=_med,top=_med,bottom=_med), Border(left=_thn,right=_thn,top=_thn,bottom=_thn)

    def sc(cell, val=None, bold=False, sz=9, color="000000", bg=None, ah="right", av="center", wrap=False):
        if val is not None: cell.value = val
        cell.font = Font(name="Arial", bold=bold, size=sz, color=color)
        cell.alignment = Alignment(horizontal=ah, vertical=av, wrapText=wrap, readingOrder=2)
        if bg: cell.fill = PatternFill("solid", fgColor=bg)
    def mc(r1,c1,r2,c2,val=None,**kw):
        ws.merge_cells(start_row=r1,start_column=c1,end_row=r2,end_column=c2)
        sc(ws.cell(r1,c1,val),**kw)

    # تجهيز بيانات الأشهر
    m_score,m_date,m_note,m_train = {},{},{},{}
    for item in monthly_scores:
        ms = item[1]
        def _v(x): return str(x).strip() if x not in (None,"nan","None") else ""
        m_score[ms] = item[2]
        m_date[ms] = _v(item[3]) if len(item)>3 else ""
        m_note[ms] = _v(item[4]) if len(item)>4 else ""
        m_train[ms] = _v(item[5]) if len(item)>5 else ""

    done = [(n,m,s) for n,m,s,*_ in monthly_scores if s>0]
    pct = sum(s for _,_,s in done)/len(done)*100 if done else 0
    verb = verbal_grade(pct)
    sc_c = "375623" if pct>=80 else ("C00000" if pct<60 else "7F6000")
    sbg = GREEN_BG if pct>=80 else (YELLOW if pct>=60 else RED_BG)

    job_kpis = [(k["KPI_Name"],k["Weight"],k.get("avg_score",0)) for k in kpis if k["KPI_Name"] not in PERSONAL_KPIS]
    per_kpis = [(k["KPI_Name"],k["Weight"],k.get("avg_score",0)) for k in kpis if k["KPI_Name"] in PERSONAL_KPIS]

    _MAR = {"Jan":"يناير","Feb":"فبراير","Mar":"مارس","Apr":"أبريل","May":"مايو","Jun":"يونيو","Jul":"يوليو","Aug":"أغسطس","Sep":"سبتمبر","Oct":"أكتوبر","Nov":"نوفمبر","Dec":"ديسمبر"}
    _company,_branch="",""
    try:
        from auth import load_app_settings as _las
        _cfg = _las()
        _company = _cfg.get("company_name","مجموعة شركات فنون")
        _branch = _cfg.get("branch_name","")
    except:
        pass
    _header = f"نموذج تقييم الأداء السنوي — {_company}" + (f" — {_branch}" if _branch else "")

    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 18
    ws.column_dimensions["D"].width = 18
    ws.column_dimensions["E"].width = 18
    ws.column_dimensions["F"].width = 18
    ws.column_dimensions["G"].width = 20
    ws.column_dimensions["H"].width = 20

    r=1
    ws.row_dimensions[1].height = 32
    mc(1,1,1,8, _header, bold=True, sz=11, color="FFFFFF", bg=DARK, ah="center")
    _logo = globals().get("LOGO_PATH","logo.png")
    if os.path.exists(_logo):
        try:
            img = XLImage(_logo)
            img.height, img.width = 70,56
            img.anchor = "A1"
            ws.add_image(img)
        except: pass

    r=2
    INFO = [("اسم الموظف", emp_name), ("رقم الموظف", employee_id), ("الوظيفة", job_title), ("القسم", dept), ("السنة", str(year)), ("اسم المقيم", manager), ("تاريخ التقييم", date.today().strftime("%d/%m/%Y"))]
    for i,(lbl,val) in enumerate(INFO):
        row = r+i
        ws.row_dimensions[row].height = 18
        sc(ws.cell(row,1,lbl), bold=True, color="FFFFFF", bg=DARK, ah="center")
        mc(row,2,row,8, val, bold=True, color="000000", bg=INFO_BG, ah="right")
    r += len(INFO)

    # نتيجة التقييم السنوي
    ws.row_dimensions[r].height = 18
    sc(ws.cell(r,1,"نتيجة التقييم السنوي"), bold=True, color="FFFFFF", bg=ORANGE, ah="center")
    mc(r,2,r,3, f"{int(round(pct))}% — {verb}", bold=True, sz=11, color=sc_c, bg=sbg, ah="center")
    r+=1

    # جدول مؤشرات الأداء الوظيفي
    ws.row_dimensions[r].height = 16
    sc(ws.cell(r,1,"مؤشرات الأداء الوظيفي"), bold=True, color="FFFFFF", bg=DARK, ah="right")
    sc(ws.cell(r,2,"الوزن النسبي %"), bold=True, color="FFFFFF", bg=DARK, ah="center")
    sc(ws.cell(r,3,"الدرجة (0-100)"), bold=True, color="FFFFFF", bg=DARK, ah="center")
    sc(ws.cell(r,4,"التقييم"), bold=True, color="FFFFFF", bg=DARK, ah="center")
    r+=1

    job_total_score, job_total_weight = 0.0, 0.0
    for i,(kname,weight,grade) in enumerate(job_kpis):
        rbg = LGRAY if i%2==0 else WHITE
        w,g = float(weight), float(grade)
        pct_val = round(kpi_score_to_pct(g,w),1)
        lbl = rating_label(pct_val)
        job_total_score += g
        job_total_weight += w
        kbg = GREEN_BG if pct_val>=80 else (YELLOW if pct_val>=60 else (RED_BG if pct_val>0 else rbg))
        ws.row_dimensions[r].height = 16
        sc(ws.cell(r,1,kname), bg=rbg, wrap=True)
        sc(ws.cell(r,2,f"{w:.1f}%"), bg=rbg, ah="center")
        sc(ws.cell(r,3,pct_val), bold=True, bg=kbg, ah="center")
        sc(ws.cell(r,4,lbl), bold=True, bg=kbg, ah="center")
        r+=1

    ws.row_dimensions[r].height = 16
    sc(ws.cell(r,1,"مجموع الأداء الوظيفي"), bold=True, color="FFFFFF", bg=MID, ah="right")
    sc(ws.cell(r,2,f"{job_total_weight:.1f}%"), bold=True, color="FFFFFF", bg=MID, ah="center")
    job_pct_total = round(kpi_score_to_pct(job_total_score, job_total_weight),1) if job_total_weight>0 else 0
    sc(ws.cell(r,3,f"{job_pct_total}%"), bold=True, color="FFFFFF", bg=MID, ah="center")
    sc(ws.cell(r,4,rating_label(job_pct_total)), bold=True, color="FFFFFF", bg=MID, ah="center")
    r+=1
    ws.row_dimensions[r].height = 3; r+=1

    # مؤشرات الصفات الشخصية
    ws.row_dimensions[r].height = 16
    mc(r,1,r,3,"مؤشرات الصفات الشخصية", bold=True, color="FFFFFF", bg=ORANGE, ah="center")
    r+=1
    ws.row_dimensions[r].height = 16
    sc(ws.cell(r,1,"المؤشر"), bold=True, color="FFFFFF", bg=MID, ah="right")
    sc(ws.cell(r,2,"الوزن النسبي %"), bold=True, color="FFFFFF", bg=MID, ah="center")
    sc(ws.cell(r,3,"الدرجة (0-100)"), bold=True, color="FFFFFF", bg=MID, ah="center")
    sc(ws.cell(r,4,"التقييم"), bold=True, color="FFFFFF", bg=MID, ah="center")
    r+=1

    per_total_score, per_total_weight = 0.0, 0.0
    for i,(kname,weight,grade) in enumerate(per_kpis):
        rbg = WARM if i%2==0 else WHITE
        w,g = float(weight), float(grade)
        pct_val = round(kpi_score_to_pct(g,w),1)
        lbl = rating_label(pct_val)
        per_total_score += g
        per_total_weight += w
        kbg = GREEN_BG if pct_val>=80 else (YELLOW if pct_val>=60 else (RED_BG if pct_val>0 else rbg))
        ws.row_dimensions[r].height = 16
        sc(ws.cell(r,1,kname), bg=rbg, wrap=True)
        sc(ws.cell(r,2,f"{w:.1f}%"), bg=rbg, ah="center")
        sc(ws.cell(r,3,pct_val), bold=True, bg=kbg, ah="center")
        sc(ws.cell(r,4,lbl), bold=True, bg=kbg, ah="center")
        r+=1

    ws.row_dimensions[r].height = 16
    sc(ws.cell(r,1,"مجموع الصفات الشخصية"), bold=True, color="FFFFFF", bg=ORANGE, ah="right")
    sc(ws.cell(r,2,f"{per_total_weight:.1f}%"), bold=True, color="FFFFFF", bg=ORANGE, ah="center")
    per_pct_total = round(kpi_score_to_pct(per_total_score, per_total_weight),1) if per_total_weight>0 else 0
    sc(ws.cell(r,3,f"{per_pct_total}%"), bold=True, color="FFFFFF", bg=ORANGE, ah="center")
    sc(ws.cell(r,4,rating_label(per_pct_total)), bold=True, color="FFFFFF", bg=ORANGE, ah="center")
    r+=1
    ws.row_dimensions[r].height = 3; r+=1

    # ملاحظات المقيم
    ws.row_dimensions[r].height = 22
    mc(r,1,r,4, f"ملاحظات المقيم: {notes or ''}", bg=NOTE_BG, wrap=True)
    r+=1

    # الاحتياجات التدريبية
    _train_vals = [v for v in m_train.values() if v and str(v).strip() not in ("","nan","None","—")]
    _train = _train_vals[0] if _train_vals else (training if training else "")
    mc(r,1,r,4, f"الاحتياجات التدريبية: {_train}" if _train else "الاحتياجات التدريبية:", bg=TRAIN_BG, wrap=True)
    r+=1
    ws.row_dimensions[r].height = 3; r+=1

    # الإجراءات التأديبية
    if disciplinary_actions is not None and not disciplinary_actions.empty:
        ws.row_dimensions[r].height = 18
        mc(r,1,r,4,"⚠️ الإجراءات التأديبية المسجلة", bold=True, color="FFFFFF", bg="B91C1C", ah="center")
        r+=1
        ws.row_dimensions[r].height = 14
        for ci,hdr in enumerate(["التاريخ","نوع الإنذار","السبب","خصم (أيام)"],1):
            sc(ws.cell(r,ci,hdr), bold=True, sz=8, color="FFFFFF", bg="7F1D1D", ah="center")
        r+=1
        for idx,(_,row_disc) in enumerate(disciplinary_actions.iterrows()):
            ws.row_dimensions[r].height = 14
            rbg = DISC_BG if idx%2==0 else "FFFFFF"
            sc(ws.cell(r,1,str(row_disc.get("action_date",""))), bg=rbg, ah="center")
            sc(ws.cell(r,2,str(row_disc.get("warning_type",""))), bg=rbg, ah="center")
            sc(ws.cell(r,3,str(row_disc.get("reason",""))), bg=rbg, wrap=True)
            sc(ws.cell(r,4,str(row_disc.get("deduction_days",0))), bg=rbg, ah="center")
            r+=1
        ws.row_dimensions[r].height = 3; r+=1

    # التوقيع
    ws.row_dimensions[r].height = 18
    sc(ws.cell(r,1,f"المسؤول المباشر: {manager}"), bold=True, ah="center")
    sc(ws.cell(r,2,"اسم الموظف"), bold=True, ah="center")
    _ec = ws.cell(r,3,emp_name)
    sc(_ec, bold=True, ah="right")
    _ec.border = Border(left=_med,right=_med,top=_med,bottom=_med)
    r+=1
    ws.row_dimensions[r].height = 18
    _t1 = ws.cell(r,1,"التوقيع: _______________")
    sc(_t1, bold=True, bg=LGRAY, ah="center")
    _t1.border = Border(left=_med,right=_med,top=_med,bottom=_med)
    ws.merge_cells(start_row=r,start_column=2,end_row=r,end_column=3)
    _t2 = ws.cell(r,2,"التوقيع: _______________")
    sc(_t2, bold=True, bg=LGRAY, ah="center")
    _t2.border = Border(left=_med,right=_med,top=_med,bottom=_med)

    # إعداد الطباعة
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = 9
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.page_margins = PageMargins(left=0.5,right=0.5,top=0.5,bottom=0.5)
    ws.print_options.horizontalCentered = True
    ws.print_options.verticalCentered = True
    return ws
