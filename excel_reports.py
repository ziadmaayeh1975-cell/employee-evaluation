import io
import os
import base64
from datetime import date
import openpyxl
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side, numbers
)
from openpyxl.utils import get_column_letter
from constants import (
    LOGO_PATH, MONTHS_AR, MONTHS_EN, MONTHS_SHORT, PERSONAL_KPIS,
    DARK, MID, LBLUE, ORANGE, YELLOW, LGRAY, GREEN_BG, RED_BG, WHITE, CREAM,
    OUTER_B, INNER_B, ROW_B,
)


# ─────────────────────────────────────────────────────────────────
# دالة بناء شيت تقييم الموظف في Excel
# ─────────────────────────────────────────────────────────────────
def build_employee_sheet(
    wb,
    emp_name, job, dept, mgr, year,
    kpis,          # list of (kpi_name, weight, grade)
    monthly_rep,   # list of (idx, short, score, eval_date, notes, training)
    notes="",
    training="",
):
    """
    يضيف شيت تقييم الموظف إلى workbook موجود.
    المعاملات:
        wb          : openpyxl Workbook
        emp_name    : اسم الموظف
        job         : المسمى الوظيفي
        dept        : القسم
        mgr         : المدير المباشر
        year        : السنة
        kpis        : قائمة (kpi_name, weight, grade)
        monthly_rep : قائمة (idx, short, score, eval_date, notes, training)
        notes       : ملاحظات إضافية
        training    : الاحتياجات التدريبية
    """
    from calculations import verbal_grade, grade_color_hex, kpi_score_to_pct, rating_label

    ws = wb.create_sheet(title=emp_name[:28])

    # ── ألوان مساعدة ──────────────────────────────────────────────
    def fill(hex_color):
        return PatternFill("solid", fgColor=hex_color.lstrip("#"))

    def font(bold=False, color="000000", size=10):
        return Font(bold=bold, color=color.lstrip("#"), size=size, name="Cairo")

    def align(h="center", v="center", wrap=False):
        return Alignment(horizontal=h, vertical=v,
                         wrap_text=wrap, readingOrder=2)

    thin = Side(style="thin",   color="AAAAAA")
    med  = Side(style="medium", color="1F3864")
    inner_b = Border(left=thin, right=thin, top=thin,  bottom=thin)
    outer_b = Border(left=med,  right=med,  top=med,   bottom=med)

    # ── عرض الأعمدة ──────────────────────────────────────────────
    col_widths = [4, 22, 8, 10, 10, 12, 14, 14]
    for i, w in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    row = 1

    # ══════════════════════════════════════════════════════════════
    # رأس التقرير
    # ══════════════════════════════════════════════════════════════
    ws.merge_cells(f"A{row}:H{row}")
    c = ws.cell(row, 1, "نموذج تقييم الأداء السنوي — مجموعة شركات فنون")
    c.font      = font(bold=True, color=WHITE, size=13)
    c.fill      = fill(DARK)
    c.alignment = align()
    ws.row_dimensions[row].height = 24
    row += 1

    # بيانات الموظف
    info = [
        ("الاسم",    emp_name),
        ("الوظيفة",  job),
        ("القسم",    dept),
        ("المدير",   mgr),
        ("السنة",    str(year)),
        ("التاريخ",  date.today().strftime("%d/%m/%Y")),
    ]
    for label, val in info:
        ws.merge_cells(f"A{row}:B{row}")
        lc = ws.cell(row, 1, label)
        lc.font = font(bold=True, color=DARK, size=10)
        lc.fill = fill(LBLUE)
        lc.alignment = align(h="right")
        ws.merge_cells(f"C{row}:H{row}")
        vc = ws.cell(row, 3, val)
        vc.font = font(size=10)
        vc.fill = fill(CREAM)
        vc.alignment = align(h="right")
        for col in range(1, 9):
            ws.cell(row, col).border = inner_b
        ws.row_dimensions[row].height = 16
        row += 1

    row += 1  # فراغ

    # ══════════════════════════════════════════════════════════════
    # جدول التقييم الشهري
    # ══════════════════════════════════════════════════════════════
    ws.merge_cells(f"A{row}:H{row}")
    c = ws.cell(row, 1, "التقييم الشهري")
    c.font = font(bold=True, color=WHITE, size=11)
    c.fill = fill(MID)
    c.alignment = align()
    ws.row_dimensions[row].height = 18
    row += 1

    headers_m = ["#", "الشهر", "الدرجة الخام", "الدرجة %", "التقييم اللفظي",
                 "تاريخ التقييم", "ملاحظات", "تدريب"]
    for ci, h in enumerate(headers_m, 1):
        c = ws.cell(row, ci, h)
        c.font = font(bold=True, color=WHITE, size=9)
        c.fill = fill(DARK)
        c.alignment = align()
        c.border = inner_b
    ws.row_dimensions[row].height = 15
    row += 1

    done_scores = []
    for idx, short, score, ev_date, nm, tr in monthly_rep:
        pct = round(score * 100, 1)
        verbal = verbal_grade(pct) if pct > 0 else ""
        bg = GREEN_BG if pct >= 70 else (YELLOW if pct >= 50 else (RED_BG if pct > 0 else "FFFFFF"))
        vals = [idx, MONTHS_AR[idx-1], round(score, 4), pct, verbal, ev_date or "", nm or "", tr or ""]
        for ci, v in enumerate(vals, 1):
            c = ws.cell(row, ci, v)
            c.font = font(size=9)
            c.fill = fill(bg)
            c.alignment = align(h="center" if ci in (1,3,4,6) else "right")
            c.border = inner_b
        ws.row_dimensions[row].height = 14
        if pct > 0:
            done_scores.append(pct)
        row += 1

    # متوسط شهري
    avg_pct = round(sum(done_scores) / len(done_scores), 1) if done_scores else 0.0
    avg_verbal = verbal_grade(avg_pct)
    avg_bg = GREEN_BG if avg_pct >= 70 else (YELLOW if avg_pct >= 50 else (RED_BG if avg_pct > 0 else LGRAY))
    ws.merge_cells(f"A{row}:C{row}")
    c = ws.cell(row, 1, "المتوسط السنوي")
    c.font = font(bold=True, color=DARK, size=10)
    c.fill = fill(LGRAY)
    c.alignment = align()
    ws.cell(row, 4, avg_pct).font = font(bold=True, size=10)
    ws.cell(row, 4).fill = fill(avg_bg)
    ws.cell(row, 4).alignment = align()
    ws.cell(row, 5, avg_verbal).font = font(bold=True, size=10)
    ws.cell(row, 5).fill = fill(avg_bg)
    ws.cell(row, 5).alignment = align()
    for ci in range(1, 9):
        ws.cell(row, ci).border = outer_b
    ws.row_dimensions[row].height = 16
    row += 2

    # ══════════════════════════════════════════════════════════════
    # جدول مؤشرات الأداء
    # ══════════════════════════════════════════════════════════════
    job_kpis  = [(k, w, g) for k, w, g in kpis if k not in PERSONAL_KPIS]
    pers_kpis = [(k, w, g) for k, w, g in kpis if k in PERSONAL_KPIS]

    for section_title, section_kpis, hdr_color in [
        ("مؤشرات الأداء الوظيفي",   job_kpis,  DARK),
        ("مؤشرات الصفات الشخصية",   pers_kpis, ORANGE),
    ]:
        if not section_kpis:
            continue
        ws.merge_cells(f"A{row}:H{row}")
        c = ws.cell(row, 1, section_title)
        c.font = font(bold=True, color=WHITE, size=11)
        c.fill = fill(hdr_color)
        c.alignment = align()
        ws.row_dimensions[row].height = 18
        row += 1

        headers_k = ["#", "اسم المؤشر", "الوزن", "الدرجة الخام",
                     "الدرجة (0-100)", "التقييم", "", ""]
        for ci, h in enumerate(headers_k, 1):
            c = ws.cell(row, ci, h)
            c.font = font(bold=True, color=WHITE, size=9)
            c.fill = fill(MID if hdr_color == DARK else "C55A11")
            c.alignment = align()
            c.border = inner_b
        ws.row_dimensions[row].height = 15
        row += 1

        for ki, (kname, weight, grade) in enumerate(section_kpis, 1):
            pct100 = round(kpi_score_to_pct(grade, weight), 1)
            rlabel = rating_label(pct100)
            bg = GREEN_BG if pct100 >= 70 else (YELLOW if pct100 >= 50 else (RED_BG if pct100 > 0 else "FFFFFF"))
            vals = [ki, kname, weight, round(grade, 3), pct100, rlabel, "", ""]
            for ci, v in enumerate(vals, 1):
                c = ws.cell(row, ci, v)
                c.font = font(size=9)
                c.fill = fill(bg)
                c.alignment = align(h="right" if ci == 2 else "center", wrap=(ci == 2))
                c.border = inner_b
            ws.row_dimensions[row].height = 14
            row += 1
        row += 1

    # ══════════════════════════════════════════════════════════════
    # ملاحظات واحتياجات تدريبية
    # ══════════════════════════════════════════════════════════════
    if notes or training:
        for label, val in [("ملاحظات المقيم", notes), ("الاحتياجات التدريبية", training)]:
            if not val:
                continue
            ws.merge_cells(f"A{row}:B{row}")
            c = ws.cell(row, 1, label)
            c.font = font(bold=True, color=DARK, size=10)
            c.fill = fill(LBLUE)
            c.alignment = align(h="right")
            ws.merge_cells(f"C{row}:H{row}")
            c = ws.cell(row, 3, val)
            c.font = font(size=9)
            c.fill = fill(CREAM)
            c.alignment = align(h="right", wrap=True)
            for ci in range(1, 9):
                ws.cell(row, ci).border = inner_b
            ws.row_dimensions[row].height = 30
            row += 1

    # ══════════════════════════════════════════════════════════════
    # النتيجة النهائية
    # ══════════════════════════════════════════════════════════════
    row += 1
    ws.merge_cells(f"A{row}:D{row}")
    c = ws.cell(row, 1, "النتيجة النهائية السنوية")
    c.font = font(bold=True, color=WHITE, size=11)
    c.fill = fill(DARK)
    c.alignment = align()
    ws.merge_cells(f"E{row}:F{row}")
    c = ws.cell(row, 5, f"{avg_pct}%")
    fin_bg = GREEN_BG if avg_pct >= 70 else (YELLOW if avg_pct >= 50 else RED_BG)
    c.font = font(bold=True, color=DARK, size=13)
    c.fill = fill(fin_bg)
    c.alignment = align()
    ws.merge_cells(f"G{row}:H{row}")
    c = ws.cell(row, 7, avg_verbal)
    c.font = font(bold=True, color=DARK, size=11)
    c.fill = fill(fin_bg)
    c.alignment = align()
    for ci in range(1, 9):
        ws.cell(row, ci).border = outer_b
    ws.row_dimensions[row].height = 22

    return ws


def excel_to_pdf_bytes(xlsx_buf):
    try:
        from reportlab.lib.pagesizes import A4, landscape
        from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
        from reportlab.lib.styles import getSampleStyleSheet
        from reportlab.lib import colors
        from reportlab.lib.units import cm
        xlsx_buf.seek(0)
        wb = openpyxl.load_workbook(xlsx_buf, data_only=True)
        pdf_buf = io.BytesIO()
        doc = SimpleDocTemplate(pdf_buf, pagesize=landscape(A4),
                                 rightMargin=1*cm, leftMargin=1*cm,
                                 topMargin=1*cm, bottomMargin=1*cm)
        story = []
        styles = getSampleStyleSheet()
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            data = []
            for row in ws.iter_rows(values_only=True):
                if any(v is not None for v in row):
                    data.append([str(v) if v is not None else "" for v in row])
            if not data:
                continue
            n_cols = max(len(r) for r in data)
            avail  = 27 * cm
            col_widths = [avail / n_cols] * n_cols
            t = Table(data, colWidths=col_widths, repeatRows=1)
            t.setStyle(TableStyle([
                ("BACKGROUND",  (0,0), (-1,0), colors.HexColor("#1F3864")),
                ("TEXTCOLOR",   (0,0), (-1,0), colors.white),
                ("FONTSIZE",    (0,0), (-1,-1), 7),
                ("ALIGN",       (0,0), (-1,-1), "CENTER"),
                ("GRID",        (0,0), (-1,-1), 0.4, colors.HexColor("#AAAAAA")),
                ("BOX",         (0,0), (-1,-1), 1.2, colors.HexColor("#1F3864")),
                ("ROWBACKGROUNDS", (0,1), (-1,-1),
                 [colors.HexColor("#F2F2F2"), colors.white]),
                ("TOPPADDING",    (0,0), (-1,-1), 2),
                ("BOTTOMPADDING", (0,0), (-1,-1), 2),
            ]))
            story.append(Paragraph(f"<b>{sheet_name}</b>", styles["Heading2"]))
            story.append(Spacer(1, 0.3*cm))
            story.append(t)
            story.append(Spacer(1, 0.5*cm))
        doc.build(story)
        pdf_buf.seek(0)
        return pdf_buf
    except Exception:
        return None


def _rgb_to_hex(color_obj):
    try:
        if color_obj and color_obj.type == "rgb":
            rgb = color_obj.rgb
            return "#" + rgb[2:]
    except Exception:
        pass
    return ""


def print_preview_html(xlsx_buf, title="\u062a\u0642\u0631\u064a\u0631", chart_b64=""):
    """
    HTML \u0644\u0644\u0637\u0628\u0627\u0639\u0629 \u2014 \u0635\u0641\u062d\u0629 \u0648\u0627\u062d\u062f\u0629 A4 Landscape \u0645\u0636\u0645\u0648\u0646\u0629
    \u0643\u0648\u062f CSS \u062b\u0627\u0628\u062a 100% \u2014 \u0628\u062f\u0648\u0646 JavaScript
    \u0627\u0644\u062d\u0644: transform:scale(0.62) \u062b\u0627\u0628\u062a \u0645\u0639 transform-origin:top right
    \u0644\u0636\u0645\u0627\u0646 \u0639\u062f\u0645 \u0627\u0644\u0642\u0637\u0639 \u0639\u0644\u0649 \u0635\u0641\u062d\u0629 \u062b\u0627\u0646\u064a\u0629
    """
    xlsx_buf.seek(0)
    wb = openpyxl.load_workbook(xlsx_buf, data_only=True)

    import base64 as _b64, os as _os3
    _logo_b64 = ""
    for _lp in [LOGO_PATH, "logo.png"]:
        if _os3.path.exists(_lp):
            with open(_lp, "rb") as _lf:
                _logo_b64 = _b64.b64encode(_lf.read()).decode()
            break

    # ══════════════════════════════════════════════════════════
    # CSS ثابت مضمون — لا يتغير بأي ظرف
    # المبدأ:
    #   - في المعاينة: الصفحة عرض 277mm كاملة
    #   - عند الطباعة: scale(0.62) يضغط التقرير ليتسع في A4
    #   - transform-origin: top right → يبدأ من الزاوية الصحيحة (RTL)
    #   - margin-bottom سالب يلغي المساحة الفارغة بعد الضغط
    #   - @page بدون margin لإعطاء أقصى مساحة للمحتوى
    # ══════════════════════════════════════════════════════════
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

/* ── ورقة المعاينة ── */
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

/* ════════════════════════════════════════════════
   طباعة — صفحة واحدة A4 Landscape مضمونة 100%

   الحساب:
   A4 Landscape = 297mm × 210mm
   هامش 6mm من كل جهة → منطقة طباعة = 285mm × 198mm
   التقرير عرضه = 277mm
   scale = 285 / 277 = 1.03 (يتسع!)
   لكن الارتفاع قد يزيد → نضغط بـ 0.62 ثابتة آمنة
   ════════════════════════════════════════════════ */
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
        /* الضغط الثابت — يضمن صفحة واحدة */
        transform: scale(0.62) !important;
        transform-origin: top right !important;

        /* إلغاء المساحة الفارغة الناتجة عن الضغط */
        /* 277mm × (1-0.62) = 277 × 0.38 = 105mm */
        margin-top: 0 !important;
        margin-bottom: -105mm !important;
        margin-right: 0 !important;
        margin-left: auto !important;

        padding: 4mm 5mm !important;
        box-shadow: none !important;
        width: 277mm !important;
        overflow: visible !important;

        /* منع أي قطع */
        page-break-after: avoid !important;
        break-after: avoid !important;
        page-break-inside: avoid !important;
        break-inside: avoid !important;
    }

    table {
        width: 100% !important;
        table-layout: fixed !important;
        font-size: 8pt !important;
    }
    td {
        font-size: 8pt !important;
        padding: 1px 3px !important;
        line-height: 1.22 !important;
    }
    tr {
        page-break-inside: avoid !important;
        break-inside: avoid !important;
    }
    img { max-height: 44px !important; object-fit: contain; }

    * {
        -webkit-print-color-adjust: exact !important;
        print-color-adjust: exact !important;
        color-adjust: exact !important;
    }

    @page {
        size: A4 landscape;
        margin: 6mm;
    }
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
                "نموذج تقييم الأداء السنوي — مجموعة شركات فنون"
                "</span>"
                f'<img src="data:image/png;base64,{_logo_b64}"'
                ' style="height:40px;width:40px;object-fit:contain;" />'
                "</div>"
            )

        # merged cells
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

        # column widths
        col_widths = {}
        for col_letter, cd in ws.column_dimensions.items():
            from openpyxl.utils import column_index_from_string
            idx = column_index_from_string(col_letter)
            col_widths[idx] = 0 if cd.hidden else max(int((cd.width or 8) * 6.5), 0)

        _hidden = {10}
        colgroup = "<colgroup>"
        for c in range(1, ws.max_column + 1):
            if c in _hidden:
                colgroup += '<col style="width:0;display:none;">'
            else:
                colgroup += f'<col style="width:{col_widths.get(c, 50)}px;">'
        colgroup += "</colgroup>"

        parts.append(f'<div class="page">{logo_tag}<table>{colgroup}')

        for r in range(1, ws.max_row + 1):
            rh = ws.row_dimensions[r].height if r in ws.row_dimensions else 13
            rh = 11 if rh is None else max(int(rh * 0.95), 11)
            parts.append(f'<tr style="height:{rh}px;">')

            for c in range(1, ws.max_column + 1):
                if col_widths.get(c, 1) == 0: continue
                if c in _hidden: continue
                if (r, c) in skip: continue

                cell = ws.cell(r, c)
                val  = cell.value
                text = "" if val is None else str(val).replace("\n", "<br>")
                style = "overflow:hidden;word-break:break-word;"

                f_obj = cell.font
                if f_obj:
                    sz = min(f_obj.size or 9, 10)
                    style += f"font-size:{sz}pt;"
                    if f_obj.bold: style += "font-weight:bold;"
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
                    rs, cs2 = merged[(r, c)]
                    if rs > 1: span += f' rowspan="{rs}"'
                    if cs2 > 1: span += f' colspan="{cs2}"'

                parts.append(f'<td style="{style}"{span}>{text}</td>')

            parts.append("</tr>")

        # الرسم البياني أسفل يسار
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

