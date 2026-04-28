from constants import MONTHS_EN, PERSONAL_KPIS

# ── سلم التقييم المئوي ─────────────────────────────────────────
RATING_SCALE = [
    {"label": "ضعيف",     "min": 0,  "max": 59,  "range": "0 – 59",   "value": 50},
    {"label": "متوسط",    "min": 60, "max": 69,  "range": "60 – 69",  "value": 65},
    {"label": "جيد",      "min": 70, "max": 79,  "range": "70 – 79",  "value": 75},
    {"label": "جيد جداً", "min": 80, "max": 89,  "range": "80 – 89",  "value": 85},
    {"label": "ممتاز",    "min": 90, "max": 100, "range": "90 – 100", "value": 95},
]

def rating_label(value: float) -> str:
    v = float(value)
    if v >= 90: return "ممتاز"
    if v >= 80: return "جيد جداً"
    if v >= 70: return "جيد"
    if v >= 60: return "متوسط"
    return "ضعيف"

def rating_label_color(label: str) -> str:
    return {
        "ممتاز":    "#15803d",
        "جيد جداً": "#1d4ed8",
        "جيد":      "#92400e",
        "متوسط":    "#b45309",
        "ضعيف":     "#b91c1c",
    }.get(label, "#374151")

def calc_kpi_score(pct_value: float, weight: float) -> float:
    """
    الدرجة الفعلية = القيمة المدخلة مباشرة (من 0 إلى الوزن)
    مثال: وزن 9.6 → المدخل من 0 إلى 9.6 → الدرجة = المدخل مباشرة
    """
    v = max(0.0, min(float(weight), float(pct_value)))
    return round(v, 2)

def kpi_score_to_pct(kpi_score: float, weight: float) -> float:
    """تحويل الدرجة المحفوظة إلى نسبة مئوية 0-100"""
    if weight == 0: return 0.0
    return round((float(kpi_score) / float(weight)) * 100.0, 1)

def kpi_score_to_label(kpi_score: float, weight: float) -> str:
    return rating_label(kpi_score_to_pct(kpi_score, weight))

def verbal_grade(pct):
    if pct >= 90: return "ممتاز"
    if pct >= 80: return "جيد جداً"
    if pct >= 70: return "جيد"
    if pct >= 60: return "متوسط"
    return "ضعيف"

def grade_color_hex(pct):
    if pct >= 90: return "#15803d"
    if pct >= 80: return "#1d4ed8"
    if pct >= 70: return "#92400e"
    if pct >= 60: return "#b45309"
    return "#b91c1c"

def calc_monthly(df_data, emp, month_en, year=None):
    mask = (df_data["EmployeeName"] == emp) & (df_data["Month"] == month_en)
    if year: mask = mask & (df_data["Year"] == int(year))
    s = df_data[mask]
    return s["KPI_%"].sum() / 100.0 if not s.empty else 0.0

def calc_monthly_personal(df_data, emp, month_en, year=None):
    mask = (df_data["EmployeeName"] == emp) & (df_data["Month"] == month_en)
    if year: mask = mask & (df_data["Year"] == int(year))
    s = df_data[mask]
    s = s[s["KPI_Name"].isin(PERSONAL_KPIS)]
    return s["KPI_%"].sum() if not s.empty else 0.0

def calc_yearly(df_data, emp, year=None):
    scores = [calc_monthly(df_data, emp, m, year) for m in MONTHS_EN]
    active = [s for s in scores if s > 0]
    return sum(active)/len(active) if active else 0.0

def calc_yearly_personal(df_data, emp, year=None):
    scores = [calc_monthly_personal(df_data, emp, m, year) for m in MONTHS_EN]
    active = [s for s in scores if s > 0]
    return sum(active)/len(active) if active else 0.0

def get_kpi_avgs(df_data, df_kpi, emp, job, months_filter=None, year=None):
    """
    ترجع قائمة من dicts تحتوي على:
      - KPI_Name  : اسم المؤشر
      - Weight    : الوزن النسبي
      - avg_score : متوسط الدرجة للأشهر المحددة
    """
    result = []
    job_kpis = df_kpi[df_kpi["JobTitle"] == job]

    for _, row in job_kpis.iterrows():
        kname  = row.get("KPI_Name")
        weight = row.get("Weight", 0)

        # تجاهل أي صف ناقص
        if not kname or str(kname).strip() in ("", "nan", "None"):
            continue

        weight = float(weight)

        # تصفية بيانات الموظف لهذا المؤشر
        sub = df_data[
            (df_data["EmployeeName"] == emp) &
            (df_data["KPI_Name"] == kname)
        ]
        if months_filter:
            sub = sub[sub["Month"].isin(months_filter)]
        if year:
            sub = sub[sub["Year"] == int(year)]

        # avg_score = متوسط KPI_% (نسبة مئوية 0-100 تمثل الدرجة من الوزن)
        avg_score = round(sub["KPI_%"].mean(), 1) if not sub.empty else 0.0

        result.append({
            "KPI_Name":  kname,
            "Weight":    weight,
            "avg_score": avg_score,
        })

    return result
