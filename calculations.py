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
    الدرجة الفعلية = (النسبة المدخلة / 100) * الوزن
    مثال: نسبة 85% ووزن 10 → درجة فعلية 8.5
    """
    v = max(0.0, min(float(weight), (float(pct_value) / 100.0) * float(weight)))
    return round(v, 2)

def kpi_score_to_pct(kpi_score: float, weight: float) -> float:
    """تحويل الدرجة الفعلية إلى نسبة مئوية 0-100"""
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
    if s.empty:
        return 0.0
    # مجموع الدرجات الفعلية مقسوم على 100 ليعطي نسبة 0-1
    total_score = s["KPI_%"].sum()
    return total_score / 100.0

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
      - Weight    : الوزن الأصلي (من 0 إلى 80 للوظيفي أو 0-20 للشخصي)
      - avg_score : متوسط الدرجة كنسبة مئوية 0-100
    """
    result = []
    
    if df_data is None or df_data.empty:
        return result
    if df_kpi is None or df_kpi.empty:
        return result
    if not emp or not job:
        return result
    
    job_kpis = df_kpi[df_kpi["JobTitle"].astype(str).str.strip() == str(job).strip()]
    
    if job_kpis.empty:
        return result
    
    for _, row in job_kpis.iterrows():
        kname = row.get("KPI_Name")
        weight = row.get("Weight", 0)
        
        if not kname or str(kname).strip() in ("", "nan", "None"):
            continue
        
        kname_str = str(kname).strip()
        weight = float(weight) if weight else 0.0
        
        sub = df_data[
            (df_data["EmployeeName"].astype(str).str.strip() == str(emp).strip()) &
            (df_data["KPI_Name"].astype(str).str.strip() == kname_str)
        ]
        
        if months_filter and len(months_filter) > 0:
            sub = sub[sub["Month"].isin(months_filter)]
        
        if year:
            sub = sub[sub["Year"] == int(year)]
        
        if not sub.empty:
            # KPI_% هي الدرجة الفعلية (مثال: إذا كان الوزن 10 والدرجة 8.5)
            avg_kpi_score = sub["KPI_%"].mean()
            if weight > 0:
                # تحويل إلى نسبة مئوية 0-100
                avg_percentage = round((avg_kpi_score / weight) * 100, 1)
            else:
                avg_percentage = 0.0
        else:
            avg_percentage = 0.0
        
        result.append({
            "KPI_Name":  kname_str,
            "Weight":    weight,
            "avg_score": avg_percentage,  # نسبة من 0-100
        })
    
    return result
