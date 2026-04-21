"""
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
PATCH لملف entry.py — رسالة تحذير عربية عند تجاوز الوزن
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

التعليمات:
  في ملف entry.py، ابحث عن المكان الذي يتم فيه عرض number_input
  لكل KPI (درجة التقييم)، وأضف الكود التالي مباشرة بعد كل number_input.

المثال: إذا كان الكود الأصلي هكذا:
─────────────────────────────────────────
    score = st.number_input("الدرجة", min_value=0.0, max_value=float(weight),
                             value=..., key=...)
─────────────────────────────────────────

استبدله بـ:
─────────────────────────────────────────
    score = st.number_input("الدرجة", min_value=0.0, value=..., key=...)
    _show_kpi_warning(score, weight)
─────────────────────────────────────────

ثم أضف الدالة التالية في أعلى entry.py (بعد الـ imports):
"""

import streamlit as st

def _show_kpi_warning(score, weight):
    """
    يعرض تحذيراً جميلاً إذا تجاوزت درجة التقييم الوزن المحدد.
    استدعِ هذه الدالة بعد كل number_input لدرجة KPI.
    """
    try:
        if float(score) > float(weight):
            st.markdown(
                f"""
                <div style="
                    display: flex;
                    align-items: center;
                    gap: 10px;
                    background: #FFF3CD;
                    border: 2px solid #FF4B4B;
                    border-radius: 10px;
                    padding: 8px 14px;
                    margin: 4px 0 8px 0;
                    direction: rtl;
                ">
                    <div style="font-size: 2rem; line-height: 1;">🚫</div>
                    <div>
                        <div style="
                            font-size: 13px;
                            font-weight: bold;
                            color: #CC0000;
                            font-family: Arial;
                        ">القيمة المدخلة أكثر من القيمة الصحيحة</div>
                        <div style="
                            font-size: 11px;
                            color: #7F4F00;
                            font-family: Arial;
                            margin-top: 2px;
                        ">الحد الأقصى المسموح: <b>{int(round(float(weight)))}</b> — 
                           القيمة المدخلة: <b>{int(round(float(score)))}</b></div>
                    </div>
                </div>
                """,
                unsafe_allow_html=True,
            )
    except (TypeError, ValueError):
        pass
