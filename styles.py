import streamlit as st

def apply_global_styles():
    st.markdown("""
  <style>
  /* ── Global font size ── */
  html, body, [class*="css"] { font-size: 15px !important; }

  /* ── Sidebar styling ── */
  section[data-testid="stSidebar"] {
      background: #F8FAFF;
      border-left: 3px solid #E2E8F0;
      min-width: 220px !important;
  }
  section[data-testid="stSidebar"] button {
      border-radius: 8px !important;
      font-size: 14px !important;
      font-weight: 500 !important;
      margin: 2px 0 !important;
      padding: 8px 12px !important;
      text-align: right !important;
      border: 1px solid #E2E8F0 !important;
      transition: all 0.2s !important;
  }
  section[data-testid="stSidebar"] button:hover {
      background: #EFF6FF !important;
      border-color: #1E3A8A !important;
      color: #1E3A8A !important;
  }

  /* ── Metric / info cards ── */
  div[data-testid="metric-container"] {
      background: white;
      border: 1px solid #E2E8F0;
      border-radius: 12px;
      padding: 16px;
      box-shadow: 0 1px 4px rgba(0,0,0,0.06);
  }

  /* ── DataFrames ── */
  div[data-testid="stDataFrame"] { border-radius: 10px; overflow: hidden; }

  /* ── Inputs & selects ── */
  div[data-baseweb="select"] > div,
  div[data-baseweb="input"] > div {
      border-radius: 8px !important;
      font-size: 14px !important;
  }

  /* ── Form submit button ── */
  div[data-testid="stFormSubmitButton"] button {
      background: #1E3A8A !important;
      color: white !important;
      border-radius: 8px !important;
      font-size: 15px !important;
      font-weight: bold !important;
      padding: 10px !important;
  }

  /* ── Download button ── */
  div[data-testid="stDownloadButton"] button {
      border-radius: 8px !important;
      font-size: 14px !important;
  }

  /* ── Expander ── */
  div[data-testid="stExpander"] {
      border: 1px solid #E2E8F0 !important;
      border-radius: 10px !important;
      margin-top: 8px !important;
  }

  /* ── Subheaders ── */
  h2 { font-size: 1.3rem !important; }
  h3 { font-size: 1.1rem !important; }
  </style>
  """, unsafe_allow_html=True)

