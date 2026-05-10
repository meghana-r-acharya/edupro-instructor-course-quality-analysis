"""
EduPro Platform v4.0 — Single File
Roles : Admin · Instructor · Student
Run   : py -m streamlit run edupro_dashboard.py
Deps  : py -m pip install streamlit pandas numpy plotly openpyxl python-dotenv groq
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
from groq import Groq
import os
from dotenv import load_dotenv
import plotly.graph_objects as go
import hashlib, datetime, time, warnings
warnings.filterwarnings("ignore")
# ================= AI HELPER =================
def ask_groq(prompt):
    try:
        resp = groq_client.chat.completions.create(
            model="llama-3.1-8b-instant",   # ✅ working model
            messages=[
                {"role": "system", "content": "You are a helpful educational AI tutor."},
                {"role": "user", "content": prompt}
            ],
            temperature=0.7,
            max_tokens=1000,
        )
        return resp.choices[0].message.content.strip()
    except Exception as e:
        return f"Error: {e}"
# 🔴 FORCE REMOVE ANY DEBUG TEXT LIKE "Groq Key Loaded"
st.markdown("""
<style>
div[data-testid="stMarkdownContainer"] p:contains("Groq Key Loaded") {
    display: none !important;
}
</style>
""", unsafe_allow_html=True)
load_dotenv()

groq_client = Groq(api_key=os.getenv("GROQ_API_KEY"))
if "doubts" not in st.session_state:
    st.session_state.doubts = []

if "login_logs" not in st.session_state:
    st.session_state.login_logs = []

if "logout_logs" not in st.session_state:
    st.session_state.logout_logs = []

def check_class_readiness(yt, meet, zoom):
    if not yt and not meet and not zoom:
        return "warning", "No class scheduled"

    if yt:
        if "youtube.com" in yt or "youtu.be" in yt:
            return "ready", "YouTube demo ready"
        else:
            return "error", "Invalid YouTube link"

    if meet:
        if "meet.google.com" in meet:
            return "ready", "Google Meet ready"
        else:
            return "error", "Invalid Meet link"

    if zoom:
        if "zoom.us" in zoom:
            return "ready", "Zoom ready"
        else:
            return "error", "Invalid Zoom link"

    return "warning", "No valid class"

st.set_page_config(
    page_title="EduPro Learning Platform",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="collapsed",
)

EXCEL_PATH     = "EduPro_Online_Platform.xlsx"

ROLE_THEMES = {
    None:         ("135deg, #0a0a1a 0%, #111132 50%, #0a0a1a 100%", "#6C63FF"),
    "Student":    ("135deg, #060d1f 0%, #0a1628 40%, #071020 100%", "#38bdf8"),
    "Instructor": ("135deg, #071a10 0%, #0a2a18 40%, #061510 100%", "#34d399"),
    "Admin":      ("135deg, #150822 0%, #1e0a30 40%, #100620 100%", "#c084fc"),
}

def _hex_to_rgb(h):
    h = h.lstrip("#")
    return ",".join(str(int(h[i:i+2], 16)) for i in (0, 2, 4))

def apply_styles(role=None):
    grad, accent = ROLE_THEMES.get(role, ROLE_THEMES[None])
    rgb = _hex_to_rgb(accent)
    st.markdown(f"""
<link href="https://fonts.googleapis.com/css2?family=Outfit:wght@300;400;500;600;700;800&display=swap" rel="stylesheet">
<style>
html, body, [class*="css"] {{ font-family: 'Outfit', sans-serif !important; color: #e2e8f0 !important; }}
.stApp {{ background: linear-gradient({grad}) !important; min-height: 100vh; }}
.block-container {{ padding-top: 20px !important; max-width: 1200px !important; }}
.block-container, [data-testid="stVerticalBlock"], [data-testid="stVerticalBlockBorderWrapper"], [data-testid="column"], div.element-container {{ background: transparent !important; }}
section[data-testid="stSidebar"] {{ background: rgba(0,0,0,0.50) !important; backdrop-filter: blur(24px) !important; border-right: 1px solid rgba(255,255,255,0.08) !important; min-width: 240px !important; }}
section[data-testid="stSidebar"] * {{ color: #e2e8f0 !important; }}
section[data-testid="stSidebar"] .stButton > button {{ background: rgba(255,255,255,0.08) !important; color: #e2e8f0 !important; border: 1px solid rgba(255,255,255,0.15) !important; border-radius: 10px !important; font-weight: 600 !important; width: 100% !important; }}
section[data-testid="stSidebar"] .stButton > button:hover {{ background: rgba(255,255,255,0.16) !important; }}
.glass {{ background: rgba(255,255,255,0.06); backdrop-filter: blur(20px); -webkit-backdrop-filter: blur(20px); border: 1px solid rgba(255,255,255,0.12); border-radius: 18px; padding: 22px 28px; margin-bottom: 14px; box-shadow: 0 8px 32px rgba(0,0,0,0.25); transition: transform .2s, box-shadow .2s, border-color .2s; }}
.glass:hover {{ transform: translateY(-2px); box-shadow: 0 12px 40px rgba(0,0,0,0.35); border-color: rgba(255,255,255,0.22); }}
.glass-sm {{ background: rgba(255,255,255,0.05); border: 1px solid rgba(255,255,255,0.10); border-radius: 14px; padding: 14px 18px; margin-bottom: 10px; }}
.card-accent {{ background: linear-gradient(135deg, rgba({rgb},0.12), rgba(255,255,255,0.04)); border: 1px solid rgba({rgb},0.30); border-radius: 18px; padding: 22px 28px; margin-bottom: 14px; box-shadow: 0 0 30px rgba({rgb},0.12); }}
.stat-pill {{ background: rgba(255,255,255,0.07); border: 1px solid rgba(255,255,255,0.13); border-radius: 16px; padding: 18px 10px; text-align: center; }}
.sp-val {{ font-size: 1.9rem; font-weight: 800; color: #fff; line-height: 1; }}
.sp-lbl {{ font-size: 0.72rem; color: rgba(255,255,255,0.5); font-weight: 500; margin-top: 4px; text-transform: uppercase; letter-spacing: 0.06em; }}
.sec-head {{ color: #fff; font-size: 1.05rem; font-weight: 700; border-left: 3px solid {accent}; padding-left: 10px; margin: 20px 0 12px 0; letter-spacing: -0.01em; }}
.stButton > button {{ background: rgba(255,255,255,0.10) !important; color: #fff !important; border: 1px solid rgba(255,255,255,0.22) !important; border-radius: 10px !important; font-family: 'Outfit', sans-serif !important; font-weight: 600 !important; transition: all 0.2s !important; }}
.stButton > button:hover {{ background: rgba(255,255,255,0.20) !important; border-color: rgba(255,255,255,0.45) !important; transform: translateY(-1px) !important; }}
.stDownloadButton > button {{ background: rgba(56,189,248,0.18) !important; color: #7dd3fc !important; border: 1px solid rgba(56,189,248,0.4) !important; border-radius: 10px !important; }}
.stTextInput > div > div > input, .stTextArea > div > div > textarea {{ background: rgba(255,255,255,0.08) !important; color: #f1f5f9 !important; border: 1px solid rgba(255,255,255,0.18) !important; border-radius: 10px !important; font-family: 'Outfit', sans-serif !important; }}
.stTextInput > div > div > input:focus, .stTextArea > div > div > textarea:focus {{ border-color: {accent} !important; box-shadow: 0 0 0 3px rgba({rgb},0.20) !important; }}
label, .stSelectbox label, .stTextInput label, .stTextArea label, .stSlider label, .stRadio label {{ color: rgba(255,255,255,0.65) !important; font-weight: 500 !important; font-size: 0.82rem !important; }}
div[data-baseweb="select"] > div {{ background: rgba(255,255,255,0.08) !important; border-color: rgba(255,255,255,0.18) !important; border-radius: 10px !important; }}
div[data-baseweb="select"] span {{ color: #e2e8f0 !important; }}
div[data-baseweb="select"] svg {{ fill: #94a3b8 !important; }}
ul[data-baseweb="menu"] {{ background: #1a1f35 !important; border: 1px solid rgba(255,255,255,0.15) !important; border-radius: 10px !important; }}
ul[data-baseweb="menu"] li {{ color: #e2e8f0 !important; }}
ul[data-baseweb="menu"] li:hover {{ background: rgba(255,255,255,0.1) !important; }}
span[data-baseweb="tag"] {{ background: rgba({rgb},0.25) !important; color: #fff !important; border-radius: 20px !important; }}
[data-testid="stMetric"] {{ background: rgba(255,255,255,0.07) !important; border: 1px solid rgba(255,255,255,0.12) !important; border-radius: 12px !important; padding: 14px 18px !important; }}
[data-testid="stMetricValue"] {{ color: #fff !important; font-weight: 800 !important; }}
[data-testid="stMetricLabel"] {{ color: rgba(255,255,255,0.5) !important; }}
[data-testid="stDataFrame"] {{ border-radius: 12px !important; border: 1px solid rgba(255,255,255,0.1) !important; overflow: hidden !important; }}
.dvn-scroller {{ background: rgba(255,255,255,0.04) !important; }}
.stSuccess {{ background: rgba(52,211,153,0.15) !important; color: #6ee7b7 !important; border: 1px solid rgba(52,211,153,0.35) !important; border-radius: 10px !important; }}
.stError {{ background: rgba(248,113,113,0.15) !important; color: #fca5a5 !important; border: 1px solid rgba(248,113,113,0.35) !important; border-radius: 10px !important; }}
.stInfo {{ background: rgba(56,189,248,0.12) !important; color: #7dd3fc !important; border: 1px solid rgba(56,189,248,0.3) !important; border-radius: 10px !important; }}
.stWarning {{ background: rgba(251,191,36,0.12) !important; color: #fde68a !important; border: 1px solid rgba(251,191,36,0.3) !important; border-radius: 10px !important; }}
.streamlit-expanderHeader {{ background: rgba(255,255,255,0.06) !important; border: 1px solid rgba(255,255,255,0.12) !important; border-radius: 10px !important; color: #e2e8f0 !important; font-weight: 600 !important; }}
.streamlit-expanderContent {{ background: rgba(255,255,255,0.03) !important; border: 1px solid rgba(255,255,255,0.08) !important; border-radius: 0 0 10px 10px !important; }}
.stTabs [data-baseweb="tab-list"] {{ background: rgba(255,255,255,0.05) !important; border-radius: 12px !important; padding: 4px !important; border: 1px solid rgba(255,255,255,0.1) !important; }}
.stTabs [data-baseweb="tab"] {{ color: rgba(255,255,255,0.55) !important; border-radius: 8px !important; font-weight: 500 !important; }}
.stTabs [aria-selected="true"] {{ background: rgba(255,255,255,0.12) !important; color: #fff !important; font-weight: 700 !important; }}

.d-card {{ background: rgba(255,255,255,0.05); border: 1px solid rgba(255,255,255,0.10); border-left: 4px solid #38bdf8; border-radius: 14px; padding: 16px; margin-bottom: 10px; }}
.d-card.done {{ border-left-color: #34d399; }}
.d-q {{ color: #fff; font-weight: 600; font-size: .92rem; }}
.d-meta {{ color: rgba(255,255,255,0.42); font-size: .76rem; margin-top: 3px; }}
.d-ans {{ background: rgba(52,211,153,0.12); border: 1px solid rgba(52,211,153,0.25); border-radius: 10px; padding: 10px 14px; margin-top: 10px; color: #a7f3d0; font-size: .87rem; }}
.log-row {{ display: flex; justify-content: space-between; align-items: center; padding: 9px 14px; background: rgba(255,255,255,0.04); border: 1px solid rgba(255,255,255,0.07); border-radius: 9px; margin-bottom: 5px; font-size: .82rem; }}
.login-in {{ color: #34d399; font-weight: 700; font-size: .75rem; text-transform: uppercase; letter-spacing: .06em; }}
.login-out {{ color: #f87171; font-weight: 700; font-size: .75rem; text-transform: uppercase; letter-spacing: .06em; }}
.rb {{ display: inline-block; padding: 3px 12px; border-radius: 20px; font-size: .72rem; font-weight: 700; text-transform: uppercase; letter-spacing: .06em; }}
.rb-S {{ background: rgba(56,189,248,0.18); color: #38bdf8; border: 1px solid rgba(56,189,248,0.4); }}
.rb-I {{ background: rgba(52,211,153,0.18); color: #34d399; border: 1px solid rgba(52,211,153,0.4); }}
.rb-A {{ background: rgba(192,132,255,0.18); color: #c084fc; border: 1px solid rgba(192,132,255,0.4); }}
.domain-chip {{ background: rgba(255,255,255,0.07); border: 1px solid rgba(255,255,255,0.14); border-radius: 12px; padding: 16px 18px; margin-bottom: 10px; transition: all .2s; cursor: default; }}
.domain-chip:hover {{ background: rgba(255,255,255,0.12); border-color: {accent}; transform: scale(1.01); }}
.dc-title {{ color: #fff; font-weight: 700; font-size: .93rem; margin-bottom: 3px; }}
.dc-sub {{ color: rgba(255,255,255,0.5); font-size: .78rem; }}
.live-tag {{ background: rgba(239,68,68,0.25); color: #fca5a5; border: 1px solid rgba(239,68,68,0.5); border-radius: 20px; padding: 3px 12px; font-size: .72rem; font-weight: 700; text-transform: uppercase; letter-spacing: .06em; animation: livePulse 1.6s infinite; display: inline-block; }}
@keyframes livePulse {{ 0%, 100% {{ opacity: 1; }} 50% {{ opacity: 0.55; }} }}
.star-row {{ color: #fbbf24; font-size: 1.1rem; }}
hr {{ border-color: rgba(255,255,255,0.10) !important; }}
#MainMenu, footer {{ visibility: hidden; }}
header[data-testid="stHeader"] {{ background: transparent !important; height: 0 !important; }}
button[kind="header"] {{ display: none !important; }}
.stSpinner > div {{ border-top-color: {accent} !important; }}
::-webkit-scrollbar {{ width: 6px; height: 6px; }}
::-webkit-scrollbar-track {{ background: rgba(255,255,255,0.03); }}
::-webkit-scrollbar-thumb {{ background: rgba(255,255,255,0.18); border-radius: 4px; }}
</style>
""", unsafe_allow_html=True)


@st.cache_data(show_spinner="Loading EduPro data…")
def load_excel():
    try:
        xl       = pd.ExcelFile(EXCEL_PATH)
        users    = pd.read_excel(xl, "Users")
        teachers = pd.read_excel(xl, "Teachers")
        courses  = pd.read_excel(xl, "Courses")
        txns     = pd.read_excel(xl, "Transactions")
        txns["TransactionDate"] = pd.to_datetime(txns["TransactionDate"])
        txns["Month"]   = txns["TransactionDate"].dt.to_period("M").astype(str)
        txns["Year"]    = txns["TransactionDate"].dt.year
        txns["Quarter"] = txns["TransactionDate"].dt.to_period("Q").astype(str)
        master = (txns.merge(courses, on="CourseID", how="left")
                      .merge(teachers, on="TeacherID", how="left",
                             suffixes=("_course", "_teacher")))
        enroll   = txns.groupby("CourseID").size().reset_index(name="EnrollmentCount")
        rev      = txns.groupby("CourseID")["Amount"].sum().reset_index(name="TotalRevenue")
        courses  = courses.merge(enroll, on="CourseID", how="left").fillna({"EnrollmentCount": 0})
        courses  = courses.merge(rev,    on="CourseID", how="left").fillna({"TotalRevenue": 0})
        t_stats  = (master.groupby("TeacherID")
                    .agg(TotalEnrollments=("TransactionID", "count"),
                         TotalRevenue=("Amount", "sum"),
                         AvgCourseRating=("CourseRating", "mean"),
                         CoursesCount=("CourseID", "nunique"))
                    .reset_index())
        teachers = teachers.merge(t_stats, on="TeacherID", how="left").fillna(0)
        return users, teachers, courses, txns, master
    except FileNotFoundError:
        st.error(f"❌ Could not find **{EXCEL_PATH}**. Place it in the same folder as this script.")
        st.stop()

users_df, teachers_df, courses_df, txns_df, master_df = load_excel()

def analytics_dashboard(df):
    """Proper analytics dashboard with charts — no raw data dumps."""
    tabs = st.tabs(["📚 Courses", "👨‍🏫 Instructors", "💰 Revenue", "📅 Trends"])

    with tabs[0]:
        c1, c2 = st.columns(2)
        with c1:
            cat_r = courses_df.groupby("CourseCategory")["CourseRating"].mean().reset_index()
            fig = px.bar(cat_r.sort_values("CourseRating"),
                         x="CourseRating", y="CourseCategory", orientation="h",
                         color="CourseRating",
                         color_continuous_scale=[[0,"#f87171"],[1,"#34d399"]],
                         title="Avg Rating by Category")
            fig.update_layout(**PLOTLY_BASE, coloraxis_showscale=False)
            st.plotly_chart(fig, use_container_width=True)
        with c2:
            level_r = courses_df.groupby("CourseLevel")["CourseRating"].mean().reset_index()
            fig = px.bar(level_r, x="CourseLevel", y="CourseRating",
                         color="CourseLevel", color_discrete_sequence=COLORS,
                         title="Avg Rating by Level")
            fig.update_layout(**PLOTLY_BASE, showlegend=False)
            st.plotly_chart(fig, use_container_width=True)

    with tabs[1]:
        top_t = teachers_df.nlargest(10, "TeacherRating")[
            ["TeacherName","Expertise","TeacherRating","YearsOfExperience","TotalEnrollments","TotalRevenue"]
        ]
        st.dataframe(top_t, use_container_width=True)
        fig = px.scatter(teachers_df, x="YearsOfExperience", y="TeacherRating",
                         color="Expertise", size="TotalEnrollments",
                         hover_name="TeacherName", trendline="ols",
                         color_discrete_sequence=COLORS, title="Experience vs Rating")
        fig.update_layout(**PLOTLY_BASE)
        st.plotly_chart(fig, use_container_width=True)

    with tabs[2]:
        total_rev = txns_df["Amount"].sum()
        c1, c2, c3 = st.columns(3)
        with c1: st.metric("Total Revenue", f"${total_rev:,.0f}")
        with c2: st.metric("Paid Txns", f"{(txns_df['Amount']>0).sum():,}")
        with c3: st.metric("Free Txns", f"{(txns_df['Amount']==0).sum():,}")
        cat_rev = (df[df["Amount"]>0]
                   .groupby("CourseCategory")["Amount"]
                   .sum().reset_index()
                   .sort_values("Amount", ascending=True))
        if not cat_rev.empty:
            fig = px.bar(cat_rev, x="Amount", y="CourseCategory", orientation="h",
                         color="Amount",
                         color_continuous_scale=[[0,"#DBEAFE"],[1,"#6C63FF"]],
                         title="Revenue by Category")
            fig.update_layout(**PLOTLY_BASE, coloraxis_showscale=False)
            st.plotly_chart(fig, use_container_width=True)

    with tabs[3]:
        gran = st.radio("Granularity", ["Monthly","Quarterly","Yearly"], horizontal=True)
        gcol = {"Monthly":"Month","Quarterly":"Quarter","Yearly":"Year"}[gran]
        enroll_t = txns_df.groupby(gcol).size().reset_index(name="Enrollments").sort_values(gcol)
        fig = px.line(enroll_t, x=gcol, y="Enrollments", markers=True,
                      color_discrete_sequence=["#6C63FF"], title="Enrollment Trend")
        fig.update_traces(line_width=3)
        fig.update_layout(**PLOTLY_BASE, xaxis_tickangle=-40)
        st.plotly_chart(fig, use_container_width=True)

@st.cache_data(show_spinner=False)
def build_auth_db(_users_df, _teachers_df):
    db = {}
    db["admin@edupro.com"] = {
        "email": "admin@edupro.com", "password": "admin123",
        "role": "Admin", "name": "EduPro Admin",
    }
    for _, row in _teachers_df.iterrows():
        name  = str(row["TeacherName"]).strip()
        # password = fullname (no spaces, lowercase) + 123
        # e.g. "Jill Day" → jillday123, "David Williams" → davidwilliams123
        pwd   = name.replace(" ", "").lower() + "123"
        email = name.replace(" ", ".").lower() + "@edupro.com"
        db[email] = {
            "email": email, "password": pwd,
            "role": "Instructor", "name": name,
            "domain": str(row["Expertise"]),
            "teacher_id": str(row["TeacherID"]),
        }
    for _, row in _users_df.iterrows():
        email = str(row["Email"]).strip().lower()
        uname = str(row["UserName"]).strip()
        db[email] = {
            "email": email, "password": uname + "123",
            "role": "Student", "name": uname,
            "user_id": str(row["UserID"]),
            "gender": str(row["Gender"]), "age": int(row["Age"]),
        }
    return db

AUTH_DB = build_auth_db(users_df, teachers_df)

DOMAIN_MAP = {}
for _, _r in teachers_df.iterrows():
    _d = str(_r["Expertise"])
    if _d not in DOMAIN_MAP:
        DOMAIN_MAP[_d] = str(_r["TeacherName"])
ALL_DOMAINS = sorted(DOMAIN_MAP.keys())

DOMAIN_ICONS = {
    "Artificial Intelligence": "🤖", "Business": "💼",
    "Cybersecurity": "🔐", "Data Science": "📊",
    "Design": "🎨", "Digital Marketing": "📣",
    "Finance": "💰", "Machine Learning": "🧠",
    "Marketing": "📢", "Programming": "💻",
    "Project Management": "📋", "Web Development": "🌐",
}

PLOTLY_BASE = dict(
    paper_bgcolor="rgba(0,0,0,0)",
    plot_bgcolor="rgba(255,255,255,0.03)",
    font=dict(family="Outfit", color="#e2e8f0"),
    margin=dict(l=10, r=10, t=40, b=10),
    xaxis=dict(gridcolor="rgba(255,255,255,0.07)",
               linecolor="rgba(255,255,255,0.12)",
               tickfont=dict(color="#94a3b8")),
    yaxis=dict(gridcolor="rgba(255,255,255,0.07)",
               linecolor="rgba(255,255,255,0.12)",
               tickfont=dict(color="#94a3b8")),
    legend=dict(bgcolor="rgba(255,255,255,0.05)",
                bordercolor="rgba(255,255,255,0.1)",
                font=dict(color="#e2e8f0")),
)
COLORS = ["#6C63FF","#34d399","#f59e0b","#38bdf8","#f87171","#c084fc","#fb923c","#4ade80"]

_DEFAULTS = {
    "user": None, "login_logs": [], "doubts": [], "chat_logs": [],
    "live_links": {}, "notes": {}, "ratings": {}, "inst_ratings": {},
    "study_paths": {}, "quiz_bank": {}, "challenge_results": [],
}
for _k, _v in _DEFAULTS.items():
    if _k not in st.session_state:
        st.session_state[_k] = _v

def now_str():
    return datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

def log_action(user, action):
    st.session_state.login_logs.append({
        "name": user["name"], "email": user["email"],
        "role": user["role"], "action": action, "time": now_str(),
    })

def logout():
    log_action(st.session_state.user, "logout")
    st.session_state.user = None
    st.rerun()

def make_uid(s):
    return hashlib.md5((s + now_str()).encode()).hexdigest()[:8]

def stars(n):
    n = max(0, min(5, int(n)))
    return "⭐" * n + "☆" * (5 - n)

def sec(title, icon=""):
    st.markdown(f'<div class="sec-head">{icon} {title}</div>', unsafe_allow_html=True)

def pill(icon, val, lbl):
    st.markdown(f"""
<div class="stat-pill">
    <div style="font-size:1.4rem;margin-bottom:4px;">{icon}</div>
    <div class="sp-val">{val}</div>
    <div class="sp-lbl">{lbl}</div>
</div>""", unsafe_allow_html=True)

def role_badge_html(role):
    cls = {"Student": "rb-S", "Instructor": "rb-I", "Admin": "rb-A"}.get(role, "rb-S")
    return f'<span class="rb {cls}">{role}</span>'

def sidebar_nav(role_icon, user, options):
    with st.sidebar:
        st.markdown(f"## {role_icon} EduPro")
        st.markdown("---")
        st.markdown(f'<b style="color:#fff;font-size:.95rem;">{user["name"]}</b>', unsafe_allow_html=True)
        st.markdown(role_badge_html(user["role"]), unsafe_allow_html=True)
        if user["role"] == "Instructor":
            st.markdown(
                f'<div style="color:rgba(255,255,255,.45);font-size:.78rem;margin-top:4px;">'
                f'📌 {user.get("domain","")}</div>', unsafe_allow_html=True)
        st.markdown("---")
        nav = st.radio("", options, label_visibility="collapsed")
        st.markdown("---")
        st.caption(f"EduPro v4.0 · {datetime.datetime.now().year}")
        if st.button("🚪 Logout", use_container_width=True):
            logout()
    return nav
# ai
def get_ai_answer(question):
    try:
        completion =groq_client.chat.completions.create(
            model="llama-3.1-8b-instant",
            messages=[
                {"role": "system", "content": "You are a helpful AI tutor."},
                {"role": "user", "content": question}
            ],
            temperature=0.7,
            max_tokens=500
        )
        return completion.choices[0].message.content

    except Exception as e:
        return f"ERROR: {str(e)}"
    
def _quiz_bank(domain):
    if "quiz_bank" not in st.session_state:
        st.session_state.quiz_bank = {}

    if domain not in st.session_state.quiz_bank:
        st.session_state.quiz_bank[domain] = []

    return st.session_state.quiz_bank[domain]

# ══════════════════════════════════════════════════════════════════════════════
# ARENA CSS (used by Quiz Attender in student page and Quiz Builder)
# ══════════════════════════════════════════════════════════════════════════════
ARENA_CSS = """
<style>
.arena-card{background:rgba(255,255,255,0.06);border:1px solid rgba(255,255,255,0.14);
 border-radius:16px;padding:18px 22px;margin-bottom:12px;}
.arena-q{color:#fff;font-size:1rem;font-weight:700;margin-bottom:12px;line-height:1.45;}
.score-band{background:linear-gradient(135deg,rgba(108,99,255,0.25),rgba(52,211,153,0.15));
 border:1px solid rgba(108,99,255,0.35);border-radius:14px;padding:18px 24px;text-align:center;margin:12px 0;}
.lb-row{display:flex;justify-content:space-between;align-items:center;padding:9px 14px;
 background:rgba(255,255,255,0.04);border:1px solid rgba(255,255,255,0.08);
 border-radius:9px;margin-bottom:5px;font-size:.84rem;}
.lb-rank{font-size:1.1rem;font-weight:800;min-width:36px;color:#fbbf24;}
.lb-name{color:#fff;font-weight:600;flex:1;margin-left:8px;}
.lb-score{color:#34d399;font-weight:700;font-size:.9rem;}
.lb-badge{background:rgba(108,99,255,0.22);color:#a78bfa;border:1px solid rgba(108,99,255,0.35);
 border-radius:20px;padding:2px 10px;font-size:.72rem;font-weight:700;margin-left:8px;}
.timer-bar{height:6px;border-radius:4px;background:rgba(255,255,255,0.12);overflow:hidden;margin:8px 0;}
.timer-fill{height:100%;border-radius:4px;background:linear-gradient(90deg,#34d399,#38bdf8);
 transition:width .5s linear;}
</style>
"""

# ══════════════════════════════════════════════════════════════════════════════
# LOGIN
# ══════════════════════════════════════════════════════════════════════════════
def page_login():
    apply_styles(None)
    _, col, _ = st.columns([1, 1.3, 1])
    with col:
        st.markdown("""
<div style="text-align:center;padding:30px 0 10px;">
    <div style="font-size:3.5rem;margin-bottom:8px;">🎓</div>
    <h1 style="color:#fff;font-family:'Outfit',sans-serif;font-size:2.4rem;font-weight:800;margin:0 0 4px;letter-spacing:-0.03em;">EduPro Learning Platform</h1>
    <p style="color:rgba(255,255,255,.45);font-size:.95rem;margin:0 0 30px;">Learn · Teach · Manage</p>
</div>""", unsafe_allow_html=True)
        st.markdown("### 🔐 Sign In")
        email    = st.text_input("Email address", placeholder="you@example.com")
        password = st.text_input("Password", placeholder="••••••••", type="password")
        login_btn = st.button("Sign In →", use_container_width=True)
        if login_btn:
            u = AUTH_DB.get(email.strip().lower())
            if u and u["password"] == password.strip():
                st.session_state.user = u
                log_action(u, "login")
                st.success(f"Welcome, {u['name']}!")
                time.sleep(0.4)
                st.rerun()
            else:
                st.error("❌ Email or password incorrect.")
        

# ══════════════════════════════════════════════════════════════════════════════
# STUDENT PAGE
# ─────────────────────────────────────────────────────────────
def page_student():

    user = st.session_state.user
    if "enrolled" not in user:
        user["enrolled"] = []

    apply_styles("Student")

    nav = sidebar_nav("🎓", user, [
        "🏠 Dashboard",
        "❓ My Doubts",
        "📡 Live Class",
        "📄 Notes & Resources",
        "📝 Quiz Attender",
        "🤖 AI Study Bot"
    ])

    # ================= DASHBOARD =================
    if "Dashboard" in nav:
        st.markdown(
            f'<div class="card-accent">'
            f'<h2 style="color:#fff;margin:0 0 4px;font-size:1.7rem;">👋 Hey, {user["name"]}!</h2>'
            f'<p style="color:rgba(255,255,255,.55);margin:0;">Ready to learn something great today?</p>'
            f'</div>', unsafe_allow_html=True)

        open_doubts = [d for d in st.session_state.doubts
                       if d.get("student_email") == user["email"] and not d.get("answer")]
        c1, c2, c3 = st.columns(3)
        with c1: pill("📚", str(len(user["enrolled"])), "Enrolled")
        with c2: pill("🌐", str(len(ALL_DOMAINS)), "Domains")
        with c3: pill("❓", str(len(open_doubts)), "Open Doubts")
        st.markdown("<br>", unsafe_allow_html=True)

        # ── Enroll section ──
        sec("Enroll in a Domain", "🎓")
        available = [d for d in ALL_DOMAINS if d not in user["enrolled"]]
        if available:
            with st.expander("➕ Enroll in a new domain"):
                pick = st.selectbox("Choose domain", available, key="enroll_pick")
                if st.button("Enroll Now", use_container_width=True, key="enroll_btn"):
                    user["enrolled"].append(pick)
                    st.success(f"Enrolled in **{pick}** 🎉")
                    st.rerun()
        else:
            st.info("You are enrolled in all available domains.")

        # ── Enrolled courses ──
        sec("My Enrolled Courses", "📚")
        if not user["enrolled"]:
            st.info("You haven't enrolled in any courses yet.")
        else:
            for dm in user["enrolled"]:
                instr    = DOMAIN_MAP.get(dm, "—")
                icon     = DOMAIN_ICONS.get(dm, "📘")
                notes_n  = len(st.session_state.notes.get(dm, []))
                live_set = "✅" if st.session_state.live_links.get(dm) else "❌"
                st.markdown(
                    f'<div class="glass">'
                    f'<div style="display:flex;justify-content:space-between;align-items:center;">'
                    f'<div>'
                    f'<div style="color:#fff;font-weight:700;font-size:1rem;">{icon} {dm}</div>'
                    f'<div style="color:rgba(255,255,255,.45);font-size:.8rem;margin-top:3px;">'
                    f'👨‍🏫 {instr} &nbsp;·&nbsp; 📄 {notes_n} notes &nbsp;·&nbsp; 📡 Live: {live_set}'
                    f'</div></div>'
                    f'<span class="rb rb-S">Enrolled ✓</span>'
                    f'</div></div>', unsafe_allow_html=True)

        # ── All domains grid ──
        st.markdown("<br>", unsafe_allow_html=True)
        sec("All Available Domains", "🌐")
        cols = st.columns(3)
        for i, (dm, instr) in enumerate(DOMAIN_MAP.items()):
            with cols[i % 3]:
                icon = DOMAIN_ICONS.get(dm, "📘")
                etag = ('<span style="background:rgba(52,211,153,0.2);color:#34d399;border-radius:10px;'
                        'padding:2px 8px;font-size:.7rem;">✓ Enrolled</span>'
                        if dm in user["enrolled"] else "")
                st.markdown(
                    f'<div class="domain-chip">'
                    f'<div class="dc-title">{icon} {dm} {etag}</div>'
                    f'<div class="dc-sub">👨‍🏫 {instr}</div>'
                    f'</div>', unsafe_allow_html=True)

    # ================= DOUBTS =================
    elif "Doubt" in nav:
        sec("My Doubts", "❓")
        my_doubts = [d for d in st.session_state.doubts
                     if d.get("student_email") == user["email"]]

        with st.form("doubt_form", clear_on_submit=True):
            d_dom = st.selectbox("Domain",
                                 user["enrolled"] if user["enrolled"] else ALL_DOMAINS,
                                 key="doubt_domain")
            d_q   = st.text_area("Your question", height=80,
                                  placeholder="Type your doubt here…")
            sub   = st.form_submit_button("Submit Doubt", use_container_width=True)

        if sub and d_q.strip():
            st.session_state.doubts.append({
                "id": make_uid(d_q),
                "student_email": user["email"],
                "student_name":  user["name"],
                "domain":  d_dom,
                "question": d_q,
                "answer":  None,
                "ai":      False,
                "time":    now_str()
            })
            st.success("Doubt submitted! Your instructor will reply soon.")
            st.rerun()

        if not my_doubts:
            st.info("No doubts submitted yet.")
        else:
            for d in reversed(my_doubts):
                cls   = "done" if d.get("answer") else ""
                ans_h = (f'<div class="d-ans">✅ <b>Reply:</b> {d["answer"]}</div>'
                         if d.get("answer") else "")
                st.markdown(
                    f'<div class="d-card {cls}">'
                    f'<div class="d-q">{d["question"]}</div>'
                    f'<div class="d-meta">{d["domain"]} · {d["time"]}</div>'
                    f'{ans_h}</div>', unsafe_allow_html=True)

    # ================= LIVE CLASS =================
    elif "Live Class" in nav:
        sec("Live Classes", "📡")
        import re as _re_live

        if not user["enrolled"]:
            st.info("Enroll in a course first to access live classes.")
        else:
            for _domain in user["enrolled"]:
                _instructor = DOMAIN_MAP.get(_domain, "—")
                _raw = st.session_state.live_links.get(_domain, "")
                if isinstance(_raw, str):
                    _lk_yt = _raw; _lk_meet = ""; _lk_zoom = ""
                else:
                    _lk_yt   = _raw.get("youtube", "")
                    _lk_meet = _raw.get("meet",    "")
                    _lk_zoom = _raw.get("zoom",    "")
                _any = _lk_yt or _lk_meet or _lk_zoom

                if _any:
                    _zm_id = ""; _mm_code = ""
                    if _lk_zoom:
                        _zm = _re_live.search(r"(?:/j/|/wc/join/)(\d+)", _lk_zoom)
                        if _zm: _zm_id = _zm.group(1)
                    if _lk_meet:
                        _mm = _re_live.search(
                            r"meet\.google\.com/([a-z]{3}-[a-z]{4}-[a-z]{3})", _lk_meet)
                        if _mm: _mm_code = _mm.group(1)

                    st.markdown(
                        f'<div class="glass">'
                        f'<div style="display:flex;align-items:center;gap:10px;margin-bottom:10px;">'
                        f'<span class="live-tag">● LIVE</span>'
                        f'<span style="color:#fff;font-weight:700;font-size:1rem;">{_domain}</span>'
                        f'</div>'
                        f'<div style="color:rgba(255,255,255,.45);font-size:.8rem;margin-bottom:14px;">'
                        f'👨‍🏫 {_instructor}</div>',
                        unsafe_allow_html=True)

                    if _lk_zoom and _zm_id:
                        st.markdown(
                            f'<div style="background:rgba(124,58,237,.12);border:1px solid rgba(167,139,250,.35);'
                            f'border-radius:12px;padding:14px 18px;margin-bottom:10px;">'
                            f'<div style="color:#a78bfa;font-weight:700;font-size:.88rem;margin-bottom:6px;">🎥 Zoom Meeting</div>'
                            f'<div style="color:rgba(255,255,255,.6);font-size:.8rem;margin-bottom:8px;">'
                            f'Copy the Meeting ID below, then click <b style="color:#fff;">Join Zoom</b>.</div>'
                            f'<div style="display:flex;align-items:center;gap:10px;flex-wrap:wrap;">'
                            f'<div style="background:rgba(0,0,0,.3);border:1px solid rgba(167,139,250,.4);'
                            f'border-radius:8px;padding:8px 16px;font-family:monospace;font-size:1.1rem;'
                            f'color:#fff;letter-spacing:.12em;font-weight:700;">🔑 {_zm_id}</div>'
                            f'<a href="https://zoom.us/wc/join/{_zm_id}" target="_blank" '
                            f'style="display:inline-block;background:linear-gradient(135deg,#7c3aed,#a78bfa);'
                            f'color:#fff;text-decoration:none;padding:9px 22px;border-radius:10px;'
                            f'font-weight:700;font-size:.88rem;">🎥 Join Zoom</a>'
                            f'</div></div>', unsafe_allow_html=True)

                    if _lk_meet and _mm_code:
                        st.markdown(
                            f'<div style="background:rgba(26,115,232,.12);border:1px solid rgba(56,189,248,.35);'
                            f'border-radius:12px;padding:14px 18px;margin-bottom:10px;">'
                            f'<div style="color:#38bdf8;font-weight:700;font-size:.88rem;margin-bottom:6px;">📹 Google Meet</div>'
                            f'<div style="color:rgba(255,255,255,.6);font-size:.8rem;margin-bottom:8px;">'
                            f'Copy the Meet code, then click <b style="color:#fff;">Join Meet</b>.</div>'
                            f'<div style="display:flex;align-items:center;gap:10px;flex-wrap:wrap;">'
                            f'<div style="background:rgba(0,0,0,.3);border:1px solid rgba(56,189,248,.4);'
                            f'border-radius:8px;padding:8px 16px;font-family:monospace;font-size:1.1rem;'
                            f'color:#fff;letter-spacing:.12em;font-weight:700;">🔑 {_mm_code}</div>'
                            f'<a href="{_lk_meet}" target="_blank" '
                            f'style="display:inline-block;background:linear-gradient(135deg,#1a73e8,#0d47a1);'
                            f'color:#fff;text-decoration:none;padding:9px 22px;border-radius:10px;'
                            f'font-weight:700;font-size:.88rem;">📹 Join Meet</a>'
                            f'</div></div>', unsafe_allow_html=True)

                    if _lk_yt:
                        st.markdown(
                            f'<div style="background:rgba(220,38,38,.12);border:1px solid rgba(248,113,113,.35);'
                            f'border-radius:12px;padding:14px 18px;margin-bottom:10px;">'
                            f'<div style="color:#f87171;font-weight:700;font-size:.88rem;margin-bottom:8px;">▶ YouTube Live</div>'
                            f'<a href="{_lk_yt}" target="_blank" '
                            f'style="display:inline-block;background:linear-gradient(135deg,#dc2626,#ef4444);'
                            f'color:#fff;text-decoration:none;padding:9px 22px;border-radius:10px;'
                            f'font-weight:700;font-size:.88rem;">▶ Watch on YouTube</a>'
                            f'</div>', unsafe_allow_html=True)

                    # ── Instructor rating ──
                    _t_email = next(
                        (e for e, u2 in AUTH_DB.items()
                         if u2.get("role") == "Instructor" and u2.get("name") == _instructor),
                        None)
                    if _t_email:
                        _existing  = st.session_state.inst_ratings.get(_t_email, [])
                        _my_rating = next(
                            (r for r in _existing if r["by"] == user["email"]), None)
                        if _my_rating:
                            _ms = _my_rating["stars"]
                            st.markdown(
                                f'<div class="star-row" style="margin:10px 0 16px;">'
                                f'{stars(_ms)}&nbsp;'
                                f'<span style="color:rgba(255,255,255,.6);font-size:.85rem;">'
                                f'You rated <b style="color:#fff;">{_instructor}</b>: '
                                f'<b style="color:#fbbf24;">{_ms} star{"s" if _ms!=1 else ""}</b>'
                                f'</span></div>', unsafe_allow_html=True)
                        else:
                            _sel = st.select_slider(
                                f"⭐ Rate {_instructor}",
                                options=[1, 2, 3, 4, 5],
                                key=f"irate_{_t_email}")
                            if st.button(
                                    f"Submit Rating for {_instructor}",
                                    key=f"iratebtn_{_t_email}"):
                                if _t_email not in st.session_state.inst_ratings:
                                    st.session_state.inst_ratings[_t_email] = []
                                st.session_state.inst_ratings[_t_email].append(
                                    {"by": user["email"], "stars": _sel, "time": now_str()})
                                st.success("Rating submitted!")
                                st.rerun()

                    st.markdown('</div>', unsafe_allow_html=True)
                else:
                    st.markdown(
                        f'<div class="glass-sm">'
                        f'<span style="color:#fff;font-weight:600;">{_domain}</span>'
                        f'<span style="color:rgba(255,255,255,.35);font-size:.82rem;margin-left:10px;">'
                        f'⏳ Instructor hasn\'t set a live link yet.</span>'
                        f'</div>', unsafe_allow_html=True)

    # ================= NOTES & RESOURCES =================
    elif "Notes" in nav:
        sec("Notes & Resources", "📄")

        if not user["enrolled"]:
            st.info("Enroll in a course first to access notes.")
        else:
            sel_domain = st.selectbox(
                "Select your domain", user["enrolled"], key="notes_domain_sel")
            notes = st.session_state.notes.get(sel_domain, [])

            if not notes:
                st.info(f"No notes uploaded yet for **{sel_domain}**.")
            else:
                TYPE_COLORS = {
                    "PDF": "#f87171", "Word Doc": "#38bdf8", "PowerPoint": "#fb923c",
                    "Excel": "#34d399", "Text": "#94a3b8", "Image": "#c084fc",
                    "Video": "#f59e0b", "Audio": "#a78bfa", "Zip/Archive": "#6ee7b7",
                }
                for note in notes:
                    nid        = note.get("id", make_uid(note.get("title","note")))
                    my_ratings = st.session_state.ratings.get(nid, {})
                    already    = user["email"] in my_ratings
                    avg_r      = (sum(my_ratings.values()) / len(my_ratings)
                                  if my_ratings else None)
                    star_disp  = (f'&nbsp;·&nbsp;<span class="star-row">'
                                  f'{stars(round(avg_r))} ({len(my_ratings)} ratings)</span>'
                                  if avg_r else "")
                    uploaded_str = note.get("uploaded") or note.get("time") or "—"
                    type_label   = note.get("type_label", "File")
                    badge_color  = TYPE_COLORS.get(type_label, "#6C63FF")
                    desc_html    = (f'<div style="color:rgba(255,255,255,.45);font-size:.78rem;margin-top:3px;">'
                                    f'{note["description"]}</div>'
                                    if note.get("description") else "")

                    st.markdown(
                        f'<div class="glass">'
                        f'<div style="margin-bottom:4px;">'
                        f'<span style="background:rgba(255,255,255,0.07);border:1px solid rgba(255,255,255,0.14);'
                        f'border-radius:6px;padding:2px 8px;font-size:.7rem;color:{badge_color};'
                        f'font-weight:700;margin-right:8px;">{type_label}</span>'
                        f'<span style="color:#fff;font-weight:700;font-size:.95rem;">'
                        f'{note.get("title","Untitled")}</span></div>'
                        f'{desc_html}'
                        f'<div style="color:rgba(255,255,255,.4);font-size:.77rem;margin-top:4px;">'
                        f'{sel_domain} · 📎 {note.get("file_name","")} · Uploaded {uploaded_str}{star_disp}</div>'
                        f'</div>', unsafe_allow_html=True)

                    col_dl, col_rt = st.columns([3, 2])
                    with col_dl:
                        if note.get("file_data"):
                            st.download_button(
                                "⬇ Download",
                                data=note["file_data"],
                                file_name=note.get("file_name") or (note.get("title","note") + ".bin"),
                                mime=note.get("file_type", "application/octet-stream"),
                                key=f"dl_{nid}",
                                use_container_width=True)
                        elif note.get("content"):
                            st.download_button(
                                "⬇ Download",
                                data=note["content"].encode(),
                                file_name=note.get("title","note").replace(" ", "_") + ".txt",
                                mime="text/plain",
                                key=f"dl_{nid}",
                                use_container_width=True)
                        else:
                            st.markdown(
                                '<div style="color:rgba(255,255,255,.35);font-size:.8rem;padding:8px 0;">'
                                'No file attached</div>', unsafe_allow_html=True)
                    with col_rt:
                        if already:
                            st.markdown(
                                f'<div class="star-row" style="padding:8px 0;">'
                                f'{stars(my_ratings[user["email"]])} &nbsp; Rated</div>',
                                unsafe_allow_html=True)
                        else:
                            r_sel = st.select_slider(
                                "Rate", [1, 2, 3, 4, 5], key=f"rs_{nid}")
                            if st.button("⭐ Submit", key=f"rb_{nid}"):
                                if nid not in st.session_state.ratings:
                                    st.session_state.ratings[nid] = {}
                                st.session_state.ratings[nid][user["email"]] = r_sel
                                st.success("Rating saved!")
                                st.rerun()

    # ================= QUIZ ATTENDER =================
    elif "Quiz Attender" in nav:
        st.markdown(ARENA_CSS, unsafe_allow_html=True)
        sec("Quiz Attender", "📝")

        if not user["enrolled"]:
            st.info("Enroll in a course first to access quizzes.")
        else:
            sel_domain = st.selectbox(
                "Select Domain", user["enrolled"], key="qa_domain_sel")
            bank = _quiz_bank(sel_domain)

            if not bank:
                st.info(f"No questions available yet for **{sel_domain}**. "
                        f"Your instructor will add questions soon.")
            else:
                _qa_key   = f"qa_started_{sel_domain}"
                _ans_key  = f"qa_answers_{sel_domain}"
                _done_key = f"qa_done_{sel_domain}"

                for _k, _dv in [(_qa_key, False), (_ans_key, {}), (_done_key, False)]:
                    if _k not in st.session_state:
                        st.session_state[_k] = _dv

                # ── Not started ──
                if not st.session_state[_qa_key]:
                    st.markdown(f"""
<div class="arena-card">
    <div style="color:#fff;font-weight:700;font-size:1rem;">📋 {sel_domain} Quiz</div>
    <div style="color:rgba(255,255,255,.5);font-size:.85rem;margin-top:8px;line-height:1.6;">
        • {len(bank)} question(s) available<br>
        • Select one option per question<br>
        • Click <b>Submit Quiz</b> when done to see your score
    </div>
</div>""", unsafe_allow_html=True)
                    if st.button("▶ Start Quiz", type="primary",
                                 use_container_width=True, key="qa_start"):
                        st.session_state[_qa_key]  = True
                        st.session_state[_ans_key] = {}
                        st.session_state[_done_key]= False
                        st.rerun()

                # ── Quiz in progress ──
                elif not st.session_state[_done_key]:
                    st.markdown(
                        f'<div style="color:rgba(255,255,255,.45);font-size:.8rem;margin-bottom:16px;">'
                        f'Domain: <b style="color:#38bdf8">{sel_domain}</b>'
                        f'&nbsp;·&nbsp; {len(bank)} question(s)'
                        f'&nbsp;·&nbsp; Select one option per question then Submit</div>',
                        unsafe_allow_html=True)

                    answers = st.session_state[_ans_key]

                    for idx, q in enumerate(bank):
                        st.markdown(
                            f'<div class="arena-card">'
                            f'<div class="arena-q">Q{idx+1}. {q["question"]}</div>'
                            f'</div>', unsafe_allow_html=True)
                        opts = [f"{k}. {v}" for k, v in q["options"].items()]
                        cur  = answers.get(idx, None)
                        def_idx = 0
                        if cur:
                            for oi, op in enumerate(opts):
                                if op.startswith(cur + "."):
                                    def_idx = oi; break
                        choice = st.radio(
                            f"q{idx}", opts,
                            index=def_idx,
                            key=f"qa_radio_{sel_domain}_{idx}",
                            label_visibility="collapsed")
                        answers[idx] = choice[0]
                        st.session_state[_ans_key] = answers
                        st.markdown("<br>", unsafe_allow_html=True)

                    col1, col2 = st.columns(2)
                    with col1:
                        if st.button("✅ Submit Quiz", type="primary",
                                     use_container_width=True, key="qa_submit"):
                            if len(answers) < len(bank):
                                st.warning("Please answer all questions before submitting.")
                            else:
                                st.session_state[_done_key] = True
                                st.rerun()
                    with col2:
                        if st.button("🔄 Reset", use_container_width=True, key="qa_reset"):
                            st.session_state[_qa_key]  = False
                            st.session_state[_ans_key] = {}
                            st.session_state[_done_key]= False
                            st.rerun()

                # ── Results ──
                else:
                    answers = st.session_state[_ans_key]
                    score   = sum(1 for i, q in enumerate(bank)
                                  if answers.get(i) == q["answer"])
                    total   = len(bank)
                    pct     = int(score / total * 100) if total else 0
                    grade   = ("🏆 Excellent!" if pct >= 80
                               else "👍 Good!" if pct >= 60
                               else "📚 Keep Practising!")

                    st.markdown(f"""
<div class="score-band">
    <div style="font-size:2.5rem;margin-bottom:6px;">{grade.split()[0]}</div>
    <div style="color:#fff;font-size:1.6rem;font-weight:800;">{score} / {total}</div>
    <div style="color:rgba(255,255,255,.55);font-size:.9rem;margin-top:4px;">
        {grade[2:]} · {pct}% · {sel_domain}
    </div>
</div>""", unsafe_allow_html=True)

                    sec("Review Answers", "📋")
                    for idx, q in enumerate(bank):
                        given   = answers.get(idx, "—")
                        correct = q["answer"]
                        icon    = "✅" if given == correct else "❌"
                        g_color = "#6ee7b7" if given == correct else "#fca5a5"
                        st.markdown(
                            f'<div class="arena-card" style="margin-bottom:8px;">'
                            f'<div class="arena-q">{icon} Q{idx+1}. {q["question"]}</div>'
                            f'<div style="font-size:.84rem;color:rgba(255,255,255,.65);margin-top:6px;">'
                            f'Your answer: <b style="color:{g_color}">'
                            f'{given}. {q["options"].get(given, "—")}</b>'
                            f'&nbsp;&nbsp;|&nbsp;&nbsp;'
                            f'Correct: <b style="color:#34d399">'
                            f'{correct}. {q["options"][correct]}</b>'
                            f'</div></div>', unsafe_allow_html=True)

                    # Save attempt
                    st.session_state.setdefault("quiz_attempts", [])
                    _attempt_key = f"_qa_saved_{sel_domain}_{score}_{total}"
                    if _attempt_key not in st.session_state:
                        st.session_state[_attempt_key] = True
                        st.session_state.quiz_attempts.append({
                            "student": user["name"], "domain": sel_domain,
                            "score": score, "total": total, "time": now_str()
                        })

                    if st.button("🔄 Retake Quiz", use_container_width=True, key="qa_retake"):
                        st.session_state[_qa_key]  = False
                        st.session_state[_ans_key] = {}
                        st.session_state[_done_key]= False
                        st.rerun()

    # ================= AI STUDY BOT =================
    elif "AI Study Bot" in nav:
        sec("AI Study Bot", "🤖")

        if "chat_history" not in st.session_state:
            st.session_state.chat_history = []

        user_q = st.text_input("Ask your question")
        if st.button("Ask"):
            if user_q.strip():
                with st.spinner("Thinking..."):
                    ai_ans = ask_groq(user_q)
                st.session_state.chat_history.append(("You", user_q))
                st.session_state.chat_history.append(("AI", ai_ans))

        for role, msg in st.session_state.chat_history:
            if role == "You":
                st.markdown(f"**🧑 You:** {msg}")
            else:
                st.markdown(f"**🤖 AI:** {msg}")



def page_instructor():

    user   = st.session_state.user
    domain = user.get("domain", ALL_DOMAINS[0] if ALL_DOMAINS else "")

    apply_styles("Instructor")

    # ================= NAV =================
    open_d = [
        d for d in st.session_state.doubts
        if d["domain"] == domain and not d.get("answer") and not d.get("ai")
    ]

    badge_txt = f"❓ Doubt Inbox ({len(open_d)} open)" if open_d else "❓ Doubt Inbox"

    nav = sidebar_nav("👨‍🏫", user, [
        "🏠 Dashboard",
        badge_txt,
        "📡 Set Live Class",
        "📄 Manage Notes",
        "📊 My Analytics",
        "🧩 Quiz Builder"
    ])

    # ================= DASHBOARD =================
    if "Dashboard" in nav:

        st.markdown(
            f'<div class="card-accent">'
            f'<h2 style="color:#fff;">👨‍🏫 Welcome, {user["name"]}</h2>'
            f'<p>Domain: <b style="color:#34d399">{domain}</b></p>'
            f'</div>',
            unsafe_allow_html=True
        )

    # ================= DOUBTS =================
    elif "Doubt" in nav:

        sec("Doubt Inbox", "❓")

        doubts = [
            d for d in st.session_state.doubts
            if d["domain"] == domain and not d.get("ai")
        ]

        if not doubts:
            st.success("No doubts 🎉")
        else:
            for d in reversed(doubts):

                st.markdown(f"**{d['student_name']}**: {d['question']}")

                if not d.get("answer"):
                    reply = st.text_input("Reply", key=d["id"])

                    if st.button("Send", key=f"btn_{d['id']}"):
                        d["answer"] = reply
                        st.success("Replied")
                        st.rerun()
                else:
                    st.success(f"Reply: {d['answer']}")

    # ================= LIVE CLASS =================
    elif "Live Class" in nav:

        sec("Set Live Class", "📡")

        # Load existing
        links = st.session_state.live_links.get(domain, {})

        meet = links.get("meet", "")
        yt   = links.get("youtube", "")

        # ---------- GOOGLE MEET ----------
        st.markdown("### 📹 Google Meet")

        c1, c2 = st.columns([3, 1])

        with c1:
            meet_input = st.text_input(
                "Paste Meet Link",
                value=meet,
                placeholder="https://meet.google.com/abc-defg-hij"
            )

        with c2:
            st.link_button(
                "🔗 Open Meet",
                "https://meet.google.com/new",
                use_container_width=True
            )

        # Show code
        if meet_input:
            import re
            match = re.search(r"meet\.google\.com/([a-z\-]+)", meet_input)
            if match:
                st.success(f"Meeting Code: {match.group(1)}")

        # ---------- YOUTUBE ----------
        st.markdown("### ▶ YouTube Live")

        yt_input = st.text_input(
            "Paste YouTube Live Link",
            value=yt,
            placeholder="https://youtube.com/watch?v=..."
        )

        # ---------- SAVE ----------
        if st.button("💾 Save Links", use_container_width=True):

            st.session_state.live_links[domain] = {
                "meet":    meet_input.strip(),
                "zoom":    "",           # kept in dict so student side doesn't break
                "youtube": yt_input.strip()
            }

            st.success("Links saved ✅")
            st.rerun()

    # ================= NOTES =================
    elif "Notes" in nav:

        sec("Manage Notes", "📄")

        ALLOWED_TYPES = {
            "PDF":        ["pdf"],
            "Word Doc":   ["doc", "docx"],
            "PowerPoint": ["ppt", "pptx"],
            "Excel":      ["xls", "xlsx"],
            "Text":       ["txt"],
            "Image":      ["png", "jpg", "jpeg", "gif", "webp"],
            "Video":      ["mp4", "mov", "avi"],
            "Audio":      ["mp3", "wav"],
            "Zip/Archive":["zip", "rar", "7z"],
            "Any":        None,  # no restriction
        }

        with st.expander("➕ Upload New Note / Resource", expanded=True):
            title     = st.text_input("Title", placeholder="e.g. Week 3 – Introduction to Neural Nets")
            file_type = st.selectbox(
                "File Type",
                list(ALLOWED_TYPES.keys()),
                index=0,
                key="note_file_type_sel"
            )
            allowed_ext = ALLOWED_TYPES[file_type]
            file = st.file_uploader(
                "Select File",
                type=allowed_ext,   # None = all types
                key="note_file_uploader"
            )
            desc = st.text_input("Description (optional)",
                                 placeholder="Brief description of this resource")

            if st.button("📤 Upload Note", use_container_width=True, type="primary"):
                if not title.strip():
                    st.warning("Please enter a title.")
                elif file is None:
                    st.warning("Please select a file to upload.")
                else:
                    note = {
                        "id":          make_uid(title),
                        "title":       title.strip(),
                        "file_name":   file.name,
                        "file_data":   file.read(),
                        "file_type":   file.type or "application/octet-stream",
                        "uploaded":    now_str(),
                        "content":     None,
                        "description": desc.strip(),
                        "type_label":  file_type,
                    }
                    st.session_state.notes.setdefault(domain, []).append(note)
                    st.success(f"✅ '{title}' uploaded successfully!")
                    st.rerun()

        # ── Show existing notes ──
        existing_notes = st.session_state.notes.get(domain, [])
        st.markdown("<br>", unsafe_allow_html=True)
        sec(f"Uploaded Notes · {domain}", "📂")

        if not existing_notes:
            st.info("No notes uploaded yet for this domain.")
        else:
            for n in existing_notes:
                type_label = n.get("type_label", "File")
                TYPE_COLORS = {
                    "PDF": "#f87171", "Word Doc": "#38bdf8", "PowerPoint": "#fb923c",
                    "Excel": "#34d399", "Text": "#94a3b8", "Image": "#c084fc",
                    "Video": "#f59e0b", "Audio": "#a78bfa", "Zip/Archive": "#6ee7b7",
                }
                badge_color = TYPE_COLORS.get(type_label, "#6C63FF")

                col_info, col_dl, col_del = st.columns([5, 2, 1])
                with col_info:
                    st.markdown(
                        f'<div style="padding:10px 0;">'
                        f'<span style="background:rgba(255,255,255,0.08);border:1px solid rgba(255,255,255,0.15);'
                        f'border-radius:6px;padding:2px 8px;font-size:.7rem;color:{badge_color};'
                        f'font-weight:700;margin-right:8px;">{type_label}</span>'
                        f'<span style="color:#fff;font-weight:600;font-size:.93rem;">{n.get("title","Untitled")}</span>'
                        f'<div style="color:rgba(255,255,255,.38);font-size:.75rem;margin-top:3px;">'
                        f'📎 {n.get("file_name","")} &nbsp;·&nbsp; 🕐 {n.get("uploaded","—")}'
                        + (f' &nbsp;·&nbsp; {n["description"]}' if n.get("description") else "") +
                        f'</div></div>',
                        unsafe_allow_html=True)
                with col_dl:
                    if n.get("file_data"):
                        st.download_button(
                            "⬇ Download",
                            data=n["file_data"],
                            file_name=n.get("file_name", n.get("title","note")),
                            mime=n.get("file_type", "application/octet-stream"),
                            key=f"inst_dl_{n['id']}",
                            use_container_width=True,
                        )
                with col_del:
                    if st.button("🗑", key=f"inst_del_{n['id']}", help="Delete this note"):
                        st.session_state.notes[domain] = [
                            x for x in st.session_state.notes[domain]
                            if x["id"] != n["id"]
                        ]
                        st.success("Note deleted.")
                        st.rerun()

    # ================= ANALYTICS =================
    elif "Analytics" in nav:
        sec("My Analytics", "📊")

        # ── Fetch this instructor's row from teachers_df ──
        t_row = teachers_df[teachers_df["Expertise"] == domain]

        # ── KPI pills ──
        t_enroll = int(t_row["TotalEnrollments"].values[0])  if not t_row.empty else 0
        t_rev    = float(t_row["TotalRevenue"].values[0])    if not t_row.empty else 0.0
        t_rating = float(t_row["TeacherRating"].values[0])   if not t_row.empty else 0.0
        t_exp    = int(t_row["YearsOfExperience"].values[0]) if not t_row.empty else 0
        t_courses= int(t_row["CoursesCount"].values[0])      if not t_row.empty else 0

        c1, c2, c3, c4, c5 = st.columns(5)
        with c1: pill("🎓", str(t_enroll),          "Enrollments")
        with c2: pill("💰", f"${t_rev:,.0f}",       "Revenue")
        with c3: pill("⭐", f"{t_rating:.1f}",       "Rating")
        with c4: pill("📅", str(t_exp) + " yrs",    "Experience")
        with c5: pill("📚", str(t_courses),          "Courses")

        st.markdown("<br>", unsafe_allow_html=True)

        # ── Domain-level: courses in this domain ──
        domain_courses = courses_df[
            courses_df["CourseID"].isin(
                master_df[master_df["Expertise"] == domain]["CourseID"]
            )
        ] if "Expertise" in master_df.columns else courses_df.head(0)

        # fallback: filter master_df by TeacherID
        if domain_courses.empty and not t_row.empty:
            tid = str(t_row["TeacherID"].values[0])
            tid_courses = master_df[master_df["TeacherID"].astype(str) == tid]["CourseID"].unique()
            domain_courses = courses_df[courses_df["CourseID"].isin(tid_courses)]

        tabs = st.tabs(["📈 Experience & Rating", "📚 My Courses", "💰 Revenue Trend"])

        # ── Tab 1: Experience vs Rating across ALL instructors, highlight this one ──
        with tabs[0]:
            st.markdown('<div class="sec-head" style="font-size:.9rem;">📊 How you compare — Experience vs Rating (all instructors)</div>', unsafe_allow_html=True)

            fig_scatter = px.scatter(
                teachers_df,
                x="YearsOfExperience",
                y="TeacherRating",
                color="Expertise",
                size="TotalEnrollments",
                hover_name="TeacherName",
                color_discrete_sequence=COLORS,
                title=f"Experience vs Rating · Your domain: {domain}",
                labels={"YearsOfExperience": "Years of Experience",
                        "TeacherRating": "Teacher Rating"}
            )
            # Highlight current instructor
            if not t_row.empty:
                fig_scatter.add_scatter(
                    x=t_row["YearsOfExperience"],
                    y=t_row["TeacherRating"],
                    mode="markers",
                    marker=dict(size=22, color="#f59e0b",
                                line=dict(width=3, color="#fff"), symbol="star"),
                    name=f"You ({user['name']})",
                    hovertext=user["name"]
                )
            fig_scatter.update_layout(**PLOTLY_BASE)
            st.plotly_chart(fig_scatter, use_container_width=True)

            # Bar: Rating by domain (highlight yours)
            domain_avg = (teachers_df.groupby("Expertise")["TeacherRating"]
                          .mean().reset_index()
                          .sort_values("TeacherRating", ascending=True))
            domain_avg["highlight"] = domain_avg["Expertise"].apply(
                lambda x: "Your Domain" if x == domain else "Others")
            fig_bar = px.bar(
                domain_avg,
                x="TeacherRating", y="Expertise", orientation="h",
                color="highlight",
                color_discrete_map={"Your Domain": "#34d399", "Others": "#6C63FF"},
                title="Average Rating by Domain",
                labels={"TeacherRating": "Avg Rating", "Expertise": "Domain"}
            )
            fig_bar.update_layout(**PLOTLY_BASE)
            st.plotly_chart(fig_bar, use_container_width=True)

        # ── Tab 2: Courses in this domain ──
        with tabs[1]:
            if domain_courses.empty:
                st.info("No course data found for your domain.")
            else:
                show_cols = [c for c in
                    ["CourseName","CourseCategory","CourseLevel","CourseRating",
                     "EnrollmentCount","TotalRevenue"]
                    if c in domain_courses.columns]
                st.dataframe(domain_courses[show_cols], use_container_width=True)

                if "CourseRating" in domain_courses.columns and len(domain_courses) > 1:
                    fig_cr = px.bar(
                        domain_courses.sort_values("CourseRating", ascending=False),
                        x="CourseName" if "CourseName" in domain_courses.columns else domain_courses.index,
                        y="CourseRating",
                        color="CourseRating",
                        color_continuous_scale=[[0,"#f87171"],[1,"#34d399"]],
                        title="Course Ratings in Your Domain"
                    )
                    fig_cr.update_layout(**PLOTLY_BASE, coloraxis_showscale=False,
                                         xaxis_tickangle=-30)
                    st.plotly_chart(fig_cr, use_container_width=True)

        # ── Tab 3: Revenue trend for this instructor's domain ──
        with tabs[2]:
            if not t_row.empty:
                tid = str(t_row["TeacherID"].values[0])
                t_txns = master_df[master_df["TeacherID"].astype(str) == tid]
                if t_txns.empty:
                    st.info("No transaction data found for your domain.")
                else:
                    rev_monthly = (t_txns.groupby("Month")["Amount"]
                                   .sum().reset_index()
                                   .sort_values("Month"))
                    fig_rev = px.line(
                        rev_monthly, x="Month", y="Amount",
                        markers=True,
                        color_discrete_sequence=["#34d399"],
                        title=f"Monthly Revenue · {domain}"
                    )
                    fig_rev.update_traces(line_width=3, marker_size=8)
                    fig_rev.update_layout(**PLOTLY_BASE, xaxis_tickangle=-40)
                    st.plotly_chart(fig_rev, use_container_width=True)

                    enroll_monthly = (t_txns.groupby("Month")
                                      .size().reset_index(name="Enrollments")
                                      .sort_values("Month"))
                    fig_en = px.bar(
                        enroll_monthly, x="Month", y="Enrollments",
                        color_discrete_sequence=["#38bdf8"],
                        title=f"Monthly Enrollments · {domain}"
                    )
                    fig_en.update_layout(**PLOTLY_BASE, xaxis_tickangle=-40)
                    st.plotly_chart(fig_en, use_container_width=True)
            else:
                st.info("Instructor data not found in the database.")

    # ================= QUIZ BUILDER =================
    elif "Quiz Builder" in nav:
        quiz_builder_instructor(user, domain)
# ══════════════════════════════════════════════════════════════════════════════
# ADMIN
# ══════════════════════════════════════════════════════════════════════════════
def page_admin():
    user = st.session_state.user
    apply_styles("Admin")

    nav = sidebar_nav("🛡", user, [
        "🏠 Platform Overview",
        "👥 User Management",
        "📋 Login / Logout Logs",
        "📊 Analytics Dashboard",
        "📡 Live Links Monitor",
        "AI Chatbot Logs",
        "📄 Notes Monitor"
    ])

    # ================= COMMON DATA =================
    total_users = len(AUTH_DB)
    total_students = len([u for u in AUTH_DB.values() if u.get("role") == "Student"])
    total_instructors = len([u for u in AUTH_DB.values() if u.get("role") == "Instructor"])

    total_doubts = len(st.session_state.get("doubts", []))
    answered_d = len([d for d in st.session_state.get("doubts", []) if d.get("answer")])

    total_logins = sum(1 for l in st.session_state.get("login_logs", []) if l["action"] == "login")
    total_logouts = sum(1 for l in st.session_state.get("login_logs", []) if l["action"] == "logout")

    # ================= OVERVIEW =================
    if "Overview" in nav:
        st.markdown(
            '<div class="card-accent">'
            '<h2 style="color:#fff;">🛡 Admin Control Center</h2>'
            '<p style="color:rgba(255,255,255,.5);">Full platform visibility & management</p>'
            '</div>',
            unsafe_allow_html=True
        )

        def pill(icon, value, label):
            st.markdown(f"""
            <div style="padding:18px;border-radius:14px;text-align:center;
                        background:rgba(255,255,255,0.05);
                        border:1px solid rgba(255,255,255,0.08);">
                <div style="font-size:1.4rem;font-weight:700;color:#fff;">
                    {icon} {value}
                </div>
                <div style="font-size:.75rem;color:rgba(255,255,255,.5);">
                    {label}
                </div>
            </div>
            """, unsafe_allow_html=True)

        # Row 1
        c1, c2, c3 = st.columns(3)
        with c1: pill("👥", total_users, "Users")
        with c2: pill("🌐", len(ALL_DOMAINS), "Domains")
        with c3: pill("❓", total_doubts, "Doubts")

        # Row 2
        c4, c5, c6 = st.columns(3)
        with c4: pill("✅", answered_d, "Answered")
        with c5: pill("🔐", total_logins, "Logins")
        with c6: pill("🚪", total_logouts, "Logouts")

        st.markdown("<br>", unsafe_allow_html=True)

        # Charts
        c_l, c_r = st.columns(2)

        with c_l:
            sec("Users by Role", "👥")
            roles = {
                "Students": total_students,
                "Instructors": total_instructors,
                "Admins": total_users - total_students - total_instructors
            }
            fig = px.pie(
                values=list(roles.values()),
                names=list(roles.keys())
            )
            st.plotly_chart(fig, use_container_width=True)

        with c_r:
            sec("Doubts by Domain", "📊")
            if st.session_state.get("doubts"):
                df = pd.DataFrame(st.session_state.doubts)
                d_cnt = df.groupby("domain").size().reset_index(name="Count")
                fig = px.bar(d_cnt, x="Count", y="domain", orientation="h")
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("No doubts yet.")

    # ================= USER MANAGEMENT =================
    elif "User Management" in nav:
        sec("User Management", "👥")

        search = st.text_input("Search users")

        for email, u in AUTH_DB.items():
            if not search or search.lower() in u["name"].lower():
                st.markdown(f"""
                <div class="glass-sm">
                    <b>{u['name']}</b> ({u['role']})<br>
                    <span style="color:gray">{email}</span>
                </div>
                """, unsafe_allow_html=True)

    # ================= LOGIN LOGS =================
    elif "Login" in nav:
        sec("Login / Logout Logs", "📋")

        logs = st.session_state.get("login_logs", [])
        if not logs:
            st.info("No logs yet")
        else:
            df = pd.DataFrame(logs)
            st.dataframe(df, use_container_width=True)

    # ================= LIVE LINKS =================
    elif "Live Links" in nav:
        sec("Live Links Monitor", "📡")

        for domain in ALL_DOMAINS:
            link = st.session_state.live_links.get(domain, "")
            st.write(domain, "→", link if link else "Not set")

    # ================= AI LOGS =================
    elif "AI Chatbot Logs" in nav:
        sec("AI Logs", "🤖")

        logs = st.session_state.get("chat_logs", [])
        if not logs:
            st.info("No logs yet")
        else:
            st.dataframe(pd.DataFrame(logs), use_container_width=True)

    elif "Analytics" in nav:
        sec("Analytics Dashboard", "📊")
        analytics_dashboard(master_df)
    elif "Notes Monitor" in nav:
        sec("Notes", "📄")

        all_notes = []
        for dom, nlist in st.session_state.notes.items():
            for n in nlist:
                all_notes.append(n)

        if all_notes:
            st.dataframe(pd.DataFrame(all_notes))
        else:
            st.info("No notes")

    # ================= CHALLENGE =================
    elif "Challenge Monitor" in nav:
        challenge_monitor_admin()
# ══════════════════════════════════════════════════════════════════════════════
# SMART STUDY PATH — helper data & functions
# ══════════════════════════════════════════════════════════════════════════════

SP_LEVELS = ["Beginner", "Intermediate", "Advanced"]

SP_LEVEL_META = {
    "Beginner":     {"color": "#34d399", "border": "rgba(52,211,153,0.35)",  "bg": "rgba(52,211,153,0.08)",  "icon": "🟢"},
    "Intermediate": {"color": "#f59e0b", "border": "rgba(245,158,11,0.35)",  "bg": "rgba(245,158,11,0.08)",  "icon": "🟡"},
    "Advanced":     {"color": "#f87171", "border": "rgba(248,113,113,0.35)", "bg": "rgba(248,113,113,0.08)", "icon": "🔴"},
}

# Default demo videos per level (reuses the YouTube demo pattern already in the app)
SP_DOMAIN_VIDEOS = {
    "Artificial Intelligence": {
        "Beginner":     [{"title": "AI For Everyone – What is AI?",           "url": "https://www.youtube.com/watch?v=mJeNghZXtMo"},
                         {"title": "Intro to Artificial Intelligence",         "url": "https://www.youtube.com/watch?v=ad79nYk2keg"}],
        "Intermediate": [{"title": "AI Search Algorithms Explained",           "url": "https://www.youtube.com/watch?v=5A7mSiDhLEQ"},
                         {"title": "AI Planning & Problem Solving",            "url": "https://www.youtube.com/watch?v=HWjNvdkaxNQ"}],
        "Advanced":     [{"title": "Reinforcement Learning – Full Course",     "url": "https://www.youtube.com/watch?v=ELE2_Mftqoc"},
                         {"title": "Large Language Models Explained",          "url": "https://www.youtube.com/watch?v=zjkBMFhNj_g"}],
    },
    "Business": {
        "Beginner":     [{"title": "Business Basics – How Companies Work",     "url": "https://www.youtube.com/watch?v=ire0bNNBhSA"},
                         {"title": "Introduction to Business Strategy",        "url": "https://www.youtube.com/watch?v=mYF2_FBCvXw"}],
        "Intermediate": [{"title": "Business Model Canvas Explained",          "url": "https://www.youtube.com/watch?v=IP0cUBWTgpY"},
                         {"title": "Operations Management Overview",           "url": "https://www.youtube.com/watch?v=OV_1GOFR-9E"}],
        "Advanced":     [{"title": "Corporate Strategy & Competitive Advantage","url": "https://www.youtube.com/watch?v=o7Ik1OB4TaE"},
                         {"title": "Mergers & Acquisitions Explained",         "url": "https://www.youtube.com/watch?v=PmHGzNm7_Gg"}],
    },
    "Cybersecurity": {
        "Beginner":     [{"title": "Cybersecurity for Beginners – Full Course","url": "https://www.youtube.com/watch?v=U_P23SqJaDc"},
                         {"title": "How the Internet Works – Security Basics", "url": "https://www.youtube.com/watch?v=AYdF7b3nMto"}],
        "Intermediate": [{"title": "Ethical Hacking Full Course",              "url": "https://www.youtube.com/watch?v=fNzpcB7ODxQ"},
                         {"title": "Network Security & Firewalls",             "url": "https://www.youtube.com/watch?v=E03gh1huvW4"}],
        "Advanced":     [{"title": "Penetration Testing Advanced Techniques",  "url": "https://www.youtube.com/watch?v=3Kq1MIfTWCE"},
                         {"title": "Malware Analysis & Reverse Engineering",   "url": "https://www.youtube.com/watch?v=ivnNOj0wc5s"}],
    },
    "Data Science": {
        "Beginner":     [{"title": "Data Science Full Course for Beginners",   "url": "https://www.youtube.com/watch?v=ua-CiDNNj30"},
                         {"title": "Statistics for Data Science",              "url": "https://www.youtube.com/watch?v=xxpc-HPKN28"}],
        "Intermediate": [{"title": "Pandas & Data Wrangling – Full Tutorial",  "url": "https://www.youtube.com/watch?v=vmEHCJofslg"},
                         {"title": "Data Visualization with Python",           "url": "https://www.youtube.com/watch?v=a9UrKTVEeZA"}],
        "Advanced":     [{"title": "Feature Engineering for Machine Learning", "url": "https://www.youtube.com/watch?v=XX6zFgz8HaM"},
                         {"title": "Big Data & Apache Spark Full Course",      "url": "https://www.youtube.com/watch?v=F8pyaR4uQ2g"}],
    },
    "Design": {
        "Beginner":     [{"title": "Graphic Design for Beginners",             "url": "https://www.youtube.com/watch?v=WONZVnlam6U"},
                         {"title": "Design Principles – Color & Typography",   "url": "https://www.youtube.com/watch?v=yNDgFK2Jj1E"}],
        "Intermediate": [{"title": "UI/UX Design Full Course",                 "url": "https://www.youtube.com/watch?v=c9Wg6Cb_YlU"},
                         {"title": "Figma Tutorial – Complete Beginner Guide", "url": "https://www.youtube.com/watch?v=FTFaQWZBqQ8"}],
        "Advanced":     [{"title": "Advanced UX Research Methods",             "url": "https://www.youtube.com/watch?v=Ovj4hFxko7c"},
                         {"title": "Design Systems & Component Libraries",     "url": "https://www.youtube.com/watch?v=EK-pHkc5EL4"}],
    },
    "Digital Marketing": {
        "Beginner":     [{"title": "Digital Marketing Full Course for Beginners","url": "https://www.youtube.com/watch?v=nU7T_E86mBE"},
                         {"title": "SEO Tutorial for Beginners",               "url": "https://www.youtube.com/watch?v=DvwS7cV9GmQ"}],
        "Intermediate": [{"title": "Google Ads Full Tutorial",                 "url": "https://www.youtube.com/watch?v=NtXxFHHBg1Y"},
                         {"title": "Social Media Marketing Strategy",          "url": "https://www.youtube.com/watch?v=q7FllhQS7TU"}],
        "Advanced":     [{"title": "Advanced SEO Techniques",                  "url": "https://www.youtube.com/watch?v=A8EI6JaFbv4"},
                         {"title": "Marketing Analytics & Attribution",        "url": "https://www.youtube.com/watch?v=OFIoQFJMGsw"}],
    },
    "Finance": {
        "Beginner":     [{"title": "Personal Finance for Beginners",           "url": "https://www.youtube.com/watch?v=HQzoZfc3GwQ"},
                         {"title": "How the Stock Market Works",               "url": "https://www.youtube.com/watch?v=p7HKvqRI_Bo"}],
        "Intermediate": [{"title": "Financial Statements Explained",           "url": "https://www.youtube.com/watch?v=lhnmCNQKElU"},
                         {"title": "Investment & Portfolio Management",        "url": "https://www.youtube.com/watch?v=q_xoFPxO_YE"}],
        "Advanced":     [{"title": "Derivatives & Options Trading",            "url": "https://www.youtube.com/watch?v=7PM4rNDr4oI"},
                         {"title": "Financial Modeling & Valuation",           "url": "https://www.youtube.com/watch?v=tHYjp7SRTEM"}],
    },
    "Machine Learning": {
        "Beginner":     [{"title": "Machine Learning for Beginners – Full Course","url": "https://www.youtube.com/watch?v=NWONeJKn6kc"},
                         {"title": "How Neural Networks Work",                 "url": "https://www.youtube.com/watch?v=aircAruvnKk"}],
        "Intermediate": [{"title": "Scikit-Learn Full Tutorial",               "url": "https://www.youtube.com/watch?v=0B5eIE_1vpU"},
                         {"title": "Decision Trees & Random Forests",          "url": "https://www.youtube.com/watch?v=J4Wdy0Wc_xQ"}],
        "Advanced":     [{"title": "Deep Learning Specialization Overview",    "url": "https://www.youtube.com/watch?v=CS4cs9xVecg"},
                         {"title": "Transformers & Attention Mechanism",       "url": "https://www.youtube.com/watch?v=4Bdc55j80l8"}],
    },
    "Marketing": {
        "Beginner":     [{"title": "Marketing 101 – Introduction",             "url": "https://www.youtube.com/watch?v=wMZM2GQfniU"},
                         {"title": "Understanding Consumer Behavior",          "url": "https://www.youtube.com/watch?v=9WiW6A6e9sU"}],
        "Intermediate": [{"title": "Content Marketing Strategy",               "url": "https://www.youtube.com/watch?v=lZNaFvSaXRc"},
                         {"title": "Brand Building & Positioning",             "url": "https://www.youtube.com/watch?v=6taVs-75Xs8"}],
        "Advanced":     [{"title": "Growth Hacking Strategies",                "url": "https://www.youtube.com/watch?v=ajccEoAhfmc"},
                         {"title": "Marketing ROI & Data-Driven Decisions",    "url": "https://www.youtube.com/watch?v=OFIoQFJMGsw"}],
    },
    "Programming": {
        "Beginner":     [{"title": "Python for Beginners – Full Course",       "url": "https://www.youtube.com/watch?v=_uQrJ0TkZlc"},
                         {"title": "How to Think Like a Programmer",           "url": "https://www.youtube.com/watch?v=azcrPFhaY9k"}],
        "Intermediate": [{"title": "Data Structures & Algorithms in Python",   "url": "https://www.youtube.com/watch?v=pkYVOmU3MgA"},
                         {"title": "Object-Oriented Programming Explained",    "url": "https://www.youtube.com/watch?v=Ej_02ICOIgs"}],
        "Advanced":     [{"title": "System Design Full Course",                "url": "https://www.youtube.com/watch?v=FSywKX3qelE"},
                         {"title": "Design Patterns in Python",                "url": "https://www.youtube.com/watch?v=bsyjSW46TDg"}],
    },
    "Project Management": {
        "Beginner":     [{"title": "Project Management for Beginners",         "url": "https://www.youtube.com/watch?v=ZKOL-rZ79gs"},
                         {"title": "Agile & Scrum Introduction",               "url": "https://www.youtube.com/watch?v=GzzkpAOxHXs"}],
        "Intermediate": [{"title": "PMP Exam Prep – Key Concepts",             "url": "https://www.youtube.com/watch?v=uWPIsaYpY7U"},
                         {"title": "Risk Management in Projects",              "url": "https://www.youtube.com/watch?v=AKCXWDGCk4Y"}],
        "Advanced":     [{"title": "Advanced Scrum & SAFe Framework",          "url": "https://www.youtube.com/watch?v=9TycLR0TqFA"},
                         {"title": "Program & Portfolio Management",           "url": "https://www.youtube.com/watch?v=Phr8NhCDPSo"}],
    },
    "Web Development": {
        "Beginner":     [{"title": "HTML & CSS Full Course for Beginners",     "url": "https://www.youtube.com/watch?v=G3e-cpL7ofc"},
                         {"title": "JavaScript for Beginners",                 "url": "https://www.youtube.com/watch?v=W6NZfCO5SIk"}],
        "Intermediate": [{"title": "React JS Full Tutorial",                   "url": "https://www.youtube.com/watch?v=bMknfKXIFA8"},
                         {"title": "Node.js & Express – Backend Development",  "url": "https://www.youtube.com/watch?v=Oe421EPjeBE"}],
        "Advanced":     [{"title": "Full-Stack Web App – End to End Project",  "url": "https://www.youtube.com/watch?v=nu_pCVPKzTk"},
                         {"title": "Web Performance Optimization",             "url": "https://www.youtube.com/watch?v=AQqFZ5t8uNc"}],
    },
}

# Fallback for any domain not in the map above
SP_DEFAULT_VIDEOS = {
    "Beginner":     [{"title": "Introduction & Core Concepts",  "url": "https://www.youtube.com/watch?v=mJeNghZXtMo"},
                     {"title": "Getting Started – Hands-On",    "url": "https://www.youtube.com/watch?v=ire0bNNBhSA"}],
    "Intermediate": [{"title": "Deep Dive – Key Techniques",    "url": "https://www.youtube.com/watch?v=IP0cUBWTgpY"},
                     {"title": "Practical Project Walkthrough", "url": "https://www.youtube.com/watch?v=vmEHCJofslg"}],
    "Advanced":     [{"title": "Advanced Topics & Research",    "url": "https://www.youtube.com/watch?v=zjkBMFhNj_g"},
                     {"title": "Real-World Implementation",     "url": "https://www.youtube.com/watch?v=FSywKX3qelE"}],
}

SP_DEFAULT_DESC = {
    "Beginner":     "No prior knowledge needed. Build a strong foundation with core concepts and guided exercises.",
    "Intermediate": "You know the basics. Now go deeper — projects, algorithms, and real-world applications.",
    "Advanced":     "For experts ready to tackle cutting-edge research, optimisation, and production-grade systems.",
}

def sp_get(domain):
    """Return the study path for a domain, initialising defaults if not set."""
    paths = st.session_state.study_paths
    if domain not in paths:
        domain_vids = SP_DOMAIN_VIDEOS.get(domain, SP_DEFAULT_VIDEOS)
        paths[domain] = {
            lvl: {
                "description": SP_DEFAULT_DESC[lvl],
                "videos":      [v.copy() for v in domain_vids[lvl]],
                "note_ids":    [],
            }
            for lvl in SP_LEVELS
        }
    return paths[domain]


# ──────────────────────────────────────────────────────────────────────────────
# STUDENT: view study path for one enrolled domain
# ──────────────────────────────────────────────────────────────────────────────
def sp_student_view(user):
    sec("Smart Study Path", "🗺️")

    enrolled = user.get("enrolled", [])
    if not enrolled:
        st.info("Enroll in a course first to view its study path.")
        return

    domain = st.selectbox("Choose your domain", enrolled, key="sp_student_domain")
    path   = sp_get(domain)
    icon   = DOMAIN_ICONS.get(domain, "📘")
    instr  = DOMAIN_MAP.get(domain, "—")

    st.markdown(f"""
<div class="glass" style="margin-bottom:20px;">
    <div style="color:#fff;font-size:1.1rem;font-weight:700;">{icon} {domain} — Study Roadmap</div>
    <div style="color:rgba(255,255,255,.5);font-size:.82rem;margin-top:3px;">👨‍🏫 {instr}</div>
</div>""", unsafe_allow_html=True)

    # Extract live links — support both legacy plain string and new dict schema
    _raw_ll = st.session_state.live_links.get(domain, {})
    if isinstance(_raw_ll, str):
        _raw_ll = {"youtube": _raw_ll, "meet": "", "zoom": ""}
    _ll_yt   = _raw_ll.get("youtube", "").strip()
    _ll_meet = _raw_ll.get("meet",    "").strip()
    _ll_zoom = _raw_ll.get("zoom",    "").strip()
    live_link = _ll_yt or _ll_meet or _ll_zoom   # any link = "live class available"
    domain_notes = st.session_state.notes.get(domain, [])

    for lvl in SP_LEVELS:
        meta = SP_LEVEL_META[lvl]
        data = path[lvl]

        # Count pinned notes that actually exist
        pinned_notes = [n for n in domain_notes if n["id"] in data["note_ids"]]

        with st.expander(
            f"{meta['icon']} {lvl}  —  {len(pinned_notes)} notes · {len(data['videos'])} videos"
                + (" · 📡 Live class available" if live_link else ""),
            expanded=(lvl == "Beginner"),
        ):
            # Description
            st.markdown(f"""
<div style="background:{meta['bg']};border:1px solid {meta['border']};
            border-radius:10px;padding:12px 16px;margin-bottom:14px;">
    <span style="color:{meta['color']};font-weight:600;">{lvl}</span>
    <span style="color:rgba(255,255,255,.65);font-size:.88rem;margin-left:8px;">{data['description']}</span>
</div>""", unsafe_allow_html=True)

            col_n, col_v = st.columns(2)

            # ── Notes column ──
            with col_n:
                st.markdown(f'<div style="color:{meta["color"]};font-size:.78rem;font-weight:700;'
                            f'text-transform:uppercase;letter-spacing:.06em;margin-bottom:8px;">📄 Notes</div>',
                            unsafe_allow_html=True)
                if pinned_notes:
                    for note in pinned_notes:
                        nid = note["id"]
                        my_r = st.session_state.ratings.get(nid, {})
                        avg  = (sum(my_r.values())/len(my_r)) if my_r else None
                        star_txt = f"{'⭐'*round(avg)} ({len(my_r)})" if avg else ""
                        st.markdown(f"""
<div class="glass-sm" style="margin-bottom:6px;">
    <div style="color:#fff;font-size:.88rem;font-weight:600;">📄 {note['title']}</div>
    <div style="color:rgba(255,255,255,.4);font-size:.74rem;">{star_txt}</div>
</div>""", unsafe_allow_html=True)
                        dl_col, rate_col = st.columns([3, 2])
                        with dl_col:
                            st.download_button(
                                "⬇ Download",
                                data=note["content"].encode(),
                                file_name=note["title"].replace(" ", "_") + ".txt",
                                mime="text/plain",
                                key=f"sp_dl_{nid}_{lvl}",
                                use_container_width=True,
                            )
                        with rate_col:
                            if user["email"] not in my_r:
                                r = st.select_slider("★", [1,2,3,4,5], key=f"sp_rs_{nid}_{lvl}")
                                if st.button("Rate", key=f"sp_rb_{nid}_{lvl}"):
                                    if nid not in st.session_state.ratings:
                                        st.session_state.ratings[nid] = {}
                                    st.session_state.ratings[nid][user["email"]] = r
                                    st.success("Rated!")
                                    st.rerun()
                            else:
                                st.markdown(
                                    f'<div style="color:#fbbf24;padding:6px 0;font-size:.82rem;">'
                                    f'{"⭐"*my_r[user["email"]]} rated</div>',
                                    unsafe_allow_html=True)
                else:
                    st.markdown('<div class="glass-sm" style="color:rgba(255,255,255,.4);'
                                'font-size:.82rem;">No notes assigned yet.</div>',
                                unsafe_allow_html=True)

            # ── Videos column ──
            with col_v:
                st.markdown(f'<div style="color:{meta["color"]};font-size:.78rem;font-weight:700;'
                            f'text-transform:uppercase;letter-spacing:.06em;margin-bottom:8px;">▶ Demo Videos</div>',
                            unsafe_allow_html=True)
                for vid in data["videos"]:
                    vid_url = vid["url"]
                    # Convert watch URL to embed URL
                    if "watch?v=" in vid_url:
                        vid_id  = vid_url.split("watch?v=")[-1].split("&")[0]
                        embed   = f"https://www.youtube.com/embed/{vid_id}"
                    elif "youtu.be/" in vid_url:
                        vid_id  = vid_url.split("youtu.be/")[-1].split("?")[0]
                        embed   = f"https://www.youtube.com/embed/{vid_id}"
                    else:
                        embed = vid_url
                    st.markdown(f"""
<div class="glass-sm" style="margin-bottom:8px;">
    <div style="color:#fff;font-size:.85rem;font-weight:600;margin-bottom:6px;">▶ {vid['title']}</div>
    <iframe width="100%" height="140" src="{embed}"
            frameborder="0" allow="accelerometer; autoplay; clipboard-write;
            encrypted-media; gyroscope; picture-in-picture" allowfullscreen
            style="border-radius:8px;"></iframe>
</div>""", unsafe_allow_html=True)

            # ── Live class — separate buttons for each active platform ──
            if live_link:
                import re as _re_sp
                _sp_btns = ""
                if _ll_zoom:
                    _zm = _re_sp.search(r"/j/(\d+)", _ll_zoom)
                    _zm_id = f" <span style='font-size:.74rem;opacity:.8;'>(ID: {_zm.group(1)})</span>" if _zm else ""
                    _sp_btns += (
                        f'<a href="{_ll_zoom}" target="_blank" '
                        f'style="display:inline-flex;align-items:center;gap:5px;'
                        f'background:linear-gradient(135deg,#7c3aed,#a78bfa);'
                        f'color:#fff;text-decoration:none;padding:9px 18px;border-radius:10px;'
                        f'font-weight:700;font-size:.84rem;margin-right:8px;margin-bottom:6px;">'
                        f'🎥 Zoom{_zm_id}</a>'
                    )
                if _ll_meet:
                    _mm = _re_sp.search(r"meet\.google\.com/([a-z]{3}-[a-z]{4}-[a-z]{3})", _ll_meet)
                    _mm_code = f" <span style='font-size:.74rem;opacity:.8;'>(Code: {_mm.group(1)})</span>" if _mm else ""
                    _sp_btns += (
                        f'<a href="{_ll_meet}" target="_blank" '
                        f'style="display:inline-flex;align-items:center;gap:5px;'
                        f'background:linear-gradient(135deg,#1a73e8,#0d47a1);'
                        f'color:#fff;text-decoration:none;padding:9px 18px;border-radius:10px;'
                        f'font-weight:700;font-size:.84rem;margin-right:8px;margin-bottom:6px;">'
                        f'📹 Google Meet{_mm_code}</a>'
                    )
                if _ll_yt:
                    _sp_btns += (
                        f'<a href="{_ll_yt}" target="_blank" '
                        f'style="display:inline-flex;align-items:center;gap:5px;'
                        f'background:linear-gradient(135deg,#dc2626,#ef4444);'
                        f'color:#fff;text-decoration:none;padding:9px 18px;border-radius:10px;'
                        f'font-weight:700;font-size:.84rem;margin-right:8px;margin-bottom:6px;">'
                        f'▶ YouTube Live</a>'
                    )
                st.markdown(f"""
<div style="margin-top:12px;">
    <div style="color:rgba(255,255,255,.5);font-size:.76rem;margin-bottom:6px;font-weight:600;">
        📡 Join Live Class for {lvl}:
    </div>
    <div style="display:flex;flex-wrap:wrap;gap:4px;">
        {_sp_btns}
    </div>
</div>""", unsafe_allow_html=True)


# ──────────────────────────────────────────────────────────────────────────────
# INSTRUCTOR: edit roadmap for their domain
# ──────────────────────────────────────────────────────────────────────────────
def sp_instructor_edit(domain):
    sec("Smart Study Path Editor", "🗺️")
    path = sp_get(domain)
    icon = DOMAIN_ICONS.get(domain, "📘")

    st.markdown(f"""
<div class="glass" style="margin-bottom:18px;">
    <div style="color:#fff;font-size:1rem;font-weight:700;">{icon} {domain} — Edit Roadmap</div>
    <div style="color:rgba(255,255,255,.45);font-size:.8rem;margin-top:3px;">
        Changes are visible to students immediately.
    </div>
</div>""", unsafe_allow_html=True)

    lvl = st.radio("Select level to edit", SP_LEVELS, horizontal=True, key="sp_inst_lvl")
    data = path[lvl]
    meta = SP_LEVEL_META[lvl]

    # ── Description ──
    st.markdown(f'<div style="color:{meta["color"]};font-weight:600;margin:10px 0 4px;">'
                f'{meta["icon"]} {lvl} — Description</div>', unsafe_allow_html=True)
    with st.form(f"sp_desc_form_{lvl}"):
        new_desc = st.text_area("Level description", value=data["description"], height=80)
        if st.form_submit_button("💾 Save Description", use_container_width=True):
            data["description"] = new_desc.strip()
            st.success("Description updated!")
            st.rerun()

    st.markdown("---")

    col_n, col_v = st.columns(2)

    # ── Pin / unpin notes ──
    with col_n:
        st.markdown(f'<div style="color:{meta["color"]};font-size:.82rem;font-weight:700;'
                    f'text-transform:uppercase;letter-spacing:.06em;margin-bottom:8px;">📄 Assign Notes</div>',
                    unsafe_allow_html=True)
        domain_notes = st.session_state.notes.get(domain, [])
        if not domain_notes:
            st.markdown('<div class="glass-sm" style="color:rgba(255,255,255,.4);font-size:.82rem;">'
                        'Upload notes first from Manage Notes.</div>', unsafe_allow_html=True)
        else:
            for note in domain_notes:
                nid     = note["id"]
                pinned  = nid in data["note_ids"]
                col_chk, col_lbl = st.columns([1, 5])
                with col_chk:
                    checked = st.checkbox("", value=pinned, key=f"sp_pin_{lvl}_{nid}")
                with col_lbl:
                    st.markdown(f'<div style="color:#fff;font-size:.86rem;padding-top:6px;">'
                                f'📄 {note["title"]}</div>', unsafe_allow_html=True)
                if checked and nid not in data["note_ids"]:
                    data["note_ids"].append(nid)
                    st.rerun()
                elif not checked and nid in data["note_ids"]:
                    data["note_ids"].remove(nid)
                    st.rerun()

    # ── Add / remove videos ──
    with col_v:
        st.markdown(f'<div style="color:{meta["color"]};font-size:.82rem;font-weight:700;'
                    f'text-transform:uppercase;letter-spacing:.06em;margin-bottom:8px;">▶ Videos</div>',
                    unsafe_allow_html=True)

        # Existing videos
        for i, vid in enumerate(data["videos"]):
            c_t, c_del = st.columns([5, 1])
            with c_t:
                st.markdown(f'<div class="glass-sm" style="margin-bottom:4px;">'
                            f'<span style="color:#fff;font-size:.84rem;">▶ {vid["title"]}</span><br>'
                            f'<span style="color:rgba(255,255,255,.4);font-size:.74rem;">{vid["url"][:50]}…</span>'
                            f'</div>', unsafe_allow_html=True)
            with c_del:
                if st.button("🗑", key=f"sp_vdel_{lvl}_{i}"):
                    data["videos"].pop(i)
                    st.rerun()

        # Add video form removed — videos are set per domain automatically


# ──────────────────────────────────────────────────────────────────────────────
# ADMIN: view + edit all domains
# ──────────────────────────────────────────────────────────────────────────────
def sp_admin_view():
    sec("Smart Study Path Manager", "🗺️")

    st.markdown("""
<div class="glass" style="margin-bottom:16px;">
    <div style="color:#fff;font-weight:700;">All Domains — Roadmap Overview</div>
    <div style="color:rgba(255,255,255,.45);font-size:.8rem;margin-top:3px;">
        Select a domain to edit its roadmap. Changes apply instantly for all students.
    </div>
</div>""", unsafe_allow_html=True)

    # Quick stats table
    rows = []
    for d in ALL_DOMAINS:
        p = sp_get(d)
        rows.append({
            "Domain":      d,
            "Instructor":  DOMAIN_MAP.get(d, "—"),
            "Beginner":    f"{len(p['Beginner']['note_ids'])}n · {len(p['Beginner']['videos'])}v",
            "Intermediate":f"{len(p['Intermediate']['note_ids'])}n · {len(p['Intermediate']['videos'])}v",
            "Advanced":    f"{len(p['Advanced']['note_ids'])}n · {len(p['Advanced']['videos'])}v",
        })
    st.dataframe(
        __import__("pandas").DataFrame(rows).set_index("Domain"),
        use_container_width=True,
    )

    st.markdown("---")
    edit_domain = st.selectbox("Edit roadmap for domain", ALL_DOMAINS, key="sp_admin_domain")
    sp_instructor_edit(edit_domain)


# ══════════════════════════════════════════════════════════════════════════════
# 🏆 PEER CHALLENGE ARENA  — NEW FEATURE
# Not covered by any of the 15 referenced papers.
# Implements active competitive learning:
#   • Instructor builds MCQ question banks per domain (Quiz Builder)
#   • Students challenge peers to timed 5-question quizzes (Challenge Arena)
#   • Both sides answer the same questions; scores logged and compared
#   • Global leaderboard tracks wins/score across all domains
#   • Admin sees full match history and platform-wide quiz stats
# ══════════════════════════════════════════════════════════════════════════════

# ── CSS additions for the arena ───────────────────────────────────────────────
ARENA_CSS = """
<style>
.arena-card{background:rgba(255,255,255,0.06);border:1px solid rgba(255,255,255,0.14);
 border-radius:16px;padding:18px 22px;margin-bottom:12px;}
.arena-q{color:#fff;font-size:1rem;font-weight:700;margin-bottom:12px;line-height:1.45;}
.arena-opt{background:rgba(255,255,255,0.08);border:1px solid rgba(255,255,255,0.15);
 border-radius:10px;padding:10px 16px;margin-bottom:6px;cursor:pointer;color:#e2e8f0;font-size:.9rem;}
.arena-opt:hover{background:rgba(255,255,255,0.16);}
.arena-opt.correct{background:rgba(52,211,153,0.20);border-color:rgba(52,211,153,0.5);color:#6ee7b7;}
.arena-opt.wrong{background:rgba(248,113,113,0.18);border-color:rgba(248,113,113,0.4);color:#fca5a5;}
.score-band{background:linear-gradient(135deg,rgba(108,99,255,0.25),rgba(52,211,153,0.15));
 border:1px solid rgba(108,99,255,0.35);border-radius:14px;padding:18px 24px;text-align:center;margin:12px 0;}
.lb-row{display:flex;justify-content:space-between;align-items:center;padding:9px 14px;
 background:rgba(255,255,255,0.04);border:1px solid rgba(255,255,255,0.08);
 border-radius:9px;margin-bottom:5px;font-size:.84rem;}
.lb-rank{font-size:1.1rem;font-weight:800;min-width:36px;color:#fbbf24;}
.lb-name{color:#fff;font-weight:600;flex:1;margin-left:8px;}
.lb-score{color:#34d399;font-weight:700;font-size:.9rem;}
.lb-badge{background:rgba(108,99,255,0.22);color:#a78bfa;border:1px solid rgba(108,99,255,0.35);
 border-radius:20px;padding:2px 10px;font-size:.72rem;font-weight:700;margin-left:8px;}
.timer-bar{height:6px;border-radius:4px;background:rgba(255,255,255,0.12);overflow:hidden;margin:8px 0;}
.timer-fill{height:100%;border-radius:4px;background:linear-gradient(90deg,#34d399,#38bdf8);
 transition:width .5s linear;}
</style>
"""

def _quiz_bank(domain):
    """Return (and initialise) the question bank for a domain."""
    if domain not in st.session_state.quiz_bank:
        st.session_state.quiz_bank[domain] = []
    return st.session_state.quiz_bank[domain]

# ─────────────────────────────────────────────────────────────────────────────
# ─────────────────────────────────────────────────────────────────────────────
# INSTRUCTOR: Quiz Builder (AI-generated questions per domain)
# ─────────────────────────────────────────────────────────────────────────────
def quiz_builder_instructor(user, domain):

    st.markdown(ARENA_CSS, unsafe_allow_html=True)
    sec("Quiz Builder — AI Question Generator", "🧩")

    bank = _quiz_bank(domain)

    # ================= HEADER =================
    st.markdown(f"""
    <div class="arena-card">
        <div style="color:#fff;font-weight:700;font-size:1rem;">
            📌 Domain: <span style="color:#34d399">{domain}</span>
        </div>
        <div style="color:rgba(255,255,255,.45);font-size:.8rem;margin-top:3px;">
            Click <b>Generate Questions</b> to create domain-specific MCQs using AI.
            Students use these in Quiz Attender.
        </div>
    </div>
    """, unsafe_allow_html=True)

    # ================= STATS =================
    total_q = len(bank)

    c1, c2, c3 = st.columns(3)
    with c1:
        pill("🧩", str(total_q), "Questions in Bank")

    with c2:
        matches = len([r for r in st.session_state.challenge_results if r["domain"] == domain])
        pill("🏆", str(matches), "Matches Played")

    with c3:
        wins = len([
            r for r in st.session_state.challenge_results
            if r["domain"] == domain and r["score_c"] > r["score_o"]
        ])
        pill("📡", str(wins), "Challenger Wins")

    st.markdown("<br>", unsafe_allow_html=True)

    # ================= GENERATOR =================
    sec("Generate AI Questions", "🤖")

    col_num, col_lvl = st.columns(2)

    with col_num:
        num_q = st.selectbox(
            "How many questions to generate?",
            [5, 10, 15],
            key="quiz_gen_num"
        )

    with col_lvl:
        q_level = st.selectbox(
            "Difficulty level",
            ["Beginner", "Intermediate", "Advanced"],
            key="quiz_gen_lvl"
        )

    # ================= BUTTON =================
    if st.button("🤖 Generate Questions with AI", type="primary", use_container_width=True):

        with st.spinner(f"Generating {num_q} {q_level} MCQ questions for {domain}…"):

            prompt = f"""
Generate exactly {num_q} multiple-choice questions (MCQs) about {domain} at {q_level} level.

Return ONLY a valid JSON array. No explanation, no markdown, no extra text.

Each object must have:
- "question"
- "A"
- "B"
- "C"
- "D"
- "answer" (A/B/C/D)

Example:
[{{"question":"What is...","A":"opt1","B":"opt2","C":"opt3","D":"opt4","answer":"A"}}]
"""

            try:
                # ✅ FIXED MODEL (IMPORTANT CHANGE)
                resp = groq_client.chat.completions.create(
                    model="llama-3.1-8b-instant",   # ✅ WORKING MODEL
                    messages=[
                        {
                            "role": "system",
                            "content": "You are a quiz generator. Output ONLY valid JSON."
                        },
                        {
                            "role": "user",
                            "content": prompt
                        }
                    ],
                    temperature=0.7,
                    max_tokens=2000,
                )

                raw = resp.choices[0].message.content.strip()

                # ================= CLEAN RESPONSE =================
                import re as _re, json as _json

                raw = _re.sub(r"^```.*?\n|```$", "", raw, flags=_re.DOTALL).strip()

                questions = _json.loads(raw)

                added = 0

                for q in questions:
                    if all(k in q for k in ["question", "A", "B", "C", "D", "answer"]):
                        bank.append({
                            "id":       make_uid(q["question"]),
                            "question": q["question"].strip(),
                            "options":  {
                                "A": q["A"],
                                "B": q["B"],
                                "C": q["C"],
                                "D": q["D"]
                            },
                            "answer":   q["answer"].strip().upper(),
                            "added_by": user["name"],
                            "level":    q_level,
                            "time":     now_str(),
                        })
                        added += 1

                st.success(f"✅ {added} questions generated and added! Total: {len(bank)}")
                st.rerun()

            except Exception as e:
                st.error(f"❌ AI generation failed: {e}")

    # ================= QUESTION BANK =================
    if bank:
        st.markdown("<br>", unsafe_allow_html=True)
        sec("Question Bank", "📋")

        for i, q in enumerate(reversed(bank), 1):

            short = q["question"][:70] + "…" if len(q["question"]) > 70 else q["question"]

            with st.expander(f"Q{len(bank)-i+1}. {short}"):

                st.markdown(
                    f"""
**A:** {q['options']['A']} &nbsp;·&nbsp;
**B:** {q['options']['B']} &nbsp;·&nbsp;
**C:** {q['options']['C']} &nbsp;·&nbsp;
**D:** {q['options']['D']}
""",
                    unsafe_allow_html=True
                )

                st.markdown(
                    f"✅ **Correct:** {q['answer']} &nbsp;·&nbsp; "
                    f"Level: {q.get('level','')} &nbsp;·&nbsp; "
                    f"Added by: {q['added_by']} &nbsp;·&nbsp; {q['time']}"
                )

                if st.button("🗑 Delete", key=f"qdel_{q['id']}"):
                    st.session_state.quiz_bank[domain] = [
                        x for x in bank if x["id"] != q["id"]
                    ]
                    st.rerun()

# ─────────────────────────────────────────────────────────────────────────────
# STUDENT: Challenge Arena
# ─────────────────────────────────────────────────────────────────────────────
def challenge_arena(user):
    st.markdown(ARENA_CSS, unsafe_allow_html=True)
    sec("Peer Challenge Arena", "🏆")

    enrolled = user.get("enrolled", [])
    if not enrolled:
        st.info("Enroll in a course first to access the Challenge Arena.")
        return

    # Init per-session challenge state
    if "arena_state" not in st.session_state:
        st.session_state.arena_state = "idle"   # idle | playing | result
    if "arena_answers" not in st.session_state:
        st.session_state.arena_answers = {}
    if "arena_q_set" not in st.session_state:
        st.session_state.arena_q_set = []
    if "arena_domain" not in st.session_state:
        st.session_state.arena_domain = ""
    if "arena_score" not in st.session_state:
        st.session_state.arena_score = 0

    # ── LEADERBOARD (always visible) ──────────────────────────────────────────
    results = st.session_state.challenge_results
    if results:
        sec("🏅 Global Leaderboard", "")
        scores = {}
        for r in results:
            for name, sc in [(r["challenger"], r["score_c"]), (r["opponent"], r["score_o"])]:
                if name not in scores:
                    scores[name] = {"played": 0, "total": 0, "wins": 0}
                scores[name]["played"] += 1
                scores[name]["total"]  += sc
                if name == r["challenger"] and r["score_c"] > r["score_o"]:
                    scores[name]["wins"] += 1
                elif name == r["opponent"] and r["score_o"] > r["score_c"]:
                    scores[name]["wins"] += 1
        sorted_lb = sorted(scores.items(), key=lambda x: (-x[1]["wins"], -x[1]["total"]))
        medals = ["🥇","🥈","🥉"]
        for rank, (name, stat) in enumerate(sorted_lb[:10], 1):
            medal = medals[rank-1] if rank <= 3 else str(rank)
            you   = " 👈 You" if name == user["name"] else ""
            badge_html = f'<span class="lb-badge">{stat["wins"]}W / {stat["played"]}P</span>'
            st.markdown(
                f'<div class="lb-row">'
                f'<span class="lb-rank">{medal}</span>'
                f'<span class="lb-name">{name}{you}</span>'
                f'<span class="lb-score">⭐ {stat["total"]} pts</span>'
                f'{badge_html}</div>',
                unsafe_allow_html=True)
        st.markdown("<br>", unsafe_allow_html=True)

    # ── IDLE STATE: configure a challenge ─────────────────────────────────────
    if st.session_state.arena_state == "idle":
        st.markdown('''
<div class="arena-card">
    <div style="color:#fff;font-weight:700;font-size:1rem;">How it works</div>
    <div style="color:rgba(255,255,255,.55);font-size:.85rem;margin-top:6px;line-height:1.6;">
        1. Choose a domain and enter your opponent's name.<br>
        2. Answer 5 questions from the instructor's question bank.<br>
        3. Your score is recorded. Your opponent can also play to compare scores.<br>
        4. The highest scorer wins the challenge. Top scorers rank on the leaderboard!
    </div>
</div>''', unsafe_allow_html=True)

        sec("Start a Challenge", "⚔️")
        col_dom, col_opp = st.columns(2)
        with col_dom:
            sel_domain = st.selectbox("Domain", enrolled, key="arena_domain_sel")
        with col_opp:
            opponent = st.text_input("Opponent name (any enrolled student)", placeholder="e.g. alice123", key="arena_opp")

        bank = _quiz_bank(sel_domain)
        if len(bank) < 5:
            st.warning(f"⚠️ Only {len(bank)} question(s) in this domain's bank. Your instructor needs to add at least 5 before you can play. You can still practice with what's available!")
            n_q = len(bank)
        else:
            n_q = 5

        if n_q > 0 and st.button("⚔️ Start Challenge!", type="primary", use_container_width=True):
            import random
            sample = random.sample(bank, min(n_q, len(bank)))
            st.session_state.arena_q_set  = sample
            st.session_state.arena_answers = {}
            st.session_state.arena_state   = "playing"
            st.session_state.arena_domain  = sel_domain
            st.session_state.arena_opp     = opponent.strip() if opponent.strip() else "Solo"
            st.session_state.arena_score   = 0
            st.session_state.arena_qidx    = 0
            st.rerun()

    # ── PLAYING STATE: answer questions one by one ────────────────────────────
    elif st.session_state.arena_state == "playing":
        q_set  = st.session_state.arena_q_set
        idx    = st.session_state.get("arena_qidx", 0)
        total  = len(q_set)

        if idx >= total:
            # All answered — compute score
            sc = 0
            for qi, q in enumerate(q_set):
                if st.session_state.arena_answers.get(qi) == q["answer"]:
                    sc += 1
            st.session_state.arena_score = sc
            st.session_state.arena_state = "result"
            st.rerun()
        else:
            q = q_set[idx]
            progress = (idx + 1) / total
            st.markdown(
                f'<div class="timer-bar"><div class="timer-fill" style="width:{int(progress*100)}%"></div></div>',
                unsafe_allow_html=True)
            st.markdown(
                f'<div style="color:rgba(255,255,255,.45);font-size:.78rem;margin-bottom:14px;">'
                f'Question {idx+1} of {total}  ·  Domain: {st.session_state.arena_domain}</div>',
                unsafe_allow_html=True)
            st.markdown(f'<div class="arena-card"><div class="arena-q">{q["question"]}</div></div>', unsafe_allow_html=True)
            chosen = None
            for letter, text in q["options"].items():
                if st.button(f"{letter}.  {text}", key=f"arena_opt_{idx}_{letter}", use_container_width=True):
                    chosen = letter
            if chosen:
                st.session_state.arena_answers[idx] = chosen
                st.session_state.arena_qidx = idx + 1
                st.rerun()

    # ── RESULT STATE: show score and log it ───────────────────────────────────
    elif st.session_state.arena_state == "result":
        sc    = st.session_state.arena_score
        total = len(st.session_state.arena_q_set)
        pct   = int(sc / total * 100) if total else 0
        grade = "🏆 Excellent!" if pct >= 80 else "👍 Good job!" if pct >= 60 else "📚 Keep practising!"

        st.markdown(f'''
<div class="score-band">
    <div style="font-size:2.8rem;margin-bottom:6px;">{grade.split()[0]}</div>
    <div style="color:#fff;font-size:1.5rem;font-weight:800;">{sc} / {total}</div>
    <div style="color:rgba(255,255,255,.55);font-size:.9rem;margin-top:4px;">{grade[2:]}  ·  {pct}%</div>
    <div style="color:rgba(255,255,255,.4);font-size:.8rem;margin-top:6px;">
        vs {st.session_state.get("arena_opp","—")}  ·  {st.session_state.arena_domain}
    </div>
</div>''', unsafe_allow_html=True)

        # Show correct answers
        sec("Review Answers", "📋")
        for qi, q in enumerate(st.session_state.arena_q_set):
            given   = st.session_state.arena_answers.get(qi, "—")
            correct = q["answer"]
            icon    = "✅" if given == correct else "❌"
            st.markdown(
                f'<div class="arena-card" style="margin-bottom:8px;">'
                f'<div class="arena-q">{icon} {q["question"]}</div>'
                f'<div style="font-size:.84rem;color:rgba(255,255,255,.6);">'
                f'Your answer: <b style="color:{("#6ee7b7" if given==correct else "#fca5a5")}">{given}</b>  '
                f'&nbsp;·&nbsp;  Correct: <b style="color:#34d399">{correct}</b>  '
                f'— {q["options"][correct]}</div></div>',
                unsafe_allow_html=True)

        # Log result only on first render of result state
        log_key = f"_arena_logged_{st.session_state.get('arena_log_id','')}"
        if log_key not in st.session_state:
            rid = make_uid(user["email"] + str(sc))
            st.session_state[f"_arena_logged_{rid}"] = True
            st.session_state.arena_log_id = rid
            st.session_state.challenge_results.append({
                "id":         rid,
                "challenger": user["name"],
                "opponent":   st.session_state.get("arena_opp", "Solo"),
                "domain":     st.session_state.arena_domain,
                "score_c":    sc,
                "score_o":    0,          # opponent fills in when they play
                "total":      total,
                "time":       now_str(),
            })

        col_r, col_n = st.columns(2)
        with col_r:
            if st.button("🔄 Play Again", use_container_width=True):
                st.session_state.arena_state = "idle"
                st.rerun()
        with col_n:
            if st.button("🏠 Back to Home", use_container_width=True):
                st.session_state.arena_state = "idle"
                st.session_state.user["enrolled"] = user.get("enrolled", [])
                st.rerun()


# ─────────────────────────────────────────────────────────────────────────────
# ADMIN: Challenge Monitor
# ─────────────────────────────────────────────────────────────────────────────
def challenge_monitor_admin():
    st.markdown(ARENA_CSS, unsafe_allow_html=True)
    sec("Challenge Arena Monitor", "🏆")

    results  = st.session_state.challenge_results
    quiz_bank = st.session_state.quiz_bank

    # KPI row
    total_q = sum(len(v) for v in quiz_bank.values())
    c1,c2,c3,c4 = st.columns(4)
    with c1: pill("🧩", str(total_q),       "Total Questions")
    with c2: pill("⚔️",  str(len(results)),  "Matches Played")
    with c3: pill("🌐", str(len(quiz_bank)), "Domains with Questions")
    with c4:
        avg_sc = (sum(r["score_c"] for r in results)/len(results) if results else 0)
        pill("⭐", f"{avg_sc:.1f}", "Avg Score")

    st.markdown("<br>", unsafe_allow_html=True)

    # Quiz bank summary
    sec("Question Bank by Domain", "🧩")
    rows = []
    for d in ALL_DOMAINS:
        q_count = len(quiz_bank.get(d, []))
        instrs  = list({q["added_by"] for q in quiz_bank.get(d, [])}) if q_count else []
        rows.append({"Domain": d, "Questions": q_count,
                     "Status": "✅ Ready" if q_count >= 5 else ("⚠️ Partial" if q_count > 0 else "❌ Empty"),
                     "Added by": ", ".join(instrs) if instrs else "—"})
    st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # Match history
    sec("Match History", "⚔️")
    if not results:
        st.info("No challenge matches played yet.")
    else:
        for r in reversed(results):
            win = "🏆" if r["score_c"] >= r["score_o"] else "📊"
            st.markdown(
                f'<div class="lb-row">'
                f'<span style="font-size:1rem;">{win}</span>'
                f'<span class="lb-name">{r["challenger"]} <span style="color:rgba(255,255,255,.35);">vs</span> {r["opponent"]}</span>'
                f'<span style="color:#38bdf8;font-size:.8rem;margin-right:12px;">{r["domain"]}</span>'
                f'<span class="lb-score">{r["score_c"]}/{r["total"]}</span>'
                f'<span style="color:rgba(255,255,255,.35);font-size:.75rem;margin-left:10px;">{r["time"]}</span>'
                f'</div>',
                unsafe_allow_html=True)

        # Export
        df_exp = pd.DataFrame(results)
        st.download_button("⬇ Export Match History CSV", df_exp.to_csv(index=False),
                           "challenge_history.csv", "text/csv")


# ─────────────────────────────────────────────────────────────────────────────
# STUDENT: Quiz Attender
# ─────────────────────────────────────────────────────────────────────────────
def quiz_attender_student(user):
    st.markdown(ARENA_CSS, unsafe_allow_html=True)
    sec("Quiz Attender", "🎯")

    enrolled = user.get("enrolled", [])
    if not enrolled:
        st.info("Enroll in a course first to access quizzes.")
        return

    # Domain selector
    sel_domain = st.selectbox("Select Domain", enrolled, key="qa_domain_sel")
    bank = _quiz_bank(sel_domain)

    if not bank:
        st.info(f"No questions available yet for **{sel_domain}**. Your instructor will add questions soon.")
        return

    # Session state keys for this quiz attempt
    _qa_key  = f"qa_attempt_{sel_domain}"
    _ans_key = f"qa_answers_{sel_domain}"
    _done_key= f"qa_done_{sel_domain}"

    if _qa_key not in st.session_state:
        st.session_state[_qa_key]  = False   # started?
    if _ans_key not in st.session_state:
        st.session_state[_ans_key] = {}      # {q_idx: chosen_option}
    if _done_key not in st.session_state:
        st.session_state[_done_key] = False   # submitted?

    # ── NOT STARTED: show start screen ────────────────────────────
    if not st.session_state[_qa_key]:
        st.markdown(f"""
<div class="arena-card">
    <div style="color:#fff;font-weight:700;font-size:1rem;">
        📋 {sel_domain} Quiz
    </div>
    <div style="color:rgba(255,255,255,.5);font-size:.85rem;margin-top:8px;line-height:1.6;">
        • {len(bank)} question(s) available<br>
        • Select one option per question<br>
        • Click <b>Submit Quiz</b> when done to see your score
    </div>
</div>""", unsafe_allow_html=True)
        if st.button("▶ Start Quiz", type="primary", use_container_width=True, key="qa_start"):
            st.session_state[_qa_key]  = True
            st.session_state[_ans_key] = {}
            st.session_state[_done_key]= False
            st.rerun()

    # ── STARTED & NOT SUBMITTED: show all questions ────────────────
    elif st.session_state[_qa_key] and not st.session_state[_done_key]:
        st.markdown(f"""
<div style="color:rgba(255,255,255,.45);font-size:.8rem;margin-bottom:16px;">
    Domain: <b style="color:#38bdf8">{sel_domain}</b>
    &nbsp;·&nbsp; {len(bank)} question(s)
    &nbsp;·&nbsp; Select one option per question, then click Submit
</div>""", unsafe_allow_html=True)

        answers = st.session_state[_ans_key]

        for idx, q in enumerate(bank):
            st.markdown(
                f'<div class="arena-card">'
                f'<div class="arena-q">Q{idx+1}. {q["question"]}</div>'
                f'</div>',
                unsafe_allow_html=True
            )
            options = [f"{k}. {v}" for k, v in q["options"].items()]
            current = answers.get(idx)
            # Build display list — pre-select if already answered
            default_idx = 0
            if current:
                # Find index of currently selected option
                for oi, opt_str in enumerate(options):
                    if opt_str.startswith(current + "."):
                        default_idx = oi
                        break
            choice = st.radio(
                f"q{idx}",
                options,
                index=default_idx,
                key=f"qa_radio_{sel_domain}_{idx}",
                label_visibility="collapsed"
            )
            # Store the letter only (first character)
            answers[idx] = choice[0]
            st.session_state[_ans_key] = answers
            st.markdown("<br>", unsafe_allow_html=True)

        col1, col2 = st.columns(2)
        with col1:
            if st.button("✅ Submit Quiz", type="primary", use_container_width=True, key="qa_submit"):
                # Check all answered
                if len(answers) < len(bank):
                    st.warning("Please answer all questions before submitting.")
                else:
                    st.session_state[_done_key] = True
                    st.rerun()
        with col2:
            if st.button("🔄 Reset", use_container_width=True, key="qa_reset"):
                st.session_state[_qa_key]  = False
                st.session_state[_ans_key] = {}
                st.session_state[_done_key]= False
                st.rerun()

    # ── SUBMITTED: show results ────────────────────────────────────
    elif st.session_state[_done_key]:
        answers = st.session_state[_ans_key]
        score   = sum(1 for i, q in enumerate(bank) if answers.get(i) == q["answer"])
        total   = len(bank)
        pct     = int(score / total * 100) if total else 0
        grade   = "🏆 Excellent!" if pct >= 80 else "👍 Good!" if pct >= 60 else "📚 Keep Practising!"

        st.markdown(f"""
<div class="score-band">
    <div style="font-size:2.5rem;margin-bottom:6px;">{grade.split()[0]}</div>
    <div style="color:#fff;font-size:1.6rem;font-weight:800;">{score} / {total}</div>
    <div style="color:rgba(255,255,255,.55);font-size:.9rem;margin-top:4px;">
        {grade[2:]}  ·  {pct}%  ·  {sel_domain}
    </div>
</div>""", unsafe_allow_html=True)

        sec("Review Answers", "📋")
        for idx, q in enumerate(bank):
            given   = answers.get(idx, "—")
            correct = q["answer"]
            icon    = "✅" if given == correct else "❌"
            g_color = "#6ee7b7" if given == correct else "#fca5a5"
            st.markdown(
                f'<div class="arena-card" style="margin-bottom:8px;">'
                f'<div class="arena-q">{icon} Q{idx+1}. {q["question"]}</div>'
                f'<div style="font-size:.84rem;color:rgba(255,255,255,.65);margin-top:6px;">'
                f'Your answer: <b style="color:{g_color}">{given}. {q["options"].get(given,"—")}</b>'
                f'&nbsp;&nbsp;|&nbsp;&nbsp;'
                f'Correct: <b style="color:#34d399">{correct}. {q["options"][correct]}</b>'
                f'</div></div>',
                unsafe_allow_html=True
            )

        if st.button("🔄 Retake Quiz", use_container_width=True, key="qa_retake"):
            st.session_state[_qa_key]  = False
            st.session_state[_ans_key] = {}
            st.session_state[_done_key]= False
            st.rerun()

# ══════════════════════════════════════════════════════════════════════════════
# ROUTER
# ══════════════════════════════════════════════════════════════════════════════
_user = st.session_state.user
if _user is None:
    page_login()
elif _user["role"] == "Student":
    page_student()
elif _user["role"] == "Instructor":
    page_instructor()
elif _user["role"] == "Admin":
    page_admin()
else:
    st.error("Unknown role.")
    if st.button("Log Out"):
        st.session_state.user = None
        st.rerun()





# dir
# cd "EduPro - Copy1(new)"
# dir
# py -m streamlit run edupro_dashboard.py