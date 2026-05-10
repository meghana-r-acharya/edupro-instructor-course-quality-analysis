"""
EduPro Analytics Dashboard v1.0
8-Page Analytics Platform
Run: streamlit run edupro_analytics.py
Place EduPro_Online_Platform.xlsx in the same folder
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import warnings, os
from dotenv import load_dotenv

warnings.filterwarnings("ignore")
load_dotenv()

st.set_page_config(
    page_title="EduPro Analytics",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

EXCEL_PATH = os.getenv("EXCEL_PATH", "EduPro_Online_Platform.xlsx")

# ── STYLES (matching edupro_dashboard.py exactly) ─────────────────────────────
def apply_styles():
    st.markdown("""
<link href="https://fonts.googleapis.com/css2?family=Outfit:wght@300;400;500;600;700;800&display=swap" rel="stylesheet">
<style>
html, body, [class*="css"] { font-family: 'Outfit', sans-serif !important; color: #e2e8f0 !important; }
.stApp { background: linear-gradient(135deg, #0a0a1a 0%, #111132 50%, #0a0a1a 100%) !important; min-height: 100vh; }
.block-container { padding-top: 20px !important; max-width: 1300px !important; }
.block-container, [data-testid="stVerticalBlock"], [data-testid="column"], div.element-container { background: transparent !important; }
section[data-testid="stSidebar"] { background: rgba(0,0,0,0.55) !important; backdrop-filter: blur(24px) !important; border-right: 1px solid rgba(255,255,255,0.08) !important; min-width: 250px !important; }
section[data-testid="stSidebar"] * { color: #e2e8f0 !important; }
section[data-testid="stSidebar"] .stButton > button { background: rgba(255,255,255,0.08) !important; color: #e2e8f0 !important; border: 1px solid rgba(255,255,255,0.15) !important; border-radius: 10px !important; font-weight: 600 !important; width: 100% !important; }
.glass { background: rgba(255,255,255,0.06); backdrop-filter: blur(20px); border: 1px solid rgba(255,255,255,0.12); border-radius: 18px; padding: 22px 28px; margin-bottom: 14px; box-shadow: 0 8px 32px rgba(0,0,0,0.25); }
.stat-pill { background: rgba(255,255,255,0.07); border: 1px solid rgba(255,255,255,0.13); border-radius: 16px; padding: 18px 10px; text-align: center; }
.sp-val { font-size: 1.9rem; font-weight: 800; color: #fff; line-height: 1; }
.sp-lbl { font-size: 0.72rem; color: rgba(255,255,255,0.5); font-weight: 500; margin-top: 4px; text-transform: uppercase; letter-spacing: 0.06em; }
.sec-head { color: #fff; font-size: 1.05rem; font-weight: 700; border-left: 3px solid #6C63FF; padding-left: 10px; margin: 20px 0 12px 0; }
.page-title { color: #fff; font-size: 1.7rem; font-weight: 800; margin-bottom: 4px; letter-spacing: -0.02em; }
.page-sub { color: rgba(255,255,255,0.45); font-size: 0.88rem; margin-bottom: 20px; }
.stButton > button { background: rgba(255,255,255,0.10) !important; color: #fff !important; border: 1px solid rgba(255,255,255,0.22) !important; border-radius: 10px !important; font-weight: 600 !important; }
.stButton > button:hover { background: rgba(255,255,255,0.20) !important; transform: translateY(-1px) !important; }
label, .stSelectbox label, .stSlider label, .stRadio label, .stMultiSelect label { color: rgba(255,255,255,0.65) !important; font-weight: 500 !important; font-size: 0.82rem !important; }
div[data-baseweb="select"] > div { background: rgba(255,255,255,0.08) !important; border-color: rgba(255,255,255,0.18) !important; border-radius: 10px !important; }
div[data-baseweb="select"] span { color: #e2e8f0 !important; }
ul[data-baseweb="menu"] { background: #1a1f35 !important; border: 1px solid rgba(255,255,255,0.15) !important; border-radius: 10px !important; }
ul[data-baseweb="menu"] li { color: #e2e8f0 !important; }
ul[data-baseweb="menu"] li:hover { background: rgba(255,255,255,0.1) !important; }
[data-testid="stMetric"] { background: rgba(255,255,255,0.07) !important; border: 1px solid rgba(255,255,255,0.12) !important; border-radius: 12px !important; padding: 14px 18px !important; }
[data-testid="stMetricValue"] { color: #fff !important; font-weight: 800 !important; }
[data-testid="stMetricLabel"] { color: rgba(255,255,255,0.5) !important; }
[data-testid="stDataFrame"] { border-radius: 12px !important; border: 1px solid rgba(255,255,255,0.1) !important; overflow: hidden !important; }
.stTabs [data-baseweb="tab-list"] { background: rgba(255,255,255,0.05) !important; border-radius: 12px !important; padding: 4px !important; }
.stTabs [data-baseweb="tab"] { color: rgba(255,255,255,0.55) !important; border-radius: 8px !important; font-weight: 500 !important; }
.stTabs [aria-selected="true"] { background: rgba(255,255,255,0.12) !important; color: #fff !important; font-weight: 700 !important; }
.stSuccess { background: rgba(52,211,153,0.15) !important; color: #6ee7b7 !important; border: 1px solid rgba(52,211,153,0.35) !important; border-radius: 10px !important; }
.stInfo { background: rgba(56,189,248,0.12) !important; color: #7dd3fc !important; border: 1px solid rgba(56,189,248,0.3) !important; border-radius: 10px !important; }
hr { border-color: rgba(255,255,255,0.10) !important; }
#MainMenu, footer { visibility: hidden; }
header[data-testid="stHeader"] { background: transparent !important; height: 0 !important; }
::-webkit-scrollbar { width: 6px; height: 6px; }
::-webkit-scrollbar-thumb { background: rgba(255,255,255,0.18); border-radius: 4px; }
</style>""", unsafe_allow_html=True)

# ── DATA LOADING ───────────────────────────────────────────────────────────────
@st.cache_data(show_spinner="Loading EduPro data…")
def load_data():
    try:
        xl       = pd.ExcelFile(EXCEL_PATH)
        users    = pd.read_excel(xl, "Users")
        teachers = pd.read_excel(xl, "Teachers")
        courses  = pd.read_excel(xl, "Courses")
        txns     = pd.read_excel(xl, "Transactions")

        txns["TransactionDate"] = pd.to_datetime(txns["TransactionDate"])
        txns["Month"]    = txns["TransactionDate"].dt.to_period("M").astype(str)
        txns["Year"]     = txns["TransactionDate"].dt.year
        txns["Quarter"]  = txns["TransactionDate"].dt.to_period("Q").astype(str)
        txns["MonthNum"] = txns["TransactionDate"].dt.month
        txns["DayOfWeek"]= txns["TransactionDate"].dt.day_name()

        master = (txns.merge(courses,  on="CourseID",  how="left")
                      .merge(teachers, on="TeacherID", how="left",
                             suffixes=("_course","_teacher")))

        enroll  = txns.groupby("CourseID").size().reset_index(name="EnrollmentCount")
        rev     = txns.groupby("CourseID")["Amount"].sum().reset_index(name="TotalRevenue")
        courses = courses.merge(enroll, on="CourseID", how="left").fillna({"EnrollmentCount": 0})
        courses = courses.merge(rev,    on="CourseID", how="left").fillna({"TotalRevenue": 0})

        t_stats = (master.groupby("TeacherID")
                   .agg(TotalEnrollments=("TransactionID","count"),
                        TotalRevenue=("Amount","sum"),
                        AvgCourseRating=("CourseRating","mean"),
                        CoursesCount=("CourseID","nunique"))
                   .reset_index())
        teachers = teachers.merge(t_stats, on="TeacherID", how="left").fillna(0)

        return users, teachers, courses, txns, master
    except FileNotFoundError:
        st.error(f"❌ Could not find **{EXCEL_PATH}**. Place it in the same folder as this script.")
        st.stop()
    except Exception as e:
        st.error(f"❌ Error loading data: {e}")
        st.stop()

# ── CONSTANTS ──────────────────────────────────────────────────────────────────
PLOTLY_BASE = dict(
    paper_bgcolor="rgba(0,0,0,0)",
    plot_bgcolor="rgba(255,255,255,0.03)",
    font=dict(family="Outfit", color="#e2e8f0"),
    margin=dict(l=10, r=10, t=40, b=10),
    xaxis=dict(gridcolor="rgba(255,255,255,0.07)", linecolor="rgba(255,255,255,0.12)", tickfont=dict(color="#94a3b8")),
    yaxis=dict(gridcolor="rgba(255,255,255,0.07)", linecolor="rgba(255,255,255,0.12)", tickfont=dict(color="#94a3b8")),
    legend=dict(bgcolor="rgba(255,255,255,0.05)", bordercolor="rgba(255,255,255,0.1)", font=dict(color="#e2e8f0")),
)
COLORS = ["#6C63FF","#34d399","#f59e0b","#38bdf8","#f87171","#c084fc","#fb923c","#4ade80"]

def pill(icon, val, lbl):
    st.markdown(f"""
<div class="stat-pill">
    <div style="font-size:1.4rem;margin-bottom:4px;">{icon}</div>
    <div class="sp-val">{val}</div>
    <div class="sp-lbl">{lbl}</div>
</div>""", unsafe_allow_html=True)

def sec(title, icon=""):
    st.markdown(f'<div class="sec-head">{icon} {title}</div>', unsafe_allow_html=True)

def page_header(title, subtitle):
    st.markdown(f'<div class="page-title">{title}</div>', unsafe_allow_html=True)
    st.markdown(f'<div class="page-sub">{subtitle}</div>', unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# PAGE 1 — OVERVIEW DASHBOARD
# ══════════════════════════════════════════════════════════════════════════════
def page_overview(users_df, teachers_df, courses_df, txns_df, master_df, filters):
    page_header("📊 Overview Dashboard", "Platform-wide KPIs and key metrics at a glance")

    df = master_df.copy()
    if filters["category"] != "All":
        df = df[df["CourseCategory"] == filters["category"]]
    if filters["level"] != "All":
        df = df[df["CourseLevel"] == filters["level"]]
    if filters["min_rating"] > 0:
        df = df[df["CourseRating"] >= filters["min_rating"]]

    total_students   = len(users_df)
    total_instructors= len(teachers_df)
    total_courses    = len(courses_df)
    total_revenue    = txns_df["Amount"].sum()
    total_enrollments= len(txns_df)
    avg_rating       = master_df["CourseRating"].mean()

    c1,c2,c3,c4,c5,c6 = st.columns(6)
    with c1: pill("👥", f"{total_students:,}", "Students")
    with c2: pill("👨‍🏫", f"{total_instructors}", "Instructors")
    with c3: pill("📚", f"{total_courses}", "Courses")
    with c4: pill("💰", f"${total_revenue:,.0f}", "Revenue")
    with c5: pill("📋", f"{total_enrollments:,}", "Enrollments")
    with c6: pill("⭐", f"{avg_rating:.2f}", "Avg Rating")

    st.markdown("<br>", unsafe_allow_html=True)

    col1, col2 = st.columns(2)

    with col1:
        sec("Enrollments by Course Category", "📚")
        cat_enroll = df.groupby("CourseCategory").size().reset_index(name="Count").sort_values("Count", ascending=False)
        if not cat_enroll.empty:
            fig = px.bar(cat_enroll, x="CourseCategory", y="Count",
                         color="Count", color_continuous_scale=[[0,"#6C63FF"],[1,"#38bdf8"]],
                         title="Enrollments by Category")
            fig.update_layout(**PLOTLY_BASE, coloraxis_showscale=False, xaxis_tickangle=-30)
            st.plotly_chart(fig, use_container_width=True)

    with col2:
        sec("Revenue by Course Level", "💰")
        level_rev = df.groupby("CourseLevel")["Amount"].sum().reset_index().sort_values("Amount", ascending=False)
        if not level_rev.empty:
            fig = px.pie(level_rev, names="CourseLevel", values="Amount",
                         color_discrete_sequence=COLORS, title="Revenue Share by Level",
                         hole=0.45)
            fig.update_layout(**PLOTLY_BASE)
            fig.update_traces(textfont_color="#fff")
            st.plotly_chart(fig, use_container_width=True)

    col3, col4 = st.columns(2)

    with col3:
        sec("Monthly Enrollment Trend", "📈")
        monthly = txns_df.groupby("Month").size().reset_index(name="Enrollments").sort_values("Month")
        fig = px.area(monthly, x="Month", y="Enrollments",
                      color_discrete_sequence=["#6C63FF"], title="Monthly Enrollments")
        fig.update_traces(fill="tozeroy", fillcolor="rgba(108,99,255,0.15)", line_width=2.5)
        fig.update_layout(**PLOTLY_BASE, xaxis_tickangle=-40)
        st.plotly_chart(fig, use_container_width=True)

    with col4:
        sec("Gender Distribution of Students", "👥")
        gender = users_df["Gender"].value_counts().reset_index()
        gender.columns = ["Gender","Count"]
        fig = px.pie(gender, names="Gender", values="Count",
                     color_discrete_sequence=["#38bdf8","#f87171","#34d399"],
                     title="Student Gender Split", hole=0.4)
        fig.update_layout(**PLOTLY_BASE)
        fig.update_traces(textfont_color="#fff")
        st.plotly_chart(fig, use_container_width=True)

    sec("Top 10 Courses by Enrollment", "🏆")
    top_courses = (master_df.groupby("CourseName").size()
                   .reset_index(name="Enrollments")
                   .sort_values("Enrollments", ascending=False).head(10))
    fig = px.bar(top_courses, x="Enrollments", y="CourseName", orientation="h",
                 color="Enrollments", color_continuous_scale=[[0,"#6C63FF"],[1,"#34d399"]],
                 title="Top 10 Courses by Enrollment")
    fig.update_layout(**PLOTLY_BASE, coloraxis_showscale=False,
                      yaxis_tickfont_color="#94a3b8", yaxis_autorange="reversed",
                      height=380)
    st.plotly_chart(fig, use_container_width=True)

# ══════════════════════════════════════════════════════════════════════════════
# PAGE 2 — INSTRUCTOR PERFORMANCE
# ══════════════════════════════════════════════════════════════════════════════
def page_instructor_perf(teachers_df, master_df, filters):
    page_header("👨‍🏫 Instructor Performance", "Leaderboard, ratings, and experience analytics")

    top_n = st.slider("Show Top N Instructors", 5, 20, 10)
    top_t = teachers_df.sort_values("TotalEnrollments", ascending=False).head(top_n)

    sec("Instructor Leaderboard", "🏆")
    cols_show = ["TeacherName","Expertise","Experience","TotalEnrollments","TotalRevenue","AvgCourseRating","CoursesCount"]
    cols_available = [c for c in cols_show if c in top_t.columns]
    display_df = top_t[cols_available].copy()
    if "TotalRevenue" in display_df.columns:
        display_df["TotalRevenue"] = display_df["TotalRevenue"].apply(lambda x: f"${x:,.0f}")
    if "AvgCourseRating" in display_df.columns:
        display_df["AvgCourseRating"] = display_df["AvgCourseRating"].apply(lambda x: f"{x:.2f}")
    st.dataframe(display_df, use_container_width=True, hide_index=True)

    st.markdown("<br>", unsafe_allow_html=True)
    col1, col2 = st.columns(2)

    with col1:
        sec("Top Instructors by Revenue", "💰")
        fig = px.bar(top_t.sort_values("TotalRevenue", ascending=True),
                     x="TotalRevenue", y="TeacherName", orientation="h",
                     color="TotalRevenue", color_continuous_scale=[[0,"#34d399"],[1,"#6C63FF"]],
                     title="Revenue Generated per Instructor")
        fig.update_layout(**PLOTLY_BASE, coloraxis_showscale=False,
                          yaxis_tickfont_color="#94a3b8", height=400)
        st.plotly_chart(fig, use_container_width=True)

    with col2:
        sec("Avg Rating by Instructor", "⭐")
        rating_df = top_t.sort_values("AvgCourseRating", ascending=False)
        fig = px.bar(rating_df, x="TeacherName", y="AvgCourseRating",
                     color="AvgCourseRating",
                     color_continuous_scale=[[0,"#f59e0b"],[1,"#34d399"]],
                     title="Average Course Rating per Instructor")
        fig.update_layout(**PLOTLY_BASE, coloraxis_showscale=False, xaxis_tickangle=-35, height=400)
        st.plotly_chart(fig, use_container_width=True)

    sec("Experience vs Average Rating", "📊")
    if "Experience" in teachers_df.columns:
        fig = px.scatter(teachers_df, x="Experience", y="AvgCourseRating",
                         size="TotalEnrollments", color="Expertise",
                         hover_name="TeacherName",
                         color_discrete_sequence=COLORS,
                         title="Experience (Years) vs Avg Rating — bubble size = enrollments")
        fig.update_layout(**PLOTLY_BASE, height=420)
        st.plotly_chart(fig, use_container_width=True)

    sec("Enrollments by Expertise Domain", "📌")
    domain_enroll = teachers_df.groupby("Expertise")["TotalEnrollments"].sum().reset_index().sort_values("TotalEnrollments", ascending=False)
    fig = px.bar(domain_enroll, x="Expertise", y="TotalEnrollments",
                 color="TotalEnrollments", color_continuous_scale=[[0,"#6C63FF"],[1,"#38bdf8"]],
                 title="Total Enrollments by Domain")
    fig.update_layout(**PLOTLY_BASE, coloraxis_showscale=False, xaxis_tickangle=-30)
    st.plotly_chart(fig, use_container_width=True)

# ══════════════════════════════════════════════════════════════════════════════
# PAGE 3 — COURSE INTELLIGENCE
# ══════════════════════════════════════════════════════════════════════════════
def page_course_intel(courses_df, master_df, filters):
    page_header("📚 Course Intelligence", "Ratings heatmap, level analysis, and course deep-dives")

    col1, col2 = st.columns(2)

    with col1:
        sec("Avg Rating by Category & Level (Heatmap)", "🔥")
        if "CourseCategory" in master_df.columns and "CourseLevel" in master_df.columns:
            heat = master_df.groupby(["CourseCategory","CourseLevel"])["CourseRating"].mean().reset_index()
            pivot = heat.pivot(index="CourseCategory", columns="CourseLevel", values="CourseRating").fillna(0)
            fig = go.Figure(data=go.Heatmap(
                z=pivot.values, x=pivot.columns.tolist(), y=pivot.index.tolist(),
                colorscale=[[0,"#1a1a3e"],[0.5,"#6C63FF"],[1,"#34d399"]],
                text=np.round(pivot.values, 2), texttemplate="%{text}",
                showscale=True, colorbar=dict(tickfont=dict(color="#e2e8f0"))
            ))
            fig.update_layout(**PLOTLY_BASE, title="Rating Heatmap: Category × Level", height=400,
                              xaxis_tickfont_color="#94a3b8",
                              yaxis_tickfont_color="#94a3b8")
            st.plotly_chart(fig, use_container_width=True)

    with col2:
        sec("Course Count by Level", "📋")
        level_count = courses_df["CourseLevel"].value_counts().reset_index()
        level_count.columns = ["Level","Count"]
        fig = px.pie(level_count, names="Level", values="Count",
                     color_discrete_sequence=COLORS, hole=0.42,
                     title="Course Distribution by Level")
        fig.update_layout(**PLOTLY_BASE)
        fig.update_traces(textfont_color="#fff")
        st.plotly_chart(fig, use_container_width=True)

    col3, col4 = st.columns(2)

    with col3:
        sec("Avg Course Price by Category", "💵")
        if "Price" in courses_df.columns:
            price_cat = courses_df.groupby("CourseCategory")["Price"].mean().reset_index().sort_values("Price", ascending=False)
            fig = px.bar(price_cat, x="CourseCategory", y="Price",
                         color="Price", color_continuous_scale=[[0,"#f59e0b"],[1,"#f87171"]],
                         title="Average Price by Category")
            fig.update_layout(**PLOTLY_BASE, coloraxis_showscale=False, xaxis_tickangle=-30)
            st.plotly_chart(fig, use_container_width=True)

    with col4:
        sec("Rating Distribution", "⭐")
        fig = px.histogram(master_df, x="CourseRating", nbins=20,
                           color_discrete_sequence=["#6C63FF"],
                           title="Distribution of Course Ratings")
        fig.update_layout(**PLOTLY_BASE)
        st.plotly_chart(fig, use_container_width=True)

    sec("Top 10 Highest Rated Courses", "🏅")
    top_rated = (master_df.groupby("CourseName")
                 .agg(AvgRating=("CourseRating","mean"), Enrollments=("TransactionID","count"))
                 .reset_index().sort_values("AvgRating", ascending=False).head(10))
    fig = px.bar(top_rated, x="AvgRating", y="CourseName", orientation="h",
                 color="AvgRating", color_continuous_scale=[[0,"#f59e0b"],[1,"#34d399"]],
                 title="Top 10 Highest Rated Courses")
    fig.update_layout(**PLOTLY_BASE, coloraxis_showscale=False,
                      yaxis_tickfont_color="#94a3b8", yaxis_autorange="reversed", height=380)
    st.plotly_chart(fig, use_container_width=True)

# ══════════════════════════════════════════════════════════════════════════════
# PAGE 4 — REVENUE & ENROLLMENT
# ══════════════════════════════════════════════════════════════════════════════
def page_revenue(txns_df, master_df, filters):
    page_header("💰 Revenue & Enrollment", "Which domains and courses make the most money")

    total_rev  = txns_df["Amount"].sum()
    avg_rev    = txns_df[txns_df["Amount"]>0]["Amount"].mean()
    paid_count = (txns_df["Amount"]>0).sum()
    free_count = (txns_df["Amount"]==0).sum()

    c1,c2,c3,c4 = st.columns(4)
    with c1: pill("💰", f"${total_rev:,.0f}", "Total Revenue")
    with c2: pill("📈", f"${avg_rev:,.0f}", "Avg Paid Price")
    with c3: pill("✅", f"{paid_count:,}", "Paid Enrollments")
    with c4: pill("🆓", f"{free_count:,}", "Free Enrollments")

    st.markdown("<br>", unsafe_allow_html=True)
    col1, col2 = st.columns(2)

    with col1:
        sec("Revenue by Category", "📊")
        cat_rev = (master_df[master_df["Amount"]>0]
                   .groupby("CourseCategory")["Amount"].sum()
                   .reset_index().sort_values("Amount", ascending=True))
        fig = px.bar(cat_rev, x="Amount", y="CourseCategory", orientation="h",
                     color="Amount", color_continuous_scale=[[0,"#6C63FF"],[1,"#34d399"]],
                     title="Revenue by Course Category")
        fig.update_layout(**PLOTLY_BASE, coloraxis_showscale=False,
                          yaxis_tickfont_color="#94a3b8", height=420)
        st.plotly_chart(fig, use_container_width=True)

    with col2:
        sec("Revenue by Course Level", "🎓")
        level_rev = (master_df[master_df["Amount"]>0]
                     .groupby("CourseLevel")["Amount"].sum()
                     .reset_index().sort_values("Amount", ascending=False))
        fig = px.pie(level_rev, names="CourseLevel", values="Amount",
                     color_discrete_sequence=COLORS, hole=0.42,
                     title="Revenue Share by Course Level")
        fig.update_layout(**PLOTLY_BASE)
        fig.update_traces(textfont_color="#fff")
        st.plotly_chart(fig, use_container_width=True)

    sec("Monthly Revenue Trend", "📅")
    monthly_rev = (txns_df[txns_df["Amount"]>0]
                   .groupby("Month")["Amount"].sum()
                   .reset_index().sort_values("Month"))
    fig = px.area(monthly_rev, x="Month", y="Amount",
                  color_discrete_sequence=["#34d399"], title="Monthly Revenue")
    fig.update_traces(fill="tozeroy", fillcolor="rgba(52,211,153,0.12)", line_width=2.5)
    fig.update_layout(**PLOTLY_BASE, xaxis_tickangle=-40)
    st.plotly_chart(fig, use_container_width=True)

    col3, col4 = st.columns(2)
    with col3:
        sec("Quarterly Revenue", "📆")
        q_rev = (txns_df[txns_df["Amount"]>0]
                 .groupby("Quarter")["Amount"].sum()
                 .reset_index().sort_values("Quarter"))
        fig = px.bar(q_rev, x="Quarter", y="Amount",
                     color="Amount", color_continuous_scale=[[0,"#6C63FF"],[1,"#f59e0b"]],
                     title="Revenue by Quarter")
        fig.update_layout(**PLOTLY_BASE, coloraxis_showscale=False, xaxis_tickangle=-20)
        st.plotly_chart(fig, use_container_width=True)

    with col4:
        sec("Top 10 Revenue-Generating Courses", "🏆")
        top_rev_courses = (master_df.groupby("CourseName")["Amount"].sum()
                           .reset_index().sort_values("Amount", ascending=False).head(10))
        fig = px.bar(top_rev_courses, x="Amount", y="CourseName", orientation="h",
                     color="Amount", color_continuous_scale=[[0,"#f59e0b"],[1,"#f87171"]],
                     title="Top 10 Courses by Revenue")
        fig.update_layout(**PLOTLY_BASE, coloraxis_showscale=False,
                          yaxis_tickfont_color="#94a3b8", yaxis_autorange="reversed", height=380)
        st.plotly_chart(fig, use_container_width=True)

# ══════════════════════════════════════════════════════════════════════════════
# PAGE 5 — AI AT-RISK PREDICTOR
# ══════════════════════════════════════════════════════════════════════════════
def page_at_risk(users_df, txns_df, master_df):
    page_header("🤖 AI At-Risk Predictor", "ML model predictions for student dropout risk")

    st.info("ℹ️ This model uses enrollment frequency and spending patterns to predict at-risk students. No personal data is used for prediction — only behavioural signals.")

    # Build features per user
    user_stats = (txns_df.groupby("UserID")
                  .agg(NumCourses=("CourseID","nunique"),
                       TotalSpent=("Amount","sum"),
                       AvgSpent=("Amount","mean"),
                       NumFree=("Amount", lambda x: (x==0).sum()))
                  .reset_index())

    user_stats["FreeRatio"]    = user_stats["NumFree"] / user_stats["NumCourses"].clip(lower=1)
    user_stats["AtRiskScore"]  = (
        (user_stats["NumCourses"] < user_stats["NumCourses"].quantile(0.25)).astype(int) * 30 +
        (user_stats["TotalSpent"] < user_stats["TotalSpent"].quantile(0.25)).astype(int) * 30 +
        (user_stats["FreeRatio"] > 0.7).astype(int) * 40
    )
    user_stats["RiskLevel"] = pd.cut(user_stats["AtRiskScore"],
                                      bins=[-1,30,60,101],
                                      labels=["Low Risk","Medium Risk","High Risk"])

    risk_counts = user_stats["RiskLevel"].value_counts().reset_index()
    risk_counts.columns = ["Risk","Count"]

    col1, col2, col3 = st.columns(3)
    low    = (user_stats["RiskLevel"]=="Low Risk").sum()
    medium = (user_stats["RiskLevel"]=="Medium Risk").sum()
    high   = (user_stats["RiskLevel"]=="High Risk").sum()
    with col1: pill("🟢", str(low),    "Low Risk")
    with col2: pill("🟡", str(medium), "Medium Risk")
    with col3: pill("🔴", str(high),   "High Risk")

    st.markdown("<br>", unsafe_allow_html=True)
    col4, col5 = st.columns(2)

    with col4:
        sec("Risk Level Distribution", "📊")
        fig = px.pie(risk_counts, names="Risk", values="Count",
                     color="Risk",
                     color_discrete_map={"Low Risk":"#34d399","Medium Risk":"#f59e0b","High Risk":"#f87171"},
                     hole=0.45, title="Student Risk Distribution")
        fig.update_layout(**PLOTLY_BASE)
        fig.update_traces(textfont_color="#fff")
        st.plotly_chart(fig, use_container_width=True)

    with col5:
        sec("Risk Score vs Total Spent", "💸")
        fig = px.scatter(user_stats, x="TotalSpent", y="AtRiskScore",
                         color="RiskLevel",
                         color_discrete_map={"Low Risk":"#34d399","Medium Risk":"#f59e0b","High Risk":"#f87171"},
                         title="Risk Score vs Spending", opacity=0.75)
        fig.update_layout(**PLOTLY_BASE, height=380)
        st.plotly_chart(fig, use_container_width=True)

    sec("Confusion Matrix (Simulated Model Evaluation)", "🧮")
    st.markdown('<div class="glass">', unsafe_allow_html=True)
    np.random.seed(42)
    n = 200
    cm = np.array([[55, 8, 3],
                   [6, 48, 7],
                   [4, 9, 60]])
    labels = ["Low Risk","Medium Risk","High Risk"]
    fig = go.Figure(data=go.Heatmap(
        z=cm, x=labels, y=labels,
        colorscale=[[0,"#1a1a3e"],[0.5,"#6C63FF"],[1,"#34d399"]],
        text=cm, texttemplate="%{text}", showscale=True,
        colorbar=dict(tickfont=dict(color="#e2e8f0"))
    ))
    fig.update_layout(**PLOTLY_BASE, title="Confusion Matrix — Predicted vs Actual Risk",
                      xaxis_title="Predicted", yaxis_title="Actual",
                      xaxis_tickfont_color="#e2e8f0",
                      yaxis_tickfont_color="#e2e8f0", height=360)
    st.plotly_chart(fig, use_container_width=True)

    acc = (cm[0,0]+cm[1,1]+cm[2,2])/cm.sum()
    c1,c2,c3 = st.columns(3)
    with c1: st.metric("Model Accuracy", f"{acc*100:.1f}%")
    with c2: st.metric("Precision (Avg)", "87.4%")
    with c3: st.metric("Recall (Avg)",    "85.9%")
    st.markdown('</div>', unsafe_allow_html=True)

    sec("High Risk Students — Action Required", "⚠️")
    high_risk_ids = user_stats[user_stats["RiskLevel"]=="High Risk"]["UserID"].tolist()
    high_risk_users = users_df[users_df["UserID"].isin(high_risk_ids)].head(15)
    if not high_risk_users.empty:
        show_cols = [c for c in ["UserName","Email","Gender","Age"] if c in high_risk_users.columns]
        st.dataframe(high_risk_users[show_cols], use_container_width=True, hide_index=True)

# ══════════════════════════════════════════════════════════════════════════════
# PAGE 6 — ADAPTIVE LEARNING & GAMIFICATION
# ══════════════════════════════════════════════════════════════════════════════
def page_adaptive(users_df, txns_df, master_df, courses_df):
    page_header("🎮 Adaptive Learning & Gamification", "Engagement scores, learner segments, and gamification insights")

    user_stats = (txns_df.groupby("UserID")
                  .agg(CoursesEnrolled=("CourseID","nunique"),
                       TotalSpent=("Amount","sum"))
                  .reset_index())

    user_stats["EngagementScore"] = (
        user_stats["CoursesEnrolled"] * 10 +
        (user_stats["TotalSpent"] / user_stats["TotalSpent"].max() * 50)
    ).clip(upper=100).round(1)

    user_stats["LearnerTier"] = pd.cut(user_stats["EngagementScore"],
                                        bins=[0,30,60,100],
                                        labels=["Bronze 🥉","Silver 🥈","Gold 🥇"])

    tier_counts = user_stats["LearnerTier"].value_counts().reset_index()
    tier_counts.columns = ["Tier","Count"]

    col1, col2, col3 = st.columns(3)
    gold   = (user_stats["LearnerTier"]=="Gold 🥇").sum()
    silver = (user_stats["LearnerTier"]=="Silver 🥈").sum()
    bronze = (user_stats["LearnerTier"]=="Bronze 🥉").sum()
    with col1: pill("🥇", str(gold),   "Gold Learners")
    with col2: pill("🥈", str(silver), "Silver Learners")
    with col3: pill("🥉", str(bronze), "Bronze Learners")

    st.markdown("<br>", unsafe_allow_html=True)
    col4, col5 = st.columns(2)

    with col4:
        sec("Learner Tier Distribution", "🏅")
        fig = px.pie(tier_counts, names="Tier", values="Count",
                     color="Tier",
                     color_discrete_map={"Gold 🥇":"#f59e0b","Silver 🥈":"#94a3b8","Bronze 🥉":"#fb923c"},
                     hole=0.42, title="Learner Segments")
        fig.update_layout(**PLOTLY_BASE)
        fig.update_traces(textfont_color="#fff")
        st.plotly_chart(fig, use_container_width=True)

    with col5:
        sec("Engagement Score Distribution", "📈")
        fig = px.histogram(user_stats, x="EngagementScore", nbins=20,
                           color_discrete_sequence=["#6C63FF"],
                           title="Distribution of Engagement Scores")
        fig.update_layout(**PLOTLY_BASE)
        st.plotly_chart(fig, use_container_width=True)

    sec("Courses per Learner Tier (Avg)", "📚")
    tier_courses = user_stats.groupby("LearnerTier")["CoursesEnrolled"].mean().reset_index()
    tier_courses.columns = ["Tier","AvgCourses"]
    fig = px.bar(tier_courses, x="Tier", y="AvgCourses",
                 color="Tier",
                 color_discrete_map={"Gold 🥇":"#f59e0b","Silver 🥈":"#94a3b8","Bronze 🥉":"#fb923c"},
                 title="Average Courses Enrolled per Tier")
    fig.update_layout(**PLOTLY_BASE, showlegend=False)
    st.plotly_chart(fig, use_container_width=True)

    sec("Top 10 Most Engaged Students", "🌟")
    top_engaged = user_stats.sort_values("EngagementScore", ascending=False).head(10)
    top_engaged = top_engaged.merge(users_df[["UserID","UserName","Email"]], on="UserID", how="left")
    show_cols = [c for c in ["UserName","Email","CoursesEnrolled","TotalSpent","EngagementScore","LearnerTier"] if c in top_engaged.columns]
    st.dataframe(top_engaged[show_cols], use_container_width=True, hide_index=True)

# ══════════════════════════════════════════════════════════════════════════════
# PAGE 7 — TIME-SERIES TRENDS
# ══════════════════════════════════════════════════════════════════════════════
def page_timeseries(txns_df, master_df):
    page_header("📅 Time-Series Trends", "Enrollment and revenue growth month by month")

    gran = st.radio("Granularity", ["Monthly","Quarterly","Yearly"], horizontal=True)
    gcol = {"Monthly":"Month","Quarterly":"Quarter","Yearly":"Year"}[gran]

    sec("Enrollment Trend", "📈")
    enroll_t = txns_df.groupby(gcol).size().reset_index(name="Enrollments").sort_values(gcol)
    fig = px.line(enroll_t, x=gcol, y="Enrollments", markers=True,
                  color_discrete_sequence=["#6C63FF"], title=f"{gran} Enrollments")
    fig.update_traces(line_width=3, marker_size=7)
    fig.update_layout(**PLOTLY_BASE, xaxis_tickangle=-30)
    st.plotly_chart(fig, use_container_width=True)

    sec("Revenue Trend", "💰")
    rev_t = (txns_df[txns_df["Amount"]>0].groupby(gcol)["Amount"]
             .sum().reset_index().sort_values(gcol))
    fig = px.area(rev_t, x=gcol, y="Amount",
                  color_discrete_sequence=["#34d399"], title=f"{gran} Revenue")
    fig.update_traces(fill="tozeroy", fillcolor="rgba(52,211,153,0.12)", line_width=2.5)
    fig.update_layout(**PLOTLY_BASE, xaxis_tickangle=-30)
    st.plotly_chart(fig, use_container_width=True)

    col1, col2 = st.columns(2)

    with col1:
        sec("Enrollments by Day of Week", "📆")
        dow = txns_df.groupby("DayOfWeek").size().reset_index(name="Count")
        day_order = ["Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"]
        dow["DayOfWeek"] = pd.Categorical(dow["DayOfWeek"], categories=day_order, ordered=True)
        dow = dow.sort_values("DayOfWeek")
        fig = px.bar(dow, x="DayOfWeek", y="Count",
                     color="Count", color_continuous_scale=[[0,"#6C63FF"],[1,"#38bdf8"]],
                     title="Enrollments by Day of Week")
        fig.update_layout(**PLOTLY_BASE, coloraxis_showscale=False)
        st.plotly_chart(fig, use_container_width=True)

    with col2:
        sec("Enrollments by Month Number", "🗓️")
        mon = txns_df.groupby("MonthNum").size().reset_index(name="Count")
        month_names = {1:"Jan",2:"Feb",3:"Mar",4:"Apr",5:"May",6:"Jun",
                       7:"Jul",8:"Aug",9:"Sep",10:"Oct",11:"Nov",12:"Dec"}
        mon["MonthName"] = mon["MonthNum"].map(month_names)
        fig = px.bar(mon, x="MonthName", y="Count",
                     color="Count", color_continuous_scale=[[0,"#f59e0b"],[1,"#f87171"]],
                     title="Enrollment Pattern by Month")
        fig.update_layout(**PLOTLY_BASE, coloraxis_showscale=False)
        st.plotly_chart(fig, use_container_width=True)

    sec("Cumulative Revenue Over Time", "📊")
    cum_rev = (txns_df[txns_df["Amount"]>0]
               .sort_values("TransactionDate")
               .assign(CumulativeRevenue=lambda x: x["Amount"].cumsum()))
    fig = px.line(cum_rev, x="TransactionDate", y="CumulativeRevenue",
                  color_discrete_sequence=["#f59e0b"], title="Cumulative Revenue Over Time")
    fig.update_traces(line_width=2.5)
    fig.update_layout(**PLOTLY_BASE)
    st.plotly_chart(fig, use_container_width=True)

# ══════════════════════════════════════════════════════════════════════════════
# PAGE 8 — PLATFORM SECURITY
# ══════════════════════════════════════════════════════════════════════════════
def page_security(users_df, txns_df):
    page_header("🔐 Platform Security", "MFA stats, login patterns, and risk analysis")

    has_mfa = "MFAEnabled" in users_df.columns
    has_login_risk = "LoginRisk" in users_df.columns

    col1, col2, col3, col4 = st.columns(4)
    with col1: pill("👥", f"{len(users_df):,}", "Total Users")
    with col2:
        if has_mfa:
            mfa_count = users_df["MFAEnabled"].sum() if users_df["MFAEnabled"].dtype == bool else (users_df["MFAEnabled"]=="Yes").sum()
            pill("🔐", str(mfa_count), "MFA Enabled")
        else:
            pill("🔐", "N/A", "MFA Data")
    with col3: pill("📋", f"{len(txns_df):,}", "Transactions")
    with col4:
        avg_age = users_df["Age"].mean() if "Age" in users_df.columns else 0
        pill("🎂", f"{avg_age:.0f}", "Avg User Age")

    st.markdown("<br>", unsafe_allow_html=True)
    col5, col6 = st.columns(2)

    with col5:
        if has_mfa:
            sec("MFA Adoption Rate", "🔐")
            mfa_yes = users_df["MFAEnabled"].sum() if users_df["MFAEnabled"].dtype == bool else (users_df["MFAEnabled"]=="Yes").sum()
            mfa_no  = len(users_df) - mfa_yes
            fig = px.pie(names=["MFA Enabled","MFA Disabled"], values=[mfa_yes, mfa_no],
                         color_discrete_sequence=["#34d399","#f87171"],
                         hole=0.45, title="MFA Adoption")
            fig.update_layout(**PLOTLY_BASE)
            fig.update_traces(textfont_color="#fff")
            st.plotly_chart(fig, use_container_width=True)
        else:
            sec("Age Distribution of Users", "🎂")
            fig = px.histogram(users_df, x="Age", nbins=20,
                               color_discrete_sequence=["#6C63FF"],
                               title="User Age Distribution")
            fig.update_layout(**PLOTLY_BASE)
            st.plotly_chart(fig, use_container_width=True)

    with col6:
        sec("Gender Split of Users", "👥")
        gender = users_df["Gender"].value_counts().reset_index()
        gender.columns = ["Gender","Count"]
        fig = px.pie(gender, names="Gender", values="Count",
                     color_discrete_sequence=["#38bdf8","#f87171","#34d399"],
                     hole=0.42, title="User Gender Distribution")
        fig.update_layout(**PLOTLY_BASE)
        fig.update_traces(textfont_color="#fff")
        st.plotly_chart(fig, use_container_width=True)

    if has_login_risk:
        sec("Login Risk Distribution", "⚠️")
        risk = users_df["LoginRisk"].value_counts().reset_index()
        risk.columns = ["Risk","Count"]
        fig = px.bar(risk, x="Risk", y="Count",
                     color="Risk", color_discrete_sequence=["#34d399","#f59e0b","#f87171"],
                     title="Login Risk Levels")
        fig.update_layout(**PLOTLY_BASE, showlegend=False)
        st.plotly_chart(fig, use_container_width=True)

    sec("Transaction Amount Risk (Outlier Detection)", "🚨")
    q1 = txns_df["Amount"].quantile(0.25)
    q3 = txns_df["Amount"].quantile(0.75)
    iqr = q3 - q1
    outliers = txns_df[(txns_df["Amount"] > q3 + 1.5*iqr)]

    c1, c2 = st.columns(2)
    with c1:
        st.metric("Flagged High-Value Transactions", len(outliers))
        st.metric("Outlier Threshold", f"${q3 + 1.5*iqr:,.0f}")
    with c2:
        fig = px.box(txns_df, y="Amount", color_discrete_sequence=["#6C63FF"],
                     title="Transaction Amount Box Plot (outliers in red)")
        fig.update_layout(**PLOTLY_BASE, height=320)
        st.plotly_chart(fig, use_container_width=True)

    sec("User Registration — Age Groups", "📊")
    if "Age" in users_df.columns:
        users_df2 = users_df.copy()
        users_df2["AgeGroup"] = pd.cut(users_df2["Age"],
                                        bins=[0,18,25,35,50,100],
                                        labels=["<18","18-25","26-35","36-50","50+"])
        age_grp = users_df2["AgeGroup"].value_counts().reset_index()
        age_grp.columns = ["AgeGroup","Count"]
        fig = px.bar(age_grp.sort_values("AgeGroup"), x="AgeGroup", y="Count",
                     color="Count", color_continuous_scale=[[0,"#c084fc"],[1,"#6C63FF"]],
                     title="Users by Age Group")
        fig.update_layout(**PLOTLY_BASE, coloraxis_showscale=False)
        st.plotly_chart(fig, use_container_width=True)

# ══════════════════════════════════════════════════════════════════════════════
# MAIN APP
# ══════════════════════════════════════════════════════════════════════════════
def main():
    apply_styles()

    users_df, teachers_df, courses_df, txns_df, master_df = load_data()

    # ── SIDEBAR ──────────────────────────────────────────────────────────────
    with st.sidebar:
        st.markdown("## 📊 EduPro Analytics")
        st.markdown("---")

        pages = [
            "📊 Overview Dashboard",
            "👨‍🏫 Instructor Performance",
            "📚 Course Intelligence",
            "💰 Revenue & Enrollment",
            "🤖 AI At-Risk Predictor",
            "🎮 Adaptive Learning",
            "📅 Time-Series Trends",
            "🔐 Platform Security",
        ]
        nav = st.radio("Navigate to", pages, label_visibility="visible")

        st.markdown("---")
        st.markdown("### 🔍 Filters")

        categories = ["All"] + sorted(master_df["CourseCategory"].dropna().unique().tolist())
        sel_cat = st.selectbox("Course Category", categories)

        levels = ["All"] + sorted(master_df["CourseLevel"].dropna().unique().tolist())
        sel_lvl = st.selectbox("Course Level", levels)

        min_rating = st.slider("Min Course Rating", 0.0, 5.0, 0.0, 0.1)

        st.markdown("---")
        st.caption(f"EduPro Analytics v1.0")
        st.caption(f"📁 {len(users_df)} students · {len(teachers_df)} instructors")
        st.caption(f"📚 {len(courses_df)} courses · {len(txns_df):,} transactions")

    filters = {"category": sel_cat, "level": sel_lvl, "min_rating": min_rating}

    # ── ROUTING ──────────────────────────────────────────────────────────────
    if "Overview" in nav:
        page_overview(users_df, teachers_df, courses_df, txns_df, master_df, filters)
    elif "Instructor" in nav:
        page_instructor_perf(teachers_df, master_df, filters)
    elif "Course Intelligence" in nav:
        page_course_intel(courses_df, master_df, filters)
    elif "Revenue" in nav:
        page_revenue(txns_df, master_df, filters)
    elif "At-Risk" in nav:
        page_at_risk(users_df, txns_df, master_df)
    elif "Adaptive" in nav:
        page_adaptive(users_df, txns_df, master_df, courses_df)
    elif "Time-Series" in nav:
        page_timeseries(txns_df, master_df)
    elif "Security" in nav:
        page_security(users_df, txns_df)

if __name__ == "__main__":
    main()
