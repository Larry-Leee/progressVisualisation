import streamlit as st
import pandas as pd
from docx import Document
import mysql.connector
from pyecharts.charts import Bar
from pyecharts import options as opts
from streamlit_echarts import st_pyecharts
from datetime import datetime

# ========== MySQL 配置 ==========
DB_CONFIG = {
    "host": "localhost",
    "user": "root",
    "password": "8brs5r00",
    "database": "water_progress",
    "charset": "utf8mb4"
}

# ========== 最近月份列表 ==========
def get_recent_months(n=12):
    today = datetime.today()
    return [
        (today.replace(day=1) - pd.DateOffset(months=i)).strftime("%Y-%m")
        for i in range(n)
    ]

# ========== 提取 Word 表格 ==========
def extract_table_3_2(docx_file):
    doc = Document(docx_file)
    for table in doc.tables:
        headers = [cell.text.strip() for cell in table.rows[0].cells]
        if "本月计划工程量" in headers and "本月完成工程量" in headers:
            df = []
            for row in table.rows[1:]:
                df.append([cell.text.strip() for cell in row.cells])
            df = pd.DataFrame(df, columns=headers)
            return df
    return None

# ========== 字段映射与清洗 ==========
def process_dataframe(df_raw, report_month, source_file):
    df = df_raw.rename(columns={
        '名称': 'project_name',
        '本月计划工程量': 'plan_amount',
        '本月完成工程量': 'actual_amount'
    })
    df['plan_amount'] = pd.to_numeric(df['plan_amount'], errors='coerce')
    df['actual_amount'] = pd.to_numeric(df['actual_amount'], errors='coerce')
    df.dropna(subset=['project_name', 'plan_amount', 'actual_amount'], inplace=True)
    df['report_month'] = report_month
    df['source_file'] = source_file
    return df[['report_month', 'project_name', 'plan_amount', 'actual_amount', 'source_file']]

# ========== 写入数据库 ==========
def insert_to_mysql(df):
    conn = mysql.connector.connect(**DB_CONFIG)
    cursor = conn.cursor()
    sql = """
        INSERT INTO water_project_progress
        (report_month, project_name, plan_amount, actual_amount, source_file)
        VALUES (%s, %s, %s, %s, %s)
    """
    for _, row in df.iterrows():
        cursor.execute(sql, tuple(row))
    conn.commit()
    cursor.close()
    conn.close()

# ========== 查询当月数据 ==========
def get_month_data(report_month):
    conn = mysql.connector.connect(**DB_CONFIG)
    df = pd.read_sql(f"""
        SELECT project_name, plan_amount, actual_amount
        FROM water_project_progress
        WHERE report_month = '{report_month}'
    """, conn)
    conn.close()
    return df

# ========== 查询累计数据 ==========
def get_cumulative_data():
    conn = mysql.connector.connect(**DB_CONFIG)
    df = pd.read_sql("""
        SELECT project_name,
               SUM(plan_amount) AS total_plan,
               SUM(actual_amount) AS total_actual
        FROM water_project_progress
        GROUP BY project_name
    """, conn)
    conn.close()
    return df

# ========== 柱状图绘制 ==========
def plot_bar_chart(x, plan, actual, title):
    bar = (
        Bar()
        .add_xaxis(x)
        .add_yaxis("计划", plan)
        .add_yaxis("实际", actual)
        .set_global_opts(
            title_opts=opts.TitleOpts(title=title),
            xaxis_opts=opts.AxisOpts(axislabel_opts={"rotate": 45}),
            tooltip_opts=opts.TooltipOpts(trigger="axis"),
            datazoom_opts=[opts.DataZoomOpts(type_="slider")],
        )
    )
    return bar

# ========== 页面逻辑 ==========
st.set_page_config(layout="wide")
st.title("📊 输水工程月报数据上传与对比分析系统")

report_month = st.selectbox("📆 请选择报表所属月份", get_recent_months())
uploaded_file = st.file_uploader("📄 上传 Word 文件（含表3.2）", type=["docx"])

if uploaded_file and report_month:
    df_raw = extract_table_3_2(uploaded_file)
    if df_raw is not None:
        df = process_dataframe(df_raw, report_month, uploaded_file.name)
        st.subheader("✅ 识别出的数据")
        st.dataframe(df)

        if st.button("💾 写入数据库"):
            insert_to_mysql(df)
            st.success("✅ 数据已写入 MySQL 数据库")

        st.subheader("📈 当月 vs 累计图表")
        col1, col2 = st.columns(2)

        with col1:
            df_month = get_month_data(report_month)
            if not df_month.empty:
                chart1 = plot_bar_chart(
                    x=df_month['project_name'].tolist(),
                    plan=df_month['plan_amount'].tolist(),
                    actual=df_month['actual_amount'].tolist(),
                    title=f"{report_month}：当月计划 vs 实际"
                )
                st_pyecharts(chart1)
            else:
                st.info("📭 当前月份暂无数据")

        with col2:
            df_cum = get_cumulative_data()
            if not df_cum.empty:
                chart2 = plot_bar_chart(
                    x=df_cum['project_name'].tolist(),
                    plan=df_cum['total_plan'].tolist(),
                    actual=df_cum['total_actual'].tolist(),
                    title="📊 累计计划 vs 实际"
                )
                st_pyecharts(chart2)
            else:
                st.info("📭 尚无累计数据")
    else:
        st.warning("⚠ 未找到表3.2，请确认 Word 表格格式是否正确")