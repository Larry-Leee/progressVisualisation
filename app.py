import streamlit as st
import pandas as pd
from docx import Document
from pyecharts.charts import Bar
from pyecharts import options as opts
from streamlit_echarts import st_pyecharts
from fpdf import FPDF
import os

def extract_table_3_2(docx_file):
    doc = Document(docx_file)
    match_tables = []

    for idx, table in enumerate(doc.tables):
        headers = [cell.text.strip() for cell in table.rows[0].cells]
        header_text = ''.join(headers)
        if ("分部" in header_text and "计划" in header_text and "完成" in header_text and "设计" in header_text and "开累" in header_text):
            rows = []
            for row in table.rows[1:]:
                rows.append([cell.text.strip() for cell in row.cells])
            df = pd.DataFrame(rows, columns=headers)

            def find_col(cols, keyword):
                return next((c for c in cols if keyword in c), None)

            col_fb = find_col(headers, "分部")
            col_sj = find_col(headers, "设计")
            col_kl = find_col(headers, "开累")
            col_jh = find_col(headers, "计划")
            col_wc = find_col(headers, "完成")

            if all([col_fb, col_sj, col_kl, col_jh, col_wc]):
                df = df[[col_fb, col_sj, col_kl, col_jh, col_wc]]
                df.columns = ['分部工程', '设计工程量', '开累完成工程量', '本月计划工程量', '本月完成工程量']
                match_tables.append(df)

    if len(match_tables) >= 2:
        return match_tables[1]
    elif match_tables:
        return match_tables[0]
    else:
        return None

def plot_plan_vs_actual(df):
    df = df[['分部工程', '本月计划工程量', '本月完成工程量']].copy()
    df['本月计划工程量'] = pd.to_numeric(df['本月计划工程量'], errors='coerce')
    df['本月完成工程量'] = pd.to_numeric(df['本月完成工程量'], errors='coerce')
    df.dropna(inplace=True)

    names = df['分部工程'].tolist()
    plan = df['本月计划工程量'].tolist()
    actual = df['本月完成工程量'].tolist()

    bar = (
        Bar()
        .add_xaxis(names)
        .add_yaxis("计划", plan)
        .add_yaxis("实际", actual)
        .set_global_opts(
            title_opts=opts.TitleOpts(title="计划工程量 vs 实际工程量"),
            tooltip_opts=opts.TooltipOpts(trigger="axis"),
            xaxis_opts=opts.AxisOpts(axislabel_opts={"rotate": 45}),
            datazoom_opts=[opts.DataZoomOpts(type_="slider")],
        )
    )
    return bar

def plot_design_vs_accum(df):
    df = df[['分部工程', '设计工程量', '开累完成工程量']].copy()
    df['设计工程量'] = pd.to_numeric(df['设计工程量'], errors='coerce')
    df['开累完成工程量'] = pd.to_numeric(df['开累完成工程量'], errors='coerce')
    df.dropna(inplace=True)

    names = df['分部工程'].tolist()
    design = df['设计工程量'].tolist()
    accum = df['开累完成工程量'].tolist()

    bar = (
        Bar()
        .add_xaxis(names)
        .add_yaxis("设计工程量", design)
        .add_yaxis("开累完成工程量", accum)
        .set_global_opts(
            title_opts=opts.TitleOpts(title="设计工程量 vs 开累完成工程量"),
            tooltip_opts=opts.TooltipOpts(trigger="axis"),
            xaxis_opts=opts.AxisOpts(axislabel_opts={"rotate": 45}),
            datazoom_opts=[opts.DataZoomOpts(type_="slider")],
        )
    )
    return bar

# ------------------ Streamlit 页面 ------------------
st.set_page_config(layout="wide")
st.markdown(
    """
    <h1 style='text-align:center; font-size:42px; color:#1ABC9C; font-weight:bold;'>
    重庆市藻渡水库隧洞进度可视化管理系统
    </h1>
    """, unsafe_allow_html=True
)

uploaded_files = st.file_uploader(
    "请上传 Word 月报文件（可批量上传 .docx）",
    type=["docx"],
    accept_multiple_files=True
)

if uploaded_files:
    month_data = {}  # 存储每个文件的数据
    st.success(f"✅ 共上传 {len(uploaded_files)} 个文件，正在解析...")

    for uploaded_file in uploaded_files:
        df = extract_table_3_2(uploaded_file)
        if df is not None:
            month_data[uploaded_file.name] = df
        else:
            st.warning(f"{uploaded_file.name} 未找到表3.2，请检查文档格式。")

    # 分月展示图表
    for month, df in month_data.items():
        with st.expander(f"📊 {month} 数据分析"):
            st.subheader("计划工程量 vs 实际工程量")
            chart1 = plot_plan_vs_actual(df)
            st_pyecharts(chart1)

            st.subheader("设计工程量 vs 开累完成工程量")
            chart2 = plot_design_vs_accum(df)
            st_pyecharts(chart2)


    if st.button("生成 PDF 报告"):
        pdf = FPDF()
        pdf.set_auto_page_break(auto=True, margin=15)

        for month, df in month_data.items():
            pdf.add_page()
            pdf.set_font("Arial", "B", 16)
            pdf.cell(0, 10, f"{month} 月报分析", ln=True, align="C")

            # 添加表格
            pdf.set_font("Arial", "", 12)
            for i, row in df.iterrows():
                line = f"{row['分部工程']} | {row['设计工程量']} | {row['开累完成工程量']} | {row['本月计划工程量']} | {row['本月完成工程量']}"
                pdf.cell(0, 8, line, ln=True)

            # 保存图表为 PNG 并插入 PDF
            chart1 = plot_plan_vs_actual(df)
            chart2 = plot_design_vs_accum(df)
            chart1.render(f"{month}_chart1.png")
            chart2.render(f"{month}_chart2.png")
            pdf.image(f"{month}_chart1.png", x=10, w=180)
            pdf.image(f"{month}_chart2.png", x=10, w=180)

        pdf_file = "批量月报分析报告.pdf"
        pdf.output(pdf_file)

        # 提供下载
        with open(pdf_file, "rb") as f:
            st.download_button(
                label="📥 下载 PDF 报告",
                data=f,
                file_name=pdf_file,
                mime="application/pdf"
            )
        st.success("PDF 生成完成 ✅")