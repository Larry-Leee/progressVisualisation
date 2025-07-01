import streamlit as st
import pandas as pd
from docx import Document
from pyecharts.charts import Bar
from pyecharts import options as opts
from streamlit_echarts import st_pyecharts


def extract_table_3_2(docx_file):
    doc = Document(docx_file)
    match_tables = []
    for table in doc.tables:
        headers = [cell.text.strip() for cell in table.rows[0].cells]
        if "分部工程" in headers and "本月计划工程量" in headers and "本月完成工程量" in headers:
            rows = []
            for row in table.rows[1:]:
                rows.append([cell.text.strip() for cell in row.cells])
            df = pd.DataFrame(rows, columns=headers)
            match_tables.append(df[['分部工程', '本月计划工程量', '本月完成工程量']])

    # 取第2个匹配表（表3.2）
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
            title_opts=opts.TitleOpts(title="表3.2 输水工程计划 vs 实际对比图"),
            tooltip_opts=opts.TooltipOpts(trigger="axis"),
            xaxis_opts=opts.AxisOpts(axislabel_opts={"rotate": 45}),
            datazoom_opts=[opts.DataZoomOpts(type_="slider")],
        )
    )
    return bar

st.set_page_config(layout="wide")
st.title("📊 表3.2 输水工程计划 vs 实际对比柱状图")

uploaded_file = st.file_uploader("请上传月报 Word 文件（.docx）", type=["docx"])

if uploaded_file:
    st.success("✅ 上传成功，正在读取表3.2…")
    df = extract_table_3_2(uploaded_file)

    if df is not None:
        st.subheader("📄 提取出的表3.2数据")
        st.dataframe(df)

        st.subheader("📈 自动生成计划 vs 实际对比柱状图")
        chart = plot_plan_vs_actual(df)
        st_pyecharts(chart)

    else:
        st.warning("⚠ 未找到表3.2，请检查文档格式是否一致。")
