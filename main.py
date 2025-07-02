import streamlit as st
import pandas as pd
from docx import Document
from pyecharts.charts import Bar
from pyecharts import options as opts
from streamlit_echarts import st_pyecharts

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

            # 定义一个列名查找函数（根据关键词）
            def find_col(cols, keyword):
                return next((c for c in cols if keyword in c), None)

            col_fb = find_col(headers, "分部")        # 例如“分部工程”
            col_sj = find_col(headers, "设计")        # 例如“设计工程量”或“设计量”
            col_kl = find_col(headers, "开累")        # 例如“开累完成工程量”或“累计完成工程量”
            col_jh = find_col(headers, "计划")        # 例如“本月计划工程量”
            col_wc = find_col(headers, "完成")        # 例如“本月完成工程量”

            if all([col_fb, col_sj, col_kl, col_jh, col_wc]):
                df = df[[col_fb, col_sj, col_kl, col_jh, col_wc]]
                df.columns = ['分部工程', '设计工程量', '开累完成工程量', '本月计划工程量', '本月完成工程量']  # 重命名统一
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

    global names
    names = df['分部工程'].tolist()
    plan = df['本月计划工程量'].tolist()
    actual = df['本月完成工程量'].tolist()

    bar = (
        Bar()
        .add_xaxis(names)
        .add_yaxis("计划", plan)
        .add_yaxis("实际", actual)
        .set_global_opts(
            title_opts=opts.TitleOpts(title="计划工程量vs 实际工程量"),
            tooltip_opts=opts.TooltipOpts(trigger="axis"),
            xaxis_opts=opts.AxisOpts(axislabel_opts={"rotate": 45}),
            datazoom_opts=[opts.DataZoomOpts(type_="slider")],
        )
    )
    return bar, names

st.set_page_config(layout="wide")
st.title("📊输水工程计划 vs 实际对比柱状图")


def bar_plan_and_accumlative(df, names):
    df = df[['分部工程', '设计工程量', '开累完成工程量']].copy()
    df['设计工程量'] = pd.to_numeric(df['设计工程量'], errors='coerce')
    df['开累完成工程量'] = pd.to_numeric(df['开累完成工程量'], errors='coerce')
    df.dropna(inplace=True)


    design_total = df['设计工程量'].tolist()
    accum = df['开累完成工程量'].tolist()

    bar = (
        Bar()
        .add_xaxis(names)
        .add_yaxis('设计工程量', design_total)
        .add_yaxis('开累完成工作量', accum)
        .set_global_opts(
            title_opts=opts.TitleOpts(title="设计工程量vs 开累完成工程量"),
            tooltip_opts=opts.TooltipOpts(trigger="axis"),
            xaxis_opts=opts.AxisOpts(axislabel_opts={"rotate": 45}),
            datazoom_opts=[opts.DataZoomOpts(type_="slider")],
        )
    )
    return bar

uploaded_file = st.file_uploader("请上传月报 Word 文件（.docx）", type=["docx"])

if uploaded_file:
    st.success("✅ 上传成功，正在解析文档")
    df = extract_table_3_2(uploaded_file)

    if df is not None:
        st.subheader("📄 提取出进度相关数据")
        st.dataframe(df)

        st.subheader("📈 自动生成计划 vs 实际对比柱状图")
        chart1, names = plot_plan_vs_actual(df)
        st_pyecharts(chart1)

        st.subheader("📉 自动生成设计工程量 vs 开累完成工程量对比柱状图")
        chart2 = bar_plan_and_accumlative(df, names)
        st_pyecharts(chart2)

    else:
        st.warning("未找到表3.2，请检查文档格式是否一致。")

# if ("分部" in ''.join(headers)) and ("计划" in ''.join(headers)) and ("完成" in ''.join(headers)):

