import os
import pandas as pd
import streamlit as st
from pyecharts import options as opts
from pyecharts.charts import Line
from streamlit_echarts import st_pyecharts

# --- 配置区 ---
OUTPUT_FILENAME = "海关统计数据汇总.xlsx"
TARGET_LOCATIONS = ['北京市', '上海市', '深圳市', '南京市', '合肥市', '浙江省']

# =============================================================================
#  Streamlit 应用主逻辑
# =============================================================================
st.set_page_config(page_title="海关进出口数据看板", layout="wide")

# --- 注入CSS以美化卡片 ---
st.markdown("""
<style>
div[data-testid="metric-container"] {
    background-color: #FFFFFF;
    border: 1px solid #E1E1E1;
    padding: 20px;
    border-radius: 10px;
    box-shadow: 0 4px 12px rgba(0,0,0,0.08);
    margin: 10px 0;
}
div[data-testid="stMetricValue"] {
    font-size: 24px;
}
</style>
""", unsafe_allow_html=True)


st.title("海关进出口数据看板")

# 使用缓存来加载数据，避免每次交互都重新读取文件
@st.cache_data
def load_data():
    if not os.path.exists(OUTPUT_FILENAME):
        return None
    try:
        xls = pd.ExcelFile(OUTPUT_FILENAME)
        data_by_location = {sheet_name: xls.parse(sheet_name) for sheet_name in xls.sheet_names}
        return data_by_location
    except Exception as e:
        st.error(f"加载Excel文件失败: {e}")
        return None

# --- 美化函数 ---
def format_metric_delta(yoy_value):
    """格式化指标卡片的同比数据"""
    if pd.isna(yoy_value):
        return "N/A"
    
    arrow = "▲" if yoy_value >= 0 else "▼"
    formatted_yoy = f"{arrow} {yoy_value:.2%}"
    return formatted_yoy

# --- 侧边栏 ---
st.sidebar.header("操作面板")


# --- 主页面 ---
data = load_data()

if data:
    # --- 数据卡片概览 ---
    st.header("最新月份数据概览")
    
    # 将6个地区分成两行，每行3个
    row1_cols = st.columns(3)
    row2_cols = st.columns(3)
    
    for i, location in enumerate(TARGET_LOCATIONS):
        col = row1_cols[i] if i < 3 else row2_cols[i-3]
        
        with col:
            df = data.get(location)
            if df is not None and not df.empty:
                # 获取最新月份的数据
                latest_data = df.iloc[df['时间'].map(pd.to_datetime).idxmax()]
                
                # 使用 st.container 创建卡片效果
                with st.container(border=True):
                    st.subheader(location)
                    st.metric(
                        label=f"进出口",
                        value=f"{latest_data['进出口']:,}",
                        delta=format_metric_delta(latest_data['进出口同比']),
                        delta_color="inverse" # 正为红，负为绿
                    )
                    st.metric(
                        label=f"进口",
                        value=f"{latest_data['进口']:,}",
                        delta=format_metric_delta(latest_data['进口同比']),
                        delta_color="inverse" # 正为红，负为绿
                    )
                    st.metric(
                        label=f"出口",
                        value=f"{latest_data['出口']:,}",
                        delta=format_metric_delta(latest_data['出口同比']),
                        delta_color="inverse" # 正为红，负为绿
                    )
            else:
                 with col:
                    st.warning(f"无 {location} 数据")


    # --- 数据详情部分 ---
    st.sidebar.markdown("---")
    selected_location = st.sidebar.selectbox(
        "请选择要查看详情的地区：",
        options=TARGET_LOCATIONS
    )
    
    st.header(f"{selected_location} - 数据详情")
    st.caption("人民币值：亿元") # 在表格旁标注单位
    
    location_df = data.get(selected_location)
    
    if location_df is not None and not location_df.empty:
        # 创建一个用于显示的副本，并格式化同比列
        display_df = location_df.copy()
        for col in ['进出口同比', '进口同比', '出口同比']:
            if col in display_df.columns:
                display_df[col] = display_df[col].apply(lambda x: f"{x:.2%}" if pd.notna(x) else 'N/A')

        # 显示表格，使用容器宽度，并隐藏索引列
        st.dataframe(display_df.sort_values(by="时间", ascending=False), use_container_width=True, hide_index=True)
        
        # --- 使用 Pyecharts 绘制图表 ---
        st.header(f"{selected_location} - 进出口走势图")
        
        # 准备数据
        x_data = location_df['时间'].tolist()
        y_jinchukou = location_df['进出口'].tolist()
        y_jinkou = location_df['进口'].tolist()
        y_chukou = location_df['出口'].tolist()

        # 创建折线图
        line_chart = (
            Line()
            .add_xaxis(xaxis_data=x_data)
            .add_yaxis(
                series_name="进出口",
                y_axis=y_jinchukou,
                label_opts=opts.LabelOpts(is_show=False),
            )
            .add_yaxis(
                series_name="进口",
                y_axis=y_jinkou,
                label_opts=opts.LabelOpts(is_show=False),
            )
            .add_yaxis(
                series_name="出口",
                y_axis=y_chukou,
                label_opts=opts.LabelOpts(is_show=False),
            )
            .set_global_opts(
                title_opts=opts.TitleOpts(title=f"{selected_location} 2024年以来进出口走势"),
                tooltip_opts=opts.TooltipOpts(trigger="axis"),
                toolbox_opts=opts.ToolboxOpts(is_show=True),
                xaxis_opts=opts.AxisOpts(type_="category", boundary_gap=False),
                yaxis_opts=opts.AxisOpts(name="人民币值：亿元"), # 为Y轴添加单位
                legend_opts=opts.LegendOpts(orient="horizontal", pos_left="center")
            )
        )
        
        # 在Streamlit中渲染图表
        st_pyecharts(line_chart, height="500px")

    else:
        st.warning(f"未找到 '{selected_location}' 的数据。")
else:
    st.info("本地没有数据文件。请手动运行 '自动更新数据.py' 脚本来获取数据。")
