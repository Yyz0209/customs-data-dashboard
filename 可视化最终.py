import os
import pandas as pd
import streamlit as st
from pyecharts import options as opts
from pyecharts.charts import Line
from streamlit_echarts import st_pyecharts
import sys
from datetime import datetime

# --- 配置区 ---
OUTPUT_FILENAME = "海关统计数据汇总.xlsx"
TARGET_LOCATIONS = ['全国', '北京市', '上海市', '深圳市', '南京市', '合肥市', '浙江省']
ZHEJIANG_CITIES = ['杭州市', '宁波市', '温州市', '湖州市', '金华市', '台州市']
ALL_LOCATIONS = TARGET_LOCATIONS + ZHEJIANG_CITIES

# =============================================================================
#  Streamlit 应用主逻辑
# =============================================================================
st.set_page_config(page_title="海关进出口数据看板", layout="wide", page_icon=None)

# --- 注入CSS以进行全面美化 (最终版) ---
st.markdown("""
<style>
    /* 全局字体和背景 */
    html, body, [class*="st-"] {
        font-family: 'Inter', 'Microsoft YaHei', sans-serif;
        background-color: #FFFFFF;
    }
    
    /* 侧边栏样式 */
    [data-testid="stSidebar"] {
        background-color: #F8F9FA;
        border-right: 1px solid #E5E7EB;
    }

    /* 侧边栏选择器卡片化 */
    [data-testid="stSelectbox"] {
        background-color: #FFFFFF;
        border: 1px solid #D1D5DB;
        border-radius: 8px;
        padding: 8px;
        box-shadow: 0 1px 2px rgba(0,0,0,0.04);
    }

    /* 主标题 */
    h1 {
        color: #1F2937;
        font-weight: 600; /* 降低字重 */
    }

    /* 副标题 */
    h2, h3 {
        color: #1F2937;
        font-weight: 600;
    }
    
    /* 指标标签 */
    div[data-testid="stMetricLabel"] {
        font-size: 15px;
        color: #6B7280;
    }
    
    /* 指标数值 (不加粗) */
    div[data-testid="stMetricValue"] {
        font-size: 28px;
        font-weight: 400;
        color: #111827;
    }

    /* 指标同比百分比 (新样式) */
    div[data-testid="stMetricDelta"] {
        padding-top: 8px;
    }

    /* 同比为正 (红色系) */
    div[data-testid="stMetricDelta"] div[style*="color: rgb(255, 75, 75)"] {
        display: inline-block;
        padding: 3px 8px;
        border-radius: 6px;
        background-color: #FEE2E2;
        color: #B91C1C !important;
    }

    /* 同比为负 (绿色系) */
    div[data-testid="stMetricDelta"] div[style*="color: rgb(34, 197, 94)"] {
        display: inline-block;
        padding: 3px 8px;
        border-radius: 6px;
        background-color: #D1FAE5;
        color: #065F46 !important;
    }

    /* 折叠面板 (Expander) */
    .stExpander {
        border: 1px solid #E5E7EB !important;
        box-shadow: none !important;
        background-color: #FFFFFF !important;
        border-radius: 8px !important;
        margin-top: 15px;
    }
    
    .stExpander header {
        font-size: 16px;
        font-weight: 600;
        color: #374151;
    }

</style>
""", unsafe_allow_html=True)


st.title("海关进出口数据看板")

# 使用缓存来加载数据
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

# --- 格式化函数 ---
def format_delta_for_metric(yoy_value):
    """将小数值格式化为带正负号的百分比字符串。"""
    if pd.isna(yoy_value):
        return None
    return f"{yoy_value:+.2%}"

def format_value(value):
    """将数值格式化为带千位分隔符的字符串。"""
    if pd.isna(value):
        return "N/A"
    return f"{value:,.0f}"

# --- 数据加载及预处理 ---
data = load_data()
latest_month_info = ""
if data:
    national_df = data.get('全国')
    if national_df is not None and not national_df.empty:
        latest_month = national_df['时间'].max()
        latest_month_info = f"数据更新至: {latest_month}"


# --- 侧边栏 ---
with st.sidebar:
    st.header("操作面板")
    st.markdown("---")
    if latest_month_info:
        st.caption(latest_month_info)
    
    st.header("地区筛选")
    selected_location = st.selectbox(
        "请选择要查看详情的地区：",
        options=ALL_LOCATIONS,
        index=ALL_LOCATIONS.index('浙江省') if '浙江省' in ALL_LOCATIONS else 0,
        label_visibility="collapsed" # 隐藏默认标签，因为我们在上面有自己的标题
    )

# --- 主页面 ---
if data:
    # --- 全国数据概览 ---
    st.subheader("全国数据概览 (万元)")
    if national_df is not None and not national_df.empty:
        latest_national_data = national_df.iloc[national_df['时间'].map(pd.to_datetime).idxmax()]
        # 使用Streamlit原生带边框的容器来创建卡片
        with st.container(border=True):
            cols = st.columns(3)
            cols[0].metric(
                label="进出口 (年初至今)", value=format_value(latest_national_data['进出口_年初至今']),
                delta=format_delta_for_metric(latest_national_data['进出口_年初至今同比']), delta_color="inverse"
            )
            cols[1].metric(
                label="进口 (年初至今)", value=format_value(latest_national_data['进口_年初至今']),
                delta=format_delta_for_metric(latest_national_data['进口_年初至今同比']), delta_color="inverse"
            )
            cols[2].metric(
                label="出口 (年初至今)", value=format_value(latest_national_data['出口_年初至今']),
                delta=format_delta_for_metric(latest_national_data['出口_年初至今同比']), delta_color="inverse"
            )

    # --- 业务地区数据概览 (每个地区一张卡片) ---
    st.subheader("业务地区数据概览 (万元)")
    locations_to_show = [loc for loc in TARGET_LOCATIONS if loc != '全国']
    
    for location in locations_to_show:
        df = data.get(location)
        if df is not None and not df.empty:
            latest_data = df.iloc[df['时间'].map(pd.to_datetime).idxmax()]
            
            # 每个地区使用一个独立的带边框容器
            with st.container(border=True):
                st.subheader(location)
                cols = st.columns(3)
                cols[0].metric(label="进出口 (年初至今)", value=format_value(latest_data['进出口_年初至今']), delta=format_delta_for_metric(latest_data['进出口_年初至今同比']), delta_color="inverse")
                cols[1].metric(label="进口 (年初至今)", value=format_value(latest_data['进口_年初至今']), delta=format_delta_for_metric(latest_data['进口_年初至今同比']), delta_color="inverse")
                cols[2].metric(label="出口 (年初至今)", value=format_value(latest_data['出口_年初至今']), delta=format_delta_for_metric(latest_data['出口_年初至今同比']), delta_color="inverse")
                
                if location == '浙江省':
                    with st.expander("展开/收起浙江省各地市数据"):
                        for city_index, city in enumerate(ZHEJIANG_CITIES):
                            city_df = data.get(city)
                            if city_df is not None and not city_df.empty:
                                latest_city_data = city_df.iloc[city_df['时间'].map(pd.to_datetime).idxmax()]
                                st.markdown(f"**{city}**")
                                city_cols = st.columns(3)
                                city_cols[0].metric(label="进出口", value=format_value(latest_city_data['进出口_年初至今']), delta=format_delta_for_metric(latest_city_data['进出口_年初至今同比']), delta_color="inverse")
                                city_cols[1].metric(label="进口", value=format_value(latest_city_data['进口_年初至今']), delta=format_delta_for_metric(latest_city_data['进口_年初至今同比']), delta_color="inverse")
                                city_cols[2].metric(label="出口", value=format_value(latest_city_data['出口_年初至今']), delta=format_delta_for_metric(latest_city_data['出口_年初至今同比']), delta_color="inverse")
                                if city_index < len(ZHEJIANG_CITIES) -1:
                                    st.markdown("---")
    
    # --- 数据详情与图表 (分离) ---
    st.header(f"{selected_location} - 数据详情")
    
    location_df = data.get(selected_location)
    
    if location_df is not None and not location_df.empty:
        # --- 图表部分 ---
        st.subheader(f"当月数据走势图")
        line_chart_month = (
            Line()
            .add_xaxis(xaxis_data=location_df['时间'].tolist())
            .add_yaxis(series_name="进出口(当月)", y_axis=location_df['进出口_当月'].tolist(), label_opts=opts.LabelOpts(is_show=False))
            .add_yaxis(series_name="进口(当月)", y_axis=location_df['进口_当月'].tolist(), label_opts=opts.LabelOpts(is_show=False))
            .add_yaxis(series_name="出口(当月)", y_axis=location_df['出口_当月'].tolist(), label_opts=opts.LabelOpts(is_show=False))
            .set_global_opts(
                title_opts=opts.TitleOpts(title=f"{selected_location} - 当月数据走势", pos_left='center', title_textstyle_opts=opts.TextStyleOpts(color="#111827")),
                tooltip_opts=opts.TooltipOpts(trigger="axis"),
                toolbox_opts=opts.ToolboxOpts(is_show=True),
                xaxis_opts=opts.AxisOpts(type_="category", boundary_gap=False),
                yaxis_opts=opts.AxisOpts(name="金额 (万元)"),
                legend_opts=opts.LegendOpts(orient="horizontal", pos_top="40")
            )
        )
        st_pyecharts(line_chart_month, height="500px")
        
        # --- 表格部分 ---
        st.subheader("详细数据表")
        st.caption("金额单位：万元")
        display_df = location_df.copy()
        for col in display_df.columns:
            if '同比' in col:
                display_df[col] = display_df[col].apply(lambda x: f"{x:.2%}" if pd.notna(x) else 'N/A')
        st.dataframe(display_df.sort_values(by="时间", ascending=False), use_container_width=True, hide_index=True)

    else:
        st.warning(f"未找到 '{selected_location}' 的数据。")
else:
    st.info("本地没有数据文件。请确保数据文件存在。")

