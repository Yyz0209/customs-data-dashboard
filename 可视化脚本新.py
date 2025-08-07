import os
import glob
import time
import pandas as pd
import streamlit as st
from playwright.sync_api import sync_playwright
from datetime import datetime
from pyecharts import options as opts
from pyecharts.charts import Line
from streamlit_echarts import st_pyecharts
import asyncio
import sys

# --- 关键修复：解决在Windows上Streamlit中运行Playwright的兼容性问题 ---
if sys.platform == 'win32':
    asyncio.set_event_loop_policy(asyncio.WindowsSelectorEventLoopPolicy())


# --- 配置区 ---
RAW_DATA_PATH = "raw_csv_data"
OUTPUT_FILENAME = "海关统计数据汇总.xlsx"
TARGET_LOCATIONS = ['北京市', '上海市', '深圳市', '南京市', '合肥市', '浙江省']
BASE_URL = "http://www.customs.gov.cn/customs/302249/zfxxgk/2799825/302274/302277/6348926/index.html"

# =============================================================================
#  第一部分：数据获取函数 (从之前的脚本整合而来)
# =============================================================================
def check_and_download_new_data():
    """检查本地已下载的文件，访问网站，只下载新增月份的原始数据。"""
    st.write("--- 开始执行数据更新检查 ---")
    os.makedirs(RAW_DATA_PATH, exist_ok=True)
    
    existing_files = set(os.listdir(RAW_DATA_PATH))
    st.write(f"本地已存在 {len(existing_files)} 个数据文件。")

    new_files_downloaded = 0
    progress_bar = st.progress(0)
    status_text = st.empty()

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True) # 改为无头模式，不在Streamlit中显示浏览器
        context = browser.new_context(user_agent='Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/99.0.4844.84 Safari/537.36')
        page = context.new_page()
        
        try:
            current_year = datetime.now().year
            for year in range(2024, current_year + 1):
                try:
                    status_text.text(f"正在检查年份: {year}...")
                    page.goto(BASE_URL, timeout=60000)
                    page.wait_for_selector("//div[@class='customs-foot']", timeout=30000)

                    year_button_selector = f"//a[contains(text(), '{year}')]"
                    page.wait_for_selector(year_button_selector, timeout=20000).click()
                    time.sleep(3)

                    table_row_selector = "//tr[contains(., '进出口商品收发货人所在地总值表')]"
                    row = page.wait_for_selector(table_row_selector, timeout=20000)

                    new_months_to_download = []
                    valid_month_links = [link for link in row.query_selector_all("a") if link.get_attribute('href')]
                    for link in valid_month_links:
                        month_text = link.inner_text()
                        if "月" in month_text:
                            month_number = int(month_text.replace("月", "").strip())
                            filename_to_check = f"{year}-{month_number:02d}.csv"
                            if filename_to_check not in existing_files:
                                new_months_to_download.append(month_text)
                    
                    if not new_months_to_download:
                        status_text.text(f"{year} 年未发现需要下载的新数据。")
                        continue
                    
                    total_new = len(new_months_to_download)
                    for i, month_text in enumerate(new_months_to_download):
                        status_text.text(f"发现新数据，正在下载: {month_text}...")
                        
                        month_link_selector = f"//tr[contains(., '进出口商品收发货人所在地总值表')]//a[text()='{month_text}']"
                        page.wait_for_selector(month_link_selector).click()
                        
                        with context.expect_page() as new_page_info:
                            detail_page = new_page_info.value
                        
                        detail_page.wait_for_load_state(timeout=60000)
                        
                        table_container_selector = "div.easysite-news-text"
                        detail_page.wait_for_selector(table_container_selector, timeout=20000)
                        table_html = detail_page.locator(table_container_selector).inner_html()
                        
                        dataframes = pd.read_html(table_html, header=[0, 1])
                        
                        if dataframes:
                            df = dataframes[0]
                            month_number = int(month_text.replace("月", "").strip())
                            filename_to_save = f"{year}-{month_number:02d}.csv"
                            file_path = os.path.join(RAW_DATA_PATH, filename_to_save)
                            df.to_csv(file_path, index=False, encoding='utf-8-sig')
                            status_text.text(f"新数据已保存至: {file_path}")
                            new_files_downloaded += 1
                        
                        detail_page.close()
                        time.sleep(2)
                        progress_bar.progress((i + 1) / total_new)

                except Exception as e:
                    st.error(f"处理年份 {year} 时出错: {e}")
                    continue
        finally:
            browser.close()
    
    status_text.text("数据更新检查完成！")
    progress_bar.empty()
    return new_files_downloaded

# =============================================================================
#  第二部分：数据处理函数 (从之前的脚本整合而来)
# =============================================================================
def process_all_data():
    """处理所有本地的原始数据，生成最终的Excel报告。"""
    st.write("\n--- 开始执行数据处理与整合 ---")
    all_csv_files = glob.glob(os.path.join(RAW_DATA_PATH, "*.csv"))
    if not all_csv_files:
        st.error(f"错误：在 '{RAW_DATA_PATH}' 文件夹中未找到任何CSV文件。")
        return

    st.write(f"找到 {len(all_csv_files)} 个原始数据文件进行处理...")
    
    all_processed_rows = []
    for file_path in all_csv_files:
        try:
            filename = os.path.basename(file_path)
            year, month = map(int, filename.replace('.csv', '').split('-'))
            data_time = f"{year}-{month:02d}"
            
            df = pd.read_csv(file_path, header=1)
            df.drop(index=0, inplace=True)
            df['时间'] = data_time
            
            required_cols_df = df.iloc[:, [0, 1, 3, 5, -1]].copy()
            required_cols_df.columns = ['地区', '进出口', '进口', '出口', '时间']
            
            filtered_df = required_cols_df[required_cols_df['地区'].isin(TARGET_LOCATIONS)]
            all_processed_rows.append(filtered_df)
        except Exception as e:
            st.warning(f"处理文件 {file_path} 时出错: {e}")
            continue

    if not all_processed_rows:
        st.error("未能处理任何数据，程序终止。")
        return

    master_df = pd.concat(all_processed_rows, ignore_index=True)

    for col in ['进出口', '进口', '出口']:
        master_df[col] = pd.to_numeric(master_df[col], errors='coerce')

    master_df['时间'] = pd.to_datetime(master_df['时间'])
    master_df.sort_values(by=['地区', '时间'], inplace=True)

    for col in ['进出口', '进口', '出口']:
        yoy_col_name = f'{col}同比'
        master_df[yoy_col_name] = master_df.groupby('地区')[col].pct_change(12)

    with pd.ExcelWriter(OUTPUT_FILENAME, engine='openpyxl') as writer:
        for location in TARGET_LOCATIONS:
            location_df = master_df[master_df['地区'] == location].copy()
            if not location_df.empty:
                location_df['时间'] = location_df['时间'].dt.strftime('%Y-%m')
                final_columns_order = ['时间', '进出口', '进出口同比', '进口', '进口同比', '出口', '出口同比']
                for col in final_columns_order:
                    if col not in location_df.columns:
                        location_df[col] = None
                location_df = location_df[final_columns_order]
                location_df.to_excel(writer, sheet_name=location, index=False)
    
    st.success(f"数据处理与整合完成！最终报告已更新: {OUTPUT_FILENAME}")

# =============================================================================
#  第三部分：Streamlit 应用主逻辑
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

# --- 核心改动：新的美化函数 ---
def format_delta_for_metric(yoy_value):
    """将小数值格式化为带正负号的百分比字符串，以便st.metric能正确上色和显示。"""
    if pd.isna(yoy_value):
        return None  # 如果没有同比数据，则不显示delta
    return f"{yoy_value:+.2%}"

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
                        delta=format_delta_for_metric(latest_data['进出口同比']),
                        delta_color="inverse" 
                    )
                    st.metric(
                        label=f"进口",
                        value=f"{latest_data['进口']:,}",
                        delta=format_delta_for_metric(latest_data['进口同比']),
                        delta_color="inverse"
                    )
                    st.metric(
                        label=f"出口",
                        value=f"{latest_data['出口']:,}",
                        delta=format_delta_for_metric(latest_data['出口同比']),
                        delta_color="inverse"
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
    st.caption("人民币值：万元") 
    
    location_df = data.get(selected_location)
    
    if location_df is not None and not location_df.empty:
        # 创建一个用于显示的副本，并格式化同比列
        display_df = location_df.copy()
        for col in ['进出口同比', '进口同比', '出口同比']:
            if col in display_df.columns:
                display_df[col] = display_df[col].apply(lambda x: f"{x:.2%}" if pd.notna(x) else 'N/A')

        # 显示表格时隐藏索引列
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
                yaxis_opts=opts.AxisOpts(name="人民币值：万元"), 
                legend_opts=opts.LegendOpts(orient="horizontal", pos_left="center")
            )
        )
        
        # 在Streamlit中渲染图表
        st_pyecharts(line_chart, height="500px")

    else:
        st.warning(f"未找到 '{selected_location}' 的数据。")
else:
    st.info("本地没有数据文件。请手动运行 '自动更新数据.py' 脚本来获取数据。")

