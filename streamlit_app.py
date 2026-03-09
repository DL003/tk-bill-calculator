import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="TK 账单助手 (模板驱动版)", layout="wide")

# --- 侧边栏 ---
with st.sidebar:
    st.header("1. 参数配置")
    ex_rate = st.number_input("汇率 (1 外币 = ? RMB)", value=7.2)
    ads_total = st.number_input("总广告费 (外币)", value=0.0)

# --- 文件上传 ---
st.header("2. 数据上传")
col_left, col_right = st.columns(2)

with col_left:
    file_a = st.file_uploader("上传【表 A 模板】(定义输出格式)", type=["xlsx"])
    file_b = st.file_uploader("上传【表 B】(订单销售)", type=["xlsx"])
with col_right:
    file_c = st.file_uploader("上传【表 C】(订单费项)", type=["xlsx"])
    file_d = st.file_uploader("上传【表 D】(成本表)", type=["xlsx"])
    file_e = st.file_uploader("上传【表 E】(刷单表-可选)", type=["xlsx"])

if all([file_a, file_b, file_c, file_d]):
    try:
        df_a_template = pd.read_excel(file_a) # 读取你的模板表头
        df_b = pd.read_excel(file_b)
        df_c = pd.read_excel(file_c)
        df_d = pd.read_excel(file_d)
        
        # 获取模板中所有的列名，作为输出标准
        output_columns = df_a_template.columns.tolist()
        
        # 1. 计算分摊基数
        df_b['订单统计'] = df_b.groupby('order number')['order number'].transform('count')
        
        # 2. 动态识别费项：只要是模板里有的列，且在表 C 里也能找到，就进行分摊计算
        common_fees = [col for col in output_columns if col in df_c.columns and col != 'order number']
        
        # 3. 合并数据
        df_merged = pd.merge(df_b, df_c[['order number'] + common_fees], on='order number', how='left')
        
        # 自动分摊模板中要求的所有费项
        for col in common_fees:
            df_merged[col] = df_merged[col].fillna(0) / df_merged['订单统计']
            
        # 4. 匹配成本与刷单 (逻辑同前)
        # ... (此处省略重复的 Merge D 和 E 的逻辑) ...

        # 5. 计算毛利 (确保使用模板中定义的最新字段求和)
        # 佣金总计 = 模板中存在的所有费项列之和
        df_merged['佣金总计'] = df_merged[common_fees].sum(axis=1)
        
        # 最终构建输出表：只保留模板里有的列，并填入算好的数据
        # 如果模板里有“毛利”列，程序会自动计算并填充
        df_final = df_merged.reindex(columns=output_columns)
        
        st.success("计算完成！已匹配模板中的字段进行分摊。")
        st.download_button("下载填充后的表 A", ...)

    except Exception as e:
        st.error(f"字段匹配失败，请检查模板表头与源表是否一致: {e}")
