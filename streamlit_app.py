import streamlit as st
import pandas as pd
import io
import re

st.set_page_config(page_title="TK 账单助手-双行表头版", layout="wide")

# --- 侧边栏 ---
with st.sidebar:
    st.header("1. 参数配置")
    rate_rmb_to_fx = st.number_input("汇率 (1 RMB = ? 外币)", value=2200.0)
    ads_total_fx = st.number_input("总广告费 (外币)", value=0.0)
    st.info("💡 汇率逻辑：RMB 成本 * 汇率 = 外币总计")

# --- 文件上传 ---
st.header("2. 数据上传")
col_a, col_b, col_c = st.columns(3)
col_d, col_e, _ = st.columns(3)

with col_a: file_a = st.file_uploader("上传【表 A 模板】(双行表头)", type=["xlsx"])
with col_b: file_b = st.file_uploader("上传【表 B】(双行表头销售)", type=["xlsx"])
with col_c: file_c = st.file_uploader("上传【表 C】(订单收入)", type=["xlsx"])
with col_d: file_d = st.file_uploader("上传【表 D】(成本表)", type=["xlsx"])
with col_e: file_e = st.file_uploader("上传【表 E】(刷单表-可选)", type=["xlsx"])

def find_col_regex(columns, pattern):
    """在列名列表中进行正则搜索"""
    for col in columns:
        if re.search(pattern, str(col), re.IGNORECASE):
            return col
    return None

if all([file_a, file_b, file_c, file_d]):
    try:
        # 1. 读取数据
        # 表 A：双表头，第二行(index 1)是 Key
        df_a_raw = pd.read_excel(file_a, header=[0, 1])
        template_keys = df_a_raw.columns.get_level_values(1).tolist()
        
        # 表 B：双表头，第一行(index 0)是 Key
        df_b_raw = pd.read_excel(file_b, header=[0, 1])
        b_keys = df_b_raw.columns.get_level_values(0).tolist()
        # 展平 B 表，使用第一行作为主列名
        df_b = df_b_raw.copy()
        df_b.columns = b_keys

        # 表 C：单表头读取
        xl_c = pd.ExcelFile(file_c)
        target_sheet = next((s for s in xl_c.sheet_names if 'order' in s.lower() and 'detail' in s.lower()), xl_c.sheet_names[0])
        df_c = pd.read_excel(file_c, sheet_name=target_sheet)
        
        # 表 D：成本表
        df_d = pd.read_excel(file_d)

        # 2. 识别核心列
        b_order_col = find_col_regex(b_keys, r'Order ID|Order Number')
        b_sku_col = find_col_regex(b_keys, r'Seller SKU')
        b_qty_col = find_col_regex(b_keys, r'Quantity')
        b_status_col = find_col_regex(b_keys, r'Status')
        
        # 售价与折扣识别 (正则匹配长描述)
        s_sub = find_col_regex(b_keys, r'Subtotal Before Discount')
        s_plat = find_col_regex(b_keys, r'Platform Discount')
        s_seller = find_col_regex(b_keys, r'Seller Discount')
        
        c_order_col = find_col_regex(df_c.columns, r'Order ID|Order Number|Order/adjustment ID')
        d_sku_col = find_col_regex(df_d.columns, r'Seller SKU|SKU')
        d_cost_col = find_col_regex(df_d.columns, r'cost|成本|价格')

        # 3. 数据清洗与统一
        df_b[b_order_col] = df_b[b_order_col].astype(str).str.strip()
        df_c[c_order_col] = df_c[c_order_col].astype(str).str.strip()
        df_b[b_sku_col] = df_b[b_sku_col].astype(str).str.strip()
        df_d[d_sku_col] = df_d[d_sku_col].astype(str).str.strip()

        # 4. 费项映射 (表 C -> 模板 A)
        fee_mapping = {}
        for c_col in df_c.columns:
            # 去掉括号内容进行模糊匹配
            clean_c = re.sub(r'\(.*?\)', '', str(c_col)).strip().lower()
            for t_key in template_keys:
                if t_key.lower() == clean_c or t_key.lower() == str(c_col).lower():
                    fee_mapping[c_col] = t_key

        # 5. 计算逻辑
        df_b['订单统计'] = df_b.groupby(b_order_col)[b_order_col].transform('count')
        
        # 合并表 C
        df = pd.merge(df_b, df_c[[c_order_col] + list(fee_mapping.keys())], left_on=b_order_col, right_on=c_order_col, how='left')
        
        # 分摊费用
        for c_col, t_key in fee_mapping.items():
            if t_key not in ['order number', 'Order ID']:
                df[t_key] = df[c_col].fillna(0) / df['订单统计']

        # 匹配成本
        df = pd.merge(df, df_d[[d_sku_col, d_cost_col]], left_on=b_sku_col, right_on=d_sku_col, how='left')

        # 实际售价公式
        v_sub = df[s_sub].fillna(0) if s_sub else 0
        v_plat = df[s_plat].fillna(0) if s_plat else 0
        v_sell = df[s_seller].fillna(0) if s_seller else 0
        df['实际售价'] = (v_sub - v_plat - v_sell) + v_plat

        # 广告分摊
        valid_mask = df[b_status_col].astype(str).str.lower() != 'canceled'
        total_sales = df.loc[valid_mask, '实际售价'].sum()
        df['广告'] = 0.0
        if total_sales > 0:
            df.loc[valid_mask, '广告'] = (df.loc[valid_mask, '实际售价'] / total_sales) * ads_total_fx

        # 成本 (RMB * 汇率)
        df['总成本'] = df.apply(lambda x: (x[b_qty_col] * (x[d_cost_col] * rate_rmb_to_fx)) if pd.notnull(x[d_cost_col]) else 0, axis=1)
        
        # 毛利
        c_total_fee_col = find_col_regex(df_c.columns, r'Total fees')
        df['毛利'] = 0.0
        if c_total_fee_col:
            df.loc[valid_mask, '毛利'] = df['实际售价'] + (df[c_total_fee_col].fillna(0)/df['订单统计']) - df['总成本'] - df['广告']

        # 6. 生成最终输出 (保留表 A 的双行表头)
        # 创建结果数据框，列名对应模板第二行
        df_final_data = df.reindex(columns=template_keys)
        # 填充核心计算结果
        core_cols = {'实际售价': '实际售价', '总成本': '总成本', '广告': '广告', '毛利': '毛利', '订单统计': '订单统计', b_order_col: 'order number'}
        for k, v in core_cols.items():
            if v in template_keys: df_final_data[v] = df[k]

        # 重新应用双层表头
        df_final_data.columns = df_a_raw.columns

        st.divider()
        st.success(f"✅ 处理完成！已自动识别双行表头。")
        st.dataframe(df_final_data.head(20))

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_final_data.to_excel(writer)
        st.download_button(label="📥 下载填充好的汇总表 A", data=output.getvalue(), file_name="Final_Report_DoubleHeader.xlsx")

    except Exception as e:
        st.error(f"❌ 运行出错: {e}")
        st.info("提示：请确保表 A 第二行和表 B 第一行包含正确的英文字段名。")
