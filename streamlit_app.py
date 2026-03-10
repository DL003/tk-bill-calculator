import streamlit as st
import pandas as pd
import io
import re

st.set_page_config(page_title="TK 账单助手-终极兼容版", layout="wide")

# --- 侧边栏 ---
with st.sidebar:
    st.header("1. 参数配置")
    rate_rmb_to_fx = st.number_input("汇率 (1 RMB = ? 外币)", value=2200.0)
    ads_total_fx = st.number_input("总广告费 (外币)", value=0.0)
    st.divider()
    st.caption("注：已增强对 'It equals SKU Subtotal...' 等长字段的识别。")

# --- 文件上传 ---
st.header("2. 数据上传")
col_a, col_b, col_c = st.columns(3)
col_d, col_e, _ = st.columns(3)

with col_a: file_a = st.file_uploader("上传【表 A 模板】", type=["xlsx"])
with col_b: file_b = st.file_uploader("上传【表 B】(订单销售)", type=["xlsx"])
with col_c: file_c = st.file_uploader("上传【表 C】(订单收入)", type=["xlsx"])
with col_d: file_d = st.file_uploader("上传【表 D】(成本表)", type=["xlsx"])
with col_e: file_e = st.file_uploader("上传【表 E】(刷单表-可选)", type=["xlsx"])

def find_col(df, keywords):
    """更强大的模糊匹配，支持正则表达式关键字"""
    for col in df.columns:
        col_str = str(col).lower()
        if any(re.search(k.lower(), col_str) for k in keywords):
            return col
    return None

if all([file_a, file_b, file_c, file_d]):
    try:
        df_template = pd.read_excel(file_a)
        df_b = pd.read_excel(file_b)
        
        # 读取表 C 的 'order details' Sheet
        xl_c = pd.ExcelFile(file_c)
        target_sheet = next((s for s in xl_c.sheet_names if 'order' in s.lower() and 'detail' in s.lower()), xl_c.sheet_names[0])
        df_c = pd.read_excel(file_c, sheet_name=target_sheet)
        df_d = pd.read_excel(file_d)

        # --- 自动识别表 B 核心列 ---
        b_order_col = find_col(df_b, [r'order id', r'order number'])
        b_sku_col = find_col(df_b, [r'seller sku', r'sku id'])
        b_qty_col = find_col(df_b, [r'sold quantity', r'qty'])
        b_status_col = find_col(df_b, [r'order status', r'status'])
        
        # 实际售价相关长字段识别
        # 原价/小计: 包含 "Subtotal Before Discount"
        s_sub = find_col(df_b, [r'subtotal before discount', r'original price'])
        # 平台折扣: 包含 "Platform Discount"
        s_plat = find_col(df_b, [r'platform discount'])
        # 卖家折扣: 包含 "Seller Discount"
        s_seller = find_col(df_b, [r'seller discount'])
        
        # --- 自动识别表 C/D 核心列 ---
        c_order_col = find_col(df_c, [r'order id', r'order number'])
        d_sku_col = find_col(df_d, [r'seller sku', r'sku'])
        d_cost_col = find_col(df_d, [r'cost', r'成本', r'价格'])

        # --- 建立费项映射 ---
        template_cols = df_template.columns.tolist()
        fee_mapping = {}
        for c_col in df_c.columns:
            clean_c = re.sub(r'\(.*?\)', '', str(c_col)).strip().lower()
            for t_col in template_cols:
                if t_col.lower() == clean_c or t_col.lower() == str(c_col).lower():
                    fee_mapping[c_col] = t_col

        # --- 数据预处理 ---
        df_b['订单统计'] = df_b.groupby(b_order_col)[b_order_col].transform('count')
        df = pd.merge(df_b, df_c[[c_order_col] + list(fee_mapping.keys())], left_on=b_order_col, right_on=c_order_col, how='left')

        # 分摊计算
        for c_col, t_col in fee_mapping.items():
            if t_col not in ['order number', 'order id']:
                df[t_col] = df[c_col].fillna(0) / df['订单统计']

        # 匹配成本
        df = pd.merge(df, df_d[[d_sku_col, d_cost_col]], left_on=b_sku_col, right_on=d_sku_col, how='left')

        # --- 关键计算：实际售价 ---
        # 逻辑：(小计 - 平台折 - 卖家折) + 平台折
        sub_val = df[s_sub].fillna(0) if s_sub else 0
        plat_val = df[s_plat].fillna(0) if s_plat else 0
        sell_val = df[s_seller].fillna(0) if s_seller else 0
        df['实际售价'] = (sub_val - plat_val - sell_val) + plat_val

        # 广告分摊
        valid_mask = df[b_status_col].astype(str).str.lower() != 'canceled'
        total_sales = df.loc[valid_mask, '实际售价'].sum()
        df['广告'] = 0.0
        if total_sales > 0:
            df.loc[valid_mask, '广告'] = (df.loc[valid_mask, '实际售价'] / total_sales) * ads_total_fx

        # 成本与刷单
        df['总成本'] = df.apply(lambda x: (x[b_qty_col] * (x[d_cost_col] * rate_rmb_to_fx)) if pd.notnull(x[d_cost_col]) else 0, axis=1)
        
        # 毛利计算
        c_total_fee_col = find_col(df_c, [r'total fees'])
        df['毛利'] = 0.0
        if c_total_fee_col:
            # 公式：实际售价 + 分摊后的Total fees - 总成本 - 广告
            df.loc[valid_mask, '毛利'] = df['实际售价'] + (df[c_total_fee_col].fillna(0)/df['订单统计']) - df['总成本'] - df['广告']

        # --- 格式化导出 ---
        df_final = df.reindex(columns=template_cols)
        core_mapping = {'实际售价': '实际售价', '总成本': '总成本', '广告': '广告', '毛利': '毛利', '订单统计': '订单统计'}
        for k, v in core_mapping.items():
            if v in template_cols: df_final[v] = df[k]

        st.divider()
        st.success(f"✅ 处理完成！自动匹配了 {len(fee_mapping)} 个费项字段。")
        st.dataframe(df_final.head(30))

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_final.to_excel(writer, index=False)
        st.download_button(label="📥 下载填充好的汇总表 A", data=output.getvalue(), file_name="Report_A_Final.xlsx")

    except Exception as e:
        st.error(f"❌ 运行错误: {e}")
        st.info("建议检查：表B中是否包含 'Subtotal Before Discount' 和 'Platform Discount' 等关键字的列。")
