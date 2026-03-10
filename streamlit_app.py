import streamlit as st
import pandas as pd
import io
import re

st.set_page_config(page_title="TK 账单助手-最终稳定版", layout="wide")

# --- 侧边栏 ---
with st.sidebar:
    st.header("1. 参数配置")
    rate_rmb_to_fx = st.number_input("汇率 (1 RMB = ? 外币)", value=2200.0)
    ads_total_fx = st.number_input("总广告费 (外币)", value=0.0)

# --- 文件上传 ---
st.header("2. 数据上传")
col_a, col_b, col_c = st.columns(3)
col_d, col_e, _ = st.columns(3)

with col_a: file_a = st.file_uploader("上传【表 A 模板】", type=["xlsx"])
with col_b: file_b = st.file_uploader("上传【表 B】(订单销售)", type=["xlsx"])
with col_c: file_c = st.file_uploader("上传【表 C】(订单收入)", type=["xlsx"])
with col_d: file_d = st.file_uploader("上传【表 D】(成本表)", type=["xlsx"])
with col_e: file_e = st.file_uploader("上传【表 E】(刷单表-可选)", type=["xlsx"])

def find_col_safe(df, pattern, df_name):
    """安全搜寻列名，找不到则报错"""
    for col in df.columns:
        if re.search(pattern, str(col), re.IGNORECASE):
            return col
    st.error(f"❌ 在 {df_name} 中找不到匹配 '{pattern}' 的列。请检查表头。")
    st.stop()

if all([file_a, file_b, file_c, file_d]):
    try:
        # 读取数据
        df_template = pd.read_excel(file_a)
        df_b = pd.read_excel(file_b)
        xl_c = pd.ExcelFile(file_c)
        target_sheet = next((s for s in xl_c.sheet_names if 'order' in s.lower() and 'detail' in s.lower()), xl_c.sheet_names[0])
        df_c = pd.read_excel(file_c, sheet_name=target_sheet)
        df_d = pd.read_excel(file_d)

        # --- 强力识别核心列 ---
        b_order_col = find_col_safe(df_b, r'Order ID|Order Number', "表 B")
        b_sku_col = find_col_safe(df_b, r'Seller SKU', "表 B")
        b_qty_col = find_col_safe(df_b, r'Quantity', "表 B")
        b_status_col = find_col_safe(df_b, r'Status', "表 B")
        
        # 售价相关关键字识别
        s_sub = find_col_safe(df_b, r'Subtotal Before Discount', "表 B")
        s_plat = find_col_safe(df_b, r'Platform Discount', "表 B")
        s_seller = find_col_safe(df_b, r'Seller Discount', "表 B")
        
        c_order_col = find_col_safe(df_c, r'Order ID|Order Number', "表 C")
        d_sku_col = find_col_safe(df_d, r'Seller SKU|SKU', "表 D")
        d_cost_col = find_col_safe(df_d, r'cost|成本|价格', "表 D")

        # --- 统一数据格式 (防止 Merge 失败) ---
        df_b[b_order_col] = df_b[b_order_col].astype(str).str.strip()
        df_c[c_order_col] = df_c[c_order_col].astype(str).str.strip()
        df_b[b_sku_col] = df_b[b_sku_col].astype(str).str.strip()
        df_d[d_sku_col] = df_d[d_sku_col].astype(str).str.strip()

        # --- 建立费项映射 ---
        template_cols = df_template.columns.tolist()
        fee_mapping = {}
        for c_col in df_c.columns:
            clean_c = re.sub(r'\(.*?\)', '', str(c_col)).strip()
            for t_col in template_cols:
                if t_col.strip().lower() == clean_c.lower() or t_col.strip().lower() == str(c_col).lower():
                    fee_mapping[c_col] = t_col

        # --- 数据计算逻辑 ---
        # 1. 订单统计
        df_b['订单统计'] = df_b.groupby(b_order_col)[b_order_col].transform('count')
        
        # 2. 合并表 C
        df = pd.merge(df_b, df_c[[c_order_col] + list(fee_mapping.keys())], left_on=b_order_col, right_on=c_order_col, how='left')
        
        # 3. 费用分摊
        for c_col, t_col in fee_mapping.items():
            if t_col not in ['order number', 'Order ID']:
                df[t_col] = df[c_col].fillna(0) / df['订单统计']

        # 4. 匹配成本
        df = pd.merge(df, df_d[[d_sku_col, d_cost_col]], left_on=b_sku_col, right_on=d_sku_col, how='left')

        # 5. 实际售价
        v_sub = df[s_sub].fillna(0)
        v_plat = df[s_plat].fillna(0)
        v_sell = df[s_seller].fillna(0)
        df['实际售价'] = (v_sub - v_plat - v_sell) + v_plat

        # 6. 广告分摊
        valid_mask = df[b_status_col].astype(str).str.lower() != 'canceled'
        total_sales = df.loc[valid_mask, '实际售价'].sum()
        df['广告'] = 0.0
        if total_sales > 0:
            df.loc[valid_mask, '广告'] = (df.loc[valid_mask, '实际售价'] / total_sales) * ads_total_fx

        # 7. 成本 (RMB * 汇率 = 外币)
        df['总成本'] = df.apply(lambda x: (x[b_qty_col] * (x[d_cost_col] * rate_rmb_to_fx)) if pd.notnull(x[d_cost_col]) else 0, axis=1)
        
        # 8. 毛利
        c_total_fee_col = find_col_safe(df_c, r'Total fees', "表 C")
        df['毛利'] = 0.0
        df.loc[valid_mask, '毛利'] = df['实际售价'] + (df[c_total_fee_col].fillna(0)/df['订单统计']) - df['总成本'] - df['广告']

        # --- 输出结果 ---
        # 补齐可能缺失的计算列
        for c in ['实际售价', '总成本', '广告', '毛利', '订单统计']:
            if c not in df.columns: df[c] = 0.0
            
        df_final = df.reindex(columns=template_cols)
        core_cols = ['实际售价', '总成本', '广告', '毛利', '订单统计']
        for c in core_cols:
            if c in template_cols:
                df_final[c] = df[c]

        st.divider()
        st.success(f"✅ 处理成功！工作表：{target_sheet}")
        st.dataframe(df_final.head(30))

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_final.to_excel(writer, index=False)
        st.download_button(label="📥 下载最终汇总表 A", data=output.getvalue(), file_name="Final_Report.xlsx")

    except Exception as e:
        st.error(f"❌ 运行遇到未知错误: {str(e)}")
