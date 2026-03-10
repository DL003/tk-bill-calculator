import streamlit as st
import pandas as pd
import io
import re

st.set_page_config(page_title="TK 账单助手-最终修复版", layout="wide")

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

def find_col(df, pattern):
    """使用正则表达式在列名中搜寻关键字"""
    for col in df.columns:
        if re.search(pattern, str(col), re.IGNORECASE):
            return col
    return None

if all([file_a, file_b, file_c, file_d]):
    try:
        # 读取数据
        df_template = pd.read_excel(file_a)
        df_b = pd.read_excel(file_b)
        xl_c = pd.ExcelFile(file_c)
        target_sheet = next((s for s in xl_c.sheet_names if 'order' in s.lower() and 'detail' in s.lower()), xl_c.sheet_names[0])
        df_c = pd.read_excel(file_c, sheet_name=target_sheet)
        df_d = pd.read_excel(file_d)

        # --- 识别核心列 (关键修复点) ---
        b_order_col = find_col(df_b, r'Order ID|Order Number')
        b_sku_col = find_col(df_b, r'Seller SKU')
        b_qty_col = find_col(df_b, r'Quantity')
        b_status_col = find_col(df_b, r'Status')
        
        # 售价相关：匹配包含特定超长描述的列
        s_sub = find_col(df_b, r'Subtotal Before Discount')
        s_plat = find_col(df_b, r'Platform Discount')
        s_seller = find_col(df_b, r'Seller Discount')
        
        c_order_col = find_col(df_c, r'Order ID|Order Number')
        d_sku_col = find_col(df_d, r'Seller SKU|SKU')
        d_cost_col = find_col(df_d, r'cost|成本|价格')

        # --- 建立费项映射 ---
        template_cols = df_template.columns.tolist()
        fee_mapping = {}
        for c_col in df_c.columns:
            # 清洗掉括号和单位，尝试与模板列名匹配
            clean_c = re.sub(r'\(.*?\)', '', str(c_col)).strip()
            for t_col in template_cols:
                if t_col.strip().lower() == clean_c.lower() or t_col.strip().lower() == str(c_col).lower():
                    fee_mapping[c_col] = t_col

        # --- 数据处理 ---
        # 1. 订单统计
        df_b['订单统计'] = df_b.groupby(b_order_col)[b_order_col].transform('count')
        
        # 2. 合并表 C
        df = pd.merge(df_b, df_c[[c_order_col] + list(fee_mapping.keys())], left_on=b_order_col, right_on=c_order_col, how='left')
        
        # 3. 费用分摊
        for c_col, t_col in fee_mapping.items():
            if t_col not in ['order number', 'Order ID']:
                df[t_col] = df[c_col].fillna(0) / df['订单统计']

        # 4. 匹配成本 (表 D)
        df = pd.merge(df, df_d[[d_sku_col, d_cost_col]], left_on=b_sku_col, right_on=d_sku_col, how='left')

        # 5. 计算实际售价 (核心公式)
        # It equals SKU Subtotal Before Discount - SKU Platform Discount - SKU Seller Discount + SKU Platform Discount
        v_sub = df[s_sub].fillna(0) if s_sub else 0
        v_plat = df[s_plat].fillna(0) if s_plat else 0
        v_sell = df[s_seller].fillna(0) if s_seller else 0
        df['实际售价'] = (v_sub - v_plat - v_sell) + v_plat

        # 6. 广告分摊
        valid_mask = df[b_status_col].astype(str).str.lower() != 'canceled'
        total_sales = df.loc[valid_mask, '实际售价'].sum()
        df['广告'] = 0.0
        if total_sales > 0:
            df.loc[valid_mask, '广告'] = (df.loc[valid_mask, '实际售价'] / total_sales) * ads_total_fx

        # 7. 成本与刷单 (RMB * 汇率 = 外币)
        df['总成本'] = df.apply(lambda x: (x[b_qty_col] * (x[d_cost_col] * rate_rmb_to_fx)) if pd.notnull(x[d_cost_col]) else 0, axis=1)
        
        # 8. 毛利计算
        c_total_fee_col = find_col(df_c, r'Total fees')
        df['毛利'] = 0.0
        if c_total_fee_col:
            # 实际售价 + 分摊后的Total fees - 总成本 - 广告
            df.loc[valid_mask, '毛利'] = df['实际售价'] + (df[c_total_fee_col].fillna(0)/df['订单统计']) - df['总成本'] - df['广告']

        # --- 导出 ---
        # 确保所有核心列都在列名列表中
        for c in ['实际售价', '总成本', '广告', '毛利', '订单统计']:
            if c not in df.columns: df[c] = 0.0
            
        df_final = df.reindex(columns=template_cols)
        
        # 强制写回计算好的数值到模板定义的列中
        core_cols = ['实际售价', '总成本', '广告', '毛利', '订单统计']
        for c in core_cols:
            if c in template_cols:
                df_final[c] = df[c]

        st.divider()
        st.success(f"✅ 处理完成！已识别工作表: {target_sheet}")
        st.dataframe(df_final.head(30))

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_final.to_excel(writer, index=False)
        st.download_button(label="📥 下载最终汇总表 A", data=output.getvalue(), file_name="Final_Report_A.xlsx")

    except Exception as e:
        st.error(f"❌ 运行出错: {e}")
