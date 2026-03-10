import streamlit as st
import pandas as pd
import io
import re

st.set_page_config(page_title="TK 账单助手-全兼容版", layout="wide")

# --- 侧边栏：参数配置 ---
with st.sidebar:
    st.header("1. 参数配置")
    rate_rmb_to_fx = st.number_input("汇率 (1 RMB = ? 外币)", value=2200.0)
    ads_total_fx = st.number_input("总广告费 (外币)", value=0.0)
    st.divider()
    st.caption("提示：已加入表头模糊匹配，支持 IDR 等带币种符号的字段。")

# --- 主界面：文件上传 ---
st.header("2. 数据上传")
col_a, col_b, col_c = st.columns(3)
col_d, col_e, _ = st.columns(3)

with col_a: file_a = st.file_uploader("上传【表 A 模板】", type=["xlsx"])
with col_b: file_b = st.file_uploader("上传【表 B】(订单销售)", type=["xlsx"])
with col_c: file_c = st.file_uploader("上传【表 C】(订单收入)", type=["xlsx"])
with col_d: file_d = st.file_uploader("上传【表 D】(成本表)", type=["xlsx"])
with col_e: file_e = st.file_uploader("上传【表 E】(刷单表-可选)", type=["xlsx"])

def find_col(df, keywords):
    """根据关键字模糊匹配列名"""
    for col in df.columns:
        if any(k.lower() in str(col).lower() for k in keywords):
            return col
    return None

if all([file_a, file_b, file_c, file_d]):
    try:
        # 1. 读取数据并锁定 Sheet
        df_template = pd.read_excel(file_a)
        df_b = pd.read_excel(file_b)
        
        xl_c = pd.ExcelFile(file_c)
        target_sheet = next((s for s in xl_c.sheet_names if 'order' in s.lower() and 'detail' in s.lower()), xl_c.sheet_names[0])
        df_c = pd.read_excel(file_c, sheet_name=target_sheet)
        df_d = pd.read_excel(file_d)

        # 2. 核心字段模糊匹配 (处理 Order ID vs order number 等差异)
        # 表B
        b_order_col = find_col(df_b, ['order id', 'order number'])
        b_sku_col = find_col(df_b, ['seller sku', 'sku id'])
        b_qty_col = find_col(df_b, ['sold quantity', 'quantity'])
        b_status_col = find_col(df_b, ['order status', 'status'])
        
        # 表C
        c_order_col = find_col(df_c, ['order id', 'order number'])
        
        # 表D
        d_sku_col = find_col(df_d, ['seller sku', 'sku'])
        d_cost_col = find_col(df_d, ['cost', '价格', '成本'])

        # 3. 字段识别逻辑：只要模板中有，且账单(C)中包含该词（如 Total fees (IDR) 匹配 Total fees）
        template_cols = df_template.columns.tolist()
        fee_mapping = {} # {账单里的列: 模板里的列}
        for c_col in df_c.columns:
            for t_col in template_cols:
                # 如果模板列名包含在账单列名中，或者账单列名去掉括号后相等
                clean_c = re.sub(r'\(.*?\)', '', str(c_col)).strip()
                if t_col.lower() == clean_c.lower() or t_col.lower() == str(c_col).lower():
                    fee_mapping[c_col] = t_col

        # 4. 预处理：计算分摊基数
        df_b['订单统计'] = df_b.groupby(b_order_col)[b_order_col].transform('count')
        
        # 合并表 C 的匹配项
        cols_to_pull = [c_order_col] + list(fee_mapping.keys())
        df = pd.merge(df_b, df_c[list(set(cols_to_pull))], left_on=b_order_col, right_on=c_order_col, how='left')
        
        # 分摊费用
        for c_col, t_col in fee_mapping.items():
            if t_col != 'order number':
                df[t_col] = df[c_col].fillna(0) / df['订单统计']

        # 5. 匹配成本 (表 D)
        df = pd.merge(df, df_d[[d_sku_col, d_cost_col]], left_on=b_sku_col, right_on=d_sku_col, how='left')

        # 6. 计算售价（根据你提供的逻辑：子总计 - 折扣）
        # 匹配销售额相关列
        s_sub = find_col(df_b, ['Before Discount'])
        s_plat = find_col(df_b, ['Platform Discount'])
        s_seller = find_col(df_b, ['Seller Discount'])
        
        df['实际售价'] = df[s_sub].fillna(0) - df[s_plat].fillna(0) - df[s_seller].fillna(0) + df[s_plat].fillna(0)

        # 7. 广告分摊
        valid_mask = df[b_status_col].astype(str).str.lower() != 'canceled'
        total_sales = df.loc[valid_mask, '实际售价'].sum()
        df['广告'] = 0.0
        if total_sales > 0:
            df.loc[valid_mask, '广告'] = (df.loc[valid_mask, '实际售价'] / total_sales) * ads_total_fx

        # 8. 成本与刷单
        df['总成本'] = df.apply(lambda x: (x[b_qty_col] * (x[d_cost_col] * rate_rmb_to_fx)) if pd.notnull(x[d_cost_col]) else 0, axis=1)
        
        # 毛利计算
        # 寻找 Total fees 列（可能带币种）
        c_total_fee_col = find_col(df_c, ['Total fees'])
        df['毛利'] = 0.0
        if c_total_fee_col:
            df.loc[valid_mask, '毛利'] = df['实际售价'] + (df[c_total_fee_col].fillna(0)/df['订单统计']) - df['总成本'] - df['广告']

        # 9. 导出结果
        df_final = df.reindex(columns=template_cols)
        # 强制写回核心计算字段（如果模板里有这些名）
        mapping_back = {'实际售价': '实际售价', '总成本': '总成本', '广告': '广告', '毛利': '毛利', '订单统计': '订单统计'}
        for k, v in mapping_back.items():
            if v in template_cols: df_final[v] = df[k]

        st.divider()
        st.success("✅ 处理完成！已自动适配表头差异。")
        st.dataframe(df_final.head(30))

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_final.to_excel(writer, index=False)
        st.download_button(label="📥 下载最终汇总表 A", data=output.getvalue(), file_name="Report_A.xlsx")

    except Exception as e:
        st.error(f"⚠️ 计算出错，可能列名差异过大: {e}")
