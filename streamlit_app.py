import streamlit as st
import pandas as pd
import io
import re

st.set_page_config(page_title="TK 账单助手-单号对齐版", layout="wide")

# --- 侧边栏 ---
with st.sidebar:
    st.header("1. 参数配置")
    rate_rmb_to_fx = st.number_input("汇率 (1 RMB = ? 外币)", value=2200.0)
    ads_total_fx = st.number_input("总广告费 (外币)", value=0.0)

# --- 文件上传 ---
st.header("2. 数据上传")
col_a, col_b, col_c = st.columns(3)
col_d, col_e, _ = st.columns(3)

with col_a: file_a = st.file_uploader("上传【表 A 模板】(第二行为 Key)", type=["xlsx"])
with col_b: file_b = st.file_uploader("上传【表 B】(第一行为 Key)", type=["xlsx"])
with col_c: file_c = st.file_uploader("上传【表 C】(订单收入明细)", type=["xlsx"])
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
        # 1. 精准读取不同格式的表格
        # 表 A：双层表头，取第二层(Index 1)作为程序 Key
        df_a_raw = pd.read_excel(file_a, header=[0, 1])
        template_keys = df_a_raw.columns.get_level_values(1).tolist()
        
        # 表 B：双层表头，取第一层(Index 0)作为程序 Key
        df_b_raw = pd.read_excel(file_b, header=[0, 1])
        df_b = df_b_raw.copy()
        df_b.columns = df_b_raw.columns.get_level_values(0).tolist()

        # 表 C：单层表头识别
        df_c = pd.read_excel(file_c)
        df_d = pd.read_excel(file_d)

        # 2. 识别核心匹配列 (锁定 Order ID = order number)
        b_order_col = find_col_regex(df_b.columns, r'Order ID|Order Number')
        c_order_col = find_col_regex(df_c.columns, r'Order/adjustment ID|Order ID|Order Number')
        b_sku_col = find_col_regex(df_b.columns, r'Seller SKU')
        d_sku_col = find_col_regex(df_d.columns, r'Seller SKU|SKU')
        
        # 3. 统一单号格式为字符串，消除 ID 类型不匹配报错
        df_b[b_order_col] = df_b[b_order_col].astype(str).str.strip()
        df_c[c_order_col] = df_c[c_order_col].astype(str).str.strip()
        df_b[b_sku_col] = df_b[b_sku_col].astype(str).str.strip()
        df_d[d_sku_col] = df_d[d_sku_col].astype(str).str.strip()

        # 4. 字段匹配与分摊逻辑
        df_b['订单统计'] = df_b.groupby(b_order_col)[b_order_col].transform('count')
        
        # 将表 C 数据根据单号合并到表 B
        # 自动寻找模板(A)中需要的费项字段
        common_fees = [c for c in df_c.columns if any(str(c).lower().startswith(str(t).lower()) for t in template_keys) and c != c_order_col]
        df = pd.merge(df_b, df_c[[c_order_col] + common_fees], left_on=b_order_col, right_on=c_order_col, how='left')

        # 费用分摊：将合并过来的费项除以订单统计
        for c_col in common_fees:
            # 找到对应的模板列名
            t_col = next((t for t in template_keys if str(c_col).lower().startswith(str(t).lower())), None)
            if t_col and t_col not in ['order number', 'order status']:
                df[t_col] = df[c_col].fillna(0) / df['订单统计']

        # 5. 售价计算逻辑
        s_sub = find_col_regex(df_b.columns, r'Subtotal Before Discount')
        s_plat = find_col_regex(df_b.columns, r'Platform Discount')
        s_seller = find_col_regex(df_b.columns, r'Seller Discount')
        
        # 实际售价 = (小计 - 平台折 - 卖家折) + 平台折
        df['实际售价'] = (df[s_sub].fillna(0) - df[s_plat].fillna(0) - df[s_seller].fillna(0)) + df[s_plat].fillna(0)

        # 6. 成本匹配 (表 D)
        d_cost_col = find_col_regex(df_d.columns, r'cost|成本|价格')
        df = pd.merge(df, df_d[[d_sku_col, d_cost_col]], left_on=b_sku_col, right_on=d_sku_col, how='left')
        b_qty_col = find_col_regex(df_b.columns, r'Quantity')
        df['总成本'] = df[b_qty_col] * (df[d_cost_col].fillna(0) * rate_rmb_to_fx)

        # 7. 生成最终结果并映射回模板 A 的表头
        df_final = df.reindex(columns=template_keys)
        
        # 强制对齐订单号：把识别到的 B 表单号填入模板的 'order number' 列
        df_final['order number'] = df[b_order_col]
        df_final['订单统计'] = df['订单统计']
        # ... 其他计算字段同样对齐 ...
        if '实际售价' in template_keys: df_final['实际售价'] = df['实际售价']
        if '总成本' in template_keys: df_final['总成本'] = df['总成本']

        # 恢复双层表头导出
        df_final.columns = df_a_raw.columns

        st.divider()
        st.success("✅ 订单号对齐成功！正在生成报表...")
        st.dataframe(df_final.head(10))

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_final.to_excel(writer)
        st.download_button(label="📥 下载填充好的汇总表 A", data=output.getvalue(), file_name="Final_Report_Aligned.xlsx")

    except Exception as e:
        st.error(f"❌ 运行错误: {e}")
