import streamlit as st
import pandas as pd
import io

# 设置页面
st.set_page_config(page_title="TK 账单助手", layout="wide")

st.title("🚀 TikTok Shop 利润核算应用")

# --- 侧边栏：全局变量 ---
with st.sidebar:
    st.header("1. 参数配置")
    ex_rate = st.number_input("汇率 (1 外币 = ? RMB)", value=7.2, help="用于将RMB成本转为本币计算")
    ads_total = st.number_input("总广告费 (外币)", value=0.0)
    st.divider()
    st.caption("提示：成本、刷单佣金(12元)默认以RMB计算")

# --- 主界面：文件上传 ---
st.header("2. 数据上传")
col_b, col_c = st.columns(2)
col_d, col_e = st.columns(2)

with col_b: file_b = st.file_uploader("表 B (订单销售)", type=["xlsx"])
with col_c: file_c = st.file_uploader("表 C (订单费项)", type=["xlsx"])
with col_d: file_d = st.file_uploader("表 D (成本表)", type=["xlsx"])
with col_e: file_e = st.file_uploader("表 E (刷单表 - 可选)", type=["xlsx"])

# --- 核心计算流程 ---
if all([file_b, file_c, file_d]):
    try:
        # 读取数据
        df_b = pd.read_excel(file_b)
        df_c = pd.read_excel(file_c)
        df_d = pd.read_excel(file_d)
        
        # 1. 订单统计 (表B匹配用)
        df_b['订单统计'] = df_b.groupby('order number')['order number'].transform('count')
        
        # 2. 字段定义 (表C的所有费项)
        fee_cols = ['Total fees', 'TikTok Shop commission fee', 'Flat fee', 'Sales fee', 'Pre-Order Service Fee', 'Mall service fee', 'Payment fee', 'Shipping cost', 'Affiliate commission', 'Affiliate partner commission', 'Affiliate Shop Ads commission', 'Affiliate Partner shop ads commission', 'Shipping Fee Program service fee', 'Dynamic Commission', 'Bonus cashback service fee', 'LIVE Specials service fee', 'Voucher Xtra service fee', 'Order processing fee', 'EAMS Program service fee', 'Brands Crazy Deals/Flash Sale service fee', 'Dilayani Tokopedia fee', 'Dilayani Tokopedia handling fee', 'PayLater program fee', 'Campaign resource fee', 'Installation service fee', 'Ajustment amount']

        # 3. 数据合并与分摊
        df = pd.merge(df_b, df_c[['order number'] + fee_cols], on='order number', how='left')
        
        # 费项平摊：每一个费项都除以“订单统计”
        for col in fee_cols:
            df[col] = df[col].fillna(0) / df['订单统计']

        # 4. 匹配成本 (表D)
        df = pd.merge(df, df_d[['Seller sku', 'cost']], on='Seller sku', how='left')

        # 5. 匹配刷单 (表E)
        if file_e:
            df_e = pd.read_excel(file_e)
            df = pd.merge(df, df_e[['order number', 'shuadan_fee']], on='order number', how='left')
            df['is_sd'] = df['shuadan_fee'].notnull()
        else:
            df['is_sd'] = False
            df['shuadan_fee'] = 0

        # 6. 计算关键数值
        # 实际售价
        df['实际售价'] = (df['It equals SKU Subtotal Before Discount - SKU Platform Discount - SKU Seller Discount.'] + 
                        df['Total platform discount in this SKU ID.'])
        
        # 广告分摊 (排除 Canceled)
        valid_mask = df['order status'] != 'Canceled'
        total_valid_sales = df.loc[valid_mask, '实际售价'].sum()
        df['广告'] = 0.0
        if total_valid_sales > 0:
            df.loc[valid_mask, '广告'] = (df.loc[valid_mask, '实际售价'] / total_valid_sales) * ads_total

        # 成本计算 (单位转换)
        # 刷单逻辑：若匹配到刷单，总成本记为0，额外记录刷单费和12元佣金
        df['总成本'] = df.apply(lambda x: (x['SKU sold quantity in the order.'] * (x['cost'] / ex_rate)) if not x['is_sd'] else 0, axis=1)
        df['刷单费用'] = df['shuadan_fee'].fillna(0) / ex_rate
        df['刷单佣金'] = df['is_sd'].apply(lambda x: 12 / ex_rate if x else 0)

        # 7. 计算毛利 (核心公式)
        def get_profit(r):
            if r['order status'] == 'Canceled': return 0
            # 实际销售 + Total fees - 总成本 - 广告 - 刷单 - 刷单佣金
            return r['实际售价'] + r['Total fees'] - r['总成本'] - r['广告'] - r['刷单费用'] - r['刷单佣金']

        df['毛利'] = df.apply(get_profit, axis=1)

        # --- 结果展示 ---
        st.divider()
        st.subheader("💰 核算结果预览")
        m1, m2, m3 = st.columns(3)
        m1.metric("总有效销售额", f"{df.loc[valid_mask, '实际售价'].sum():.2f}")
        m2.metric("预估总毛利", f"{df['毛利'].sum():.2f}")
        m3.metric("订单行数", len(df))

        st.dataframe(df.head(50))

        # 导出 Excel
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name='Sheet1')
        
        st.download_button(
            label="📥 下载最终汇总表 A",
            data=output.getvalue(),
            file_name="Table_A_Calculated.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"发生错误：{e}")
