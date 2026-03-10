import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="TK 账单助手-精准 Sheet 版", layout="wide")

# --- 侧边栏：参数配置 ---
with st.sidebar:
    st.header("1. 参数配置")
    # 汇率逻辑：1 RMB = ? 外币
    rate_rmb_to_fx = st.number_input("汇率 (1 RMB = ? 外币)", value=2200.0, help="例如印尼站填2200，美区填0.14")
    ads_total_fx = st.number_input("总广告费 (外币)", value=0.0)
    st.divider()
    st.caption("注：程序会自动读取表C中的 'order detail' 工作表。")

# --- 主界面：文件上传 ---
st.header("2. 数据上传")
col_a, col_b, col_c = st.columns(3)
col_d, col_e, _ = st.columns(3)

with col_a: file_a = st.file_uploader("上传【表 A 模板】", type=["xlsx"])
with col_b: file_b = st.file_uploader("上传【表 B】(订单销售)", type=["xlsx"])
with col_c: file_c = st.file_uploader("上传【表 C】(订单收入/多Sheet版)", type=["xlsx"])
with col_d: file_d = st.file_uploader("上传【表 D】(成本表)", type=["xlsx"])
with col_e: file_e = st.file_uploader("上传【表 E】(刷单表-可选)", type=["xlsx"])

if all([file_a, file_b, file_c, file_d]):
    try:
        # 1. 读取数据
        df_template = pd.read_excel(file_a)
        df_b = pd.read_excel(file_b)
        
        # 核心：指定读取表 C 的 'order detail' 工作表
        try:
            df_c = pd.read_excel(file_c, sheet_name='order detail')
        except Exception:
            st.error("❌ 在表 C 中未找到名为 'order detail' 的工作表，请检查原始文件。")
            st.stop()
            
        df_d = pd.read_excel(file_d)
        
        # 2. 字段识别逻辑
        template_cols = df_template.columns.tolist()
        # 自动识别：模板中有且表C也有的费项（排除单号）
        fee_cols_to_process = [col for col in df_c.columns if col in template_cols and col != 'order number']
        
        # 3. 预处理与合并
        df_b['订单统计'] = df_b.groupby('order number')['order number'].transform('count')
        df = pd.merge(df_b, df_c[['order number'] + fee_cols_to_process], on='order number', how='left')
        
        # 费项平摊计算
        for col in fee_cols_to_process:
            df[col] = df[col].fillna(0) / df['订单统计']

        # 4. 关联成本与刷单
        df = pd.merge(df, df_d[['Seller sku', 'cost']], on='Seller sku', how='left')

        if file_e:
            df_e = pd.read_excel(file_e)
            df = pd.merge(df, df_e[['order number', 'shuadan_fee']], on='order number', how='left')
            df['is_sd'] = df['shuadan_fee'].notnull()
        else:
            df['is_sd'], df['shuadan_fee'] = False, 0

        # 5. 核心计算 (单位：外币 FX)
        # 实际售价
        df['实际售价'] = (df['It equals SKU Subtotal Before Discount - SKU Platform Discount - SKU Seller Discount.'] + 
                        df['Total platform discount in this SKU ID.'])
        
        # 广告分摊 (排除 Canceled)
        valid_mask = df['order status'] != 'Canceled'
        total_valid_sales = df.loc[valid_mask, '实际售价'].sum()
        df['广告'] = 0.0
        if total_valid_sales > 0:
            df.loc[valid_mask, '广告'] = (df.loc[valid_mask, '实际售价'] / total_valid_sales) * ads_total_fx

        # 成本与刷单：RMB * 汇率 = 外币
        df['总成本'] = df.apply(lambda x: (x['SKU sold quantity in the order.'] * (x['cost'] * rate_rmb_to_fx)) if not x['is_sd'] else 0, axis=1)
        df['刷单费用'] = df['shuadan_fee'].fillna(0) * rate_rmb_to_fx
        df['刷单佣金'] = df['is_sd'].apply(lambda x: (12 * rate_rmb_to_fx) if x else 0)

        # 6. 毛利计算 (实际销售 + Total fees - 总成本 - 广告 - 刷单 - 刷单佣金)
        def calc_profit(r):
            if r['order status'] == 'Canceled': return 0
            t_fee = r['Total fees'] if 'Total fees' in r else 0
            return r['实际售价'] + t_fee - r['总成本'] - r['广告'] - r['刷单费用'] - r['刷单佣金']

        df['毛利'] = df.apply(calc_profit, axis=1)

        # 7. 构建最终表 A
        df_final = df.reindex(columns=template_cols)
        # 强制填充计算好的核心列
        core_cols = ['实际售价', '总成本', '广告', '刷单费用', '刷单佣金', '毛利', '订单统计']
        for c in core_cols:
            if c in template_cols: df_final[c] = df[c]

        # --- 展示与导出 ---
        st.divider()
        st.success(f"✅ 计算完成！已从 'order detail' 读取并处理了 {len(fee_cols_to_process)} 项费用。")
        st.dataframe(df_final.head(30))

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_final.to_excel(writer, index=False)
        st.download_button(label="📥 下载最终汇总表 A", data=output.getvalue(), file_name="Final_A_Report.xlsx")

    except Exception as e:
        st.error(f"⚠️ 处理出错，请检查字段名。错误详情: {e}")
