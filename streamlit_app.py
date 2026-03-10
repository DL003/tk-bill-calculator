import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="TK 账单助手-Sheet兼容版", layout="wide")

# --- 侧边栏：参数配置 ---
with st.sidebar:
    st.header("1. 参数配置")
    # 汇率逻辑：1 RMB = ? 外币
    rate_rmb_to_fx = st.number_input("汇率 (1 RMB = ? 外币)", value=2200.0)
    ads_total_fx = st.number_input("总广告费 (外币)", value=0.0)
    st.divider()
    st.caption("提示：程序会自动匹配 'order details' 或 'order detail' 工作表。")

# --- 主界面：文件上传 ---
st.header("2. 数据上传")
col_a, col_b, col_c = st.columns(3)
col_d, col_e, _ = st.columns(3)

with col_a: file_a = st.file_uploader("上传【表 A 模板】", type=["xlsx"])
with col_b: file_b = st.file_uploader("上传【表 B】(订单销售数据)", type=["xlsx"])
with col_c: file_c = st.file_uploader("上传【表 C】(订单收入/多Sheet版)", type=["xlsx"])
with col_d: file_d = st.file_uploader("上传【表 D】(成本表)", type=["xlsx"])
with col_e: file_e = st.file_uploader("上传【表 E】(刷单表-可选)", type=["xlsx"])

if all([file_a, file_b, file_c, file_d]):
    try:
        # 读取模板和基础数据
        df_template = pd.read_excel(file_a)
        df_b = pd.read_excel(file_b)
        df_d = pd.read_excel(file_d)
        
        # --- 智能读取表 C ---
        xl_c = pd.ExcelFile(file_c)
        sheet_names = xl_c.sheet_names
        
        # 模糊匹配：寻找包含 'order' 和 'detail' 的 Sheet (不区分大小写，自动处理 details/detail)
        target_sheet = next((s for s in sheet_names if 'order' in s.lower() and 'detail' in s.lower()), None)
        
        if target_sheet:
            df_c = pd.read_excel(file_c, sheet_name=target_sheet)
            st.info(f"💡 已成功识别工作表: **{target_sheet}**")
        else:
            st.error(f"❌ 未能找到工作表。该文件包含的 Sheet 有: {', '.join(sheet_names)}")
            st.stop()
            
        # --- 字段识别与计算逻辑 ---
        template_cols = df_template.columns.tolist()
        # 自动识别：模板中有且表C也有的费项
        fee_cols_to_process = [col for col in df_c.columns if col in template_cols and col != 'order number']
        
        # 预处理：计算分摊基数 (订单统计)
        df_b['订单统计'] = df_b.groupby('order number')['order number'].transform('count')
        
        # 合并表 C 费项
        df = pd.merge(df_b, df_c[['order number'] + fee_cols_to_process], on='order number', how='left')
        
        # 自动分摊所有匹配到的费项
        for col in fee_cols_to_process:
            df[col] = df[col].fillna(0) / df['订单统计']

        # 关联成本 (表 D)
        df = pd.merge(df, df_d[['Seller sku', 'cost']], on='Seller sku', how='left')

        # 关联刷单 (表 E)
        if file_e:
            df_e = pd.read_excel(file_e)
            df = pd.merge(df, df_e[['order number', 'shuadan_fee']], on='order number', how='left')
            df['is_sd'] = df['shuadan_fee'].notnull()
        else:
            df['is_sd'], df['shuadan_fee'] = False, 0

        # --- 核心计算 (外币 FX) ---
        # 实际售价
        df['实际售价'] = (df['It equals SKU Subtotal Before Discount - SKU Platform Discount - SKU Seller Discount.'] + 
                        df['Total platform discount in this SKU ID.'])
        
        # 广告分摊 (排除 Canceled)
        valid_mask = df['order status'] != 'Canceled'
        total_valid_sales = df.loc[valid_mask, '实际售价'].sum()
        df['广告'] = 0.0
        if total_valid_sales > 0:
            df.loc[valid_mask, '广告'] = (df.loc[valid_mask, '实际售价'] / total_valid_sales) * ads_total_fx

        # 成本与刷单 (RMB * 汇率 = 外币)
        df['总成本'] = df.apply(lambda x: (x['SKU sold quantity in the order.'] * (x['cost'] * rate_rmb_to_fx)) if not x['is_sd'] else 0, axis=1)
        df['刷单费用'] = df['shuadan_fee'].fillna(0) * rate_rmb_to_fx
        df['刷单佣金'] = df['is_sd'].apply(lambda x: (12 * rate_rmb_to_fx) if x else 0)

        # 毛利公式
        def calc_profit(r):
            if r['order status'] == 'Canceled': return 0
            t_fee = r['Total fees'] if 'Total fees' in r else 0
            return r['实际售价'] + t_fee - r['总成本'] - r['广告'] - r['刷单费用'] - r['刷单佣金']

        df['毛利'] = df.apply(calc_profit, axis=1)

        # --- 构建输出结果 ---
        # 按照模板列顺序排序
        df_final = df.reindex(columns=template_cols)
        # 强制更新计算出来的核心业务字段
        core_cols = ['实际售价', '总成本', '广告', '刷单费用', '刷单佣金', '毛利', '订单统计']
        for c in core_cols:
            if c in template_cols:
                df_final[c] = df[c]

        st.divider()
        st.success(f"✅ 处理完成！已识别并计算了 {len(fee_cols_to_process)} 项费项。")
        st.dataframe(df_final.head(30))

        # 导出
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_final.to_excel(writer, index=False)
        st.download_button(label="📥 下载填充好的表 A 汇总", data=output.getvalue(), file_name="Table_A_Final.xlsx")

    except Exception as e:
        st.error(f"⚠️ 遇到错误，请检查上传表格的字段名。报错信息: {e}")
