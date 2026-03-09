import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="TK 账单助手-模板驱动版", layout="wide")

# --- 侧边栏：参数配置 ---
with st.sidebar:
    st.header("1. 参数配置")
    # 汇率改为 1RMB = ? 外币
    rate_rmb_to_fx = st.number_input("汇率 (1 RMB = ? 外币)", value=2200.0, help="例如印尼盾填写2200，美金填写0.14")
    ads_total_fx = st.number_input("总广告费 (外币单位)", value=0.0)
    st.divider()
    st.caption("提示：成本和刷单佣金(12元)为RMB，将乘以该汇率转为外币。")

# --- 主界面：文件上传 ---
st.header("2. 数据上传")
col_a, col_b, col_c = st.columns(3)
col_d, col_e, _ = st.columns(3)

with col_a: file_a = st.file_uploader("上传【表 A 模板】(确定输出列)", type=["xlsx"])
with col_b: file_b = st.file_uploader("上传【表 B】(订单销售)", type=["xlsx"])
with col_c: file_c = st.file_uploader("上传【表 C】(订单费项)", type=["xlsx"])
with col_d: file_d = st.file_uploader("上传【表 D】(成本表)", type=["xlsx"])
with col_e: file_e = st.file_uploader("上传【表 E】(刷单表-可选)", type=["xlsx"])

if all([file_a, file_b, file_c, file_d]):
    try:
        # 读取数据
        df_template = pd.read_excel(file_a)
        df_b = pd.read_excel(file_b)
        df_c = pd.read_excel(file_c)
        df_d = pd.read_excel(file_d)
        
        # 1. 确定需要分摊的费项字段
        # 逻辑：只要是模板(A)里有的列，且在账单(C)里也存在的，都自动纳入计算
        template_cols = df_template.columns.tolist()
        fee_cols_to_process = [col for col in df_c.columns if col in template_cols and col != 'order number']
        
        # 2. 预处理表 B：计算分摊基数
        df_b['订单统计'] = df_b.groupby('order number')['order number'].transform('count')
        
        # 3. 合并表 C 费项
        df = pd.merge(df_b, df_c[['order number'] + fee_cols_to_process], on='order number', how='left')
        
        # 自动分摊所有匹配到的费项
        for col in fee_cols_to_process:
            df[col] = df[col].fillna(0) / df['订单统计']

        # 4. 匹配成本 (表 D)
        df = pd.merge(df, df_d[['Seller sku', 'cost']], on='Seller sku', how='left')

        # 5. 匹配刷单 (表 E)
        if file_e:
            df_e = pd.read_excel(file_e)
            df = pd.merge(df, df_e[['order number', 'shuadan_fee']], on='order number', how='left')
            df['is_sd'] = df['shuadan_fee'].notnull()
        else:
            df['is_sd'] = False
            df['shuadan_fee'] = 0

        # 6. 核心数值计算 (全部转为外币 FX)
        # 实际售价
        df['实际售价'] = (df['It equals SKU Subtotal Before Discount - SKU Platform Discount - SKU Seller Discount.'] + 
                        df['Total platform discount in this SKU ID.'])
        
        # 广告分摊 (排除 Canceled)
        valid_mask = df['order status'] != 'Canceled'
        total_valid_sales = df.loc[valid_mask, '实际售价'].sum()
        df['广告'] = 0.0
        if total_valid_sales > 0:
            df.loc[valid_mask, '广告'] = (df.loc[valid_mask, '实际售价'] / total_valid_sales) * ads_total_fx

        # 成本转换：RMB * 汇率 = 外币
        df['总成本'] = df.apply(lambda x: (x['SKU sold quantity in the order.'] * (x['cost'] * rate_rmb_to_fx)) if not x['is_sd'] else 0, axis=1)
        df['刷单费用'] = (df['shuadan_fee'].fillna(0) * rate_rmb_to_fx)
        df['刷单佣金'] = df['is_sd'].apply(lambda x: (12 * rate_rmb_to_fx) if x else 0)

        # 7. 计算毛利
        # 提取 Total fees (如果有的话)
        t_fees = df['Total fees'] if 'Total fees' in df.columns else 0
        
        def get_profit(r):
            if r['order status'] == 'Canceled': return 0
            # 公式：实际销售 + Total fees - 总成本 - 广告 - 刷单 - 刷单佣金
            return r['实际售价'] + (r['Total fees'] if 'Total fees' in r else 0) - r['总成本'] - r['广告'] - r['刷单费用'] - r['刷单佣金']

        df['毛利'] = df.apply(get_profit, axis=1)

        # 8. 按照模板列顺序导出
        # 如果模板中有计算结果中没有的列，填充空值；如果有计算好的列，自动填充
        df_final = df.reindex(columns=template_cols)
        
        # 自动把算好的“毛利”、“广告”、“总成本”等填回模板对应列（如果模板里有这些列名）
        for special_col in ['实际售价', '总成本', '广告', '刷单费用', '刷单佣金', '毛利', '订单统计']:
            if special_col in template_cols:
                df_final[special_col] = df[special_col]

        # --- 页面展示 ---
        st.divider()
        st.success("✅ 填充完成！已根据模板识别并计算了 {} 项费项。".format(len(fee_cols_to_process)))
        
        st.dataframe(df_final.head(50))

        # 导出
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_final.to_excel(writer, index=False)
        
        st.download_button(label="📥 下载填充好的表 A", data=output.getvalue(), file_name="Final_Report_A.xlsx")

    except Exception as e:
        st.error(f"发生错误：{e}")
