import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="TK 账单极速核算系统", layout="wide")

st.title("📊 TikTok Shop 财务核算系统")

# ==========================================
# 1. 侧边栏：核心参数与多币种广告配置
# ==========================================
with st.sidebar:
    st.header("⚙️ 汇率与成本配置")
    rate_rmb_to_fx = st.number_input("RMB 兑 目标外币 汇率 (如 1 RMB = 2200 IDR)", value=2200.0)
    
    st.header("📢 广告费多币种录入")
    st.caption("支持混录不同币种，系统将自动统一为目标外币进行分摊。")
    ad_target_currency = st.number_input("目标外币广告费 (如直接充值的印尼盾)", value=0.0)
    ad_other_currency = st.number_input("其他外币广告费 (如充值的美金)", value=0.0)
    rate_other_to_target = st.number_input("其他外币 兑 目标外币 汇率 (如 1 USD = 15000 IDR)", value=15000.0)
    
    # 计算总广告费
    total_ads_fx = ad_target_currency + (ad_other_currency * rate_other_to_target)
    st.success(f"参与分摊的广告总额: {total_ads_fx:,.2f}")

# ==========================================
# 2. 文件上传区
# ==========================================
st.header("📂 数据源上传")
col1, col2, col3 = st.columns(3)
col4, col5, _ = st.columns(3)

with col1: file_a = st.file_uploader("表 A (模板表 - 读第2行)", type=["xlsx", "csv"])
with col2: file_b = st.file_uploader("表 B (销售表 - Order details)", type=["xlsx", "csv"])
with col3: file_c = st.file_uploader("表 C (收入表)", type=["xlsx", "csv"])
with col4: file_d = st.file_uploader("表 D (成本表)", type=["xlsx", "csv"])
with col5: file_e = st.file_uploader("表 E (刷单表 - 可选)", type=["xlsx", "csv"])

# 辅助函数：安全转小写匹配
def to_key(s):
    return str(s).strip().lower()

if all([file_a, file_b, file_c, file_d]):
    try:
        # ==========================================
        # 第一步：精准读取与基础字段映射
        # ==========================================
        # 表 A：真正的表头在第二排 (header=1)
        df_a_raw = pd.read_excel(file_a, header=1) if file_a.name.endswith('xlsx') else pd.read_csv(file_a, header=1)
        
        # 表 B：真正的表头在第一排 (header=0)，且指定读取 'Order details' sheet
        if file_b.name.endswith('xlsx'):
            df_b = pd.read_excel(file_b, header=0, sheet_name='Order details')
        else:
            df_b = pd.read_csv(file_b, header=0)
            
        df_c = pd.read_excel(file_c) if file_c.name.endswith('xlsx') else pd.read_csv(file_c)
        df_d = pd.read_excel(file_d) if file_d.name.endswith('xlsx') else pd.read_csv(file_d)

        # 统一核心 ID 格式（转字符串、去空格、转小写，实现不区分大小写匹配）
        b_order_col = next((c for c in df_b.columns if to_key(c) == 'order id'), None)
        c_order_col = next((c for c in df_c.columns if to_key(c) == 'order/adjustment id'), None)
        
        df_b['match_id'] = df_b[b_order_col].astype(str).str.strip().str.lower()
        df_c['match_id'] = df_c[c_order_col].astype(str).str.strip().str.lower()
        
        # 提取表 B 规定字段
        b_cols_needed = ['order status', 'Seller sku', 'Quantity', 'Sku Quantity of return', 
                         'SKU Platform Discount', 'SKU Seller Discount', 'SKU Subtotal After Discount']
        # 找到表 B 中实际对应的列名 (忽略大小写)
        b_col_map = {to_key(c): c for c in df_b.columns}
        
        # 构建基础操作表
        df = pd.DataFrame()
        df['order number'] = df_b[b_order_col] # 还原原始单号大小写
        df['match_id'] = df_b['match_id']
        
        for col in b_cols_needed:
            actual_col = b_col_map.get(to_key(col))
            df[col] = df_b[actual_col] if actual_col else 0

        # ==========================================
        # 第二步：计算“订单统计”与“实际售价”
        # ==========================================
        df['订单统计'] = df.groupby('match_id')['match_id'].transform('count')
        
        v_plat = pd.to_numeric(df['SKU Platform Discount'], errors='coerce').fillna(0)
        v_sub_after = pd.to_numeric(df['SKU Subtotal After Discount'], errors='coerce').fillna(0)
        
        df['实际售价'] = v_plat + v_sub_after
        # 状态判断 (忽略大小写)
        df.loc[df['order status'].astype(str).str.lower() == 'canceled', '实际售价'] = 0.0

        # ==========================================
        # 第三步 & 第四步：匹配表 C 费用并计算佣金总计
        # ==========================================
        c_col_map = {to_key(c): c for c in df_c.columns}
        fee_columns = [
            'Total fees', 'TikTok Shop commission fee', 'Flat fee', 'Sales fee', 'Pre-Order Service Fee', 
            'Mall service fee', 'Payment fee', 'Shipping cost', 'Affiliate commission', 'Affiliate partner commission', 
            'Affiliate Shop Ads commission', 'Affiliate Partner shop ads commission', 'Shipping Fee Program service fee', 
            'Dynamic Commission', 'Bonus cashback service fee', 'LIVE Specials service fee', 'Voucher Xtra service fee', 
            'Order processing fee', 'EAMS Program service fee', 'Brands Crazy Deals/Flash Sale service fee', 
            'Dilayani Tokopedia fee', 'Dilayani Tokopedia handling fee', 'PayLater program fee', 'Campaign resource fee', 
            'Installation service fee', 'Ajustment amount'
        ]
        
        # 匹配并将 C 表数据拉过来，除以订单统计进行分摊
        df_c_unique = df_c.drop_duplicates(subset=['match_id'])
        df = pd.merge(df, df_c_unique[['match_id'] + [c_col_map[to_key(f)] for f in fee_columns if to_key(f) in c_col_map]], on='match_id', how='left')
        
        佣金计算项 = []
        for fee in fee_columns:
            actual_c_col = c_col_map.get(to_key(fee))
            if actual_c_col:
                # 分摊费项
                df[fee] = pd.to_numeric(df[actual_c_col], errors='coerce').fillna(0) / df['订单统计']
                if fee != 'Total fees':
                    佣金计算项.append(fee)
            else:
                df[fee] = 0.0
                
        df['佣金总计'] = df[佣金计算项].sum(axis=1)

        # ==========================================
        # 第五步：匹配表 D 成本表 (RMB 转 外币)
        # ==========================================
        d_col_map = {to_key(c): c for c in df_d.columns}
        d_sku_actual = d_col_map.get('nomor referensi sku')
        d_cost_actual = d_col_map.get('cost') or d_col_map.get('成本') or d_col_map.get('价格')
        
        if d_sku_actual and d_cost_actual:
            # 统一 SKU 大小写进行匹配
            df_d['match_sku'] = df_d[d_sku_actual].astype(str).str.strip().str.lower()
            df['match_sku'] = df['Seller sku'].astype(str).str.strip().str.lower()
            
            df_d_unique = df_d.drop_duplicates(subset=['match_sku'])
            df = pd.merge(df, df_d_unique[['match_sku', d_cost_actual]], on='match_sku', how='left')
            df['RMB成本'] = pd.to_numeric(df[d_cost_actual], errors='coerce').fillna(0)
        else:
            df['RMB成本'] = 0.0

        # ==========================================
        # 第六步：匹配表 E 刷单表
        # ==========================================
        df['刷单'] = 0.0
        df['刷单佣金'] = 0.0
        df['is_sd'] = False
        
        if file_e:
            df_e = pd.read_excel(file_e) if file_e.name.endswith('xlsx') else pd.read_csv(file_e)
            e_col_map = {to_key(c): c for c in df_e.columns}
            e_order_actual = e_col_map.get('order id') or e_col_map.get('order number')
            e_fee_actual = next((c for c in df_e.columns if 'fee' in to_key(c) or '刷单' in to_key(c) or '费用' in to_key(c)), None)
            
            if e_order_actual and e_fee_actual:
                df_e['match_id'] = df_e[e_order_actual].astype(str).str.strip().str.lower()
                df_e_unique = df_e.drop_duplicates(subset=['match_id'])
                df = pd.merge(df, df_e_unique[['match_id', e_fee_actual]], on='match_id', how='left')
                
                df['is_sd'] = df[e_fee_actual].notnull()
                df['刷单'] = pd.to_numeric(df[e_fee_actual], errors='coerce').fillna(0) * rate_rmb_to_fx
                df['刷单佣金'] = df['is_sd'].apply(lambda x: 12.0 * rate_rmb_to_fx if x else 0.0)

        # ==========================================
        # 第七步：计算总成本
        # ==========================================
        qty = pd.to_numeric(df['Quantity'], errors='coerce').fillna(0)
        df['总成本'] = qty * df['RMB成本'] * rate_rmb_to_fx
        
        # 刷单或取消，总成本记为 0
        df.loc[df['is_sd'] == True, '总成本'] = 0.0
        df.loc[df['order status'].astype(str).str.lower() == 'canceled', '总成本'] = 0.0

        # ==========================================
        # 第八步：按比例分摊混合广告费
        # ==========================================
        valid_sales_mask = df['order status'].astype(str).str.lower() != 'canceled'
        total_valid_sales = df.loc[valid_sales_mask, '实际售价'].sum()
        
        df['广告'] = 0.0
        if total_valid_sales > 0:
            df.loc[valid_sales_mask, '广告'] = (df.loc[valid_sales_mask, '实际售价'] / total_valid_sales) * total_ads_fx

        # ==========================================
        # 第九步：计算最终毛利
        # ==========================================
        df['毛利'] = df['实际售价'] + df['Total fees'] - df['总成本'] - df['广告'] - df['刷单'] - df['刷单佣金']
        
        # 对于 Canceled 订单，毛利强制作 0 处理（如果你的业务规则需要）
        # df.loc[~valid_sales_mask, '毛利'] = 0.0 

        # ==========================================
        # 最终组装回表 A 模板格式
        # ==========================================
        # 读取原始表 A 的列名顺序
        template_columns = df_a_raw.columns.tolist()
        df_final = pd.DataFrame(columns=template_columns)
        
        # 将我们计算好的字典，按名字映射进去 (不区分大小写映射)
        df_cols_map = {to_key(c): c for c in df.columns}
        
        for t_col in template_columns:
            if pd.isna(t_col): continue
            match_col = df_cols_map.get(to_key(t_col))
            if match_col:
                df_final[t_col] = df[match_col]

        st.divider()
        st.success("✅ 全新核算系统运行成功！数据结构清爽无报错。")
        st.dataframe(df_final.head(15))

        # 导出
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_final.to_excel(writer, index=False)
        st.download_button(label="📥 一键下载最终财务报表", data=output.getvalue(), file_name="TK_Financial_Report.xlsx")

    except Exception as e:
        st.error(f"❌ 运行发生异常，请检查表格格式: {e}")
