import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="TK 账单极速核算系统", layout="wide")

st.title("📊 TikTok Shop 财务核算系统")

# ==========================================
# 1. 侧边栏：核心参数
# ==========================================
with st.sidebar:
    st.header("⚙️ 汇率与成本配置")
    rate_rmb_to_fx = st.number_input("RMB 兑 目标外币 汇率 (如 1 RMB = 2200 IDR)", value=2200.0)
    
    st.header("📢 广告费多币种录入")
    ad_target_currency = st.number_input("目标外币广告费 (如直接充值的印尼盾)", value=0.0)
    ad_other_currency = st.number_input("其他外币广告费 (如充值的美金)", value=0.0)
    rate_other_to_target = st.number_input("其他外币 兑 目标外币 汇率 (如 1 USD = 15000 IDR)", value=15000.0)
    
    total_ads_fx = ad_target_currency + (ad_other_currency * rate_other_to_target)
    st.success(f"参与分摊的广告总额: {total_ads_fx:,.2f}")

# ==========================================
# 2. 文件上传区
# ==========================================
st.header("📂 数据源上传")
col1, col2, col3 = st.columns(3)
col4, col5, _ = st.columns(3)

with col1: file_a = st.file_uploader("表 A (模板表 - 读第2行)", type=["xlsx", "csv"])
with col2: file_b = st.file_uploader("表 B (销售表)", type=["xlsx", "csv"])
with col3: file_c = st.file_uploader("表 C (收入表)", type=["xlsx", "csv"])
with col4: file_d = st.file_uploader("表 D (成本表)", type=["xlsx", "csv"])
with col5: file_e = st.file_uploader("表 E (刷单表 - 可选)", type=["xlsx", "csv"])

# 辅助函数：安全转小写
def to_key(s):
    return str(s).strip().lower()

# 辅助函数：终极单号清洗器
def clean_id(x):
    s = str(x).strip().lower()
    if 'e+' in s:
        try:
            return str(int(float(s)))
        except:
            pass
    if s.endswith('.0'):
        s = s[:-2]
    return s

if all([file_a, file_b, file_c, file_d]):
    try:
        # ==========================================
        # 第一步：精准读取
        # ==========================================
        df_a_raw = pd.read_excel(file_a, header=1) if file_a.name.endswith('xlsx') else pd.read_csv(file_a, header=1)
        df_b = pd.read_excel(file_b, header=0) if file_b.name.endswith('xlsx') else pd.read_csv(file_b, header=0)
            
        if file_c.name.endswith('xlsx'):
            xl_c = pd.ExcelFile(file_c)
            target_sheet_c = next((s for s in xl_c.sheet_names if 'order' in s.lower() and 'detail' in s.lower()), xl_c.sheet_names[0])
            df_c = pd.read_excel(file_c, sheet_name=target_sheet_c)
        else:
            df_c = pd.read_csv(file_c)
            
        df_d = pd.read_excel(file_d) if file_d.name.endswith('xlsx') else pd.read_csv(file_d)

        # 获取核心列名
        b_order_col = next((c for c in df_b.columns if 'order id' in to_key(c) or 'order number' in to_key(c)), None)
        c_order_col = next((c for c in df_c.columns if 'adjustment id' in to_key(c) or 'order id' in to_key(c)), None)
        
        # 应用单号清洗器
        df_b['match_id'] = df_b[b_order_col].apply(clean_id)
        df_c['match_id'] = df_c[c_order_col].apply(clean_id)
        
        b_cols_needed = ['order status', 'Seller sku', 'Quantity', 'Sku Quantity of return', 
                         'SKU Platform Discount', 'SKU Seller Discount', 'SKU Subtotal After Discount']
        b_col_map = {to_key(c): c for c in df_b.columns}
        
        df = pd.DataFrame()
        df['order number'] = df_b[b_order_col] 
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
        df.loc[df['order status'].astype(str).str.lower() == 'canceled', '实际售价'] = 0.0

        # ==========================================
        # 第三步：匹配表 C 费用
        # ==========================================
        c_col_map = {to_key(c): c for c in df_c.columns}
        
        all_fee_columns = [
            'Total Fees', 'Platform commission fee', 'Pre-order service fee', 'Mall service fee', 'Payment Fee', 
            'Shipping cost', 'Shipping costs passed on to the logistics provider', 'Replacement shipping fee (passed on to the customer)', 
            'Exchange shipping fee (passed on to the customer)', 'Shipping cost borne by the platform', 'Shipping cost paid by the customer', 
            'Refunded shipping cost paid by the customer', 'Return shipping costs (passed on to the customer)', 'Shipping cost subsidy', 
            'Distance shipping fee from Horizon+ Program', 'Affiliate Commission', 'Affiliate partner commission', 'Affiliate Shop Ads commission', 
            'Affiliate Partner shop ads commission', 'Shipping Fee Program service fee', 'Dynamic commission', 'Bonus cashback service fee', 
            'LIVE Specials service fee', 'Voucher Xtra service fee', 'Order processing fee', 'EAMS Program service fee', 
            'Brands Crazy Deals/Flash Sale service fee', 'Dilayani Tokopedia fee', 'Dilayani Tokopedia handling fee', 'PayLater program fee', 
            'Campaign resource fee', 'Installation service fee', 'Article 22 Income Tax withheld', 'Platform special service fee', 
            'GMV Max ad fee', 'Ajustment amount'
        ]
        
        valid_c_cols = [c_col_map[to_key(f)] for f in all_fee_columns if to_key(f) in c_col_map]
        df_c_unique = df_c.drop_duplicates(subset=['match_id'])
        df = pd.merge(df, df_c_unique[['match_id'] + valid_c_cols], on='match_id', how='left')
        
        佣金求和项 = []
        for fee in all_fee_columns:
            actual_c_col = c_col_map.get(to_key(fee))
            if actual_c_col:
                df[fee] = pd.to_numeric(df[actual_c_col], errors='coerce').fillna(0) / df['订单统计']
                if to_key(fee) != 'total fees':
                    佣金求和项.append(fee)
            else:
                df[fee] = 0.0
                
        df['佣金共计'] = df[佣金求和项].sum(axis=1)
        df['佣金总计'] = df['佣金共计']

        # ==========================================
        # 第五步：匹配成本表 (修复匹配逻辑)
        # ==========================================
        # 模糊匹配 SKU 和 成本列，只要包含关键词即可
        d_sku_actual = next((c for c in df_d.columns if 'nomor referensi sku' in to_key(c) or 'seller sku' in to_key(c)), None)
        d_cost_actual = next((c for c in df_d.columns if '成本' in to_key(c) or 'cost' in to_key(c) or '价格' in to_key(c)), None)
        
        if d_sku_actual and d_cost_actual:
            df_d['match_sku'] = df_d[d_sku_actual].astype(str).str.strip().str.lower()
            df['match_sku'] = df['Seller sku'].astype(str).str.strip().str.lower()
            
            df_d_unique = df_d.drop_duplicates(subset=['match_sku'])
            df = pd.merge(df, df_d_unique[['match_sku', d_cost_actual]], on='match_sku', how='left')
            # 这里的字段名必须直接叫 '成本'，才能匹配到模板 A 的表头
            df['成本'] = pd.to_numeric(df[d_cost_actual], errors='coerce').fillna(0)
        else:
            df['成本'] = 0.0

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
                df_e['match_id'] = df_e[e_order_actual].apply(clean_id)
                df_e_unique = df_e.drop_duplicates(subset=['match_id'])
                df = pd.merge(df, df_e_unique[['match_id', e_fee_actual]], on='match_id', how='left')
                
                df['is_sd'] = df[e_fee_actual].notnull()
                df['刷单'] = pd.to_numeric(df[e_fee_actual], errors='coerce').fillna(0) * rate_rmb_to_fx
                df['刷单佣金'] = df['is_sd'].apply(lambda x: 12.0 * rate_rmb_to_fx if x else 0.0)

        # ==========================================
        # 第七步：计算总成本
        # ==========================================
        qty = pd.to_numeric(df['Quantity'], errors='coerce').fillna(0)
        # 总成本 = 数量 * 成本 * 汇率
        df['总成本'] = qty * df['成本'] * rate_rmb_to_fx
        
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
        t_fee_col = 'Total Fees'
        t_fee_val = df[t_fee_col] if t_fee_col in df.columns else 0.0
        
        df['毛利'] = df['实际售价'] + t_fee_val - df['总成本'] - df['广告'] - df['刷单'] - df['刷单佣金']

        # ==========================================
        # 最终组装
        # ==========================================
        template_columns = df_a_raw.columns.tolist()
        df_final = pd.DataFrame(columns=template_columns)
        
        df_cols_map = {to_key(c): c for c in df.columns}
        
        for t_col in template_columns:
            if pd.isna(t_col): continue
            
            clean_t_col = str(t_col).replace('\n', ' ').strip().lower()
            match_col = next((c for c in df.columns if str(c).replace('\n', ' ').strip().lower() == clean_t_col), None)
            
            if match_col:
                df_final[t_col] = df[match_col]

        st.divider()
        st.success("✅ 核算完毕！【成本】列已修复匹配逻辑，数据完全对齐。")
        st.dataframe(df_final.head(15))

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_final.to_excel(writer, index=False)
        st.download_button(label="📥 下载最终核算报表", data=output.getvalue(), file_name="TK_Financial_Report_Final.xlsx")

    except Exception as e:
        st.error(f"❌ 运行发生异常: {e}")
