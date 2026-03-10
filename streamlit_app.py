import streamlit as st
import pandas as pd
import io
import re

st.set_page_config(page_title="TK 财务全能核算系统", layout="wide")

st.title("📊 TikTok Shop 财务核算与多维分析系统")

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
        # 第五步：匹配成本表
        # ==========================================
        d_sku_actual = next((c for c in df_d.columns if 'nomor referensi sku' in to_key(c) or 'seller sku' in to_key(c)), None)
        d_cost_actual = next((c for c in df_d.columns if '成本' in to_key(c) or 'cost' in to_key(c) or '价格' in to_key(c)), None)
        
        if d_sku_actual and d_cost_actual:
            df_d['match_sku'] = df_d[d_sku_actual].astype(str).str.strip().str.lower()
            df['match_sku'] = df['Seller sku'].astype(str).str.strip().str.lower()
            
            df_d_unique = df_d.drop_duplicates(subset=['match_sku'])
            df = pd.merge(df, df_d_unique[['match_sku', d_cost_actual]], on='match_sku', how='left')
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
        df['Quantity_num'] = pd.to_numeric(df['Quantity'], errors='coerce').fillna(0)
        df['总成本'] = df['Quantity_num'] * df['成本'] * rate_rmb_to_fx
        
        df.loc[df['is_sd'] == True, '总成本'] = 0.0
        df.loc[df['order status'].astype(str).str.lower() == 'canceled', '总成本'] = 0.0

        # ==========================================
        # 第八步 & 第九步：广告与毛利
        # ==========================================
        valid_sales_mask = df['order status'].astype(str).str.lower() != 'canceled'
        total_valid_sales = df.loc[valid_sales_mask, '实际售价'].sum()
        
        df['广告'] = 0.0
        if total_valid_sales > 0:
            df.loc[valid_sales_mask, '广告'] = (df.loc[valid_sales_mask, '实际售价'] / total_valid_sales) * total_ads_fx

        t_fee_col = 'Total Fees'
        t_fee_val = df[t_fee_col] if t_fee_col in df.columns else 0.0
        df['毛利'] = df['实际售价'] + t_fee_val - df['总成本'] - df['广告'] - df['刷单'] - df['刷单佣金']

        # ==========================================
        # 🟢 新增模块 1：构建 Sheet 1 (店铺汇总)
        # ==========================================
        total_qty = df['Quantity_num'].sum()
        total_sales = df['实际售价'].sum()
        total_ads = df['广告'].sum()
        total_cost = df['总成本'].sum()
        total_profit = df['毛利'].sum()

        def pct(val, base):
            return f"{(val/base)*100:.2f}%" if base > 0 else "0.00%"

        summary_data = [
            {'分析指标': '销售总额 (Actual Sales)', '金额/数值': total_sales, '占销售额百分比': '-'},
            {'分析指标': '销售总数量 (Qty)', '金额/数值': total_qty, '占销售额百分比': '-'},
            {'分析指标': '广告费用 (Ads)', '金额/数值': total_ads, '占销售额百分比': pct(total_ads, total_sales)},
            {'分析指标': '总成本 (Cost)', '金额/数值': total_cost, '占销售额百分比': pct(total_cost, total_sales)},
            {'分析指标': '毛利总额 (Gross Profit)', '金额/数值': total_profit, '占销售额百分比': pct(total_profit, total_sales)},
        ]
        
        # 将各佣金项自动追加进汇总表
        for fee in all_fee_columns:
            if fee in df.columns:
                fee_sum = df[fee].sum()
                if fee_sum != 0:
                    summary_data.append({
                        '分析指标': f"【费项】{fee}",
                        '金额/数值': fee_sum,
                        '占销售额百分比': pct(fee_sum, total_sales)
                    })
        
        df_summary = pd.DataFrame(summary_data)

        # ==========================================
        # 🟢 新增模块 2：构建 Sheet 2 (物流运费分析)
        # ==========================================
        shipping_cols = [
            'Payment Fee', 'Shipping cost', 'Shipping costs passed on to the logistics provider', 
            'Replacement shipping fee (passed on to the customer)', 'Exchange shipping fee (passed on to the customer)', 
            'Shipping cost borne by the platform', 'Shipping cost paid by the customer', 
            'Refunded shipping cost paid by the customer', 'Return shipping costs (passed on to the customer)', 
            'Shipping cost subsidy', 'Distance shipping fee from Horizon+ Program'
        ]
        valid_ship_cols = [c for c in shipping_cols if c in df.columns]
        
        df_shipping = df.groupby('Seller sku')[valid_ship_cols].sum().reset_index()
        df_shipping['物流相关费用总计'] = df_shipping[valid_ship_cols].sum(axis=1)
        
        # 按费用总计倒序排列，找出物流大头
        df_shipping = df_shipping.sort_values(by='物流相关费用总计', ascending=True)

        # ==========================================
        # 🟢 新增模块 3：构建 Sheet 3 (SKU 深度分析)
        # ==========================================
        df_sku = df.groupby('Seller sku').agg(
            销量=('Quantity_num', 'sum'),
            销售额=('实际售价', 'sum'),
            广告花费=('广告', 'sum'),
            总成本=('总成本', 'sum'),
            毛利=('毛利', 'sum')
        ).reset_index()

        df_sku['成本占比'] = df_sku.apply(lambda x: f"{(x['总成本']/x['销售额'])*100:.2f}%" if x['销售额']>0 else "0.00%", axis=1)
        df_sku['广告占比'] = df_sku.apply(lambda x: f"{(x['广告花费']/x['销售额'])*100:.2f}%" if x['销售额']>0 else "0.00%", axis=1)
        df_sku['毛利率'] = df_sku.apply(lambda x: f"{(x['毛利']/x['销售额'])*100:.2f}%" if x['销售额']>0 else "0.00%", axis=1)
        
        # 按照销售额从大到小排序
        df_sku = df_sku.sort_values(by='销售额', ascending=False)

        # ==========================================
        # 构建 Sheet 4 (原有的模板明细表)
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

        # ==========================================
        # 打包导出多 Sheet Excel
        # ==========================================
        st.divider()
        st.success("✅ 核算完毕！已成功生成包含四大分析维度的数据报表。")
        st.dataframe(df_summary)

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            # 写入四大工作表
            df_summary.to_excel(writer, sheet_name='店铺汇总', index=False)
            df_sku.to_excel(writer, sheet_name='SKU分析', index=False)
            df_shipping.to_excel(writer, sheet_name='物流费分析', index=False)
            df_final.to_excel(writer, sheet_name='账单明细(表A)', index=False)
            
        st.download_button(label="📥 下载多维数据报表 (包含 4 个 Sheet)", data=output.getvalue(), file_name="TK_Dashboard_Report.xlsx")

    except Exception as e:
        st.error(f"❌ 运行发生异常: {e}")
