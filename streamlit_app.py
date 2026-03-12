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

# 辅助函数：正则匹配列名 (免疫换行和空格)
def find_col_regex(df_cols, pattern):
    for c in df_cols:
        if pd.notna(c) and re.search(pattern, str(c).replace('\n', ' '), re.IGNORECASE):
            return c
    return None

# 辅助函数：终极单号清洗器
def clean_id(x):
    if pd.isna(x): return ""
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
        # 第一步：精准读取与清洗
        # ==========================================
        # 【关键修复】保留双层表头，保证导出结果格式不被破坏
        df_a_raw = pd.read_excel(file_a, header=[0, 1]) if file_a.name.endswith('xlsx') else pd.read_csv(file_a, header=[0, 1])
        df_b = pd.read_excel(file_b, header=0) if file_b.name.endswith('xlsx') else pd.read_csv(file_b, header=0)
            
        if file_c.name.endswith('xlsx'):
            xl_c = pd.ExcelFile(file_c)
            target_sheet_c = next((s for s in xl_c.sheet_names if 'order' in s.lower() and 'detail' in s.lower()), xl_c.sheet_names[0])
            df_c = pd.read_excel(file_c, sheet_name=target_sheet_c)
        else:
            df_c = pd.read_csv(file_c)
            
        df_d = pd.read_excel(file_d) if file_d.name.endswith('xlsx') else pd.read_csv(file_d)

        b_order_col = find_col_regex(df_b.columns, r'order id|order number')
        c_order_col = find_col_regex(df_c.columns, r'adjustment id|order id|order number')
        
        # 删除 TikTok 导出的第一行解释文本 (如 "Current order status.")
        if b_order_col:
            df_b = df_b[~df_b[b_order_col].astype(str).str.contains(r'Platform unique|Transaction|Order', flags=re.IGNORECASE, na=False)].reset_index(drop=True)
        if c_order_col:
            df_c = df_c[~df_c[c_order_col].astype(str).str.contains(r'Platform unique|Transaction|Order', flags=re.IGNORECASE, na=False)].reset_index(drop=True)

        df_b['match_id'] = df_b[b_order_col].apply(clean_id) if b_order_col else ""
        df_c['match_id'] = df_c[c_order_col].apply(clean_id) if c_order_col else ""
        
        # 锁定表 B 的各种字段
        b_status = find_col_regex(df_b.columns, r'order status|status')
        b_sku = find_col_regex(df_b.columns, r'seller sku|sku id')
        b_qty = find_col_regex(df_b.columns, r'^quantity$|sold quantity|qty')
        b_return = find_col_regex(df_b.columns, r'quantity of return|return quantity|return qty')
        b_plat_disc = find_col_regex(df_b.columns, r'platform discount')
        b_sell_disc = find_col_regex(df_b.columns, r'seller discount')
        b_sub_after = find_col_regex(df_b.columns, r'subtotal after discount')

        # ==========================================
        # 第二步：计算“订单统计”与“实际售价”
        # ==========================================
        df_b['订单统计'] = df_b.groupby('match_id')['match_id'].transform('count')
        
        v_plat = pd.to_numeric(df_b[b_plat_disc], errors='coerce').fillna(0) if b_plat_disc else 0
        v_sub_after = pd.to_numeric(df_b[b_sub_after], errors='coerce').fillna(0) if b_sub_after else 0
        
        df_b['实际售价'] = v_plat + v_sub_after
        if b_status:
            df_b.loc[df_b[b_status].astype(str).str.lower() == 'canceled', '实际售价'] = 0.0

        # ==========================================
        # 第三步：匹配表 C 费用
        # ==========================================
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
        
        fee_cols_in_c = {}
        for fee in all_fee_columns:
            c_col = find_col_regex(df_c.columns, r'(?i)' + re.escape(fee))
            if c_col: fee_cols_in_c[fee] = c_col

        df_c_unique = df_c.drop_duplicates(subset=['match_id'])
        df_b = pd.merge(df_b, df_c_unique[['match_id'] + list(fee_cols_in_c.values())], on='match_id', how='left')
        
        佣金求和项 = []
        for fee in all_fee_columns:
            if fee in fee_cols_in_c:
                c_col = fee_cols_in_c[fee]
                df_b[fee] = pd.to_numeric(df_b[c_col], errors='coerce').fillna(0) / df_b['订单统计']
                if fee.lower() != 'total fees': 佣金求和项.append(fee)
            else:
                df_b[fee] = 0.0
                
        df_b['佣金共计'] = df_b[佣金求和项].sum(axis=1)

        # ==========================================
        # 第五步：匹配成本表
        # ==========================================
        d_sku_actual = find_col_regex(df_d.columns, r'nomor referensi sku|seller sku')
        d_cost_actual = find_col_regex(df_d.columns, r'成本|cost|价格')
        
        if d_sku_actual and d_cost_actual and b_sku:
            df_d['match_sku'] = df_d[d_sku_actual].astype(str).str.strip().str.lower()
            df_b['match_sku'] = df_b[b_sku].astype(str).str.strip().str.lower()
            df_d_unique = df_d.drop_duplicates(subset=['match_sku'])
            df_b = pd.merge(df_b, df_d_unique[['match_sku', d_cost_actual]], on='match_sku', how='left')
            df_b['成本'] = pd.to_numeric(df_b[d_cost_actual], errors='coerce').fillna(0)
        else:
            df_b['成本'] = 0.0

        # ==========================================
        # 第六步：匹配表 E 刷单表
        # ==========================================
        df_b['刷单'] = 0.0
        df_b['刷单佣金'] = 0.0
        df_b['is_sd'] = False
        if file_e:
            df_e = pd.read_excel(file_e) if file_e.name.endswith('xlsx') else pd.read_csv(file_e)
            e_order_actual = find_col_regex(df_e.columns, r'order id|order number|单号')
            e_fee_actual = find_col_regex(df_e.columns, r'fee|刷单|费用')
            if e_order_actual and e_fee_actual:
                df_e['match_id'] = df_e[e_order_actual].apply(clean_id)
                df_e_unique = df_e.drop_duplicates(subset=['match_id'])
                df_b = pd.merge(df_b, df_e_unique[['match_id', e_fee_actual]], on='match_id', how='left')
                df_b['is_sd'] = df_b[e_fee_actual].notnull()
                df_b['刷单'] = pd.to_numeric(df_b[e_fee_actual], errors='coerce').fillna(0) * rate_rmb_to_fx
                df_b['刷单佣金'] = df_b['is_sd'].apply(lambda x: 12.0 * rate_rmb_to_fx if x else 0.0)

        # ==========================================
        # 第七步：计算总成本
        # ==========================================
        qty_series = pd.to_numeric(df_b[b_qty], errors='coerce').fillna(0) if b_qty else pd.Series(0, index=df_b.index)
        df_b['总成本'] = qty_series * df_b['成本'] * rate_rmb_to_fx
        df_b.loc[df_b['is_sd'] == True, '总成本'] = 0.0
        if b_status:
            df_b.loc[df_b[b_status].astype(str).str.lower() == 'canceled', '总成本'] = 0.0

        # ==========================================
        # 第八步 & 第九步：广告与毛利
        # ==========================================
        valid_sales_mask = df_b[b_status].astype(str).str.lower() != 'canceled' if b_status else pd.Series(True, index=df_b.index)
        total_valid_sales = df_b.loc[valid_sales_mask, '实际售价'].sum()
        df_b['广告'] = 0.0
        if total_valid_sales > 0:
            df_b.loc[valid_sales_mask, '广告'] = (df_b.loc[valid_sales_mask, '实际售价'] / total_valid_sales) * total_ads_fx

        df_b['毛利'] = df_b['实际售价'] + df_b['Total Fees'] - df_b['总成本'] - df_b['广告'] - df_b['刷单'] - df_b['刷单佣金']

        # ==========================================
        # 🟢 终极强制映射：携带双层表头组装表A
        # ==========================================
        # 创建一个空表，完全继承原表 A 的列数和双排表头结构
        df_final = pd.DataFrame(index=df_b.index, columns=df_a_raw.columns)

        def get_t_col_tuple(keywords):
            for col in df_final.columns:
                # 把第一行中文和第二行英文合并起来，无论搜什么语言的词都能找到这列
                search_text = " ".join([str(x) for x in col if pd.notna(x)]).lower().replace('\n', ' ')
                for k in keywords:
                    if re.search(k, search_text, re.IGNORECASE):
                        return col
            return None

        # 暴力赋值器：只要找到了对应列，直接把 B 表算好的数据覆盖过去
        def assign_to_final(df_col_name, keywords):
            t_col = get_t_col_tuple(keywords)
            if t_col and df_col_name in df_b.columns:
                df_final[t_col] = df_b[df_col_name]

        # 映射基础销售字段
        if b_order_col: assign_to_final(b_order_col, [r'^order number$', r'^order id$', r'订单号'])
        if b_status: assign_to_final(b_status, [r'order status', r'status', r'订单状态'])
        if b_sku: assign_to_final(b_sku, [r'seller sku', r'sku'])
        if b_qty: assign_to_final(b_qty, [r'^quantity$', r'数量', r'count'])
        if b_return: assign_to_final(b_return, [r'quantity of return', r'退货', r'取消数量'])
        if b_plat_disc: assign_to_final(b_plat_disc, [r'platform discount', r'平台折扣'])
        if b_sell_disc: assign_to_final(b_sell_disc, [r'seller discount', r'买家折扣', r'卖家折扣'])
        if b_sub_after: assign_to_final(b_sub_after, [r'subtotal after discount', r'售价'])

        # 映射核算出来的业务字段
        assign_to_final('订单统计', [r'订单统计', r'计数'])
        assign_to_final('实际售价', [r'实际售价', r'销售计算'])
        assign_to_final('佣金共计', [r'佣金共计', r'佣金总计'])
        assign_to_final('成本', [r'^成本$'])
        assign_to_final('总成本', [r'^总成本$'])
        assign_to_final('广告', [r'^广告$'])
        assign_to_final('刷单', [r'^刷单$', r'刷单费用'])
        assign_to_final('刷单佣金', [r'^刷单佣金$'])
        assign_to_final('毛利', [r'^毛利$'])

        # 映射 36 项明细费项
        for fee in all_fee_columns:
            assign_to_final(fee, [r'(?i)' + re.escape(fee)])

        # ==========================================
        # 🟢 全局统计表生成
        # ==========================================
        total_qty = qty_series.sum() if isinstance(qty_series, pd.Series) else 0
        total_sales = df_b['实际售价'].sum()
        total_ads = df_b['广告'].sum()
        total_cost = df_b['总成本'].sum()
        total_profit = df_b['毛利'].sum()

        def pct(val, base): return f"{(val/base)*100:.2f}%" if base > 0 else "0.00%"

        summary_data = [
            {'分析指标': '销售总额 (Actual Sales)', '金额/数值': total_sales, '占销售额百分比': '-'},
            {'分析指标': '销售总数量 (Qty)', '金额/数值': total_qty, '占销售额百分比': '-'},
            {'分析指标': '广告费用 (Ads)', '金额/数值': total_ads, '占销售额百分比': pct(total_ads, total_sales)},
            {'分析指标': '总成本 (Cost)', '金额/数值': total_cost, '占销售额百分比': pct(total_cost, total_sales)},
            {'分析指标': '毛利总额 (Gross Profit)', '金额/数值': total_profit, '占销售额百分比': pct(total_profit, total_sales)},
        ]
        
        for fee in all_fee_columns:
            if fee in df_b.columns:
                fee_sum = df_b[fee].sum()
                if fee_sum != 0:
                    summary_data.append({'分析指标': f"【费项】{fee}", '金额/数值': fee_sum, '占销售额百分比': pct(fee_sum, total_sales)})
        df_summary = pd.DataFrame(summary_data)

        shipping_cols = [
            'Payment Fee', 'Shipping cost', 'Shipping costs passed on to the logistics provider', 
            'Replacement shipping fee (passed on to the customer)', 'Exchange shipping fee (passed on to the customer)', 
            'Shipping cost borne by the platform', 'Shipping cost paid by the customer', 
            'Refunded shipping cost paid by the customer', 'Return shipping costs (passed on to the customer)', 
            'Shipping cost subsidy', 'Distance shipping fee from Horizon+ Program'
        ]
        valid_ship_cols = [c for c in shipping_cols if c in df_b.columns]
        
        if b_sku and b_sku in df_b.columns:
            df_shipping = df_b.groupby(b_sku)[valid_ship_cols].sum().reset_index()
            df_shipping.rename(columns={b_sku: 'Seller sku'}, inplace=True)
            df_shipping['物流相关费用总计'] = df_shipping[valid_ship_cols].sum(axis=1)
            df_shipping = df_shipping.sort_values(by='物流相关费用总计', ascending=True)

            df_sku = df_b.groupby(b_sku).agg(
                销量=(b_qty if b_qty else 'Quantity', 'sum'),
                销售额=('实际售价', 'sum'),广告花费=('广告', 'sum'),
                总成本=('总成本', 'sum'),毛利=('毛利', 'sum')
            ).reset_index()
            df_sku.rename(columns={b_sku: 'Seller sku'}, inplace=True)
            df_sku['成本占比'] = df_sku.apply(lambda x: f"{(x['总成本']/x['销售额'])*100:.2f}%" if x['销售额']>0 else "0.00%", axis=1)
            df_sku['广告占比'] = df_sku.apply(lambda x: f"{(x['广告花费']/x['销售额'])*100:.2f}%" if x['销售额']>0 else "0.00%", axis=1)
            df_sku['毛利率'] = df_sku.apply(lambda x: f"{(x['毛利']/x['销售额'])*100:.2f}%" if x['销售额']>0 else "0.00%", axis=1)
            df_sku = df_sku.sort_values(by='销售额', ascending=False)
        else:
            df_shipping = pd.DataFrame()
            df_sku = pd.DataFrame()

        # ==========================================
        # 打包导出
        # ==========================================
        st.divider()
        st.success("✅ 数据合并大成功！导出的 Excel 将完美保留原有的双层表头结构。")
        st.dataframe(df_final.head(15))

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_final.to_excel(writer, sheet_name='账单明细(表A)', index=False)
            df_summary.to_excel(writer, sheet_name='店铺汇总', index=False)
            if not df_sku.empty: df_sku.to_excel(writer, sheet_name='SKU分析', index=False)
            if not df_shipping.empty: df_shipping.to_excel(writer, sheet_name='物流费分析', index=False)
            
        st.download_button(label="📥 下载带有双层表头的多维数据报表", data=output.getvalue(), file_name="TK_Dashboard_Report_V2.xlsx")

    except Exception as e:
        st.error(f"❌ 运行发生异常: {e}")
