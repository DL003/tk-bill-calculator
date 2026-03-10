import streamlit as st
import pandas as pd
import io
import re

st.set_page_config(page_title="TK 账单助手-空值免疫版", layout="wide")

# --- 侧边栏 ---
with st.sidebar:
    st.header("1. 参数配置")
    rate_rmb_to_fx = st.number_input("汇率 (1 RMB = ? 外币)", value=2200.0)
    ads_total_fx = st.number_input("总广告费 (外币)", value=0.0)

# --- 文件上传 ---
st.header("2. 数据上传")
col_a, col_b, col_c = st.columns(3)
col_d, col_e, _ = st.columns(3)

with col_a: file_a = st.file_uploader("上传【表 A 模板】", type=["xlsx", "csv"])
with col_b: file_b = st.file_uploader("上传【表 B】(销售表)", type=["xlsx", "csv"])
with col_c: file_c = st.file_uploader("上传【表 C】(收入表)", type=["xlsx", "csv"])
with col_d: file_d = st.file_uploader("上传【表 D】(成本表)", type=["xlsx", "csv"])
with col_e: file_e = st.file_uploader("上传【表 E】(刷单-可选)", type=["xlsx", "csv"])

def find_col_regex(columns, pattern):
    for col in columns:
        if pd.notna(col) and re.search(pattern, str(col), re.IGNORECASE):
            return col
    return None

def load_file(file_obj, header_config):
    if file_obj.name.endswith('.csv'):
        return pd.read_csv(file_obj, header=header_config)
    return pd.read_excel(file_obj, header=header_config)

if all([file_a, file_b, file_c, file_d]):
    try:
        # 1. 精准读取
        df_a_raw = load_file(file_a, header=[0, 1])
        template_keys = df_a_raw.columns.get_level_values(1).tolist()
        
        df_b_raw = load_file(file_b, header=[0, 1])
        df_b = df_b_raw.copy()
        df_b.columns = df_b_raw.columns.get_level_values(0).tolist()

        if file_c.name.endswith('.csv'):
            df_c = pd.read_csv(file_c)
        else:
            xl_c = pd.ExcelFile(file_c)
            target_sheet = next((s for s in xl_c.sheet_names if 'order' in s.lower() and 'detail' in s.lower()), xl_c.sheet_names[0])
            df_c = pd.read_excel(file_c, sheet_name=target_sheet)
            
        df_d = load_file(file_d, header=0)

        # 2. 识别核心匹配列
        b_order_col = find_col_regex(df_b.columns, r'Order ID|Order Number')
        c_order_col = find_col_regex(df_c.columns, r'Order/adjustment ID|Order ID|Order Number')
        b_sku_col = find_col_regex(df_b.columns, r'Seller SKU')
        d_sku_col = find_col_regex(df_d.columns, r'Seller SKU|SKU')
        b_qty_col = find_col_regex(df_b.columns, r'sold quantity|quantity')
        b_status_col = find_col_regex(df_b.columns, r'Status')
        
        # 3. 统一单号格式为字符串，消除 ID 类型不匹配
        df_b[b_order_col] = df_b[b_order_col].astype(str).str.strip()
        df_c[c_order_col] = df_c[c_order_col].astype(str).str.strip()
        df_b[b_sku_col] = df_b[b_sku_col].astype(str).str.strip()
        df_d[d_sku_col] = df_d[d_sku_col].astype(str).str.strip()

        # 4. 费项映射 (强制字符串转换防范空表头 NaN 报错)
        fee_mapping = {}
        for c_col in df_c.columns:
            clean_c = re.sub(r'\(.*?\)', '', str(c_col)).strip().lower()
            for t_key in template_keys:
                if pd.isna(t_key): continue # 忽略空表头
                t_str = str(t_key).strip().lower()
                if t_str == clean_c or t_str == str(c_col).lower():
                    fee_mapping[c_col] = t_key

        # 5. 合并与计算
        df_b['订单统计'] = df_b.groupby(b_order_col)[b_order_col].transform('count')
        
        common_fees = [c for c in fee_mapping.keys() if c != c_order_col]
        df = pd.merge(df_b, df_c[[c_order_col] + common_fees], left_on=b_order_col, right_on=c_order_col, how='left')

        # 分摊费用
        for c_col in common_fees:
            t_key = fee_mapping[c_col]
            # 【修复点】强制转为 str 后再 lower()，防止 float 对象报错
            if str(t_key).lower() not in ['order number', 'order id']:
                numeric_vals = pd.to_numeric(df[c_col], errors='coerce').fillna(0)
                df[t_key] = numeric_vals / df['订单统计']

        # 售价计算
        s_sub = find_col_regex(df_b.columns, r'Subtotal Before Discount')
        s_plat = find_col_regex(df_b.columns, r'Platform Discount')
        s_seller = find_col_regex(df_b.columns, r'Seller Discount')
        
        v_sub = pd.to_numeric(df[s_sub], errors='coerce').fillna(0) if s_sub else 0
        v_plat = pd.to_numeric(df[s_plat], errors='coerce').fillna(0) if s_plat else 0
        v_sell = pd.to_numeric(df[s_seller], errors='coerce').fillna(0) if s_seller else 0
        df['实际售价'] = (v_sub - v_plat - v_sell) + v_plat

        # 成本匹配
        d_cost_col = find_col_regex(df_d.columns, r'cost|成本|价格')
        df = pd.merge(df, df_d[[d_sku_col, d_cost_col]], left_on=b_sku_col, right_on=d_sku_col, how='left')
        
        # 刷单处理
        df['is_sd'] = False
        df['刷单费用'] = 0.0
        if file_e:
            df_e = load_file(file_e, header=0)
            e_order_col = find_col_regex(df_e.columns, r'Order ID|Order Number|单号')
            e_fee_col = find_col_regex(df_e.columns, r'fee|费用|刷单')
            if e_order_col and e_fee_col:
                df_e[e_order_col] = df_e[e_order_col].astype(str).str.strip()
                df = pd.merge(df, df_e[[e_order_col, e_fee_col]], left_on=b_order_col, right_on=e_order_col, how='left')
                df['is_sd'] = df[e_fee_col].notnull()
                df['刷单费用'] = pd.to_numeric(df[e_fee_col], errors='coerce').fillna(0) * rate_rmb_to_fx
        
        df['刷单佣金'] = df['is_sd'].apply(lambda x: 12.0 * rate_rmb_to_fx if x else 0.0)

        # 【修复点】矢量化成本与毛利计算，避免 apply 和索引错误
        qty_numeric = pd.to_numeric(df[b_qty_col], errors='coerce').fillna(0)
        cost_numeric = pd.to_numeric(df[d_cost_col], errors='coerce').fillna(0)
        
        df['总成本'] = qty_numeric * cost_numeric * rate_rmb_to_fx
        # 若是刷单，成本设为0
        df.loc[df['is_sd'] == True, '总成本'] = 0.0

        valid_mask = df[b_status_col].astype(str).str.lower() != 'canceled'
        total_sales = df.loc[valid_mask, '实际售价'].sum()
        df['广告'] = 0.0
        if total_sales > 0:
            df.loc[valid_mask, '广告'] = (df.loc[valid_mask, '实际售价'] / total_sales) * ads_total_fx

        c_total_fee_col = find_col_regex(df_c.columns, r'Total fees')
        df['毛利'] = 0.0
        if c_total_fee_col:
            fee_vals = pd.to_numeric(df[c_total_fee_col], errors='coerce').fillna(0)
            df.loc[valid_mask, '毛利'] = df['实际售价'] + (fee_vals/df['订单统计']) - df['总成本'] - df['广告'] - df['刷单费用'] - df['刷单佣金']

        # 6. 列拼接构建最终表
        out_cols = []
        for t_key in template_keys:
            if pd.isna(t_key):
                # 遇到空白列，填充空数据保持结构完整
                out_cols.append(pd.Series(None, index=df.index, name=t_key))
                continue
                
            clean_key = str(t_key).strip()
            if clean_key == 'order number':
                out_cols.append(df[b_order_col].rename(t_key))
            elif clean_key == 'order status':
                out_cols.append(df[b_status_col].rename(t_key))
            elif clean_key == 'Seller sku':
                out_cols.append(df[b_sku_col].rename(t_key))
            elif clean_key in ['订单统计', '实际售价', '总成本', '广告', '毛利', '刷单', '刷单佣金']:
                col_name_to_extract = '刷单费用' if clean_key == '刷单' else clean_key
                out_cols.append(df[col_name_to_extract].rename(t_key))
            else:
                match = next((c for c in df.columns if str(c).strip().lower() == clean_key.lower()), None)
                if match:
                    out_cols.append(df[match].rename(t_key))
                else:
                    out_cols.append(pd.Series(None, index=df.index, name=t_key))

        df_final_data = pd.concat(out_cols, axis=1)
        df_final_data.columns = df_a_raw.columns

        st.divider()
        st.success("✅ 数据处理完毕，空白表头已自动忽略！")
        st.dataframe(df_final_data.head(10))

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_final_data.to_excel(writer)
        st.download_button(label="📥 下载最终汇总表", data=output.getvalue(), file_name="Final_Report_Resolved.xlsx")

    except Exception as e:
        st.error(f"❌ 运行错误: {e}")
