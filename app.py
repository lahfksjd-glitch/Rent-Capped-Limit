import os
import glob
import pandas as pd
import streamlit as st

# 頁面配置
st.set_page_config(page_title="案件撥款查詢站", layout="wide", initial_sidebar_state="expanded")

# 自訂樣式：為深色模式優化，房東姓名與地址加上白底黑字
st.markdown("""
<style>
    .info-label { font-size: 20px; font-weight: bold; color: #94a3b8; margin-bottom: 12px; }
    .info-value-white { font-size: 22px; font-weight: bold; color: #0f172a; background-color: #ffffff; padding: 4px 12px; border-radius: 6px; display: inline-block; border: 1px solid #cbd5e1; }
    .info-id { font-size: 22px; font-weight: bold; color: #1d4ed8; background-color: #eff6ff; padding: 4px 12px; border-radius: 6px; display: inline-block; }
    .info-balance { font-size: 28px; font-weight: bold; color: #16a34a; background-color: #f0fdf4; padding: 2px 12px; border-radius: 6px; display: inline-block; }
    .info-date { font-size: 22px; font-weight: bold; color: #0f172a; background-color: #ffffff; padding: 4px 12px; border-radius: 6px; display: inline-block; border: 1px solid #cbd5e1; }
</style>
""", unsafe_allow_html=True)

st.title("📊 社會住宅經費控管與撥款查詢站")

# 民國年自動轉換函式
def to_roc_date(val):
    if pd.isna(val) or val is None:
        return '-'
    val_str = str(val).strip()
    if val_str in ['-', 'nan', 'None', '']:
        return '-'
    
    if ' ' in val_str:
        val_str = val_str.split(' ')[0]
        
    val_str_clean = val_str.replace('.', '/').replace('-', '/')
    parts = val_str_clean.split('/')
    
    if len(parts) == 3:
        try:
            y, m, d = int(parts[0]), int(parts[1]), int(parts[2])
            if y > 1900:
                y -= 1911
            return f"{y:03d}/{m:02d}/{d:02d}"
        except Exception:
            pass
            
    return val_str

@st.cache_data
def load_all_data():
    all_records = []
    excel_files = glob.glob("*.xlsx")
    
    for file_path in excel_files:
        try:
            xls = pd.ExcelFile(file_path)
            for sheet_name in ['代租 ', '包租 ']:
                if sheet_name in xls.sheet_names:
                    df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
                    
                    col_map = {}
                    for c in range(df.shape[1]):
                        header_vals = [str(df.iloc[r, c]) for r in range(min(4, len(df)))]
                        header_text = " ".join(header_vals)
                        
                        if '媒合編號' in header_text and 'match_id' not in col_map:
                            col_map['match_id'] = c
                        if ('房東姓名' in header_text or '房東' in header_text) and 'landlord' not in col_map:
                            if '代表人' in header_text or '房東姓名(代表人)' in header_text:
                                col_map['landlord'] = c
                        if ('房屋座落地址' in header_text or '座落地址' in header_text or '地址' in header_text) and 'address' not in col_map:
                            col_map['address'] = c
                        if ('租約起日' in header_text or '約起日' in header_text) and 'start_date' not in col_map:
                            col_map['start_date'] = c
                        if ('租約訖日' in header_text or '約訖日' in header_text) and 'end_date' not in col_map:
                            col_map['end_date'] = c
                        if '修繕費餘額' in header_text and 'repair_balance' not in col_map:
                            col_map['repair_balance'] = c
                        
                        if ('公證費(1)' in header_text or '公證費' in header_text) and 'notary_fee' not in col_map:
                            if '金額' in header_text or '公證費' in header_vals[1] or '公證費' in header_vals[2]:
                                col_map['notary_fee'] = c
                                if c + 1 < df.shape[1]:
                                    col_map['notary_date'] = c + 1

                    if 'landlord' not in col_map:
                        for c in range(df.shape[1]):
                            header_text = " ".join([str(df.iloc[r, c]) for r in range(min(4, len(df)))])
                            if '房東' in header_text:
                                col_map['landlord'] = c
                                break

                    match_col = col_map.get('match_id')
                    
                    if match_col is not None:
                        for row_idx in range(4, len(df)):
                            match_id = str(df.iloc[row_idx, match_col]).strip()
                            if match_id and match_id not in ['nan', 'None', '', '媒合編號']:
                                row = df.iloc[row_idx]
                                
                                landlord = row.iloc[col_map['landlord']] if 'landlord' in col_map else '未提供'
                                address = row.iloc[col_map['address']] if 'address' in col_map else '未提供'
                                start_date = row.iloc[col_map['start_date']] if 'start_date' in col_map else '-'
                                end_date = row.iloc[col_map['end_date']] if 'end_date' in col_map else '-'
                                repair_balance = row.iloc[col_map['repair_balance']] if 'repair_balance' in col_map else '0'
                                
                                notary_fee = row.iloc[col_map['notary_fee']] if 'notary_fee' in col_map else '-'
                                notary_date = row.iloc[col_map['notary_date']] if 'notary_date' in col_map else '-'

                                repair_list = []
                                for c in range(df.shape[1]):
                                    h_vals = [str(df.iloc[r, c]) for r in range(min(4, len(df)))]
                                    h_str = " ".join(h_vals)
                                    if '修繕費' in h_str and '餘額' not in h_str and '修繕費' in h_vals[-1]:
                                        val = row.iloc[c]
                                        date_val = row.iloc[c+1] if c+1 < df.shape[1] else '-'
                                        if pd.notna(val) and str(val) not in ['nan', 'None', '-']:
                                            repair_list.append((str(val), to_roc_date(date_val)))
                                
                                all_records.append({
                                    'match_id': match_id,
                                    'file_name': os.path.basename(file_path),
                                    'type': sheet_name.strip(),
                                    'landlord': str(landlord) if pd.notna(landlord) else '未提供',
                                    'address': str(address) if pd.notna(address) else '未提供',
                                    'start_date': to_roc_date(start_date),
                                    'end_date': to_roc_date(end_date),
                                    'repair_balance': str(repair_balance) if pd.notna(repair_balance) else '0',
                                    'notary_fee': str(notary_fee) if pd.notna(notary_fee) else '-',
                                    'notary_date': to_roc_date(notary_date),
                                    'repair_list': repair_list,
                                    'row_data': row
                                })
        except Exception as e:
            st.error(f"讀取檔案 {file_path} 失敗: {e}")
            
    return pd.DataFrame(all_records)

df_records = load_all_data()

st.sidebar.header("📁 資料庫狀態")
st.sidebar.success(f"已成功載入 {len(df_records)} 筆案件總資料")

st.subheader("🔍 案件速查看板（對照試算表頂部卡片）")
search_id = st.text_input("請輸入或貼上「媒合編號」（例如：力群新北B2M30500038）：", value="").strip()

if search_id:
    matched_df = df_records[df_records['match_id'].str.contains(search_id, case=False, na=False)]
    
    if not matched_df.empty:
        item = matched_df.iloc[0]
        
        col_left, col_right = st.columns([1, 1])
        
        with col_left:
            st.info("📋 **基本資訊**")
            
            # 使用自訂高對比白底樣式
            st.markdown(f"<div class='info-label'>媒 合 編 號： <span class='info-id'>{item['match_id']}</span></div>", unsafe_allow_html=True)
            st.markdown(f"<div class='info-label'>房 東 姓 名： <span class='info-value-white'>{item['landlord']}</span></div>", unsafe_allow_html=True)
            st.markdown(f"<div class='info-label'>可 用 餘 額： <span class='info-balance'>${item['repair_balance']}</span></div>", unsafe_allow_html=True)
            st.markdown(f"<div class='info-label'>到   期   日： <span class='info-date'>{item['end_date']}</span></div>", unsafe_allow_html=True)
            st.markdown(f"<div class='info-label'>房 屋 地 址： <span class='info-value-white'>{item['address']}</span></div>", unsafe_allow_html=True)
            
        with col_right:
            st.warning("💸 **修繕/公證費 & 撥款時間**")
            
            items = ["公證費"]
            status_list = [item['notary_fee']]
            dates_list = [item['notary_date']]
            
            repairs = item['repair_list']
            for idx in range(4):
                items.append(f"修繕費 (第{idx+1}次)")
                if idx < len(repairs):
                    status_list.append(repairs[idx][0])
                    dates_list.append(repairs[idx][1])
                else:
                    status_list.append("-")
                    dates_list.append("-")
            
            grant_data = {
                "項目": items,
                "撥款狀態/金額": status_list,
                "撥款日期": dates_list
            }
            st.table(pd.DataFrame(grant_data))

    else:
        st.error("❌ 查無相對應的媒合編號，請檢查輸入是否正確。")

st.markdown("---")

st.subheader("📋 案件經費控管總清單")
search_filter = st.text_input("🔎 清單關鍵字過濾（可輸入房東姓名、編號或房屋地址）：", key="list_filter")

display_df = df_records[['match_id', 'landlord', 'address', 'start_date', 'end_date', 'repair_balance', 'type']].copy()
display_df.columns = ['媒合編號', '房東姓名', '房屋地址', '租約起日', '租約訖日', '修繕餘額', '類別']

if search_filter:
    mask = display_df.apply(lambda row: row.astype(str).str.contains(search_filter, case=False).any(), axis=1)
    display_df = display_df[mask]

st.dataframe(display_df, use_container_width=True, height=400)

csv_data = display_df.to_csv(index=False).encode('utf-8-sig')
st.download_button(
    label="📥 下載/匯出目前篩選的 Excel (CSV) 報表",
    data=csv_data,
    file_name="social_housing_grants.csv",
    mime="text/csv"
)
