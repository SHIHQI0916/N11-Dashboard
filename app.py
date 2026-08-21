import streamlit as st
import pandas as pd
import sqlite3
import os
import datetime
import time
import plotly.express as px

# ==========================================
# 1. 基本設定與資料庫初始化
# ==========================================
st.set_page_config(page_title="📦 N11 進貨營運監控", layout="wide")

DATA_DIR = "Data"
FILE_PATH = os.path.join(DATA_DIR, "進貨管理.xlsx")
DB_PATH = os.path.join(DATA_DIR, "notes.db")

if not os.path.exists(DATA_DIR):
    os.makedirs(DATA_DIR)

def init_db():
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS notes (arrival_no TEXT, area TEXT, note TEXT, status TEXT DEFAULT '處理中', update_time TEXT, PRIMARY KEY(arrival_no, area))''')
    try: c.execute("ALTER TABLE notes ADD COLUMN area TEXT DEFAULT 'RPH'")
    except sqlite3.OperationalError: pass
    try: c.execute("ALTER TABLE notes ADD COLUMN status TEXT DEFAULT '處理中'")
    except sqlite3.OperationalError: pass
    try: c.execute("ALTER TABLE notes ADD COLUMN update_time TEXT")
    except sqlite3.OperationalError: pass
    try: c.execute("ALTER TABLE notes ADD COLUMN author TEXT DEFAULT ''")
    except sqlite3.OperationalError: pass
    conn.commit()
    conn.close()

init_db()

def get_active_notes(area):
    conn = sqlite3.connect(DB_PATH)
    df = pd.read_sql(f"SELECT arrival_no, note, author FROM notes WHERE (status = '處理中' OR status IS NULL) AND area = '{area}'", conn)
    conn.close()
    if not df.empty:
        df['note'] = df.apply(lambda x: f"[{x['author']}] {x['note']}" if pd.notna(x['author']) and str(x['author']).strip() != '' else x['note'], axis=1)
    return df[['arrival_no', 'note']]

def get_all_notes(area):
    conn = sqlite3.connect(DB_PATH)
    df = pd.read_sql(f"SELECT * FROM notes WHERE area = '{area}'", conn)
    conn.close()
    return df

def save_note(arrival_no, note_text, area, author):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    c.execute('REPLACE INTO notes (arrival_no, area, note, status, update_time, author) VALUES (?, ?, ?, ?, ?, ?)', (arrival_no, area, note_text, '處理中', now_str, author))
    conn.commit()
    conn.close()

def resolve_note(arrival_no, area):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    c.execute("UPDATE notes SET status = '已解決', update_time = ? WHERE arrival_no = ? AND area = ?", (now_str, arrival_no, area))
    conn.commit()
    conn.close()

def revert_note(arrival_no, area):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    c.execute("UPDATE notes SET status = '處理中', update_time = ? WHERE arrival_no = ? AND area = ?", (now_str, arrival_no, area))
    conn.commit()
    conn.close()

def delete_note(arrival_no, area):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("DELETE FROM notes WHERE arrival_no = ? AND area = ?", (arrival_no, area))
    conn.commit()
    conn.close()

# ==========================================
# 2. 核心大腦：SLA 時間計算邏輯
# ==========================================
def calculate_sla(start_time, priority, holidays_set):
    if pd.isna(start_time): return pd.NaT
    if not isinstance(start_time, pd.Timestamp): start_time = pd.to_datetime(start_time)
    
    prio_str = str(priority).lower()
    if 'emergency' in prio_str: add_m = 60
    elif 'priority' in prio_str: add_m = 480
    else: add_m = 840
        
    curr = start_time
    while add_m > 0:
        curr += pd.Timedelta(minutes=1)
        if curr.date() in holidays_set or curr.weekday() >= 5:
            curr = pd.Timestamp(curr.date()) + pd.Timedelta(days=1, hours=9)
            continue
        
        time_mins = curr.hour * 60 + curr.minute
        if time_mins < 540: curr = curr.replace(hour=9, minute=0, second=0)
        elif 720 <= time_mins < 780: curr = curr.replace(hour=13, minute=0, second=0)
        elif time_mins > 1080: curr = pd.Timestamp(curr.date()) + pd.Timedelta(days=1, hours=9)
        else: add_m -= 1
    return curr

def format_time_diff(diff_minutes):
    if pd.isna(diff_minutes): return ""
    hrs = int(abs(diff_minutes) // 60)
    mins = int(abs(diff_minutes) % 60)
    return f"{hrs}H{mins}M"

# ==========================================
# 3. 資料處理引擎 
# ==========================================
@st.cache_data(ttl=300) 
def load_and_process_data(file_path, area):
    sheet_name = f'Raw_{area}_IB'
    try:
        df_raw = pd.read_excel(file_path, sheet_name=sheet_name)
        df_hol = pd.read_excel(file_path, sheet_name='假日表')
    except ValueError:
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), False
        
    holidays_set = set(pd.to_datetime(df_hol.iloc[:, 0].dropna()).dt.date)
    
    df_raw['Status'] = df_raw['Status'].fillna('未作業')
    df_raw['Print Status'] = df_raw['Print Status'].fillna('N')
    
    df_base = df_raw[df_raw['Seq'].isin([0, '0', '0000', 0.0]) & (df_raw['Status'] != '確定完成')].copy()
    
    if df_base.empty: return pd.DataFrame(), pd.DataFrame(), df_raw, True
    
    now_time = pd.Timestamp(datetime.datetime.now())
    results_ongoing, results_not_arrived = [], []
    
    grouped = df_base.groupby(['DN', 'Arrival No.'])
    for (dn, arr_no), group in grouped:
        total_l = len(group)
        comp_l = (group['Status'] == '檢品完成').sum()
        srv = group['Service Level'].iloc[0]
        is_printed = (group['Print Status'] == 'Y').any()
        
        supp_raw = group['Supplier Name1'].iloc[0]
        supplier_name = str(supp_raw) if pd.notna(supp_raw) and str(supp_raw).strip() != "" else "未知廠商"
        
        if is_printed:
            cat = "待上架 (檢品完)" if comp_l == total_l else "進行中 (已到貨)"
            prog_val = comp_l / total_l
            prog_str = f"{int(prog_val*100)}% ({comp_l}/{total_l})"
            
            unfin_lines = group.loc[group['Status'] == '未作業', 'Line'].astype(str).tolist()
            unfin_str = ", ".join(unfin_lines)
            line_msg = f"未上架: {', '.join(group['Line'].astype(str).tolist())}" if cat == "待上架 (檢品完)" else f"未收料: {unfin_str}"
            
            recv_d = pd.to_datetime(group['Arrival Receiving Date'].iloc[0], errors='coerce')
            gr_d = pd.to_datetime(group['GR Time'], errors='coerce').max()
            base_t = gr_d if cat == "待上架 (檢品完)" else recv_d
            base_t_str = base_t.strftime("%Y/%m/%d %H:%M") if pd.notna(base_t) else ("未填寫到貨時間" if cat=="進行中 (已到貨)" else "-")
            
            kpi_t = calculate_sla(base_t, srv, holidays_set) if cat == "進行中 (已到貨)" else pd.NaT
            kpi_t_str = kpi_t.strftime("%Y/%m/%d %H:%M") if pd.notna(kpi_t) else "-"
            
            wait_minutes = (now_time - base_t).total_seconds() / 60 if pd.notna(base_t) else 0
            sla_diff_m = 9999 
            
            if cat == "待上架 (檢品完)":
                msg = f"已待上架 {format_time_diff(wait_minutes)} ⚠️" if pd.notna(base_t) else "缺少收料時間"
                sort_key = (base_t.timestamp() - now_time.timestamp()) + 1000000000 if pd.notna(base_t) else 1999999999
            else:
                if pd.notna(kpi_t):
                    sla_diff_m = (kpi_t - now_time).total_seconds() / 60
                    t_str = format_time_diff(sla_diff_m)
                    msg = f"逾時 {t_str} 🚨" if sla_diff_m < 0 else (f"剩餘 {t_str} 🚨" if sla_diff_m < 60 else f"剩餘 {t_str} ✅")
                    sort_key = sla_diff_m
                else:
                    msg, sort_key = "無倒數 (缺時間)", 999999999

            results_ongoing.append({
                "狀態類別": cat, "Service Level": srv, "來源供應商": supplier_name, 
                "DN": dn, "Arrival No.": arr_no, "總Line數": total_l, "完成率": prog_val,
                "進度": prog_str, "🔴 需關注的 Line": line_msg, "起算基準": base_t_str, "KPI時間": kpi_t_str, 
                "狀態倒數/提示": msg, "_sort": sort_key, "_wait_m": wait_minutes, "SLA_剩餘分鐘": sla_diff_m
            })
        else:
            scdl_d = pd.to_datetime(group['Scdl Date'].iloc[0], errors='coerce')
            line_str = ", ".join(group['Line'].astype(str).tolist())
            if pd.notna(scdl_d):
                wait_days = (now_time.date() - scdl_d.date()).days
                msg, sort_key = f"已開單 {wait_days} 天未到 ⚠️", scdl_d.timestamp()
            else:
                msg, sort_key = "無開單日", 9999999999
                
            results_not_arrived.append({
                "狀態類別": "未到貨 (未印單)", "Service Level": srv, "來源供應商": supplier_name,
                "DN": dn, "Arrival No.": arr_no, "總Line數": total_l, "🔴 未到貨 Line 明細": f"未到貨: {line_str}", 
                "開單日期": scdl_d.strftime("%Y/%m/%d") if pd.notna(scdl_d) else "-", 
                "異常提示": msg, "_sort": sort_key
            })
            
    df_og = pd.DataFrame(results_ongoing).sort_values('_sort').drop(columns=['_sort']) if results_ongoing else pd.DataFrame()
    df_na = pd.DataFrame(results_not_arrived).sort_values('_sort').drop(columns=['_sort']) if results_not_arrived else pd.DataFrame()
    return df_og, df_na, df_raw, True

# ==========================================
# 4. 網頁 UI 呈現與側邊欄
# ==========================================
with st.sidebar:
    st.header("📍 系統作業區域切換")
    selected_area = st.radio("選擇您目前要監控的區域：", ["RPH", "RTH"])
    
    st.markdown("---")
    st.header("📥 資料更新")
    uploaded_file = st.file_uploader("上傳最新 Raw Data (覆蓋)", type=["xlsx"])
    if uploaded_file:
        with open(FILE_PATH, "wb") as f: f.write(uploaded_file.getbuffer())
        st.cache_data.clear()
        st.success("✅ 檔案已上傳！")
        time.sleep(0.5)
        st.rerun()
        
    st.markdown("---")
    st.header(f"📝 {selected_area} 異常備註管理")
    edit_arr = st.text_input("輸入單號 (Arrival No.)")
    edit_author = st.text_input("經手人/登記者 (必填，如：王大明)")
    edit_note = st.text_area("備註內容 (若要撤銷/刪除可留空)")
    
    col1, col2 = st.columns(2)
    if col1.button("💾 儲存 (處理中)"):
        if edit_arr and edit_author: 
            save_note(edit_arr.strip(), edit_note, selected_area, edit_author.strip())
            st.success("已儲存！")
            time.sleep(0.5)
            st.rerun()
        else: st.warning("單號與經手人不可空白！")
        
    if col2.button("✅ 標記已解決"):
        if edit_arr:
            resolve_note(edit_arr.strip(), selected_area)
            st.success("已歸檔至歷史庫！")
            time.sleep(0.5)
            st.rerun()
        else: st.warning("請輸入單號")
        
    st.markdown("<br>", unsafe_allow_html=True)
    st.caption("⚠️ 登記錯誤修改區")
    col3, col4 = st.columns(2)
    
    if col3.button("↩️ 撤銷回處理中"):
        if edit_arr:
            revert_note(edit_arr.strip(), selected_area)
            st.success("已恢復至現場清單！")
            time.sleep(0.5)
            st.rerun()
        else: st.warning("請輸入單號")
        
    if col4.button("🗑️ 完全刪除紀錄"):
        if edit_arr:
            delete_note(edit_arr.strip(), selected_area)
            st.error("已徹底刪除此備註！")
            time.sleep(0.5)
            st.rerun()
        else: st.warning("請輸入單號")

st.title(f"📦 N11 [{selected_area}] 現場進貨戰情儀表板")
st.caption(f"即時更新時間：{datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')} (每10分鐘自動更新 / 按 F5 手動重整)")

if not os.path.exists(FILE_PATH):
    st.warning("⚠️ 找不到資料，請由左側上傳。")
    st.stop()

df_og, df_na, df_raw, sheet_exists = load_and_process_data(FILE_PATH, selected_area)

if not sheet_exists:
    st.error(f"❌ 警告：在您上傳的 Excel 檔案中，找不到名為 `Raw_{selected_area}_IB` 的工作表。請確認檔案是否正確。")
    st.stop()

df_active_notes = get_active_notes(selected_area)
df_all_notes = get_all_notes(selected_area)

if not df_og.empty:
    df_og = pd.merge(df_og, df_active_notes, how='left', left_on='Arrival No.', right_on='arrival_no').drop(columns=['arrival_no'])
    df_og.rename(columns={'note': '📝 系統備註'}, inplace=True)
if not df_na.empty:
    df_na = pd.merge(df_na, df_active_notes, how='left', left_on='Arrival No.', right_on='arrival_no').drop(columns=['arrival_no'])
    df_na.rename(columns={'note': '📝 系統備註'}, inplace=True)

SL_COLORS = {
    'Emergency (S4)': '#d62728', 
    'Priority (S2)': '#ff7f0e',  
    'Routine (S1)': '#1f77b4',   
    'Routine (S3)': '#1f77b4',   
    '未標示': '#7f7f7f'
}

def style_kpi(val):
    val_str = str(val)
    if '🚨' in val_str: return 'background-color: #5c1818; color: #ffcccc; font-weight: bold;' 
    if '⚠️' in val_str: return 'background-color: #59470c; color: #ffe599; font-weight: bold;' 
    if '✅' in val_str: return 'background-color: #123b1f; color: #b3ffcc; font-weight: bold;' 
    return ''

def style_sl(val):
    val_str = str(val).lower()
    if 'emergency' in val_str or 's4' in val_str: return 'color: #ff4b4b; font-weight: bold;'
    if 'priority' in val_str or 's2' in val_str: return 'color: #ffc107; font-weight: bold;'
    if 'routine' in val_str or 's1' in val_str: return 'color: #21c354;'
    return ''

# ==========================================
# 5. 四大分頁顯示區 (✅ 加入全域搜尋列)
# ==========================================
t1, t2, t3, t4 = st.tabs(["🔴 進行中/待上架清單 (現場)", "⚠️ 異常未到貨清單", "📚 備註追蹤與歷史庫", "📊 主管與現場戰情總覽"])

with t1:
    if df_og.empty: 
        st.info(f"目前 {selected_area} 區域無進行中單據")
    else: 
        search_t1 = st.text_input("🔍 快速搜尋 (可輸入單號 / 供應商 / 狀態 / 等級...)", key="search_t1")
        df_display = df_og.drop(columns=['_wait_m', '完成率', 'SLA_剩餘分鐘'])
        
        # 搜尋篩選邏輯
        if search_t1:
            mask = df_display.astype(str).apply(lambda x: x.str.contains(search_t1, case=False, na=False)).any(axis=1)
            df_display = df_display[mask]
            
        styled_og = df_display.style.map(style_kpi, subset=['狀態倒數/提示']).map(style_sl, subset=['Service Level'])
        st.dataframe(styled_og, use_container_width=True, hide_index=True, height=600)
        st.caption("💡 提示：若單據太多，可以將滑鼠移到表格右上角，點擊 **⛶ (Fullscreen)** 圖示將表格放大至全螢幕。")

with t2:
    if df_na.empty: 
        st.info(f"目前 {selected_area} 區域無未到貨單據")
    else: 
        search_t2 = st.text_input("🔍 快速搜尋 (可輸入單號 / 供應商 / 等級...)", key="search_t2")
        df_na_display = df_na
        
        if search_t2:
            mask = df_na_display.astype(str).apply(lambda x: x.str.contains(search_t2, case=False, na=False)).any(axis=1)
            df_na_display = df_na_display[mask]
            
        styled_na = df_na_display.style.map(style_kpi, subset=['異常提示']).map(style_sl, subset=['Service Level'])
        st.dataframe(styled_na, use_container_width=True, hide_index=True, height=600)

with t3:
    st.subheader(f"📝 {selected_area} 備註案件追蹤與歷史檔案庫")
    if df_all_notes.empty: 
        st.success("此區域目前沒有備註紀錄。")
    else:
        search_t3 = st.text_input("🔍 歷史紀錄搜尋 (可輸入單號 / 備註內容 / 經手人 / 狀態...)", key="search_t3")
        
        df_status = df_raw.drop_duplicates(subset=['Arrival No.'])[['Arrival No.', 'Status']]
        df_track = pd.merge(df_all_notes, df_status, how='left', left_on='arrival_no', right_on='Arrival No.')
        df_track = df_track[['arrival_no', 'status', 'Status', 'note', 'author', 'update_time']]
        df_track.columns = ['Arrival No.', '備註狀態', '現場最新進度', '📝 備註內容', '經手人', '最後更新時間']
        df_track['現場最新進度'] = df_track['現場最新進度'].fillna('無資料/已完成')
        
        # 搜尋篩選邏輯
        if search_t3:
            mask = df_track.astype(str).apply(lambda x: x.str.contains(search_t3, case=False, na=False)).any(axis=1)
            df_track = df_track[mask]
        
        st.markdown("##### 🚨 處理中案件")
        st.dataframe(df_track[df_track['備註狀態'] == '處理中'], use_container_width=True, hide_index=True)
        st.markdown("##### ✅ 已解決歷史")
        st.dataframe(df_track[df_track['備註狀態'] == '已解決'], use_container_width=True, hide_index=True)

# ==========================================
# 分頁 4：動態模組化戰情室
# ==========================================
with t4:
    dynamic_sl_colors = {}
    if not df_og.empty:
        for sl in df_og['Service Level'].unique():
            sl_str = str(sl).lower()
            if 'emergency' in sl_str or 's4' in sl_str: dynamic_sl_colors[sl] = '#ff4b4b'
            elif 'priority' in sl_str or 's2' in sl_str: dynamic_sl_colors[sl] = '#ffc107'
            elif 'routine' in sl_str or 's1' in sl_str or 's3' in sl_str: dynamic_sl_colors[sl] = '#21c354'
            else: dynamic_sl_colors[sl] = '#7f7f7f'
    
    total_ongoing = len(df_og)
    total_lines = df_og['總Line數'].sum() if not df_og.empty else 0
    emergencies = len(df_og[df_og['Service Level'].str.contains('Emergency', na=False)]) if not df_og.empty else 0
    wait_putaway = len(df_og[df_og['狀態類別'] == '待上架 (檢品完)']) if not df_og.empty else 0
    overdue = len(df_og[df_og['狀態倒數/提示'].str.contains('🚨', na=False)]) if not df_og.empty else 0
    
    st.markdown("### 🎯 現場量能與風險指標")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("待處理/作業中單據", f"{total_ongoing} 單", f"共包含 {total_lines} Lines", delta_color="off")
    c2.metric("最急件 (Emergency)", f"{emergencies} 單", "最高優先" if emergencies>0 else "", delta_color="inverse")
    c3.metric("待上架 (清空暫存區)", f"{wait_putaway} 單")
    c4.metric("逾時告警危機 🚨", f"{overdue} 單", "- 需主管介入" if overdue>0 else "狀況良好", delta_color="inverse")
    
    if not df_og.empty:
        df_overdue = df_og[(df_og['狀態類別'] == '進行中 (已到貨)') & (df_og['狀態倒數/提示'].str.contains('逾時', na=False))]
        if not df_overdue.empty:
            with st.expander(f"🚨 點擊展開查看：今日逾時單據明細 (共 {len(df_overdue)} 單，已排除待上架)"):
                st.dataframe(df_overdue[['DN', 'Arrival No.', 'Service Level', '來源供應商', '總Line數', '狀態倒數/提示', '🔴 需關注的 Line']], use_container_width=True, hide_index=True)
        else:
            with st.expander("🚨 今日逾時單據明細 (目前無逾時單據 ✅)"):
                st.success("太棒了！目前沒有任何進行中且逾時的單據。")
    
    st.markdown("---")
    
    st.markdown("##### 🛠️ 請選擇您需要的監控視角 (支援拖曳排序)")
    all_charts = [
        "【現場盯單】🚨 SLA 倒數雷達 (抓急件必備)",
        "【現場盯單】🎯 進行中大單進度 (含 SLA 狀態)",
        "【現場盯單】🏆 神隊友雙榜單 (GR檢品 vs 上架)",
        "【主管監控】📈 今日每小時 GR檢品 產能趨勢",
        "【主管監控】⏳ 待上架滯留分析 (看暫存區塞車)", 
        "【主管監控】🧩 供應商作業佔比 (板塊樹狀圖)", 
        "【主管監控】🍩 緊急度佔比 (中空圓餅圖)"
    ]
    selected_charts = st.multiselect("圖表庫清單：", all_charts, 
                                     default=["【現場盯單】🚨 SLA 倒數雷達 (抓急件必備)", "【現場盯單】🎯 進行中大單進度 (含 SLA 狀態)", "【現場盯單】🏆 神隊友雙榜單 (GR檢品 vs 上架)"])
    
    def render_sla_radar():
        st.markdown("###### 🚨 現場 SLA 倒數雷達 (氣泡圖)")
        if not df_og.empty:
            df_alert = df_og[(df_og['狀態類別'] == '進行中 (已到貨)') & (df_og['SLA_剩餘分鐘'] != 9999)].copy()
            if not df_alert.empty:
                fig = px.scatter(df_alert, x='SLA_剩餘分鐘', y='Arrival No.', size='總Line數', 
                                 color='Service Level', color_discrete_map=dynamic_sl_colors,
                                 hover_name='Arrival No.',
                                 hover_data={'SLA_剩餘分鐘': False, '來源供應商': True, '狀態倒數/提示': True, '總Line數': True},
                                 text='狀態倒數/提示')
                fig.update_traces(textposition='top right', textfont_size=13)
                fig.add_vline(x=0, line_width=2, line_dash="dash", line_color="red", annotation_text="🚨 逾時線 (0分)", annotation_position="top left")
                fig.update_layout(xaxis=dict(autorange="reversed"), 
                                  xaxis_title="SLA 剩餘時間 (分鐘) ➡️ 越往右邊越接近超時", yaxis_title="", 
                                  height=400, margin=dict(t=20, b=0, l=0, r=0),
                                  legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
                st.plotly_chart(fig, use_container_width=True)
                return
        st.info("目前無正在倒數的進行中單據")

    def render_progress():
        st.markdown("###### 🎯 進行中大單進度排行 (Top 10)")
        if not df_og.empty:
            df_prog = df_og[df_og['狀態類別'] == '進行中 (已到貨)'].sort_values('總Line數', ascending=False).head(10)
            if not df_prog.empty:
                df_prog['進度與警示'] = (df_prog['完成率'] * 100).astype(int).astype(str) + "% | " + df_prog['狀態倒數/提示']
                fig = px.bar(df_prog, y='Arrival No.', x='完成率', text='進度與警示', orientation='h', 
                             hover_data=['來源供應商', '總Line數', 'Service Level'], 
                             color='Service Level', color_discrete_map=dynamic_sl_colors)
                fig.update_traces(textposition='outside', textfont_size=14)
                fig.update_layout(yaxis={'categoryorder':'total ascending'}, 
                                  xaxis_tickformat='.0%', xaxis_range=[0, 1.3],
                                  height=400, margin=dict(t=20, b=0, l=0, r=0), 
                                  legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
                st.plotly_chart(fig, use_container_width=True)
                return
        st.info("無進行中單據")

    def render_hourly_throughput():
        st.markdown("###### 📈 今日每小時 [GR檢品] 產能趨勢")
        df_raw['GR_DT'] = pd.to_datetime(df_raw['GR Time'], errors='coerce')
        seq_1_mask = df_raw['Seq'].isin([1, '1', '0001', 1.0])
        df_comp = df_raw[seq_1_mask].dropna(subset=['GR_DT']).copy()
        if not df_comp.empty:
            today_date = datetime.datetime.now().date()
            df_today = df_comp[df_comp['GR_DT'].dt.date == today_date].copy()
            if df_today.empty:
                latest_date = df_comp['GR_DT'].dt.date.max()
                df_today = df_comp[df_comp['GR_DT'].dt.date == latest_date].copy()
                st.caption(f"💡 今日尚無紀錄，自動顯示最近作業日 ({latest_date}) 數據")
            if not df_today.empty:
                df_today['Hour'] = df_today['GR_DT'].dt.strftime('%H:00')
                hourly_counts = df_today.groupby('Hour').size().reset_index(name='檢品Line數')
                fig = px.bar(hourly_counts, x='Hour', y='檢品Line數', text='檢品Line數', color_discrete_sequence=['#2ca02c'])
                fig.update_traces(textposition='outside', textfont_size=18)
                fig.update_layout(xaxis_title="時段 (基於 Seq=0001)", yaxis_title="完成檢品 Line 數", height=350, margin=dict(t=20, b=0, l=0, r=0))
                st.plotly_chart(fig, use_container_width=True)
                return
        st.info("尚無 GR 作業數據可供分析")

    def render_worker_leaderboard():
        st.markdown("###### 🏆 現場神隊友雙榜單 (今日 Top 10)")
        df_raw['GR_DT'] = pd.to_datetime(df_raw['GR Time'], errors='coerce')
        if 'Arrival Shelving Date' in df_raw.columns:
            df_raw['Putaway_DT'] = pd.to_datetime(df_raw['Arrival Shelving Date'], errors='coerce')
        else:
            df_raw['Putaway_DT'] = pd.NaT
            
        seq_1_mask = df_raw['Seq'].isin([1, '1', '0001', 1.0])
        seq_not_0_mask = ~df_raw['Seq'].isin([0, '0', '0000', 0.0])
        
        today_date = datetime.datetime.now().date()
        active_gr = df_raw[seq_1_mask].dropna(subset=['GR_DT'])
        
        if active_gr.empty:
            st.info("尚無作業數據可供分析")
            return
            
        if active_gr[active_gr['GR_DT'].dt.date == today_date].empty:
            today_date = active_gr['GR_DT'].dt.date.max()
            st.caption(f"💡 顯示基準日: {today_date}")
            
        col_gr, col_put = st.columns(2)
        
        with col_gr:
            st.markdown("**📦 檢品 (GR) 達人** (計算 Seq=0001)")
            df_today_gr = active_gr[active_gr['GR_DT'].dt.date == today_date].copy()
            df_today_gr = df_today_gr.dropna(subset=['Worker'])
            df_today_gr = df_today_gr[df_today_gr['Worker'].astype(str).str.strip() != '']
            
            if not df_today_gr.empty:
                gr_counts = df_today_gr.groupby('Worker').size().reset_index(name='GR次數').sort_values('GR次數', ascending=False).head(10)
                fig1 = px.bar(gr_counts, y='Worker', x='GR次數', text='GR次數', orientation='h', color_discrete_sequence=['#ff7f0e'])
                fig1.update_traces(textposition='outside', textfont_size=14)
                fig1.update_layout(yaxis={'categoryorder':'total ascending'}, xaxis_title="", yaxis_title="", height=280, margin=dict(t=0, b=0, l=0, r=0))
                st.plotly_chart(fig1, use_container_width=True)
            else:
                st.info("無檢品數據")

        with col_put:
            st.markdown("**🏗️ 上架 (Putaway) 達人** (計算所有分板 Seq)")
            if 'Location Entry User' in df_raw.columns:
                df_putaway = df_raw[(df_raw['Status'] == '確定完成') & (df_raw['Putaway_DT'].dt.date == today_date) & seq_not_0_mask].copy()
                df_putaway = df_putaway.dropna(subset=['Location Entry User'])
                df_putaway = df_putaway[df_putaway['Location Entry User'].astype(str).str.strip() != '']
                
                if not df_putaway.empty:
                    df_putaway['上架員'] = df_putaway['Location Entry User'].astype(str).apply(lambda x: x.split('@')[0])
                    put_counts = df_putaway.groupby('上架員').size().reset_index(name='上架次數').sort_values('上架次數', ascending=False).head(10)
                    fig2 = px.bar(put_counts, y='上架員', x='上架次數', text='上架次數', orientation='h', color_discrete_sequence=['#1f77b4'])
                    fig2.update_traces(textposition='outside', textfont_size=14)
                    fig2.update_layout(yaxis={'categoryorder':'total ascending'}, xaxis_title="", yaxis_title="", height=280, margin=dict(t=0, b=0, l=0, r=0))
                    st.plotly_chart(fig2, use_container_width=True)
                else:
                    st.info("尚無上架完成數據")
            else:
                st.warning("Excel 中缺少 'Location Entry User' 欄位")

    def render_aging():
        st.markdown("###### ⏳ 待上架區 滯留時間分析 (Aging)")
        if not df_og.empty:
            df_wait = df_og[df_og['狀態類別'] == '待上架 (檢品完)'].copy()
            if not df_wait.empty:
                bins = [-1, 60, 180, 99999]
                labels = ['1小時內 (剛放)', '1~3小時 (留意)', '3小時以上 🚨']
                df_wait['滯留區間'] = pd.cut(df_wait['_wait_m'], bins=bins, labels=labels)
                aging_counts = df_wait['滯留區間'].value_counts().reset_index()
                color_map = {'1小時內 (剛放)': '#21c354', '1~3小時 (留意)': '#ffa421', '3小時以上 🚨': '#ff4b4b'}
                fig = px.bar(aging_counts, x='滯留區間', y='count', text='count', color='滯留區間', color_discrete_map=color_map)
                fig.update_traces(textposition='outside', textfont_size=18)
                fig.update_layout(xaxis_title="", yaxis_title="單據數量", showlegend=False, height=350, margin=dict(t=20, b=0, l=0, r=0))
                st.plotly_chart(fig, use_container_width=True)
                return
        st.info("目前無待上架單據")

    def render_treemap():
        st.markdown("###### 🧩 供應商作業量佔比 (板塊樹狀圖)")
        if not df_og.empty:
            df_tree = df_og.copy()
            df_tree['Root'] = '總作業量'
            df_tree['Service Level'] = df_tree['Service Level'].fillna('未標示')
            fig = px.treemap(df_tree, path=['Root', '來源供應商', 'Service Level'], values='總Line數',
                             color='Service Level', color_discrete_map=dynamic_sl_colors)
            fig.update_layout(height=350, margin=dict(t=20, b=0, l=0, r=0))
            st.plotly_chart(fig, use_container_width=True)
        else: st.info("無進行中單據")
        
    def render_donut():
        st.markdown("###### 🍩 各緊急度 Line 數分佈 (圓餅圖)")
        if not df_og.empty:
            sl_counts = df_og.groupby('Service Level')['總Line數'].sum().reset_index()
            fig = px.pie(sl_counts, names='Service Level', values='總Line數', hole=0.4, 
                         color='Service Level', color_discrete_map=dynamic_sl_colors)
            fig.update_traces(textposition='inside', textinfo='percent+label', textfont_size=16)
            fig.update_layout(height=350, margin=dict(t=20, b=0, l=0, r=0), showlegend=False)
            st.plotly_chart(fig, use_container_width=True)
        else: st.info("無進行中單據")

    chart_funcs = {
        "【現場盯單】🚨 SLA 倒數雷達 (抓急件必備)": render_sla_radar,
        "【現場盯單】🎯 進行中大單進度 (含 SLA 狀態)": render_progress,
        "【現場盯單】🏆 神隊友雙榜單 (GR檢品 vs 上架)": render_worker_leaderboard,
        "【主管監控】📈 今日每小時 GR檢品 產能趨勢": render_hourly_throughput,
        "【主管監控】⏳ 待上架滯留分析 (看暫存區塞車)": render_aging,
        "【主管監控】🧩 供應商作業佔比 (板塊樹狀圖)": render_treemap,
        "【主管監控】🍩 緊急度佔比 (中空圓餅圖)": render_donut
    }
    
    st.markdown("<br>", unsafe_allow_html=True)
    
    for chart_name in selected_charts:
        if "雙榜單" in chart_name:
            st.markdown("---")
            chart_funcs[chart_name]()
            st.markdown("---")
            selected_charts.remove(chart_name)
            break
            
    for i in range(0, len(selected_charts), 2):
        cols = st.columns(2)
        with cols[0]:
            chart_funcs[selected_charts[i]]()
        if i + 1 < len(selected_charts):
            with cols[1]:
                chart_funcs[selected_charts[i+1]]()