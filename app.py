import streamlit as st
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd
from datetime import datetime
import calendar
import json
import os
import time

# ==========================================
# 1. DATABASE CONNECTION (STREAMLIT SECRETS)
# ==========================================
def connect_sheet(sheet_name):
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    try:
        # GitHub/Streamlit Cloud deployment ke liye secrets
        creds_info = st.secrets["gcp_service_account"]
        creds_info["private_key"] = creds_info["private_key"].replace("\\n", "\n")
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_info, scope)
        client = gspread.authorize(creds)
        return client.open("Bike Check-In (Responses)").worksheet(sheet_name)
    except Exception as e:
        st.error(f"🔴 Connection Error: {sheet_name}")
        return None

def connect_sheet_by_url(sheet_url, sheet_name):
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    try:
        creds_info = st.secrets["gcp_service_account"]
        creds_info["private_key"] = creds_info["private_key"].replace("\\n", "\n")
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_info, scope)
        client = gspread.authorize(creds)
        return client.open_by_url(sheet_url).worksheet(sheet_name)
    except Exception as e:
        st.error(f"🔴 Connection Error: {str(e)}")
        return None

# Page Setup
st.set_page_config(page_title="Munich Motorrad Local", page_icon="🏍️", layout="wide")

# ==========================================
# 2. PRO CSS (OFFLINE OPTIMIZED)
# ==========================================
st.markdown("""
    <style>
    .stTabs [data-baseweb="tab-list"] { gap: 15px; background-color: #f0f2f6; padding: 10px; border-radius: 12px; }
    .stTabs [data-baseweb="tab"] { 
        height: 55px; background-color: white; border: 2px solid #0056b3 !important; 
        border-radius: 10px; color: #0056b3; font-weight: bold; padding: 0 25px;
    }
    .stTabs [aria-selected="true"] { background-color: #0056b3 !important; color: white !important; }
    
    /* Workshop Card Styling */
    .stExpander { border: 1px solid #ddd !important; border-radius: 12px !important; margin-bottom: 15px !important; }
    
    /* Big Update Button */
    button[key^="up_"] {
        height: 75px !important; background-color: #0056b3 !important; color: white !important;
        font-size: 22px !important; font-weight: 900 !important; border-radius: 15px !important;
    }
    
    /* Calendar Grid Scroll Fix */
    [data-testid="stVerticalBlockBorderWrapper"] > div > div {
        min-height: 250px !important; max-height: 300px !important; overflow-y: auto !important;
    }
    </style>
    """, unsafe_allow_html=True)

st.title("🏁 Munich Motorrad Workshop Management")

tab_ws, tab_app, tab_call, tab_kpi, tab_parts = st.tabs([
    "🔧 WORKSHOP MANAGER", 
    "📅 APPOINTMENTS", 
    "📞 SERVICE CALLING", 
    "📊 KPI DASHBOARD",
    "📦 PARTS"
])
# ------------------------------------------
# TAB 1: WORKSHOP MANAGER
# ------------------------------------------
with tab_ws:
    try:
        ws_sheet = connect_sheet("Digital JC")
        ws_data = ws_sheet.get_all_values()
        
        if len(ws_data) > 1:
            header = ws_data[0]
            all_items = []
            for i, r in enumerate(ws_data[1:]):
                all_items.append({'row_idx': i + 2, 'data': r})

            sub1, sub2 = st.tabs(["📂 Open JCs", "✅ History (Delivered)"])
            
            with sub1:
                s_o = st.text_input("🔍 Search Open JCs", key="ws_so")
                open_list = [item for item in all_items if len(item['data']) > 18 and item['data'][18] != "Delivered"]
                
                def workshop_sort(item):
                    status = item['data'][18].strip() if len(item['data']) > 18 else ""
                    if status == "On Hold": return 1
                    if status == "In Process": return 2
                    if status == "Complete": return 3
                    return 4
                
                open_list.sort(key=workshop_sort)

                for item in open_list:
                    r = item['data']
                    idx = item['row_idx']
                    stat = r[18].strip() if len(r) > 18 else "In Process"
                    
                    if stat == "On Hold":
                        display_stat, title_color = f"🔴 {stat}", "#FF4B4B"
                    elif stat == "Complete":
                        display_stat, title_color = f"🟢 {stat}", "#28a745"
                    else:
                        display_stat, title_color = f"🟠 {stat}", "#FF8C00"

                    if s_o.lower() in str(r).lower():
                        with st.expander(f"{r[1]} | {r[5]} | {display_stat}"):
                            st.markdown(f"<h2 style='color:{title_color}; margin-top:0;'>{display_stat}</h2>", unsafe_allow_html=True)
                            c1, c2, c3 = st.columns(3)
                            with c1:
                                st.markdown(f"**👤 Customer:**\n{r[1]}\n---\n**📱 Phone:**\n{r[2]}\n---\n**🆔 VIN:**\n{r[8]}")
                            with c2:
                                st.markdown(f"**🛠️ Staff:**\n{r[9]}\n---\n**🛣️ Odometer (KM):**\n{r[7] if len(r)>7 else 'N/A'}\n---\n")
                                st.info(f"**🚨 Issues:**\n{r[15] if len(r)>15 else 'N/A'}")
                            with c3:
                                n_rem = st.text_area("Workshop Remark", r[17] if len(r)>17 else "", key=f"r_{idx}", height=120)
                                st_opts = ["On Hold", "In Process", "Complete", "Delivered"]
                                try: curr_sel = st_opts.index(stat)
                                except: curr_sel = 1
                                n_stat = st.selectbox("Update Status", st_opts, index=curr_sel, key=f"s_{idx}")
                                if st.button("SAVE CHANGES", key=f"up_{idx}", use_container_width=True):
                                    ws_sheet.update_cell(idx, 18, n_rem)
                                    ws_sheet.update_cell(idx, 19, n_stat)
                                    st.rerun()

            with sub2:
                s_h = st.text_input("🔍 Search Delivered History", key="ws_sh")
                history_list = [item for item in all_items if len(item['data']) > 18 and item['data'][18] == "Delivered"]
                for item in history_list:
                    r = item['data']
                    if s_h.lower() in str(r).lower():
                        with st.expander(f"✅ {r[1]} | {r[5]} | Delivered"):
                            hc1, hc2, hc3 = st.columns(3)
                            with hc1: st.markdown(f"**👤 Customer:** {r[1]}\n**📱 Phone:** {r[2] if len(r)>2 else 'N/A'}\n**📧 Email:** {r[3] if len(r)>3 else 'N/A'}\n**📅 Date:** {r[0]}")
                            with hc2: st.markdown(f"**🆔 VIN:** {r[8] if len(r)>8 else 'N/A'}\n**🛠️ Staff:** {r[9] if len(r)>9 else 'N/A'}\n**🛣️ Odo:** {r[7] if len(r)>7 else 'N/A'} KM")
                            with hc3:
                                st.warning(f"**🚨 Issues:**\n{r[15] if len(r)>15 else 'N/A'}")
                                st.success(f"**✍️ Final Remark:**\n{r[17] if len(r)>17 else 'N/A'}")
                                st.info(f"**Status:** Delivered")
    except Exception as e: st.error(f"Workshop Error: {e}")
# ------------------------------------------
# TAB 2: APPOINTMENT CALENDAR
# ------------------------------------------
with tab_app:
    try:
        app_sheet = connect_sheet("Appointments")
        app_vals = app_sheet.get_all_values()
        
        if len(app_vals) > 1:
            df = pd.DataFrame([{"row": i+2, "N": r[1], "D": r[3], "M": r[4], "P": r[6], "R": r[7] if len(r)>7 else ""} for i, r in enumerate(app_vals[1:])])
            df['D'] = pd.to_datetime(df['D'], errors='coerce')
            df = df.dropna(subset=['D'])

            cy, cm = st.columns(2)
            y = cy.selectbox("Year", [2025, 2026], index=1)
            mn = cm.selectbox("Month", list(calendar.month_name)[1:], index=datetime.now().month-1)
            midx = list(calendar.month_name).index(mn)

            today = datetime.now().date()
            month_df = df[(df['D'].dt.month == midx) & (df['D'].dt.year == y)].copy()
            
            total_apps = len(month_df)
            reported_apps = len(month_df[month_df['R'].str.contains("Bike Reported", na=False)])
            pending_apps = len(month_df[(month_df['D'].dt.date < today) & (~month_df['R'].str.contains("Bike Reported", na=False))])

            st.markdown("---")
            s1, s2, s3 = st.columns(3)
            s1.metric("Total Appointments", total_apps)
            s2.metric("✅ Reported (Done)", reported_apps)
            s3.metric("⚠️ Missed/Pending", pending_apps)
            st.markdown("---")

            cal = calendar.monthcalendar(y, midx)
            cols_header = st.columns(7)
            days_list = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]
            for j, d in enumerate(days_list): 
                cols_header[j].markdown(f"<p style='text-align:center;font-weight:bold;color:#0056b3;'>{d}</p>", unsafe_allow_html=True)

            for week in cal:
                cols = st.columns(7)
                for i, day in enumerate(week):
                    if day != 0:
                        with cols[i].container(border=True):
                            st.markdown(f"<div style='text-align:right; font-weight:bold; color:#555;'>{day}</div>", unsafe_allow_html=True)
                            day_data = month_df[month_df['D'].dt.day == day]
                            for _, row in day_data.iterrows():
                                is_rep = "Bike Reported" in str(row['R'])
                                if is_rep: b_c = "#28a745"
                                elif row['D'].date() < today: b_c = "#d9534f"
                                else: b_c = "#007bff"
                                
                                st.markdown(f"""<style>div.stButton > button[key="trig_{row['row']}"] {{ border: 2.5px solid {b_c} !important; min-height: 60px !important; margin-bottom: 5px !important; }}</style>""", unsafe_allow_html=True)
                                if st.button(f"{'✅' if is_rep else '⭕'} {row['N']}", key=f"trig_{row['row']}", use_container_width=True):
                                    app_sheet.update_cell(row['row'], 8, "Bike Reported" if not is_rep else "")
                                    st.rerun()
                    else:
                        cols[i].markdown("<div style='height:100px; opacity:0.1;'></div>", unsafe_allow_html=True)
    except Exception as e: st.error(f"Calendar Error: {e}")
# ------------------------------------------
# TAB 3: SERVICE CALLING
# ------------------------------------------
with tab_call:
    try:
        @st.cache_data(ttl=60)
        def load_service_calling_data():
            dash_url = "https://docs.google.com/spreadsheets/d/1Roc_HIQLxsqRoxwCZVEkRgnQ0koe8A7wJOdJBWPWVas"
            dash_sheet = connect_sheet_by_url(dash_url, "Dashboard")
            main_sheet = connect_sheet_by_url(dash_url, "Main Service Sheet")
            stats = dash_sheet.get("F2:G8")
            raw_data = main_sheet.get_all_values()
            return stats, main_sheet, raw_data

        stats_data, call_sheet, raw_data = load_service_calling_data()
        df_raw = pd.DataFrame(raw_data[1:], columns=raw_data[0])
        df = df_raw.loc[:, ~df_raw.columns.duplicated()].copy() 
        df["__row"] = df.index + 2
        
        target_col = "Sold by Other Dealer" 
        mask = (df["Status"].str.lower().str.contains("pending", na=False)) & (df[target_col].fillna("").str.strip() == "")
        df_filtered = df[mask]

        t1, t2 = st.tabs(["🏙️ INDORE CALLING", "🌍 OUTSTATION CALLING"])

        def render_dashboard(df_part, mode):
            mode_idx = 0 if mode == "Indore" else 1
            total_pending = stats_data[6][mode_idx] if len(stats_data) > 6 else "0"
            st.markdown(f"<div style='background-color:#ff8c00; padding:15px; border-radius:15px; text-align:center; margin-bottom:20px;'><h1 style='color:white; margin:0;'>{total_pending}</h1><p style='color:white; margin:0;'>{mode.upper()} PENDING</p></div>", unsafe_allow_html=True)

            cols = st.columns(6)
            svc_key = f"svc_sel_{mode}"
            if svc_key not in st.session_state: st.session_state[svc_key] = "All"
            labels = ["1st", "2nd", "3rd", "4th", "5th", "6th"]
            for i, label in enumerate(labels, 1):
                count_val = stats_data[i-1][mode_idx] if len(stats_data) > i-1 else "0"
                if cols[i-1].button(f"{label} Service\n{count_val}", key=f"btn_{mode}_{i}"):
                    st.session_state[svc_key] = str(i)
                    st.rerun()

            current = st.session_state[svc_key]
            final_df = df_part if current == "All" else df_part[df_part["Service Count"].astype(str).str.startswith(current)]
            
            edited = st.data_editor(final_df, use_container_width=True, height=500, disabled=[c for c in final_df.columns if c not in ["Remark"]], key=f"ed_{mode}_{current}")
            for r in range(len(final_df)):
                if edited.iloc[r]["Remark"] != final_df.iloc[r]["Remark"]:
                    row_to_update = int(edited.iloc[r]["__row"])
                    call_sheet.update(f"T{row_to_update}", [[edited.iloc[r]["Remark"]]])
                    st.toast(f"Saved Row {row_to_update} ✅")

        with t1: render_dashboard(df_filtered[df_filtered["City"].str.lower() == "indore"], "Indore")
        with t2: render_dashboard(df_filtered[df_filtered["City"].str.lower() != "indore"], "Outstation")
    except Exception as e: st.error(f"🔴 System Error: {str(e)}")
# ------------------------------------------
# TAB 4: KPI DASHBOARD
# ------------------------------------------
with tab_kpi:
    try:
        KPI_SHEET_URL = "https://docs.google.com/spreadsheets/d/1zfAgfW4VbpLMJhHbRckRrK4PokA4jEYSfgxr6R800Fc/edit"
        kpi_sheet = connect_sheet_by_url(KPI_SHEET_URL, "Score Board")
        kpi_data = kpi_sheet.get_all_values()

        if len(kpi_data) > 2:
            def safe_float(val):
                try: 
                    v = str(val).replace('%', '').replace(',', '').replace('₹', '').strip()
                    return float(v) if v else 0.0
                except: return 0.0

            main_scores = [safe_float(kpi_data[i][5]) for i in range(2, 11) if i < len(kpi_data)]
            neg_scores = [safe_float(kpi_data[i][5]) for i in [13, 14] if i < len(kpi_data)]
            total_sum = sum(main_scores) + sum(neg_scores)

            st.markdown(f"<div style='background-color:#0056b3; padding:15px; border-radius:10px; text-align:center; color:white;'><h3>🏆 KPI SCORE: {total_sum:.1f} / 100</h3></div>", unsafe_allow_html=True)
            
            st.subheader("Score Board Detail")
            main_html = """<table style="width:100%; border-collapse: collapse; color: black; margin-top:10px;">
                <tr style="background-color: #0056b3; color: white;"><th>S/N</th><th>Parameter</th><th>Target</th><th>Actual</th><th>Score</th></tr>"""
            for r in kpi_data[2:11]:
                main_html += f'<tr><td>{r[0]}</td><td>{r[2]}</td><td>{r[3]}</td><td>{r[4]}</td><td>{r[5]}</td></tr>'
            st.markdown(main_html + "</table>", unsafe_allow_html=True)
    except Exception as e: st.error(f"🔴 KPI Error: {e}")

# ------------------------------------------
# TAB 5: PARTS
# ------------------------------------------
with tab_parts:
    try:
        p_sheet = connect_sheet_by_url(KPI_SHEET_URL, "Target VS Achivement")
        p_all = p_sheet.get_all_values()
        
        c1, c2, c3 = st.columns(3)
        c1.metric("Wholesale Achievement", p_all[8][24])
        c2.metric("Inventory Value", f"₹ {p_all[27][25]}")
        c3.metric("Gap Status", p_all[20][22])

        st.markdown("### Detailed Inventory Table")
        st.table(pd.DataFrame(p_all[3:22], columns=p_all[2]).iloc[1:])
    except Exception as e: st.error(f"🔴 Parts Error: {e}")
