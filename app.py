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
# 1. MAIN SHEET CONNECTION (WORKSHOP + APPOINTMENT)
# ==========================================
def connect_sheet(sheet_name):
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    try:
        # MAGIC FIX: Secrets ko dict mein convert kiya taaki error na aaye
        creds_dict = dict(st.secrets["gcp_service_account"])
        if "private_key" in creds_dict:
            creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")

        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
        client = gspread.authorize(creds)

        return client.open("Bike Check-In (Responses)").worksheet(sheet_name)

    except Exception as e:
        st.error(f"🔴 Connection Error: {sheet_name} | {e}")
        return None

# ==========================================
# SERVICE CALLING SHEET
# ==========================================
def connect_sheet_by_url(sheet_url, sheet_name):
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    try:
        creds_dict = dict(st.secrets["gcp_service_account"])
        if "private_key" in creds_dict:
            creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")

        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
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

st.markdown("""
    <div style="display: flex; align-items: center; gap: 15px; margin-bottom: 20px;">
        <img src="https://upload.wikimedia.org/wikipedia/commons/4/44/BMW.svg" width="55">
        <h1 style="margin: 0; color: #111;">Munich Motorrad Workshop Management</h1>
    </div>
""", unsafe_allow_html=True)

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
                all_items.append({
                    'row_idx': i + 2, 
                    'data': r
                })

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
                        display_stat = f"🔴 {stat}"
                        title_color = "#FF4B4B"
                    elif stat == "Complete":
                        display_stat = f"🟢 {stat}"
                        title_color = "#28a745"
                    else:
                        display_stat = f"🟠 {stat}"
                        title_color = "#FF8C00"

                    if s_o.lower() in str(r).lower():
                        with st.expander(f"{r[1]} | {r[5]} | {display_stat}"):
                            st.markdown(f"<h2 style='color:{title_color}; margin-top:0;'>{display_stat}</h2>", unsafe_allow_html=True)
                            
                            c1, c2, c3 = st.columns(3)
                            with c1:
                                st.markdown(f"**👤 Customer:**\n{r[1]}")
                                st.markdown("---")
                                st.markdown(f"**📱 Phone:**\n{r[2]}")
                                st.markdown("---")
                                st.markdown(f"**🆔 VIN:**\n{r[8]}")
                            with c2:
                                st.markdown(f"**🛠️ Staff:**\n{r[9]}")
                                st.markdown("---")
                                st.markdown(f"**🛣️ Odometer (KM):**\n{r[7] if len(r)>7 else 'N/A'}")
                                st.markdown("---")
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
                            with hc1:
                                st.markdown(f"**👤 Customer:** {r[1]}")
                                st.markdown(f"**📱 Phone:** {r[2] if len(r)>2 else 'N/A'}")
                                st.markdown(f"**📧 Email:** {r[3] if len(r)>3 else 'N/A'}")
                                st.markdown(f"**📅 Date:** {r[0]}")
                            with hc2:
                                st.markdown(f"**🆔 VIN:** {r[8] if len(r)>8 else 'N/A'}")
                                st.markdown(f"**🛠️ Staff:** {r[9] if len(r)>9 else 'N/A'}")
                                st.markdown(f"**🛣️ Odo:** {r[7] if len(r)>7 else 'N/A'} KM")
                            with hc3:
                                st.warning(f"**🚨 Issues:**\n{r[15] if len(r)>15 else 'N/A'}")
                                st.success(f"**✍️ Final Remark:**\n{r[17] if len(r)>17 else 'N/A'}")
                                st.info(f"**Status:** Delivered")

    except Exception as e:
        st.error(f"Workshop Load Error: {e}")
# ------------------------------------------
# TAB 2: APPOINTMENT CALENDAR (NO HIGHLIGHT)
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
            pending_apps = len(month_df[
                (month_df['D'].dt.date < today) & 
                (~month_df['R'].str.contains("Bike Reported", na=False))
            ])

            st.markdown("---")
            s1, s2, s3 = st.columns(3)
            
            s1.markdown(f"""<div style="background:#f0f2f6; padding:15px; border-radius:10px; text-align:center; border-left: 5px solid #0056b3;">
                <p style="margin:0; font-size:14px; color:#555;">Total Appointments</p>
                <h3 style="margin:0; color:#0056b3;">{total_apps}</h3>
            </div>""", unsafe_allow_html=True)
            
            s2.markdown(f"""<div style="background:#e7f3ef; padding:15px; border-radius:10px; text-align:center; border-left: 5px solid #28a745;">
                <p style="margin:0; font-size:14px; color:#555;">✅ Reported (Done)</p>
                <h3 style="margin:0; color:#28a745;">{reported_apps}</h3>
            </div>""", unsafe_allow_html=True)
            
            p_color = "#d9534f" if pending_apps > 0 else "#ff8c00"
            s3.markdown(f"""<div style="background:#fff4e6; padding:15px; border-radius:10px; text-align:center; border-left: 5px solid {p_color};">
                <p style="margin:0; font-size:14px; color:#555;">⚠️ Missed/Pending</p>
                <h3 style="margin:0; color:{p_color};">{pending_apps}</h3>
            </div>""", unsafe_allow_html=True)
            
            st.markdown("---")

            st.markdown("""<style>
                [data-testid="stVerticalBlockBorderWrapper"] > div > div { min-height: 250px !important; max-height: 250px !important; overflow-y: auto !important; }
                div.stButton > button p { white-space: normal !important; word-wrap: break-word !important; font-size: 11px !important; font-weight: bold !important; }
            </style>""", unsafe_allow_html=True)

            cal = calendar.monthcalendar(y, midx)
            cols_header = st.columns(7)
            for j, d in enumerate(["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]): 
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
                                
                                if is_rep:
                                    b_c = "#28a745"
                                elif row['D'].date() < today:
                                    b_c = "#d9534f"
                                else:
                                    b_c = "#007bff"
                                
                                st.markdown(f"""<style>div.stButton > button[key="trig_{row['row']}"] {{ border: 2.5px solid {b_c} !important; min-height: 60px !important; margin-bottom: 5px !important; }}</style>""", unsafe_allow_html=True)
                                
                                if st.button(f"{'✅' if is_rep else '⭕'} {row['N']}", key=f"trig_{row['row']}", use_container_width=True):
                                    app_sheet.update_cell(row['row'], 8, "Bike Reported" if not is_rep else "")
                                    st.rerun()
                    else:
                        cols[i].markdown("<div style='height:100px; opacity:0.1;'></div>", unsafe_allow_html=True)

    except Exception as e: 
        st.error(f"Calendar Error: {e}")

# ------------------------------------------
# TAB 3: SERVICE CALLING (FIXED UI & TEXT)
# ------------------------------------------
with tab_call:
    try:
        dash_url = "https://docs.google.com/spreadsheets/d/1Roc_HIQLxsqRoxwCZVEkRgnQ0koe8A7wJOdJBWPWVas"
        
        @st.cache_data(ttl=60)
        def load_service_calling_data():
            dash_sheet = connect_sheet_by_url(dash_url, "Dashboard")
            main_sheet = connect_sheet_by_url(dash_url, "Main Service Sheet")
            stats = dash_sheet.get("F2:G8")
            raw_data = main_sheet.get_all_values()
            # FIX: Sirf data return kar rahe hain, connection object nahi
            return stats, raw_data

        stats_data, raw_data = load_service_calling_data()
        
        # FIX: Connection object ko cache ke bahar rakha hai taaki wo expire na ho
        call_sheet = connect_sheet_by_url(dash_url, "Main Service Sheet")

        df_raw = pd.DataFrame(raw_data[1:], columns=raw_data[0])
        df = df_raw.loc[:, ~df_raw.columns.duplicated()].copy() 
        df["__row"] = df.index + 2
        
        target_col = "Sold by Other Dealer" 
        mask = (df["Status"].str.lower().str.contains("pending", na=False)) & \
               (df[target_col].fillna("").str.strip() == "")
        df_filtered = df[mask]

        st.markdown("""<style>
            div.stButton > button { 
                width: 100%; height: 110px !important; border-radius: 15px; 
                background-color: white; border: 1px solid #ddd; transition: 0.3s;
                white-space: pre-wrap !important;
            }
            div.stButton > button:hover { border: 2px solid #ff8c00 !important; background-color: #fffaf5 !important; }
            div.stButton > button p {
                font-size: 16px !important; font-weight: 800 !important;
                line-height: 1.3 !important; color: #333;
            }
        </style>""", unsafe_allow_html=True)

        t1, t2 = st.tabs(["🏙️ INDORE CALLING", "🌍 OUTSTATION CALLING"])

        def render_dashboard(df_part, mode):
            mode_idx = 0 if mode == "Indore" else 1
            
            total_pending = stats_data[6][mode_idx] if len(stats_data) > 6 else "0"
            st.markdown(f"""
                <div style="background-color:#ff8c00; padding:15px; border-radius:15px; text-align:center; margin-bottom:20px; border: 2px solid #fff; box-shadow: 0 4px 10px rgba(0,0,0,0.1);">
                    <p style="color:white; margin:0; font-size:14px; font-weight:bold;">{mode.upper()} TOTAL PENDING</p>
                    <h1 style="color:white; margin:0; font-size:40px;">{total_pending}</h1>
                </div>
            """, unsafe_allow_html=True)

            cols = st.columns(6)
            svc_key = f"svc_sel_{mode}"
            if svc_key not in st.session_state: st.session_state[svc_key] = "All"

            labels = ["1st", "2nd", "3rd", "4th", "5th", "6th"]
            for i, label in enumerate(labels, 1):
                count_val = stats_data[i-1][mode_idx] if len(stats_data) > i-1 else "0"
                is_sel = st.session_state[svc_key] == str(i)
                display_text = f"{label} Service\nPending: {count_val}"
                
                if is_sel:
                    st.markdown(f"<style>div.stButton > button[key*='btn_{mode}_{i}'] {{ background-color: #1E7E34 !important; border: 2px solid #FFD700 !important; }} div.stButton > button[key*='btn_{mode}_{i}'] p {{ color: white !important; }}</style>", unsafe_allow_html=True)
                
                if cols[i-1].button(display_text, key=f"btn_{mode}_{i}"):
                    st.session_state[svc_key] = "All" if is_sel else str(i)
                    st.rerun()

            current = st.session_state[svc_key]
            final_df = df_part if current == "All" else df_part[df_part["Service Count"].astype(str).str.startswith(current)]

            if final_df.empty:
                st.info(f"No {current if current != 'All' else ''} pending calls.")
            else:
                st.write(f"📋 Showing **{len(final_df)}** records")
                edited = st.data_editor(
                    final_df, 
                    use_container_width=True, height=500,
                    disabled=[c for c in final_df.columns if c not in ["Remark"]],
                    column_config={"__row": None, "Remark": st.column_config.TextColumn(width="large")},
                    key=f"ed_{mode}_{current}"
                )

                for r in range(len(final_df)):
                    if edited.iloc[r]["Remark"] != final_df.iloc[r]["Remark"]:
                        row_to_update = int(edited.iloc[r]["__row"])
                        call_sheet.update(f"T{row_to_update}", [[edited.iloc[r]["Remark"]]])
                        st.toast(f"Saved Row {row_to_update} ✅")

        with t1: render_dashboard(df_filtered[df_filtered["City"].str.lower() == "indore"], "Indore")
        with t2: render_dashboard(df_filtered[df_filtered["City"].str.lower() != "indore"], "Outstation")

    except Exception as e:
        st.error(f"🔴 System Error: {str(e)}")

# ------------------------------------------
# TAB 4: KPI DASHBOARD (COMPLETE FIXED CODE)
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

            st.markdown(f"""
                <div style="background-color:#0056b3; padding:5px 10px; border-radius:10px; text-align:center; border: 2px solid #FFD700; margin-bottom:10px;">
                    <p style="color:white; margin:0; font-size:14px; font-weight:bold;">🏆 TOTAL KPI SCORE (Q2)</p>
                    <h1 style="color:#FFD700; font-size:38px; margin:-5px 0; font-weight:bold;">{total_sum:.1f} <span style="font-size:16px; color:white;">/ 100</span></h1>
                </div>
            """, unsafe_allow_html=True)

            s1, s2 = st.tabs(["🎯 SCORE BOARD", "📊 QUARTERLY STATS"])

            with s1:
                st.subheader("Quarter 2: Achievement Score")
                main_html = """<table style="width:100%; border-collapse: collapse; color: black; margin-bottom:20px;">
                    <tr style="background-color: #0056b3; color: white; font-weight: bold;">
                        <th style="padding: 8px; border: 1px solid #ddd;">S/N</th>
                        <th style="padding: 8px; border: 1px solid #ddd;">Parameters</th>
                        <th style="padding: 8px; border: 1px solid #ddd;">Target</th>
                        <th style="padding: 8px; border: 1px solid #ddd;">Actual</th>
                        <th style="padding: 8px; border: 1px solid #ddd;">Score Achieved</th>
                        <th style="padding: 8px; border: 1px solid #ddd;">Scoring Criteria</th>
                        <th style="padding: 8px; border: 1px solid #ddd;">Score</th>
                        <th style="padding: 8px; border: 1px solid #ddd;">Planning</th>
                    </tr>"""
                for r in kpi_data[2:11]:
                    if len(r) > 8 and r[0].strip():
                        main_html += f'<tr style="background-color: white;">' \
                                     f'<td style="padding:8px; border:1px solid #ddd;">{r[0]}</td>' \
                                     f'<td style="padding:8px; border:1px solid #ddd;">{r[2]}</td>' \
                                     f'<td style="padding:8px; border:1px solid #ddd;">{r[3]}</td>' \
                                     f'<td style="padding:8px; border:1px solid #ddd;">{r[4]}</td>' \
                                     f'<td style="padding:8px; border:1px solid #ddd;">{r[5]}</td>' \
                                     f'<td style="padding:8px; border:1px solid #ddd;">{r[6]}</td>' \
                                     f'<td style="padding:8px; border:1px solid #ddd;">{r[7]}</td>' \
                                     f'<td style="padding:8px; border:1px solid #ddd; font-weight:bold;">{r[8]}</td></tr>'
                main_html += "</table>"
                st.markdown(main_html, unsafe_allow_html=True)

                st.markdown("---")
                st.markdown("<h3 style='color:#FF4B4B;'>⚠️ Negative Marking</h3>", unsafe_allow_html=True)
                neg_html = """<table style="width:100%; border-collapse: collapse; color: black;">
                    <tr style="background-color: #0056b3; color: white; font-weight: bold;">
                        <th style="padding: 8px; border: 1px solid #ddd;">S/N</th>
                        <th style="padding: 8px; border: 1px solid #ddd;">Parameters</th>
                        <th style="padding: 8px; border: 1px solid #ddd;">Target</th>
                        <th style="padding: 8px; border: 1px solid #ddd;">Actual</th>
                        <th style="padding: 8px; border: 1px solid #ddd;">Score Achieved</th>
                        <th style="padding: 8px; border: 1px solid #ddd;">Scoring Criteria</th>
                        <th style="padding: 8px; border: 1px solid #ddd;">Score</th>
                        <th style="padding: 8px; border: 1px solid #ddd;">Planning</th>
                    </tr>"""
                for i in [13, 14]:
                    if i < len(kpi_data):
                        r = kpi_data[i]
                        neg_html += f'<tr style="background-color: white;">' \
                                     f'<td style="padding:8px; border:1px solid #ddd;">{r[0]}</td>' \
                                     f'<td style="padding:8px; border:1px solid #ddd;">{r[2]}</td>' \
                                     f'<td style="padding:8px; border:1px solid #ddd;">{r[3]}</td>' \
                                     f'<td style="padding:8px; border:1px solid #ddd;">{r[4]}</td>' \
                                     f'<td style="padding:8px; border:1px solid #ddd;">{r[5]}</td>' \
                                     f'<td style="padding:8px; border:1px solid #ddd;">{r[6]}</td>' \
                                     f'<td style="padding:8px; border:1px solid #ddd;">{r[7]}</td>' \
                                     f'<td style="padding:8px; border:1px solid #ddd; font-weight:bold;">{r[8]}</td></tr>'
                neg_html += "</table>"
                st.markdown(neg_html, unsafe_allow_html=True)

            with s2:
                st.subheader("📊 Q2 STATS")
                stats_html = """<table style="width:100%; border-collapse: collapse; color: black;">
                    <tr style="background-color: #0056b3; color: white; font-weight: bold;">
                        <th style="padding: 8px; border: 1px solid #ddd;">Parameters</th>
                        <th style="padding: 8px; border: 1px solid #ddd;">Total</th>
                        <th style="padding: 8px; border: 1px solid #ddd;">Apr</th>
                        <th style="padding: 8px; border: 1px solid #ddd;">May</th>
                        <th style="padding: 8px; border: 1px solid #ddd;">Jun</th>
                    </tr>"""
                for i in range(1, 9):
                    try:
                        r = kpi_data[i]
                        stats_html += f'<tr>' \
                                      f'<td style="padding:8px; border:1px solid #ddd;">{r[10]}</td>' \
                                      f'<td style="padding:8px; border:1px solid #ddd;">{r[11]}</td>' \
                                      f'<td style="padding:8px; border:1px solid #ddd;">{r[15]}</td>' \
                                      f'<td style="padding:8px; border:1px solid #ddd;">{r[16]}</td>' \
                                      f'<td style="padding:8px; border:1px solid #ddd;">{r[17]}</td>' \
                                      f'</tr>'
                    except: pass
                stats_html += "</table>"
                st.markdown(stats_html, unsafe_allow_html=True)

                st.markdown("---")
                st.markdown("### 🏍️ Munich Motorrad Indore JC 25 vs 26")

                try:
                    jc_needed = kpi_data[1][25]
                    st.markdown(f'<div style="background-color:#fff3cd; padding:8px 15px; border-radius:8px; border: 1px solid #ffa000; display: inline-block; margin-bottom:15px;"><span style="color:#856404; font-weight:bold;">🎯 JC NEEDED TO CLOSE: </span><span style="color:#000; font-size:18px; font-weight:bold;">{jc_needed}</span></div>', unsafe_allow_html=True)
                except: pass

                growth_rows = [kpi_data[i] for i in range(2, 16) if i < len(kpi_data) and kpi_data[i][21].strip()]
                
                growth_html = """<table style="width:100%; border-collapse: collapse; color: black;">
                    <tr style="background-color: #0056b3; color: white;">
                        <th style="padding: 10px; border: 1px solid #ddd;">Month</th>
                        <th style="padding: 10px; border: 1px solid #ddd;">JC - 2025</th>
                        <th style="padding: 10px; border: 1px solid #ddd;">JC - 2026</th>
                        <th style="padding: 10px; border: 1px solid #ddd;">Growth %</th>
                    </tr>"""
                
                for idx, r in enumerate(growth_rows):
                    is_last = (idx == len(growth_rows) - 1)
                    bg = "#FFD700" if is_last else "white"
                    weight = "bold" if is_last else "normal"
                    growth_html += f'<tr style="background-color:{bg}; font-weight:{weight};">' \
                                   f'<td style="padding:8px; border:1px solid #ddd;">{r[21]}</td>' \
                                   f'<td style="padding:8px; border:1px solid #ddd;">{r[22]}</td>' \
                                   f'<td style="padding:8px; border:1px solid #ddd;">{r[23]}</td>' \
                                   f'<td style="padding:8px; border:1px solid #ddd;">{r[24]}</td></tr>'
                growth_html += "</table>"
                st.markdown(growth_html, unsafe_allow_html=True)

    except Exception as e:
        st.error(f"🔴 KPI Error: {e}")
# ==========================================
# MAIN TAB 5: PARTS (STABLE VERSION)
# ==========================================
with tab_parts:
    # 1. SHARED FUNCTIONS
    def get_clean_range(values, r1, r2, c1, c2):
        try:
            sc, ec = ord(c1.upper())-65, ord(c2.upper())-64
            data = [r[sc:ec] for r in values[r1-1:r2]]
            if not data or len(data) < 1: return pd.DataFrame()
            raw_h = data[0]
            new_h = []
            for i, h in enumerate(raw_h):
                val = str(h).strip()
                if val == "" or val == "None":
                    if i == 0: val = "Month"
                    elif i < 6: val = f"P_Col_{i}"
                    elif i < 11: val = f"G_Col_{i}"
                    else: val = f"T_Col_{i}"
                if val == "Wty": val = "Wty (P)" if i < 8 else "Wty (G)"
                if val == "Achv. W/O Wty": val = "Achv(P)" if i < 8 else "Achv(G)"
                if val in new_h: val = f"{val}_{i}"
                new_h.append(val)
            return pd.DataFrame(data[1:], columns=new_h)
        except: return pd.DataFrame()

    def get_val(values, r, c):
        try:
            col_idx = ord(c.upper()) - 65
            val = values[r-1][col_idx]
            return val if val else "0"
        except: return "0"

    try:
        p_sheet = connect_sheet_by_url(KPI_SHEET_URL, "Target VS Achivement")
        p_all = p_sheet.get_all_values()

        # --- CSS FOR BOX TILES & TABLES ---
        st.markdown("""
            <style>
                .metric-container {
                    background-color: #ffffff; padding: 15px; border-radius: 10px;
                    border-left: 5px solid #44546a; box-shadow: 2px 2px 10px rgba(0,0,0,0.1);
                    text-align: center; margin-bottom: 10px;
                }
                .metric-label { font-size: 12px; color: #555; font-weight: bold; margin-bottom: 5px; }
                .metric-value { font-size: 18px; color: #111; font-weight: bold; }
                .parts-box { width: 98% !important; margin: auto; }
                .stTable { border: 1px solid #333 !important; }
                .stTable td, .stTable th { border: 1px solid #333 !important; font-size: 10px !important; padding: 2px !important; }
                
                /* FIX: .stTable lagaya hai taaki Tab 4 mein color na faile */
                .stTable th:nth-child(1) { background-color: #f2f2f2 !important; color: black !important; }
                .stTable th:nth-child(n+2):nth-child(-n+6) { background-color: #ffff00 !important; color: black !important; }
                .stTable th:nth-child(n+7):nth-child(-n+11) { background-color: #fce4d6 !important; color: black !important; }
                .stTable th:nth-child(n+12):nth-child(-n+15) { background-color: #ed7d31 !important; color: white !important; }
                .stTable th:nth-child(n+16) { background-color: #92d050 !important; color: black !important; }
                .gap-table th:last-child { background-color: #fce4d6 !important; }
            </style>
        """, unsafe_allow_html=True)

        # --- TOP TILES ---
        t1, t2, t3 = st.columns(3)
        with t1:
            st.markdown(f'<div class="metric-container"><div class="metric-label">YTD WHOLESALE / RETAIL</div><div class="metric-value">{get_val(p_all, 9, "Y")} / {get_val(p_all, 9, "Z")}</div></div>', unsafe_allow_html=True)
        with t2:
            st.markdown(f'<div class="metric-container"><div class="metric-label">INVENTORY VALUE</div><div class="metric-value">₹ {get_val(p_all, 28, "Z")}</div></div>', unsafe_allow_html=True)
        with t3:
            st.markdown(f'<div class="metric-container"><div class="metric-label">YTD GAPS (PART | ACC | TOTAL)</div><div class="metric-value">{get_val(p_all, 21, "W")} | {get_val(p_all, 21, "Z")} | {get_val(p_all, 16, "W")}</div></div>', unsafe_allow_html=True)

        # --- SUB TABS ---
        p_sub1, p_sub2, p_sub3 = st.tabs(["📊 Purchase & Sale", "🔄 Dealer Transfer", "📦 Inventory"])

        with p_sub1:
            def excel_style(row):
                m = str(row.iloc[0]).upper().strip()
                # Sabhi returns mein 'text-align: center' add kiya hai
                if any(x in m for x in ["QTR", "QR"]): 
                    return ['background-color: #bfbfbf; color: black; font-weight: bold; text-align: center'] * len(row)
                elif any(x in m for x in ["H1", "H2", "TOTAL"]): 
                    return ['background-color: #44546a; color: white; font-weight: bold; text-align: center'] * len(row)
                else:
                    res = ['background-color: #f2f2f2; font-weight: bold; text-align: center']
                    res += ['background-color: #deeaf6; color: black; text-align: center'] * (len(row)-1)
                    return res

            # Header style variable mein center alignment add kiya
            h_style = [{'selector': 'th', 'props': [('font-weight', 'bold'), ('color', 'black'), ('border', '1px solid #333'), ('text-align', 'center')]}]

            st.markdown('<div class="parts-box">', unsafe_allow_html=True)
            
            # 1. Wholesale Achievement
            df_wh = get_clean_range(p_all, 3, 22, 'A', 'R')
            if not df_wh.empty:
                st.markdown("##### 🔹 Wholesale Achievement")
                st.table(df_wh.style.apply(excel_style, axis=1).set_table_styles(h_style))

            # 2. Retail Achievement
            df_rt = get_clean_range(p_all, 27, 46, 'A', 'R')
            if not df_rt.empty:
                st.markdown("---")
                st.markdown("##### 🔸 Retail Achievement")
                st.table(df_rt.style.apply(excel_style, axis=1).set_table_styles(h_style))

            # 3. Summary Gaps
            st.markdown("#### 📊 Summary Gaps")
            cs1, cs2 = st.columns(2)
            
            with cs1:
                df1 = get_clean_range(p_all, 13, 16, 'T', 'W')
                if not df1.empty:
                    df1 = df1.set_index(df1.columns[0])
                    df1.index.name = None
                    st.markdown('<div class="gap-table">', unsafe_allow_html=True)
                    # Gap tables ke data ko bhi center karne ke liye .set_properties use kiya
                    st.table(df1.style.set_table_styles(h_style).set_properties(**{'text-align': 'center'}))
                    st.markdown('</div>', unsafe_allow_html=True)
            
            with cs2:
                df2 = get_clean_range(p_all, 18, 21, 'T', 'Z')
                if not df2.empty:
                    df2 = df2.set_index(df2.columns[0])
                    df2.index.name = None
                    st.markdown('<div class="gap-table">', unsafe_allow_html=True)
                    st.table(df2.style.set_table_styles(h_style).set_properties(**{'text-align': 'center'}))
                    st.markdown('</div>', unsafe_allow_html=True)
            
            st.markdown('</div>', unsafe_allow_html=True)

        with p_sub2:
            try:
                st.markdown("### **🔄 Dealer Transfer Analysis**")
                
                # 1. Summary Table (Range: T8:W11)
                st.markdown("#### **📈 Dealer Transfer Summary (H1/H2)**")
                df_dt_sum = get_clean_range(p_all, 8, 11, 'T', 'W')
                
                if not df_dt_sum.empty:
                    # HEADING RENAME: P_Col_3 ko 'Total Amount' karna
                    df_dt_sum.columns = ["Month", "Part Dealer Transfer", "Acc Dealer Transfer", "Total Amount"]
                    
                    # STYLING: Bold Heading + Yellow Background
                    st.table(df_dt_sum.style.set_table_styles([
                        {'selector': 'th', 'props': [('background-color', '#ffff00'), ('color', 'black'), ('font-weight', 'bold'), ('border', '1px solid #333')]}
                    ]).set_properties(**{
                        'background-color': '#f2f2f2', 
                        'border': '1px solid #333',
                        'color': 'black',
                        'font-weight': 'bold'
                    }))
                
                st.markdown("---")
                
                # 2. Raw Data (Dealer Transfer Sheet)
                dt_sheet = connect_sheet_by_url(KPI_SHEET_URL, "Dealer Transfer")
                dt_all = dt_sheet.get_all_values()
                
                if len(dt_all) > 1:
                    raw_df = pd.DataFrame(dt_all[1:], columns=dt_all[0])
                    df_dt = raw_df.iloc[:, [2, 4, 16]].copy()
                    df_dt.columns = ['Customer', 'Date', 'Amount']
                    
                    df_dt['Amount'] = pd.to_numeric(df_dt['Amount'].str.replace(',', ''), errors='coerce').fillna(0)
                    df_dt['Date'] = pd.to_datetime(df_dt['Date'], errors='coerce')
                    df_dt['Month'] = df_dt['Date'].dt.strftime('%B %Y')
                    
                    # Highlights
                    t_amt = df_dt['Amount'].sum()
                    u_cust = df_dt['Customer'].nunique()
                    st.markdown(f"**Total Transfer Amount:** ₹ {t_amt:,.2f} | **Customers:** {u_cust}")

                    # Monthly Table
                    st.markdown("#### **🗓️ Month-wise Customer Transfer Totals**")
                    summary_dt = df_dt.groupby(['Month', 'Customer'])['Amount'].sum().reset_index()
                    summary_dt['SortOrder'] = pd.to_datetime(summary_dt['Month'], format='%B %Y')
                    
                    # Sort karne ke baad index reset kiya taaki S.No. sahi aaye (1, 2, 3...)
                    summary_dt = summary_dt.sort_values(['SortOrder', 'Amount'], ascending=[True, False]).drop('SortOrder', axis=1)
                    summary_dt = summary_dt.reset_index(drop=True)
                    summary_dt.index = summary_dt.index + 1 
                    
                    summary_dt['Amount'] = summary_dt['Amount'].apply(lambda x: f"₹ {x:,.2f}")
                    
                    # Is table ki heading bhi Bold aur Yellow
                    st.table(summary_dt.style.set_table_styles([
                        {'selector': 'th', 'props': [('background-color', '#ffff00'), ('color', 'black'), ('font-weight', 'bold')]}
                    ]).set_properties(**{
                        'background-color': '#deeaf6',
                        'color': 'black',
                        'border': '1px solid #333'
                    }))
                    
                else:
                    st.warning("Dealer Transfer sheet khali mili.")

            except Exception as e_dt:
                st.error(f"🔴 Dealer Transfer Error: {e_dt}")

        with p_sub3:
            try:
                st.markdown("### **📦 Inventory Analysis**")
                
                # --- DATA FETCHING ---
                # Z27 = Last Check Date (Row 27, Col Z/25)
                # Z28 = Total Inventory Value (Row 28, Col Z/25)
                check_date = p_all[26][25] if len(p_all) > 26 else "N/A"
                
                raw_total = p_all[27][25] if len(p_all) > 27 else "0"
                # Numeric conversion for color logic
                try:
                    num_total = float(str(raw_total).replace(',', '').replace('₹', '').strip())
                except:
                    num_total = 0
                
                # Color Logic: 35L se kam toh Green, zyada toh Red
                text_color = "#2e7d32" if num_total <= 3500000 else "#d32f2f"
                
                # Top Info Bar for Date
                st.info(f"📅 **Last Inventory Check Date:** {check_date}")
                
                # Data rows (X28:AB32) fetch karna
                rows = []
                for i in range(27, 32): 
                    rows.append(p_all[i][23:28])
                
                if rows:
                    html_table = f"""
                    <style>
                        .inv-table {{ width: 100%; border-collapse: collapse; font-family: sans-serif; border: 2px solid #333; }}
                        .inv-table th {{ background-color: #ffff00; color: black; font-weight: bold; padding: 10px; border: 1px solid #333; text-align: center; }}
                        .inv-table td {{ padding: 8px; border: 1px solid #333; text-align: center; font-weight: bold; background-color: #f2f2f2; }}
                        .merged-cell {{ 
                            background-color: #ffffff; 
                            font-size: 24px; 
                            color: {text_color};
                            vertical-align: middle; 
                            border: 1px solid #333; 
                            text-align: center;
                            font-weight: 900;
                        }}
                    </style>
                    <table class="inv-table">
                        <thead>
                            <tr>
                                <th>Category</th>
                                <th>Stock Value</th>
                                <th>Inventory Value (Total)</th>
                                <th>Dead Stock</th>
                                <th>Target</th>
                            </tr>
                        </thead>
                        <tbody>
                            <tr>
                                <td>{rows[0][0]}</td><td>{rows[0][1]}</td>
                                <td rowspan="5" class="merged-cell">₹ {raw_total}</td>
                                <td>{rows[0][3]}</td><td>{rows[0][4]}</td>
                            </tr>
                            <tr><td>{rows[1][0]}</td><td>{rows[1][1]}</td><td>{rows[1][3]}</td><td>{rows[1][4]}</td></tr>
                            <tr><td>{rows[2][0]}</td><td>{rows[2][1]}</td><td>{rows[2][3]}</td><td>{rows[2][4]}</td></tr>
                            <tr><td>{rows[3][0]}</td><td>{rows[3][1]}</td><td>{rows[3][3]}</td><td>{rows[3][4]}</td></tr>
                            <tr><td>{rows[4][0]}</td><td>{rows[4][1]}</td><td>{rows[4][3]}</td><td>{rows[4][4]}</td></tr>
                        </tbody>
                    </table>
                    """
                    st.markdown(html_table, unsafe_allow_html=True)
                else:
                    st.error("Sheet se data rows nahi mil rahi hain.")

            except Exception as e_inv:
                st.error(f"🔴 Inventory Error: {e_inv}")

    except Exception as e:
        st.error(f"🔴 Main Error: {e}")
