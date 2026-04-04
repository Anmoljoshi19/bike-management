import streamlit as st
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd
from datetime import datetime
import calendar

# ==========================================
# 1. DATABASE CONNECTION (FIXED)
# ==========================================
def connect_sheet(sheet_name):
    scope = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/drive"
    ]
    try:
        creds = ServiceAccountCredentials.from_json_keyfile_dict(
            st.secrets["gcp_service_account"], scope
        )
        client = gspread.authorize(creds)

        return client.open("Bike Check-In (Responses)").worksheet(sheet_name)

    except Exception as e:
        st.error(f"🔴 Connection Failed: {str(e)}")
        return None


# Page Setup
st.set_page_config(page_title="Munich Motorrad Local", page_icon="🏍️", layout="wide")

# ==========================================
# 2. PRO CSS (OFFLINE OPTIMIZED)
# ==========================================
st.markdown("""
<style>

/* MOBILE CALENDAR FIX */
[data-testid="stVerticalBlockBorderWrapper"] > div > div {
    min-height: 180px !important;
    max-height: 220px !important;
    overflow-y: auto !important;
    padding: 5px !important;
}

/* Button text wrap (important for phone) */
div.stButton > button p {
    white-space: normal !important;
    word-wrap: break-word !important;
    font-size: 10px !important;
    font-weight: bold !important;
    line-height: 1.2 !important;
}

/* Button size mobile friendly */
div.stButton > button {
    padding: 8px !important;
    border-radius: 10px !important;
}

/* Calendar day box */
[data-testid="stVerticalBlock"] {
    gap: 4px !important;
}

/* Horizontal scroll for mobile */
section.main > div {
    overflow-x: auto;
}

/* Day title */
.calendar-day {
    font-size: 13px;
}

/* Improve tap feel */
button {
    touch-action: manipulation;
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
            # Header extract karein
            header = ws_data[0]
            
            # Saare rows ko original index ke saath list mein daalein
            all_items = []
            for i, r in enumerate(ws_data[1:]):
                all_items.append({
                    'row_idx': i + 2, 
                    'data': r
                })

            sub1, sub2 = st.tabs(["📂 Open JCs", "✅ History (Delivered)"])
            
            with sub1:
                s_o = st.text_input("🔍 Search Open JCs", key="ws_so")
                
                # Filter out Delivered bikes
                open_list = [item for item in all_items if len(item['data']) > 18 and item['data'][18] != "Delivered"]
                
                # --- ARRANGE LOGIC (HOLD -> PROCESS -> COMPLETE) ---
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
                    
                    # Status visuals (Emoji for Title)
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
                        # Display Card
                        with st.expander(f"{r[1]} | {r[5]} | {display_stat}"):
                            # Colored Heading inside
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
                # Filter Delivered bikes
                history_list = [item for item in all_items if len(item['data']) > 18 and item['data'][18] == "Delivered"]
                
                for item in history_list:
                    r = item['data']
                    if s_h.lower() in str(r).lower():
                        with st.expander(f"✅ {r[1]} | {r[5]} | Delivered"):
                            # Pura data History section mein (245 lines goal)
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
            y, mn = cy.selectbox("Year", [2025, 2026], 1), cm.selectbox("Month", list(calendar.month_name)[1:], datetime.now().month-1)
            midx = list(calendar.month_name).index(mn)

            st.markdown("""<style>[data-testid="stVerticalBlockBorderWrapper"] > div > div { min-height: 250px !important; max-height: 250px !important; overflow-y: auto !important; }
            div.stButton > button p { white-space: normal !important; word-wrap: break-word !important; font-size: 11px !important; font-weight: bold !important; }</style>""", unsafe_allow_html=True)

            cal = calendar.monthcalendar(y, midx)
            cols = st.columns(7)
            for j, d in enumerate(["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]): 
                cols[j].markdown(f"<p style='text-align:center;font-weight:bold;color:#0056b3;'>{d}</p>", unsafe_allow_html=True)

            for week in cal:
                cols = st.columns(7)
                for i, day in enumerate(week):
                    if day != 0:
                        with cols[i].container(border=True):
                            st.markdown(f"<p style='text-align:right; font-weight:bold; color:#555;'>{day}</p>", unsafe_allow_html=True)
                            day_data = df[(df['D'].dt.day == day) & (df['D'].dt.month == midx)]
                            for _, row in day_data.iterrows():
                                is_rep = "Bike Reported" in str(row['R'])
                                b_c = "#28a745" if is_rep else "#000"
                                st.markdown(f"""<style>div.stButton > button[key="trig_{row['row']}"] {{ border: 2.5px solid {b_c} !important; min-height: 60px !important; margin-bottom: 5px !important; }}</style>""", unsafe_allow_html=True)
                                if st.button(f"{'✅' if is_rep else '⭕'} {row['N']}", key=f"trig_{row['row']}", use_container_width=True):
                                    app_sheet.update_cell(row['row'], 8, "Bike Reported" if not is_rep else "")
                                    st.rerun()
                    else:
                        cols[i].markdown("<div style='height:100px; opacity:0.2;'></div>", unsafe_allow_html=True)
    except Exception as e: st.error(f"Calendar Error: {e}")

# ------------------------------------------
# TAB 3: SERVICE CALLING (FIXED UI & TEXT)
# ------------------------------------------
with tab_call:
    try:
        # 1. LOAD DATA (URL aur Connection Setup)
        @st.cache_data(ttl=60)
        def load_service_calling_data():
            # URL fix for Service Sheet
            dash_url = "https://docs.google.com/spreadsheets/d/1Roc_HIQLxsqRoxwCZVEkRgnQ0koe8A7wJOdJBWPWVas"
            dash_sheet = connect_sheet_by_url(dash_url, "Dashboard")
            main_sheet = connect_sheet_by_url(dash_url, "Main Service Sheet")
            
            if not dash_sheet or not main_sheet:
                return None, None, None
                
            stats = dash_sheet.get("F2:G8")
            raw_data = main_sheet.get_all_values()
            return stats, main_sheet, raw_data

        stats_data, call_sheet, raw_data = load_service_calling_data()

        if raw_data:
            # 2. DATA CLEANING
            df_raw = pd.DataFrame(raw_data[1:], columns=raw_data[0])
            df = df_raw.loc[:, ~df_raw.columns.duplicated()].copy() 
            df["__row"] = df.index + 2
            
            # 3. GLOBAL FILTERS (Pending aur Blank 'Sold by Other Dealer')
            target_col = "Sold by Other Dealer" 
            mask = (df["Status"].str.lower().str.contains("pending", na=False)) & \
                   (df[target_col].fillna("").str.strip() == "")
            df_filtered = df[mask]

            # 4. UI STYLING (Bada Box aur Bold Text)
            st.markdown("""<style>
                div.stButton > button { 
                    width: 100%; 
                    height: 110px !important; 
                    border-radius: 15px; 
                    background-color: white; 
                    border: 1px solid #ddd; 
                    transition: 0.3s;
                    white-space: pre-wrap !important;
                }
                div.stButton > button:hover { border: 2px solid #ff8c00 !important; background-color: #fffaf5 !important; }
                div.stButton > button p {
                    font-size: 16px !important;
                    font-weight: 800 !important;
                    line-height: 1.3 !important;
                    color: #333;
                }
            </style>""", unsafe_allow_html=True)

            t1, t2 = st.tabs(["🏙️ INDORE CALLING", "🌍 OUTSTATION CALLING"])

            # 5. RENDER ENGINE (Har city ke liye alag filter)
            def render_dashboard(df_part, mode):
                mode_idx = 0 if mode == "Indore" else 1
                
                # --- SUMMARY TILE ---
                total_pending = stats_data[6][mode_idx] if len(stats_data) > 6 else "0"
                st.markdown(f"""
                    <div style="background-color:#ff8c00; padding:15px; border-radius:15px; text-align:center; margin-bottom:20px; border: 2px solid #fff; box-shadow: 0 4px 10px rgba(0,0,0,0.1);">
                        <p style="color:white; margin:0; font-size:14px; font-weight:bold;">{mode.upper()} TOTAL PENDING</p>
                        <h1 style="color:white; margin:0; font-size:40px;">{total_pending}</h1>
                    </div>
                """, unsafe_allow_html=True)

                # --- SERVICE SELECTOR TILES ---
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

                # --- DATA TABLE ---
                current = st.session_state[svc_key]
                final_df = df_part if current == "All" else df_part[df_part["Service Count"].astype(str).str.startswith(current)]

                if final_df.empty:
                    st.info(f"No {current if current != 'All' else ''} pending calls.")
                else:
                    st.write(f"📋 Showing **{len(final_df)}** records")
                    edited = st.data_editor(
                        final_df, 
                        use_container_width=True, 
                        height=500,
                        disabled=[c for c in final_df.columns if c not in ["Remark"]],
                        column_config={"__row": None, "Remark": st.column_config.TextColumn(width="large")},
                        key=f"ed_{mode}_{current}"
                    )

                    # --- INSTANT SAVE ---
                    for r in range(len(final_df)):
                        if edited.iloc[r]["Remark"] != final_df.iloc[r]["Remark"]:
                            row_to_update = int(edited.iloc[r]["__row"])
                            call_sheet.update(f"T{row_to_update}", [[edited.iloc[r]["Remark"]]])
                            st.toast(f"Saved Row {row_to_update} ✅")
        else:
            st.warning("Data load nahi ho pa raha hai. Sheet URL aur Permissions check karein.")

    except Exception as e:
        st.error(f"🔴 System Error: {str(e)}")
