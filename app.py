import streamlit as st
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd
from datetime import datetime
import calendar

# ==========================================
# 1. CONNECTION
# ==========================================
def connect_sheet(sheet_name):
    # Secrets se key uthana
    creds_dict = dict(st.secrets["gcp_service_account"])
    
    # Ye step JWT error aur base64 error dono ko khatam kar dega
    # Hum key ke ander ke slash-n ko asli newline mein badal rahe hain
    fixed_key = creds_dict["private_key"].replace("\\n", "\n")
    creds_dict["private_key"] = fixed_key
    
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
    client = gspread.authorize(creds)
    
    return client.open("Bike Check-In (Responses)").worksheet(sheet_name)
st.set_page_config(page_title="Munich Motorrad Management", layout="wide")

# ==========================================
# 2. GLOBAL CSS (FIXED)
# ==========================================
st.markdown("""
    <style>
    /* BLUE TABS STYLE */
    .stTabs [data-baseweb="tab-list"] {
        gap: 12px;
        background-color: #e9ecef;
        padding: 10px;
        border-radius: 12px;
    }
    .stTabs [data-baseweb="tab"] {
        height: 55px;
        background-color: white;
        border: 2.5px solid #0056b3 !important;
        border-radius: 10px;
        padding: 0 25px;
        color: #0056b3;
        font-weight: bold;
    }
    .stTabs [aria-selected="true"] {
        background-color: #0056b3 !important;
        color: white !important;
    }

    /* SAVE BUTTON STYLING */
    button[key^="up_"] {
        height: 75px !important;
        background-color: #0056b3 !important;
        color: white !important;
        font-size: 24px !important;
        font-weight: 900 !important;
        border-radius: 12px !important;
    }

    /* EXPANDER BORDER */
    .stExpander {
        border: 1px solid #ddd !important;
        border-radius: 10px !important;
        margin-bottom: 10px !important;
    }
    </style>
    """, unsafe_allow_html=True)

tab_ws, tab_app = st.tabs(["🔧 Workshop Manager", "📅 Appointment Calendar"])

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
