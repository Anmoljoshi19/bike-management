import streamlit as st
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd
from datetime import datetime
import calendar
import json
import os
import time
import re
import base64

# ==========================================
# PAGE SETUP (FULL SCREEN WIDE)
# ==========================================
st.set_page_config(page_title="BMW Munich Motorrad Indore", page_icon="🏍️", layout="wide")

def get_base64_image(image_path):
    if not os.path.exists(image_path):
        st.error(f"❌ ERROR: Image not found! Ensure it's uploaded to GitHub: {image_path}")
        return ""
    try:
        with open(image_path, "rb") as img_file:
            return base64.b64encode(img_file.read()).decode('utf-8').replace('\n', '')
    except Exception as e:
        st.error(f"❌ ERROR: Error reading image: {e}")
        return ""

# ==========================================
# GLOBAL CSS & FONT INJECTION
# ==========================================
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Roboto+Condensed:wght@400;700;900&family=Roboto:wght@300;400;700&display=swap');
    
    header { visibility: hidden; }
    * { font-family: 'Roboto', sans-serif; }
    </style>
""", unsafe_allow_html=True)

# ==========================================
# GOOGLE SHEETS CONNECTION (UPDATED FOR CLOUD)
# ==========================================
def connect_sheet(sheet_name):
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    try:
        # Using Streamlit Secrets instead of local creds.json
        creds_dict = dict(st.secrets["gcp_service_account"])
        if "private_key" in creds_dict:
            creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")

        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
        client = gspread.authorize(creds)
        return client.open("Bike Check-In (Responses)").worksheet(sheet_name)
    except Exception as e:
        st.error(f"🔴 Connection Error: {sheet_name} | {e}")
        return None

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

@st.cache_data(ttl=300)
def get_cached_sheet_data(url, sheet_name):
    try:
        ws = connect_sheet_by_url(url, sheet_name)
        if ws: return ws.get_all_values()
        return []
    except Exception:
        time.sleep(2)
        ws = connect_sheet_by_url(url, sheet_name)
        if ws: return ws.get_all_values()
        return []

# ==========================================
# SESSION NAVIGATION STATE
# ==========================================
if 'active_section' not in st.session_state:
    st.session_state['active_section'] = 'Landing'

# ==========================================
# A. LANDING PAGE (BMW MOTORRAD THEME)
# ==========================================
if st.session_state['active_section'] == 'Landing':

    # 🔴 RELATIVE PATHS USED HERE - Ensure these files are uploaded to GitHub! 🔴
    sales_bg = get_base64_image("Showroom.png") 
    service_bg = get_base64_image("Workshop.png")
    hero_bg = get_base64_image("GS Image.jpg")

    st.markdown("""
        <style>
        .block-container { padding: 0 !important; max-width: 100% !important; overflow-x: hidden; }
        .stApp { background-color: #f4f4f4; }
        
        .bmw-hero {
            position: relative; width: 100%; height: 70vh;
            display: flex; align-items: center;
        }
        .bmw-hero::before {
            content: ''; position: absolute; top: 0; left: 0; right: 0; bottom: 0;
            background: linear-gradient(90deg, rgba(10,10,10,0.9) 0%, rgba(10,10,10,0.3) 100%);
        }
        .bmw-content { position: relative; z-index: 2; padding-left: 8%; color: white; }
        .bmw-slogan { font-family: 'Roboto Condensed', sans-serif; font-size: 16px; letter-spacing: 5px; color: #d6d6d6; margin-bottom: 10px; }
        .bmw-title { font-family: 'Roboto Condensed', sans-serif; font-size: 70px; font-weight: 900; line-height: 1.1; text-transform: uppercase; margin: 0; }
        .bmw-subtitle { font-size: 20px; font-weight: 300; margin-top: 15px; color: #0066b1; font-weight: bold; }
        
        .module-title { text-align: center; font-family: 'Roboto Condensed', sans-serif; font-size: 32px; font-weight: 900; color: #111; margin: 60px 0 30px 0; text-transform: uppercase; }

        div.stButton > button {
            height: 450px !important; border-radius: 0 !important; border: none !important; color: white !important;
            background-color: #111 !important; transition: all 0.5s cubic-bezier(0.25, 0.46, 0.45, 0.94) !important;
            display: flex !important; align-items: flex-end !important; justify-content: flex-start !important; padding: 40px !important; width: 100% !important;
        }

        div.stButton > button:hover { transform: scale(1.02) !important; box-shadow: 0 20px 40px rgba(0,0,0,0.3) !important; }
        
        div.stButton > button p {
            font-family: 'Roboto Condensed', sans-serif !important; font-size: 36px !important; font-weight: 900 !important; 
            text-transform: uppercase !important; margin: 0 !important; position: relative; z-index: 10;
        }
        div.stButton > button:hover p::after { content: ' ❯' !important; color: #0066b1 !important; }
        </style>
    """, unsafe_allow_html=True)

    st.markdown(f"""
        <style>
        .bmw-hero {{
            background: url('data:image/jpeg;base64,{hero_bg}') no-repeat center center !important;
            background-size: cover !important;
        }}
        div[data-testid="column"]:nth-child(2) div.stButton > button,
        div[data-testid="stColumn"]:nth-child(2) div.stButton > button {{
            background-image: linear-gradient(to top, rgba(0,0,0,0.9) 0%, rgba(0,0,0,0) 50%), url('data:image/png;base64,{sales_bg}') !important;
            background-size: cover !important;
            background-position: center !important;
        }}
        div[data-testid="column"]:nth-child(4) div.stButton > button,
        div[data-testid="stColumn"]:nth-child(4) div.stButton > button {{
            background-image: linear-gradient(to top, rgba(0,0,0,0.9) 0%, rgba(0,0,0,0) 50%), url('data:image/png;base64,{service_bg}') !important;
            background-size: cover !important;
            background-position: center !important;
        }}
        </style>
    """, unsafe_allow_html=True)

    st.markdown("""
        <div class="bmw-hero">
            <div class="bmw-content">
                <p class="bmw-slogan">MAKE LIFE A RIDE.</p>
                <h1 class="bmw-title">MUNICH MOTORRAD<br>INDORE</h1>
                <p class="bmw-subtitle">INTEGRATED DEALERSHIP PLATFORM</p>
            </div>
        </div>
        
        <h2 class="module-title">SELECT DEPARTMENT</h2>
    """, unsafe_allow_html=True)

    _, col1, _, col2, _ = st.columns([1, 4, 0.5, 4, 1])

    with col1:
        if st.button("SHOWROOM & SALES", key="btn_sales", use_container_width=True):
            st.session_state['active_section'] = 'Sales'
            st.rerun()

    with col2:
        if st.button("WORKSHOP & SERVICE", key="btn_service", use_container_width=True):
            st.session_state['active_section'] = 'Service'
            st.rerun()
            
    st.markdown("<br><br><br>", unsafe_allow_html=True)

# ==========================================
# B. SALES DEPARTMENT SECTION
# ==========================================
elif st.session_state['active_section'] == 'Sales':
    
    st.markdown("<style>.block-container { padding: 3rem !important; max-width: 95% !important; margin: 0 auto; }</style>", unsafe_allow_html=True)
    
    st.markdown("""
        <style>
        .back-btn button { background-color: transparent !important; color: #111 !important; border: 2px solid #111 !important; font-weight: bold !important; height: 45px !important; border-radius: 0px !important; text-transform: uppercase; font-family: 'Roboto Condensed', sans-serif; }
        .back-btn button:hover { background-color: #111 !important; color: white !important; }
        </style>
    """, unsafe_allow_html=True)

    st.markdown('<div class="back-btn">', unsafe_allow_html=True)
    if st.button("❮ BACK TO HOME", key="back_sales"):
        st.session_state['active_section'] = 'Landing'
        st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)
        
    st.markdown("<h1 style='font-family: Roboto Condensed, sans-serif; text-transform: uppercase; font-weight: 900;'>Sales Department</h1>", unsafe_allow_html=True)
    st.info("Yahan apna Sales ka code daal dena...")

# ==========================================
# C. SERVICE DEPARTMENT SECTION
# ==========================================
elif st.session_state['active_section'] == 'Service':
    
    st.markdown("<style>.block-container { padding: 3rem !important; max-width: 95% !important; margin: 0 auto; }</style>", unsafe_allow_html=True)
    
    st.markdown("""
        <style>
        .back-btn button { background-color: transparent !important; color: #111 !important; border: 2px solid #111 !important; font-weight: bold !important; height: 45px !important; border-radius: 0px !important; text-transform: uppercase; font-family: 'Roboto Condensed', sans-serif; }
        .back-btn button:hover { background-color: #111 !important; color: white !important; }
        
        .stTabs [data-baseweb="tab-list"] { gap: 15px; background-color: #f0f2f6; padding: 10px; border-radius: 12px; }
        .stTabs [data-baseweb="tab"] { height: 55px; background-color: white; border: 2px solid #0056b3 !important; border-radius: 10px; color: #0056b3; font-weight: bold; padding: 0 25px; }
        .stTabs [aria-selected="true"] { background-color: #0056b3 !important; color: white !important; }
        .stExpander { border: 1px solid #ddd !important; border-radius: 12px !important; margin-bottom: 15px !important; }
        
        button[key^="up_"] { height: 75px !important; background-color: #0056b3 !important; color: white !important; font-size: 22px !important; font-weight: 900 !important; border-radius: 15px !important; }
        [data-testid="stVerticalBlockBorderWrapper"] > div > div { min-height: 250px !important; max-height: 300px !important; overflow-y: auto !important; }
        </style>
        """, unsafe_allow_html=True)

    st.markdown('<div class="back-btn">', unsafe_allow_html=True)
    if st.button("❮ BACK TO HOME", key="back_service"):
        st.session_state['active_section'] = 'Landing'
        st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown("""
        <div style="display: flex; align-items: center; gap: 15px; margin-bottom: 20px; margin-top: 15px;">
            <img src="https://upload.wikimedia.org/wikipedia/commons/4/44/BMW.svg" width="55">
            <h1 style="margin: 0; color: #111; font-family: 'Roboto Condensed', sans-serif; font-weight: 900; text-transform: uppercase;">Workshop Management</h1>
        </div>
    """, unsafe_allow_html=True)

    tab_ws, tab_crm, tab_kpi, tab_parts, tab_parked, tab_rev = st.tabs([
        "🔧 WORKSHOP MANAGER", 
        "🤝 CRM",  
        "📊 KPI DASHBOARD",
        "📦 PARTS",
        "🅿️ BIKE PARKED",
        "💰 REVENUE"
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

            # --- 1. CALCULATION FOR TILES ---
                # Status index 19 (Column T)
                counts = {"Delivered": 0, "Complete": 0, "In Process": 0, "On Hold": 0}
                total_projection_revenue = 0.0
                for item in all_items:
                    row_data = item['data']
                    status = row_data[19].strip() if len(row_data) > 19 else "In Process"
                    
                    if status in counts:
                        counts[status] += 1

                    # Agar status "Delivered" nahi hai, toh Column X (index 23) ka sum lena hai
                    if status != "Delivered" and len(row_data) > 23:
                        # Amount mein se comma hata kar float mein convert karna zaroori hai
                        revenue_str = str(row_data[23]).replace(',', '').strip()
                        
                        # Check karna ki value empty na ho aur number hi ho
                        try:
                            if revenue_str: 
                                total_projection_revenue += float(revenue_str)
                        except ValueError:
                            pass # Agar galti se text likha ho cell me toh ignore karega

                # --- 2. DISPLAY SUMMARY TILES ---
                t_col1, t_col2, t_col3, t_col4,t_col5 = st.columns(5)
                tile_style = """
                    <div style="background-color: {color}; padding: 15px; border-radius: 10px; text-align: center; color: {text_color}; box-shadow: 2px 2px 5px rgba(0,0,0,0.1);">
                        <p style="margin: 0; font-size: 14px; font-weight: bold; text-transform: uppercase;">{label}</p>
                        <h2 style="margin: 0; font-size: 24px;">{value}</h2>
                    </div>
                """

                with t_col1:
                    st.markdown(tile_style.format(color="#E3F2FD", text_color="#0D47A1", label="Delivered", value=counts["Delivered"]), unsafe_allow_html=True)
                with t_col2:
                    st.markdown(tile_style.format(color="#E8F5E9", text_color="#1B5E20", label="Complete", value=counts["Complete"]), unsafe_allow_html=True)
                with t_col3:
                    st.markdown(tile_style.format(color="#FFFDE7", text_color="#F57F17", label="In Process", value=counts["In Process"]), unsafe_allow_html=True)
                with t_col4:
                    st.markdown(tile_style.format(color="#FFEBEE", text_color="#B71C1C", label="On Hold", value=counts["On Hold"]), unsafe_allow_html=True)
                with t_col5:
                    # Revenue ko thoda achhe format me dikhane ke liye (eg. 150000 -> ₹1,50,000)
                    formatted_revenue = f"₹{total_projection_revenue:,.0f}"
                    st.markdown(tile_style.format(color="#F3E5F5", text_color="#4A148C", label="Projected Revenue", value=formatted_revenue), unsafe_allow_html=True)

                st.markdown("<br>", unsafe_allow_html=True) # Gap between tiles and tabs


                sub1, sub2 = st.tabs(["📂 Open JCs", "✅ History (Delivered)"])
                
                with sub1:
                    s_o = st.text_input("🔍 Search Open JCs", key="ws_so")
                    
                    # Filter out Delivered bikes (Status ab index 19 par hai)
                    open_list = [item for item in all_items if len(item['data']) > 19 and item['data'][19] != "Delivered"]
                    
                    # --- ARRANGE LOGIC (HOLD -> PROCESS -> COMPLETE) ---
                    def workshop_sort(item):
                        status = item['data'][19].strip() if len(item['data']) > 19 else ""
                        if status == "On Hold": return 1
                        if status == "In Process": return 2
                        if status == "Complete": return 3
                        return 4
                    
                    open_list.sort(key=workshop_sort)

                    for item in open_list:
                        r = item['data']
                        idx = item['row_idx']
                        stat = r[19].strip() if len(r) > 19 else "In Process"
                        
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
                                    st.markdown("---")
                                    st.markdown(f"**💵 Estimated Total Revenue:**\n ₹ {r[23]}")
                                with c2:
                                    st.markdown(f"**🛠️ Staff:**\n{r[9]}")
                                    st.markdown("---")
                                    st.markdown(f"**👨‍🔧 Technician:**\n{r[17] if len(r)>17 and r[17].strip() != '' else 'Not Assigned'}")
                                    st.markdown("---")
                                    st.markdown(f"**🛣️ Odometer (KM):**\n{r[7] if len(r)>7 else 'N/A'}")
                                    st.markdown("---")
                                    st.info(f"**🚨 Issues:**\n{r[15] if len(r)>15 else 'N/A'}")
                                with c3:
                                    # Remark ab index 18 par hai
                                    n_rem = st.text_area("Workshop Remark", r[18] if len(r)>18 else "", key=f"r_{idx}", height=120)
                                    st_opts = ["On Hold", "In Process", "Complete", "Delivered"]
                                    try: curr_sel = st_opts.index(stat)
                                    except: curr_sel = 1
                                    
                                    n_stat = st.selectbox("Update Status", st_opts, index=curr_sel, key=f"s_{idx}")
                                    
                                    if st.button("SAVE CHANGES", key=f"up_{idx}", use_container_width=True):
                                        # Sheet mein S column = 19, T column = 20
                                        ws_sheet.update_cell(idx, 19, n_rem)
                                        ws_sheet.update_cell(idx, 20, n_stat)
                                        st.rerun()

                with sub2:
                    s_h = st.text_input("🔍 Search Delivered History", key="ws_sh")
                    # Filter Delivered bikes (Status ab index 19 par hai)
                    history_list = [item for item in all_items if len(item['data']) > 19 and item['data'][19] == "Delivered"]
                    
                    for item in history_list:
                        r = item['data']
                        if s_h.lower() in str(r).lower():
                            with st.expander(f"✅ {r[1]} | {r[5]} | Delivered"):
                                # Pura data History section mein
                                hc1, hc2, hc3 = st.columns(3)
                                with hc1:
                                    st.markdown(f"**👤 Customer:** {r[1]}")
                                    st.markdown(f"**📱 Phone:** {r[2] if len(r)>2 else 'N/A'}")
                                    st.markdown(f"**📧 Email:** {r[3] if len(r)>3 else 'N/A'}")
                                    st.markdown(f"**📅 Date:** {r[0]}")
                                with hc2:
                                    st.markdown(f"**🆔 VIN:** {r[8] if len(r)>8 else 'N/A'}")
                                    st.markdown(f"**🛠️ Staff:** {r[9] if len(r)>9 else 'N/A'}")
                                    st.markdown(f"**👨‍🔧 Technician:**\n{r[17] if len(r)>17 and r[17].strip() != '' else 'Not Assigned'}")
                                    st.markdown(f"**🛣️ Odo:** {r[7] if len(r)>7 else 'N/A'} KM")
                                    st.markdown(f"**💵 Parts Revenue (Without Tax):** ₹ {r[21] if len(r)>21 else 'N/A'}")
                                    st.markdown(f"**💵 Labor Revenue (With Tax):** ₹ {r[22] if len(r)>22 else 'N/A'}")
                                with hc3:
                                    st.warning(f"**🚨 Issues:**\n{r[15] if len(r)>15 else 'N/A'}")
                                    # Remark ab index 18 par hai
                                    st.success(f"**✍️ Final Remark:**\n{r[18] if len(r)>18 else 'N/A'}")
                                    st.info(f"**Status:** Delivered")

        except Exception as e:
            st.error(f"Workshop Load Error: {e}")

    # ==========================================
    # 🤝 TAB 2: CRM (Appointments, Calling, VOC)
    # ==========================================
    with tab_crm: 
        st.header("🤝 Customer Relationship Management (CRM)")
        
        # --- 🎨 UI FIX: ROYAL BLUE FILLED BORDER TABS ---
        st.markdown("""
            <style>
            /* Radio circles ko hide karna */
            [data-testid="stRadio"] div[role="radiogroup"] > label > div:first-child {
                display: none !important;
            }
            /* Buttons ko line mein lana aur gap dena */
            [data-testid="stRadio"] div[role="radiogroup"] {
                flex-direction: row;
                gap: 15px;
                margin-bottom: 15px;
            }
            /* Normal State: Royal Blue Border aur White background */
            [data-testid="stRadio"] div[role="radiogroup"] > label {
                background-color: #ffffff !important;
                border: 2px solid #4169E1 !important; /* Royal Blue */
                padding: 10px 25px !important;
                border-radius: 8px !important;
                cursor: pointer;
                transition: all 0.3s ease;
            }
            /* Normal Text: Royal Blue color */
            [data-testid="stRadio"] div[role="radiogroup"] > label p {
                color: #4169E1 !important;
                font-size: 16px !important;
                font-weight: 700 !important;
                margin: 0 !important;
            }
            /* Hover Effect: Halka blue background */
            [data-testid="stRadio"] div[role="radiogroup"] > label:hover {
                background-color: #f0f4ff !important;
            }
            /* ACTIVE/SELECTED State: Pura Royal Blue fill aur White Text */
            div[data-testid="stRadio"] > div[role="radiogroup"] > label:has(input:checked) {
                background-color: #4169E1 !important;
                box-shadow: 0px 4px 6px rgba(65, 105, 225, 0.3);
            }
            div[data-testid="stRadio"] > div[role="radiogroup"] > label:has(input:checked) p {
                color: white !important;
            }
            </style>
        """, unsafe_allow_html=True)
        
        # 🔘 CRM Main Navigation
        crm_nav = st.radio(
            "Navigate CRM:",
            ["📅 Appointments", "📞 Service Calling", "🗣️ VOC"],
            horizontal=True,
            label_visibility="collapsed"
        )
        
        st.markdown("---")

        # ==========================================
        # 1️⃣ MODULE: APPOINTMENTS
        # ==========================================
        if crm_nav == "📅 Appointments":
            try:
                app_sheet = connect_sheet("Appointments")
                app_vals = app_sheet.get_all_values()
                
                if len(app_vals) > 1:
                    df = pd.DataFrame([{
                        "row": i+2, 
                        "Name": r[1] if len(r)>1 else "", 
                        "Phone": r[2] if len(r)>2 else "", 
                        "D": r[3] if len(r)>3 else "", 
                        "Model": r[4] if len(r)>4 else "", 
                        "VIN": r[5] if len(r)>5 else "", 
                        "P": r[6] if len(r)>6 else "", 
                        "R": r[7] if len(r)>7 else ""
                    } for i, r in enumerate(app_vals[1:])])
                    
                    df['D'] = pd.to_datetime(df['D'], errors='coerce')
                    df = df.dropna(subset=['D'])
                    today = datetime.now().date()

                    tab_cal, tab_missed = st.tabs(["📅 Calendar View", "⚠️ Missed Appointments"])

                    with tab_cal:
                        cy, cm = st.columns(2)
                        y = cy.selectbox("Year", [2025, 2026], index=1)
                        mn = cm.selectbox("Month", list(calendar.month_name)[1:], index=datetime.now().month-1)
                        midx = list(calendar.month_name).index(mn)

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
                                            
                                            if is_rep: b_c = "#28a745"
                                            elif row['D'].date() < today: b_c = "#d9534f"
                                            else: b_c = "#007bff"
                                            
                                            st.markdown(f"""<style>div.stButton > button[key="trig_{row['row']}"] {{ border: 2.5px solid {b_c} !important; min-height: 60px !important; margin-bottom: 5px !important; }}</style>""", unsafe_allow_html=True)
                                            
                                            if st.button(f"{'✅' if is_rep else '⭕'} {row['Name']}", key=f"trig_{row['row']}", use_container_width=True):
                                                app_sheet.update_cell(row['row'], 8, "Bike Reported" if not is_rep else "")
                                                import time
                                                time.sleep(0.5)
                                                st.rerun()
                                else:
                                    cols[i].markdown("<div style='height:100px; opacity:0.1;'></div>", unsafe_allow_html=True)

                    with tab_missed:
                        st.subheader("⚠️ Action Required: Missed Appointments")
                        missed_df = df[(df['D'].dt.date < today) & (~df['R'].str.contains("Bike Reported", na=False))].sort_values(by='D', ascending=False)
                        
                        if missed_df.empty:
                            st.success("🎉 Perfect! Koi bhi Missed Appointment pending nahi hai.")
                        else:
                            for _, row in missed_df.iterrows():
                                with st.container(border=True):
                                    c1, c2, c3, c4 = st.columns([2, 2, 2, 2.5])
                                    with c1:
                                        st.markdown(f"**👤 Name:** {row['Name']}")
                                        st.markdown(f"**📞 Contact:** {row['Phone']}")
                                    with c2:
                                        st.markdown(f"**🏍️ Model:** {row['Model']}")
                                        st.markdown(f"**🔢 VIN:** {row['VIN']}")
                                    with c3:
                                        st.markdown(f"**🔧 Purpose:** {row['P']}")
                                        st.markdown(f"**📅 Old Date:** <span style='color:#d9534f; font-weight:bold;'>{row['D'].strftime('%d %b %Y')}</span>", unsafe_allow_html=True)
                                    with c4:
                                        new_date = st.date_input("Select New Date", value=today, key=f"new_dt_{row['row']}")
                                        st.markdown("<br>", unsafe_allow_html=True)
                                        if st.button("🔄 Reschedule", key=f"resched_{row['row']}", use_container_width=True):
                                            app_sheet.update_cell(row['row'], 4, new_date.strftime("%d-%b-%Y"))
                                            st.success(f"✅ Appointment rescheduled to {new_date.strftime('%d %b')}!")
                                            import time
                                            time.sleep(0.5)
                                            st.rerun()
            except Exception as e: 
                st.error(f"Calendar Error: {e}")

        # ==========================================
        # 2️⃣ MODULE: SERVICE CALLING
        # ==========================================
        elif crm_nav == "📞 Service Calling":
            try:
                dash_url = "https://docs.google.com/spreadsheets/d/1Roc_HIQLxsqRoxwCZVEkRgnQ0koe8A7wJOdJBWPWVas"

                @st.cache_data(ttl=60)
                def load_calling_data():
                    d_sheet = connect_sheet_by_url(dash_url, "Dashboard")
                    m_sheet = connect_sheet_by_url(dash_url, "Main Service Sheet")
                    return d_sheet.get("F2:G8"), m_sheet.get_all_values()

                stats_data, raw_data = load_calling_data()

                df_raw = pd.DataFrame(raw_data[1:], columns=raw_data[0])
                df = df_raw.loc[:, ~df_raw.columns.duplicated()].copy() 
                df["__row"] = df.index + 2
                
                target_col = "Sold by Other Dealer" 
                mask = (df["Status"].str.lower().str.contains("pending", na=False)) & \
                    (df[target_col].fillna("").str.strip() == "")
                df_filtered = df[mask]

                st.markdown("""<style>
                    div.stButton > button { width: 100%; height: 110px !important; border-radius: 15px; background-color: white; border: 1px solid #ddd; transition: 0.3s; white-space: pre-wrap !important; }
                    div.stButton > button:hover { border: 2px solid #ff8c00 !important; background-color: #fffaf5 !important; }
                    div.stButton > button p { font-size: 16px !important; font-weight: 800 !important; line-height: 1.3 !important; color: #333; }
                    div.stButton > button[kind="primary"] { background-color: #28a745 !important; border: 3px solid #1e7e34 !important; }
                    div.stButton > button[kind="primary"] p { color: white !important; }
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
                        btn_type = "primary" if is_sel else "secondary"
                        
                        if cols[i-1].button(display_text, key=f"btn_{mode}_{i}", type=btn_type):
                            st.session_state[svc_key] = "All" if is_sel else str(i)
                            st.rerun()

                    current = st.session_state[svc_key]
                    final_df = df_part if current == "All" else df_part[df_part["Service Count"].astype(str).str.startswith(current)]

                    if final_df.empty:
                        st.info(f"No {current if current != 'All' else ''} pending calls.")
                    else:
                        st.write(f"📋 Showing **{len(final_df)}** records")
                        edited = st.data_editor(
                            final_df, use_container_width=True, height=500,
                            disabled=[c for c in final_df.columns if c not in ["Remark"]],
                            column_config={"__row": None, "Remark": st.column_config.TextColumn(width="large")},
                            key=f"ed_{mode}_{current}"
                        )

                        if st.button("💾 SAVE ALL CHANGES", key=f"save_btn_{mode}_{current}", type="primary"):
                            updates_count = 0
                            with st.spinner("⏳ Saving to Google Sheets..."):
                                fresh_call_sheet = connect_sheet_by_url(dash_url, "Main Service Sheet")
                                for r in range(len(final_df)):
                                    if edited.iloc[r]["Remark"] != final_df.iloc[r]["Remark"]:
                                        row_to_update = int(edited.iloc[r]["__row"])
                                        new_remark = edited.iloc[r]["Remark"]
                                        fresh_call_sheet.update(f"T{row_to_update}", [[new_remark]])
                                        updates_count += 1
                            
                            if updates_count > 0:
                                st.success(f"✅ {updates_count} records updated successfully!")
                                load_calling_data.clear() 
                                import time
                                time.sleep(1)
                                st.rerun()
                            else:
                                st.info("No new changes detected to save.")

                with t1: render_dashboard(df_filtered[df_filtered["City"].str.lower() == "indore"], "Indore")
                with t2: render_dashboard(df_filtered[df_filtered["City"].str.lower() != "indore"], "Outstation")

            except Exception as e:
                st.error(f"🔴 System Error: {str(e)}")

        # ==========================================
        # 3️⃣ MODULE: VOC (Voice of Customer)
        # ==========================================
        elif crm_nav == "🗣️ VOC":
            st.subheader("🗣️ Voice of Customer (VOC)")
            st.info("🚧 VOC Dashboard is currently under construction. Please add your sheet connection and logic here.")
    # ------------------------------------------
    # TAB 3: KPI DASHBOARD 
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

                # --- TOP TILE (COMPACT) ---
                st.markdown(f"""
                    <div style="background-color:#0056b3; padding:5px 10px; border-radius:10px; text-align:center; border: 2px solid #FFD700; margin-bottom:10px;">
                        <p style="color:white; margin:0; font-size:14px; font-weight:bold;">🏆 TOTAL KPI SCORE (Q2)</p>
                        <h1 style="color:#FFD700; font-size:38px; margin:-5px 0; font-weight:bold;">{total_sum:.1f} <span style="font-size:16px; color:white;">/ 100</span></h1>
                    </div>
                """, unsafe_allow_html=True)

                s1, s2 = st.tabs(["🎯 SCORE BOARD", "📊 QUARTERLY STATS"])

                # --- TAB 1: SCORE BOARD (WITH COLUMN I) ---
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

                # --- TAB 2: QUARTERLY STATS ---
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
                        except: 
                            pass
                    stats_html += "</table>"
                    st.markdown(stats_html, unsafe_allow_html=True)

                    st.markdown("---")
                    st.markdown("### 🏍️ Munich Motorrad Indore JC 25 vs 26")

                    try:
                        jc_needed = kpi_data[1][25]
                        st.markdown(f'<div style="background-color:#fff3cd; padding:8px 15px; border-radius:8px; border: 1px solid #ffa000; display: inline-block; margin-bottom:15px;"><span style="color:#856404; font-weight:bold;">🎯 JC NEEDED TO CLOSE: </span><span style="color:#000; font-size:18px; font-weight:bold;">{jc_needed}</span></div>', unsafe_allow_html=True)
                    except: 
                        pass

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
    # TAB 4: PARTS 
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

    # ==========================================
    # TAB 5: BIKE PARKED 
    # ==========================================
    with tab_parked:
        try:
            # 1. LOAD DATA 
            parked_url = "https://docs.google.com/spreadsheets/d/1Roc_HIQLxsqRoxwCZVEkRgnQ0koe8A7wJOdJBWPWVas"
            
            @st.cache_data(ttl=60)
            def load_parked_data():
                d_sheet = connect_sheet_by_url(parked_url, "Dashboard")
                m_sheet = connect_sheet_by_url(parked_url, "Main Service Sheet")
                o_sheet = connect_sheet_by_url(parked_url, "Oil") 
                return d_sheet.get_all_values(), m_sheet.get_all_values(), o_sheet.get_all_values()

            bp_data, main_sheet_data, oil_data = load_parked_data()

            # Helper function for Dashboard cells
            def get_bp_val(row_idx, col_letter):
                try:
                    c_idx = ord(col_letter.upper()) - 65
                    val = bp_data[row_idx - 1][c_idx]
                    return val if val else "0"
                except:
                    return "0"
                    
            def safe_int(val):
                try:
                    return int(str(val).replace(',', '').strip())
                except:
                    return 0

            # --- PYTHON LOGIC FOR "TOTAL BIKE PARKED YTD" ---
            ytd_parked_count = 0
            active_parked_ids = set() 
            
            for row in main_sheet_data[1:]: 
                col_c = str(row[2]) if len(row) > 2 else "" # Col C = index 2
                col_v = str(row[21]) if len(row) > 21 else "" # Col V = index 21
                
                if col_c.strip() != "" and col_v == "":
                    ytd_parked_count += 1
                    # FIX 1: Python ko case-insensitive banane ke liye .upper() lagaya
                    active_parked_ids.add(col_c.strip().upper()) 

            # --- PYTHON LOGIC FOR "INDORE OIL CHANGED (1 Yr)" ---
            indore_oil_count = 0
            try:
                # FIX 2: Exact Google Sheet jaisa date logic (bina time ke)
                today = pd.Timestamp.today().normalize()
                one_year_ago = today - pd.Timedelta(days=365)
                unique_oil_ids = set()
                
                for row in oil_data[1:]: 
                    # FIX 1: Oil ID ko bhi upper case kiya match karne ke liye
                    oil_id = str(row[0]).strip().upper() if len(row) > 0 else "" # Col A
                    oil_date_str = str(row[4]).strip() if len(row) > 4 else "" # Col E
                    
                    if oil_id in active_parked_ids and oil_date_str:
                        try:
                            # Date ko normalize kiya taaki time ka jhanjhat na rahe
                            oil_date = pd.to_datetime(oil_date_str, errors='coerce', dayfirst=True).normalize()
                            
                            # FIX 2: >= TODAY()-365 AND <= TODAY() wala exact Google logic
                            if pd.notna(oil_date) and (one_year_ago <= oil_date <= today):
                                unique_oil_ids.add(oil_id)
                        except:
                            pass
                indore_oil_count = len(unique_oil_ids)
            except Exception as e:
                indore_oil_count = 0

            # --- CSS FOR THIS TAB ---
            st.markdown("""
                <style>
                    .bp-tile { background-color: #f8f9fa; border-left: 5px solid #0056b3; padding: 15px; border-radius: 8px; text-align: center; box-shadow: 2px 2px 5px rgba(0,0,0,0.05); margin-bottom: 15px; }
                    .bp-tile-title { font-size: 13px; color: #555; font-weight: bold; margin-bottom: 5px; text-transform: uppercase; }
                    .bp-tile-value { font-size: 24px; color: #111; font-weight: 900; margin: 0; }
                    
                    .parked-table { width: 100%; border-collapse: collapse; text-align: center; margin-bottom: 20px; }
                    .parked-table th, .parked-table td { border: 1px solid #ddd; padding: 8px; }
                    .parked-table th { background-color: #0056b3 !important; color: white !important; font-weight: bold !important; }
                </style>
            """, unsafe_allow_html=True)

            # ---------------------------------------------------------
            # SECTION 1: TOP SUMMARY TILES
            # ---------------------------------------------------------
            st.markdown("### 📊 Bike Parked Summary")
            t1, t2, t3, t4 = st.columns(4)
            
            with t1:
                st.markdown(f'<div class="bp-tile" style="border-color: #0056b3;"><div class="bp-tile-title">TOTAL BIKE PARKED (YTD)</div><div class="bp-tile-value">{ytd_parked_count}</div></div>', unsafe_allow_html=True)
            with t2:
                st.markdown(f'<div class="bp-tile" style="border-color: #28a745;"><div class="bp-tile-title">BBND</div><div class="bp-tile-value">{get_bp_val(2, "I")}</div></div>', unsafe_allow_html=True)
            with t3:
                st.markdown(f'<div class="bp-tile" style="border-color: #ff8c00;"><div class="bp-tile-title">SERVICE RETENTION</div><div class="bp-tile-value">{get_bp_val(14, "J")}</div></div>', unsafe_allow_html=True)
            with t4:
                st.markdown(f'<div class="bp-tile" style="border-color: #17a2b8;"><div class="bp-tile-title">TOTAL BIKE PARKED OIL CHANGE (1 YR)</div><div class="bp-tile-value">{indore_oil_count}</div></div>', unsafe_allow_html=True)

            st.markdown("---")

            # ---------------------------------------------------------
            # SECTION 2: INACTIVE Bikes & CORE VS SMART
            # ---------------------------------------------------------
            col_A, col_B = st.columns(2)

            with col_A:
                st.markdown("#### 🚫 Inactive Bikes")
                st.info("Vehicles beyond service scope - relocated, sold, or total loss.")
                
                oot_html = f"""
                <table class="parked-table">
                    <tr><th>Category</th><th>Count</th></tr>
                    <tr><td>Relocated</td><td><b>{get_bp_val(5, "J")}</b></td></tr>
                    <tr><td>Sold / Contact N.A.</td><td><b>{get_bp_val(6, "J")}</b></td></tr>
                    <tr><td>Total Loss</td><td><b>{get_bp_val(7, "J")}</b></td></tr>
                    <tr style="background-color:#ffeeba;"><td><b>Lost Customer</b></td><td><b style="color:red; font-size: 18px;">{get_bp_val(4, "J")}</b></td></tr>
                </table>
                """
                st.markdown(oot_html, unsafe_allow_html=True)

            with col_B:
                st.markdown("#### 🏍️ Core & Smart Bikes Status")
                st.success("Bifurcation of Total Bikes Parked.")
                
                total_core = sum(safe_int(get_bp_val(i, "B")) for i in range(14, 19))
                total_smart = sum(safe_int(get_bp_val(i, "C")) for i in range(14, 19)) + safe_int(get_bp_val(2, "I"))

                core_smart_html = f"""
                <table class="parked-table">
                    <tr><th>Status</th><th>Core</th><th>Smart</th></tr>
                    <tr><td>Service Done</td><td>{get_bp_val(14, "B")}</td><td>{get_bp_val(14, "C")}</td></tr>
                    <tr><td>Service Pending</td><td>{get_bp_val(15, "B")}</td><td>{get_bp_val(15, "C")}</td></tr>
                    <tr><td>Relocated</td><td>{get_bp_val(16, "B")}</td><td>{get_bp_val(16, "C")}</td></tr>
                    <tr><td>Sold / Contact N.A.</td><td>{get_bp_val(17, "B")}</td><td>{get_bp_val(17, "C")}</td></tr>
                    <tr><td>Total Loss</td><td>{get_bp_val(18, "B")}</td><td>{get_bp_val(18, "C")}</td></tr>
                    <tr style="background-color:#d4edda;"><td><b style="font-size:16px;">Total</b></td><td><b style="font-size:16px;">{total_core}</b></td><td><b style="font-size:16px;">{total_smart}</b></td></tr>
                </table>
                """
                st.markdown(core_smart_html, unsafe_allow_html=True)

            st.markdown("---")

            # ---------------------------------------------------------
            # SECTION 3: SERVICE PIPELINE & AREA-WISE BACKLOG
            # ---------------------------------------------------------
            st.markdown("#### 📅 Service Pipeline & Area-wise Backlog")
            
            pipe_col1, pipe_gap, pipe_col2 = st.columns([1, 0.1, 0.8])
            
            with pipe_col1:
                pipe_html = f"""
                <table class="parked-table">
                    <tr><th>Service No.</th><th>Service Pending</th><th>Service Coming</th><th>Service Done</th></tr>
                    <tr><td><b>1st Service</b></td><td>{get_bp_val(2, "B")}</td><td>{get_bp_val(2, "C")}</td><td>{get_bp_val(2, "D")}</td></tr>
                    <tr><td><b>2nd Service</b></td><td>{get_bp_val(3, "B")}</td><td>{get_bp_val(3, "C")}</td><td>{get_bp_val(3, "D")}</td></tr>
                    <tr><td><b>3rd Service</b></td><td>{get_bp_val(4, "B")}</td><td>{get_bp_val(4, "C")}</td><td>{get_bp_val(4, "D")}</td></tr>
                    <tr><td><b>4th Service</b></td><td>{get_bp_val(5, "B")}</td><td>{get_bp_val(5, "C")}</td><td>{get_bp_val(5, "D")}</td></tr>
                    <tr><td><b>5th Service</b></td><td>{get_bp_val(6, "B")}</td><td>{get_bp_val(6, "C")}</td><td>{get_bp_val(6, "D")}</td></tr>
                    <tr><td><b>6th Service</b></td><td>{get_bp_val(7, "B")}</td><td>{get_bp_val(7, "C")}</td><td>{get_bp_val(7, "D")}</td></tr>
                    <tr style="background-color:#fff3cd;"><td><b>Total Pending</b></td><td><b>{get_bp_val(8, "B")}</b></td><td><b>{get_bp_val(8, "C")}</b></td><td><b>{get_bp_val(8, "D")}</b></td></tr>
                    <tr style="background-color:#d4edda;"><td><b>Percentage</b></td><td><b>{get_bp_val(9, "B")}</b></td><td><b>{get_bp_val(9, "C")}</b></td><td><b>{get_bp_val(9, "D")}</b></td></tr>
                </table>
                """
                st.markdown(pipe_html, unsafe_allow_html=True)
                
            with pipe_col2:
                city_html = f"""
                <table class="parked-table">
                    <tr><th>Service No.</th><th>Indore Pending</th><th>Outstation Pending</th></tr>
                    <tr><td><b>1st Service</b></td><td>{get_bp_val(2, "F")}</td><td>{get_bp_val(2, "G")}</td></tr>
                    <tr><td><b>2nd Service</b></td><td>{get_bp_val(3, "F")}</td><td>{get_bp_val(3, "G")}</td></tr>
                    <tr><td><b>3rd Service</b></td><td>{get_bp_val(4, "F")}</td><td>{get_bp_val(4, "G")}</td></tr>
                    <tr><td><b>4th Service</b></td><td>{get_bp_val(5, "F")}</td><td>{get_bp_val(5, "G")}</td></tr>
                    <tr><td><b>5th Service</b></td><td>{get_bp_val(6, "F")}</td><td>{get_bp_val(6, "G")}</td></tr>
                    <tr><td><b>6th Service</b></td><td>{get_bp_val(7, "F")}</td><td>{get_bp_val(7, "G")}</td></tr>
                    <tr style="background-color:#fff3cd;"><td><b>Total</b></td><td><b>{get_bp_val(8, "F")}</b></td><td><b>{get_bp_val(8, "G")}</b></td></tr>
                </table>
                """
                st.markdown(city_html, unsafe_allow_html=True)

            st.markdown("---")

            # ---------------------------------------------------------
            # SECTION 4: HIGH-RISK FLEET
            # ---------------------------------------------------------
            st.markdown("#### ⚠️ High-Risk Bikes (1.5+ Years Inactive)")
            st.error("Customers highly likely to be servicing at unauthorized workshops.")
            
            risk_html = f"""
            <table class="parked-table" style="width:60%; margin:auto;">
                <tr><th>Service Category</th><th>Indore</th><th>Outstation</th><th>Total</th></tr>
                <tr><td><b>1st Service</b></td><td>{get_bp_val(23, "D")}</td><td>{get_bp_val(23, "C")}</td><td><b>{get_bp_val(23, "B")}</b></td></tr>
                <tr><td><b>2nd Service</b></td><td>{get_bp_val(24, "D")}</td><td>{get_bp_val(24, "C")}</td><td><b>{get_bp_val(24, "B")}</b></td></tr>
                <tr><td><b>3rd Service</b></td><td>{get_bp_val(25, "D")}</td><td>{get_bp_val(25, "C")}</td><td><b>{get_bp_val(25, "B")}</b></td></tr>
                <tr><td><b>4th Service</b></td><td>{get_bp_val(26, "D")}</td><td>{get_bp_val(26, "C")}</td><td><b>{get_bp_val(26, "B")}</b></td></tr>
                <tr style="background-color:#f8d7da;"><td><b>GRAND TOTAL</b></td><td><b style="color:red; font-size:18px;">{get_bp_val(27, "D")}</b></td><td><b style="color:red; font-size:18px;">{get_bp_val(27, "C")}</b></td><td><b style="color:red; font-size:18px;">{get_bp_val(27, "B")}</b></td></tr>
            </table>
            """
            st.markdown(risk_html, unsafe_allow_html=True)

        except Exception as e:
            st.error(f"🔴 Bike Parked Data Error: {str(e)}")

    # ==========================================
    # 6. REVENUE TAB (STRUCTURE ONLY)
    # ==========================================
    with tab_rev:
        st.title("💰 Revenue Analytics")
        
        # Sub-tabs create karna
        st_rev_1, st_rev_2= st.tabs([
            "📅 CURRENT MONTH", 
            "📊 YoY COMPARISON", 
                    
        ])

        # --- SUB TAB 1: CURRENT MONTH ---
        with st_rev_1:
            
            # 🗓️ MONTH & YEAR SELECTOR
            col_m, col_y, _ = st.columns([1, 1, 3]) 
            with col_m:
                months_list = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
                current_m_idx = datetime.now().month - 1
                sel_month = st.selectbox("📅 Select Month", months_list, index=current_m_idx)
            with col_y:
                sel_year = st.selectbox("🗓️ Select Year", [2025, 2026, 2027], index=1) # 2026 default

            # 🔄 DYNAMIC DATES CALCULATION
            m_num = months_list.index(sel_month) + 1
            start_dt = pd.to_datetime(f"{sel_year}-{m_num:02d}-01")
            
            
            report_dates = pd.date_range(start=start_dt, end=start_dt + pd.offsets.MonthEnd(0))
            
            start_date_range = report_dates[0].strftime("%d %b %Y")
            end_date_range = report_dates[-1].strftime("%d %b %Y")
            
            # 1. 🏆 PREMIUM STYLED TITLE
            st.markdown(f"""
                <div style="
                    background: linear-gradient(90deg, #1e3a8a 0%, #3b82f6 100%);
                    padding: 20px;
                    border-radius: 15px;
                    text-align: center;
                    box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.2);
                    margin-top: 10px;
                    margin-bottom: 30px;
                    border: 2px solid #000;
                ">
                    <h1 style="color: white; margin: 0; font-size: 28px; font-family: 'Segoe UI', sans-serif;">
                        MUNICH MOTORRAD INDORE W/S REPORT
                    </h1>
                    <p style="color: #e0e7ff; margin: 5px 0 0 0; font-size: 18px; font-weight: 500;">
                        Period: {start_date_range} to {end_date_range}
                    </p>
                </div>
            """, unsafe_allow_html=True)

            url_2026 = "https://docs.google.com/spreadsheets/d/1XYZWk94-RCqABmzCd40-v6RLlXcwbOaxPMK9jE8x9No/edit"
            ws_2026 = connect_sheet_by_url(url_2026, "2026")

            if ws_2026:
                raw_data = ws_2026.get_all_values()
                
                if len(raw_data) > 1:
                    headers = [str(h).strip() for h in raw_data[0]]
                    df_2026 = pd.DataFrame(raw_data[1:], columns=headers)

                    # Column Config
                    col_date, col_vin, col_cat = 'Document Date', 'VIN #', 'Item Type'
                    col_labour_amt = 'Final Amount with tax'     # AN
                    col_parts_amt = 'Final Amount without tax'   # AG
                    col_dnp = 'DNP'                              # AT

                    if col_date in df_2026.columns:
                        # 1. DATA CLEANING
                        cols_to_clean = [col_cat, 'Type', 'Module Name', 'Customer Type', 'Part Type', 'Description']
                        for col in cols_to_clean:
                            if col in df_2026.columns:
                                df_2026[col + '_Clean'] = df_2026[col].astype(str).str.strip().str.lower()
                            else:
                                df_2026[col + '_Clean'] = ''
                        
                        def clean_amt(series):
                            return pd.to_numeric(series.astype(str).str.replace(',', '').str.strip(), errors='coerce').fillna(0)
                        
                        df_2026['LabAmt'] = clean_amt(df_2026[col_labour_amt]) 
                        df_2026['PartAmt'] = clean_amt(df_2026[col_parts_amt])
                        df_2026['DnpAmt'] = clean_amt(df_2026[col_dnp])

                        df_2026['Date_DT'] = pd.to_datetime(df_2026[col_date], format='mixed', dayfirst=True, errors='coerce').dt.date
                        df_2026['VIN_Clean'] = df_2026[col_vin].astype(str).str.strip().str.upper()
                        invalid_vins = {'', 'NAN', 'NONE', '- NONE -'}

                        rows = []

                        # Calculate Table 1 (Throughput)
                        for d in report_dates:
                            target_date = d.date()
                            df_day = df_2026[df_2026['Date_DT'] == target_date]
                            day_vins = df_day[~df_day['VIN_Clean'].isin(invalid_vins)]['VIN_Clean'].unique()
                            ach_thru = len(day_vins)
                            
                            ach_lab = df_day.loc[df_day['Item Type_Clean'].isin(["labor", "warranty handling charges"]), 'LabAmt'].sum()
                            ach_part = df_day.loc[df_day['Item Type_Clean'].isin(["parts", "part"]), 'PartAmt'].sum()

                            rows.append({
                                "Date": f"{d.month}/{d.day}/{d.year}",
                                "T_Throughput": 3, "T_Labour": 7142, "T_Parts": 28571,
                                "A_Throughput": ach_thru, "A_Labour": ach_lab, "A_Parts": ach_part,
                                "G_Throughput": 3 - ach_thru, "G_Labour": 7142 - ach_lab, "G_Parts": 28571 - ach_part
                            })

                        df_data = pd.DataFrame(rows)
                        total_row = {"Date": "Total"}
                        for c in df_data.columns[1:]: total_row[c] = df_data[c].sum()
                        df_final = pd.concat([pd.DataFrame([total_row]), df_data], ignore_index=True)
                        df_final.columns = pd.MultiIndex.from_tuples([("", "Date"), ("Target", "Throughput"), ("Target", "Labour"), ("Target", "Parts"), ("Achievement", "Throughput"), ("Achievement", "Labour With Tax"), ("Achievement", "Parts W/O Tax"), ("Gap", "Throughput"), ("Gap", "Labour"), ("Gap", "Parts")])

                        # 🎨 OPTIMIZED CSS 
                        custom_css = [
                            {'selector': 'table', 'props': [('border', '2px solid black'), ('border-collapse', 'collapse'), ('width', '100%'), ('margin-bottom', '20px'), ('font-size', '13px')]},
                            {'selector': 'th', 'props': [('background-color', 'royalblue'), ('color', 'white'), ('font-weight', 'bold'), ('text-align', 'center'), ('border', '1px solid black'), ('padding', '6px 4px'), ('white-space', 'nowrap')]},
                            {'selector': 'td', 'props': [('text-align', 'center'), ('border', '1px solid black'), ('padding', '6px 4px'), ('font-weight', '500'), ('white-space', 'nowrap')]}
                        ]

                        # --- SIDE BY SIDE LAYOUT START ---
                        col1, col2 = st.columns([1.6, 1]) 

                        with col1:
                            st.markdown("### 📊 Monthly Achievement")
                            fmt_map = {c: "{:.2f}" for c in df_final.columns[1:]}
                            for c in df_final.columns:
                                if 'Throughput' in c[1]: fmt_map[c] = "{:.0f}"
                            
                            styled_t1 = df_final.style.format(fmt_map).apply(lambda r: ['background-color: #8db4e2; font-weight: bold; color: black']*len(r) if r[("", "Date")] == "Total" else ['']*len(r), axis=1).set_table_styles(custom_css).hide(axis="index")
                            
                            st.markdown(f'<div style="overflow-x: auto; max-width: 100%;">{styled_t1.to_html()}</div>', unsafe_allow_html=True)

                        with col2:
                            st.markdown("### 💰 Monthly Revenue Summary")
                            df_month = df_2026[df_2026['Date_DT'].isin([d.date() for d in report_dates])]
                            try: tiles_sum = counts.get("Complete", 0) + counts.get("In Process", 0) + counts.get("On Hold", 0)
                            except NameError: tiles_sum = 0

                            # Vehicle Rec logic (Agar select kiya hua mahina current month nahi hai, to in-process wali tiles ko ignore karna better hai, par filhal wahi rakha hai)
                            v_comp = df_data["A_Throughput"].sum()
                            v_rec = v_comp + tiles_sum
                            t_bikes = len(df_month[df_month['Description_Clean'].str.contains("oil filter", na=False) & df_month['VIN_Clean'].str.contains("WB", na=False)])

                            def sm(t): return df_month['Type_Clean'].str.contains('invoice', na=False) if t == 'inv' else df_month['Customer Type_Clean'].str.contains(t, na=False)
                            def get_rev(m_it, m_ct):
                                mask = df_month['Item Type_Clean'].str.contains(m_it, na=False) & sm('inv') & sm(m_ct)
                                return df_month.loc[mask, 'LabAmt'].sum(), df_month.loc[mask, 'PartAmt'].sum(), df_month.loc[mask, 'DnpAmt'].sum()

                            mask_ws_labour = df_month['Item Type_Clean'].str.contains('labor', na=False) & df_month['Customer Type_Clean'].str.contains('external', na=False)
                            ws_l_wt = df_month.loc[mask_ws_labour, 'LabAmt'].sum()
                            ws_l_wot = df_month.loc[mask_ws_labour, 'PartAmt'].sum()
                            war_l1_wt, war_l1_wot, _ = get_rev('warranty handling', 'inv') 
                            war_l2_wt, war_l2_wot, _ = get_rev('labor', 'warranty')
                            war_l_wt, war_l_wot = war_l1_wt + war_l2_wt, war_l1_wot + war_l2_wot
                            
                            m_spare = df_month['Item Type_Clean'].str.contains('part', na=False)
                            m_ws_p = df_month['Part Type_Clean'].str.contains('1 - normal|9 - miscelleneous|local', na=False)
                            ws_s_wt = df_month.loc[m_spare & sm('external') & m_ws_p, 'LabAmt'].sum()
                            ws_s_wot = df_month.loc[m_spare & sm('external') & m_ws_p, 'PartAmt'].sum()
                            ws_s_dnp = df_month.loc[m_spare & sm('external') & m_ws_p, 'DnpAmt'].sum()
                            
                            war_s_wt = df_month.loc[m_spare & sm('warranty'), 'LabAmt'].sum()
                            war_s_wot = df_month.loc[m_spare & sm('warranty'), 'PartAmt'].sum()
                            war_s_dnp = df_month.loc[m_spare & sm('warranty'), 'DnpAmt'].sum()
                            
                            m_acc = df_month['Part Type_Clean'].str.contains('7 - bmw|3 - retrofit', na=False)
                            acc_wt = df_month.loc[m_spare & sm('external') & m_acc, 'LabAmt'].sum()
                            acc_wot = df_month.loc[m_spare & sm('external') & m_acc, 'PartAmt'].sum()
                            acc_dnp = df_month.loc[m_spare & sm('external') & m_acc, 'DnpAmt'].sum()

                            def f_v(v, dnp_l=False): return "" if (dnp_l or v==0) else f"₹ {v:,.2f}"
                            
                            summary_data = [
                                {"Particulars": "Vehicle received", "With Tax": str(int(v_rec)), "Without Tax": "", "DNP": ""},
                                {"Particulars": "Vehicle completed", "With Tax": str(int(v_comp)), "Without Tax": "", "DNP": ""},
                                {"Particulars": "Total Bikes serviced", "With Tax": str(int(t_bikes)), "Without Tax": "", "DNP": ""},
                                {"Particulars": "W/S LABOUR", "With Tax": f_v(ws_l_wt), "Without Tax": f_v(ws_l_wot), "DNP": ""},
                                {"Particulars": "WARRANTY LABOUR", "With Tax": f_v(war_l_wt), "Without Tax": f_v(war_l_wot), "DNP": ""},
                                {"Particulars": "TOTAL LABOUR", "With Tax": f_v(ws_l_wt+war_l_wt), "Without Tax": f_v(ws_l_wot+war_l_wot), "DNP": ""},
                                {"Particulars": "W/S SPARE", "With Tax": f_v(ws_s_wt), "Without Tax": f_v(ws_s_wot), "DNP": f_v(ws_s_dnp)},
                                {"Particulars": "WARRANTY SPARE", "With Tax": f_v(war_s_wt), "Without Tax": f_v(war_s_wot), "DNP": f_v(war_s_dnp)},
                                {"Particulars": "TOTAL SPARE", "With Tax": f_v(ws_s_wt+war_s_wt), "Without Tax": f_v(ws_s_wot+war_s_wot), "DNP": f_v(ws_s_dnp+war_s_dnp)},
                                {"Particulars": "TOTAL Acc & GG", "With Tax": f_v(acc_wt), "Without Tax": f_v(acc_wot), "DNP": f_v(acc_dnp)},
                                {"Particulars": "GRAND REVENUE", "With Tax": f_v(ws_l_wt+war_l_wt+ws_s_wt+war_s_wt+acc_wt), "Without Tax": f_v(ws_l_wot+war_l_wot+ws_s_wot+war_s_wot+acc_wot), "DNP": f_v(ws_s_dnp+war_s_dnp+acc_dnp)}
                            ]
                            
                            styled_t2 = pd.DataFrame(summary_data).style.apply(lambda r: ['background-color: #f0f2f6; font-weight: bold; color: black']*len(r) if "TOTAL" in r["Particulars"] or "GRAND" in r["Particulars"] else ['']*len(r), axis=1).set_table_styles(custom_css).hide(axis="index")
                            
                            st.markdown(f'<div style="overflow-x: auto; max-width: 100%;">{styled_t2.to_html()}</div>', unsafe_allow_html=True)
                        
                    else: st.error("Document Date not found.")
                else: st.warning("Sheet is empty.")
            else: st.error("Connection Error.")

    # ==========================================
        # 📈 TAB 6-2: YTD & YoY COMPARISON (2022 - 2026)
        # ==========================================
        with st_rev_2:
            st.subheader("📊 Year on Year Comparison")
            
            col_v, col_p, _ = st.columns([1, 1, 2])
            view_type = col_v.selectbox("View Type", ["Monthly", "Quarterly", "Half Yearly", "Annually"])
            
            # 🛠️ FIX: Yahan default_index ko pehle hi 0 assign kar dein
            default_index = 0 
            
            if view_type == "Monthly":
                options = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
                current_month_str = datetime.now().strftime("%b")
                if current_month_str in options:
                    default_index = options.index(current_month_str)
            elif view_type == "Quarterly":
                options = ["Q1 (Jan-Mar)", "Q2 (Apr-Jun)", "Q3 (Jul-Sep)", "Q4 (Oct-Dec)"]
            elif view_type == "Half Yearly":
                options = ["H1 (Jan-Jun)", "H2 (Jul-Dec)"]
            else:
                options = ["Full Year"]
                
            selected_period = col_p.selectbox("Select Period", options, index=default_index)

            # --- 1. DIRECT DATA FETCH ---
            url_master = "https://docs.google.com/spreadsheets/d/1XYZWk94-RCqABmzCd40-v6RLlXcwbOaxPMK9jE8x9No/edit"
            
            if 'summary_sheet_data' not in st.session_state:
                with st.spinner("Fetching Pre-Calculated Summary from Sheet..."):
                    ws_sum = connect_sheet_by_url(url_master, "SImple Format Without Tax")
                    if ws_sum:
                        st.session_state['summary_sheet_data'] = ws_sum.get_all_values()
                    else:
                        st.session_state['summary_sheet_data'] = []

            raw_data = st.session_state.get('summary_sheet_data', [])

            year_cols = {'2022': (2, 3), '2023': (4, 5), '2024': (6, 7), '2025': (8, 9), '2026': (10, 11)}
            years_to_fetch = list(year_cols.keys())
            month_names = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
            
            parsed_data = {y: {m: {'spare': 0, 'acc': 0, 'war': 0, 'lab': 0, 'jc': 0, 'srv': 0} for m in month_names} for y in years_to_fetch}

            # --- 2. PARSING LOGIC ---
            def clean_val(v):
                try:
                    return float(str(v).replace(',', '').replace('₹', '').strip())
                except:
                    return 0.0

            def extract_jc_srv(txt):
                match = re.search(r'JC Closed\s*-\s*(\d+)[,\s]*Service Done\s*-\s*(\d+)', str(txt), re.IGNORECASE)
                if match:
                    return int(match.group(1)), int(match.group(2))
                return 0, 0

            current_month = None

            if raw_data:
                for row in raw_data:
                    if len(row) < 12: continue
                    
                    m_val = str(row[0]).strip()
                    if m_val in month_names:
                        current_month = m_val
                        
                    category = str(row[1]).strip()
                    
                    if current_month and category in ['Spare Parts', 'ACC & GG', 'Warranty Spare Parts', 'Labor']:
                        for y, (val_col, txt_col) in year_cols.items():
                            amt = clean_val(row[val_col])
                            
                            if category == 'Spare Parts': parsed_data[y][current_month]['spare'] = amt
                            elif category == 'ACC & GG': parsed_data[y][current_month]['acc'] = amt
                            elif category == 'Warranty Spare Parts': parsed_data[y][current_month]['war'] = amt
                            elif category == 'Labor':
                                parsed_data[y][current_month]['lab'] = amt
                                jc_count, srv_count = extract_jc_srv(row[txt_col])
                                parsed_data[y][current_month]['jc'] = jc_count
                                parsed_data[y][current_month]['srv'] = srv_count

            # --- 3. AGGREGATING FOR SELECTED PERIOD ---
            month_map_nums = {"Jan": [1], "Feb": [2], "Mar": [3], "Apr": [4], "May": [5], "Jun": [6], "Jul": [7], "Aug": [8], "Sep": [9], "Oct": [10], "Nov": [11], "Dec": [12],
                        "Q1 (Jan-Mar)": [1,2,3], "Q2 (Apr-Jun)": [4,5,6], "Q3 (Jul-Sep)": [7,8,9], "Q4 (Oct-Dec)": [10,11,12],
                        "H1 (Jan-Jun)": [1,2,3,4,5,6], "H2 (Jul-Dec)": [7,8,9,10,11,12], "Full Year": list(range(1, 13))}
            
            target_month_nums = month_map_nums.get(selected_period, [])
            num_to_name = {i+1: name for i, name in enumerate(month_names)}
            target_month_names = [num_to_name[m] for m in target_month_nums]

            agg_data = {}
            for y in years_to_fetch:
                s_sum = sum(parsed_data[y][m]['spare'] for m in target_month_names)
                a_sum = sum(parsed_data[y][m]['acc'] for m in target_month_names)
                w_sum = sum(parsed_data[y][m]['war'] for m in target_month_names)
                l_sum = sum(parsed_data[y][m]['lab'] for m in target_month_names)
                
                jc_sum = sum(parsed_data[y][m]['jc'] for m in target_month_names)
                srv_sum = sum(parsed_data[y][m]['srv'] for m in target_month_names)
                
                agg_data[y] = {
                    'spare': s_sum, 'acc': a_sum, 'war_spare': w_sum, 'labor': l_sum,
                    'total_parts': s_sum + a_sum + w_sum,
                    'jc_text': f"JC Closed - {jc_sum}, Service Done - {srv_sum}" if (jc_sum > 0 or srv_sum > 0) else ""
                }

            # ==========================================
            # 🚀 4. DYNAMIC METRIC CARDS
            # ==========================================
            if raw_data:
                cy_rev = agg_data['2026']['total_parts'] + agg_data['2026']['labor']
                py_rev = agg_data['2025']['total_parts'] + agg_data['2025']['labor']
                
                if py_rev > 0:
                    delta_pct = ((cy_rev - py_rev) / py_rev) * 100
                    delta_str = f"{delta_pct:.2f}%"
                else:
                    delta_str = "0%"

                st.markdown("<br>", unsafe_allow_html=True)
                col_y1, col_y2 = st.columns(2)
                
                col_y1.metric(f"Current Year (2026) - {selected_period}", f"₹ {cy_rev:,.0f}", delta=delta_str)
                col_y2.metric(f"Previous Year (2025) - {selected_period}", f"₹ {py_rev:,.0f}")
                st.markdown("<br>", unsafe_allow_html=True)

                st.write("📊 **Comparison Summary**")

                # --- 5. HTML TABLE GENERATION ---
                clr_spare, clr_acc, clr_war, clr_lab, clr_head = "#e2f0d9", "#dae3f3", "#fce4d6", "#ffffff", "#f0f2f6"
                label_disp = selected_period.split()[0]

                html_table = f"""
                <div style="overflow-x: auto; margin-top: 10px;">
                    <table style="border-collapse: collapse; width: 100%; text-align: center; font-family: sans-serif; font-size: 13px; border: 2px solid black;">
                        <thead>
                            <tr>
                                <th style="border: 1px solid black; background-color: {clr_head}; width: 50px;"></th>
                                <th style="border: 1px solid black; background-color: {clr_head}; width: 150px;"></th>
                """
                for y in years_to_fetch:
                    html_table += f'<th colspan="2" style="border: 1px solid black; background-color: {clr_head}; font-weight: bold; font-size: 15px; padding: 8px;">{y}</th>'
                
                html_table += "</tr></thead><tbody>"

                def get_row(label, key, color, has_total=False):
                    row_html = f'<tr>'
                    if label == "Spare Parts":
                        row_html += f'<td rowspan="4" style="border: 1px solid black; font-weight: bold; vertical-align: middle; background: white;">{label_disp}</td>'
                    
                    row_html += f'<td style="border: 1px solid black; background-color: {color}; text-align: left; padding-left: 10px; font-weight: 500;">{label}</td>'
                    for y in years_to_fetch:
                        val = agg_data[y][key]
                        row_html += f'<td style="border: 1px solid black; background-color: {color}; padding: 6px;">{val:,.0f}</td>'
                        if has_total:
                            tot = agg_data[y]['total_parts']
                            row_html += f'<td rowspan="3" style="border: 1px solid black; background-color: #e2f0d9; font-weight: bold; vertical-align: middle; font-size: 14px;">{tot:,.0f}</td>'
                    return row_html + "</tr>"

                html_table += get_row("Spare Parts", "spare", clr_spare, True)
                html_table += get_row("ACC & GG", "acc", clr_acc)
                html_table += get_row("Warranty Spare Parts", "war_spare", clr_war)
                
                html_table += f'<tr><td style="border: 1px solid black; background-color: {clr_lab}; text-align: left; padding-left: 10px;">Labor</td>'
                for y in years_to_fetch:
                    val, txt = agg_data[y]['labor'], agg_data[y]['jc_text']
                    html_table += f'<td style="border: 1px solid black; background-color: {clr_lab}; padding: 6px;">{val:,.0f}</td>'
                    html_table += f'<td style="border: 1px solid black; background-color: {clr_lab}; font-size: 11px;">{txt}</td>'
                html_table += "</tr></tbody></table></div>"

                st.markdown(html_table, unsafe_allow_html=True)

                # ==========================================
                # 📉 6. GAP ANALYSIS TABLE (N2:U6)
                # ==========================================
                if len(raw_data) >= 6:
                    st.markdown("<br><h3>📉 Overall Gap Analysis</h3>", unsafe_allow_html=True)
                    
                    gap_html = f"""
                    <div style="overflow-x: auto; margin-top: 10px;">
                        <table style="border-collapse: collapse; width: 100%; font-family: sans-serif; font-size: 14px; border: 2px solid black;">
                            <thead>
                                <tr>
                    """
                    
                    header_row = raw_data[1]
                    if len(header_row) < 21:
                        header_row += [""] * (21 - len(header_row))
                        
                    for i, h in enumerate(header_row[13:21]):
                        bg = "#ffff00" if i == 5 else "#f0f2f6" 
                        gap_html += f'<th style="border: 1px solid black; background-color: {bg}; font-weight: bold; padding: 8px; text-align: center;">{h}</th>'
                        
                    gap_html += "</tr></thead><tbody>"
                    
                    for row in raw_data[2:6]:
                        if len(row) < 21:
                            row += [""] * (21 - len(row))
                            
                        gap_html += "<tr>"
                        for i, val in enumerate(row[13:21]):
                            bg = "#ffff00" if i == 5 else "#ffffff"
                            align = "left" if i == 0 else "right"
                            fw = "bold" if i == 0 else "normal"
                            gap_html += f'<td style="border: 1px solid black; background-color: {bg}; padding: 6px; text-align: {align}; font-weight: {fw};">{val}</td>'
                        gap_html += "</tr>"
                        
                    gap_html += "</tbody></table></div>"
                    st.markdown(gap_html, unsafe_allow_html=True)

                # ==========================================
                # 📈 7. YoY COMPARISON % TABLE (N10:R14)
                # ==========================================
                if len(raw_data) >= 14:
                    st.markdown("<br><h3>📈 YoY Comparison (%)</h3>", unsafe_allow_html=True)
                    
                    comp_html = f"""
                    <div style="overflow-x: auto; margin-top: 10px;">
                        <table style="border-collapse: collapse; width: 100%; font-family: sans-serif; font-size: 14px; border: 2px solid black;">
                            <thead>
                                <tr>
                    """
                    
                    # Headers from Row 10 (index 9)
                    header_row_2 = raw_data[9]
                    if len(header_row_2) < 18:
                        header_row_2 += [""] * (18 - len(header_row_2))
                        
                    for i, h in enumerate(header_row_2[13:18]):
                        comp_html += f'<th style="border: 1px solid black; background-color: #f0f2f6; font-weight: bold; padding: 8px; text-align: center;">{h}</th>'
                        
                    comp_html += "</tr></thead><tbody>"
                    
                    # Data Rows from Row 11 to 14 (index 10 to 14)
                    for row in raw_data[10:14]:
                        if len(row) < 18:
                            row += [""] * (18 - len(row))
                            
                        comp_html += "<tr>"
                        for i, val in enumerate(row[13:18]):
                            align = "left" if i == 0 else "right"
                            fw = "bold" if i == 0 else "normal"
                            
                            # Smart Color Logic for Percentages
                            text_color = "black"
                            if "%" in str(val):
                                if "-" in str(val):
                                    text_color = "#d9534f" # Red for drop
                                else:
                                    text_color = "#28a745" # Green for growth
                                    
                            comp_html += f'<td style="border: 1px solid black; background-color: #ffffff; padding: 6px; text-align: {align}; font-weight: {fw}; color: {text_color};">{val}</td>'
                        comp_html += "</tr>"
                        
                    comp_html += "</tbody></table></div>"
                    st.markdown(comp_html, unsafe_allow_html=True)

            else:
                st.error("Sheet data could not be loaded. Please check the sheet name and access.")
