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
# 2. PRO CSS (NO CHANGE)
# ==========================================
st.markdown("""
<style>
.stTabs [data-baseweb="tab-list"] { gap: 15px; background-color: #f0f2f6; padding: 10px; border-radius: 12px; }
.stTabs [data-baseweb="tab"] { 
    height: 55px; background-color: white; border: 2px solid #0056b3 !important; 
    border-radius: 10px; color: #0056b3; font-weight: bold; padding: 0 25px;
}
.stTabs [aria-selected="true"] { background-color: #0056b3 !important; color: white !important; }

.stExpander { border: 1px solid #ddd !important; border-radius: 12px !important; margin-bottom: 15px !important; }

button[key^="up_"] {
    height: 75px !important; background-color: #0056b3 !important; color: white !important;
    font-size: 22px !important; font-weight: 900 !important; border-radius: 15px !important;
}

[data-testid="stVerticalBlockBorderWrapper"] > div > div {
    min-height: 250px !important; max-height: 300px !important; overflow-y: auto !important;
}
</style>
""", unsafe_allow_html=True)

st.title("🏁 Munich Motorrad Workshop Management")

tab_ws, tab_app = st.tabs(["🔧 WORKSHOP MANAGER", "📅 APPOINTMENTS"])

# ------------------------------------------
# TAB 1: WORKSHOP
# ------------------------------------------
with tab_ws:
    ws_sheet = connect_sheet("Digital JC")

    if ws_sheet is None:
        st.stop()   # 🔥 IMPORTANT FIX

    try:
        ws_data = ws_sheet.get_all_values()

        if len(ws_data) > 1:
            all_items = [{"row_idx": i+2, "data": r} for i, r in enumerate(ws_data[1:])]

            sub1, sub2 = st.tabs(["📂 Open JCs", "✅ History (Delivered)"])

            # ---------------- OPEN ----------------
            with sub1:
                s_o = st.text_input("🔍 Search Open JCs", key="ws_so")

                open_list = [i for i in all_items if len(i['data']) > 18 and i['data'][18] != "Delivered"]

                def workshop_sort(item):
                    status = item['data'][18].strip() if len(item['data']) > 18 else ""
                    return {"On Hold":1,"In Process":2,"Complete":3}.get(status,4)

                open_list.sort(key=workshop_sort)

                for item in open_list:
                    r = item['data']
                    idx = item['row_idx']
                    stat = r[18].strip() if len(r) > 18 else "In Process"

                    if stat == "On Hold":
                        display_stat, color = f"🔴 {stat}", "#FF4B4B"
                    elif stat == "Complete":
                        display_stat, color = f"🟢 {stat}", "#28a745"
                    else:
                        display_stat, color = f"🟠 {stat}", "#FF8C00"

                    if s_o.lower() in str(r).lower():
                        with st.expander(f"{r[1]} | {r[5]} | {display_stat}"):

                            st.markdown(f"<h2 style='color:{color}'>{display_stat}</h2>", unsafe_allow_html=True)

                            c1, c2, c3 = st.columns(3)

                            with c1:
                                st.markdown(f"**👤 Customer:** {r[1]}")
                                st.markdown(f"**📱 Phone:** {r[2]}")
                                st.markdown(f"**🆔 VIN:** {r[8]}")

                            with c2:
                                st.markdown(f"**🛠️ Staff:** {r[9]}")
                                st.markdown(f"**🛣️ Odo:** {r[7] if len(r)>7 else 'N/A'}")
                                st.info(f"{r[15] if len(r)>15 else 'N/A'}")

                            with c3:
                                n_rem = st.text_area("Remark", r[17] if len(r)>17 else "", key=f"r_{idx}")
                                opts = ["On Hold","In Process","Complete","Delivered"]

                                n_stat = st.selectbox("Status", opts, index=opts.index(stat) if stat in opts else 1, key=f"s_{idx}")

                                if st.button("SAVE CHANGES", key=f"up_{idx}", use_container_width=True):
                                    ws_sheet.update_cell(idx, 18, n_rem)
                                    ws_sheet.update_cell(idx, 19, n_stat)
                                    st.rerun()

            # ---------------- HISTORY ----------------
            with sub2:
                s_h = st.text_input("🔍 Search Delivered", key="ws_sh")

                history_list = [i for i in all_items if len(i['data']) > 18 and i['data'][18] == "Delivered"]

                for item in history_list:
                    r = item['data']
                    if s_h.lower() in str(r).lower():
                        with st.expander(f"✅ {r[1]} | {r[5]} | Delivered"):

                            c1, c2, c3 = st.columns(3)

                            with c1:
                                st.markdown(f"**👤 {r[1]}**")
                                st.markdown(f"📱 {r[2]}")
                                st.markdown(f"📅 {r[0]}")

                            with c2:
                                st.markdown(f"VIN: {r[8]}")
                                st.markdown(f"Staff: {r[9]}")

                            with c3:
                                st.success(r[17] if len(r)>17 else "No Remark")

    except Exception as e:
        st.error(f"Workshop Error: {e}")


# ------------------------------------------
# TAB 2: APPOINTMENTS
# ------------------------------------------
with tab_app:
    app_sheet = connect_sheet("Appointments")

    if app_sheet is None:
        st.stop()   # 🔥 IMPORTANT FIX

    try:
        app_vals = app_sheet.get_all_values()

        if len(app_vals) > 1:
            df = pd.DataFrame([
                {"row": i+2, "N": r[1], "D": r[3], "R": r[7] if len(r)>7 else ""}
                for i, r in enumerate(app_vals[1:])
            ])

            df['D'] = pd.to_datetime(df['D'], errors='coerce')
            df = df.dropna(subset=['D'])

            y = st.selectbox("Year", [2025, 2026], 1)
            mn = st.selectbox("Month", list(calendar.month_name)[1:], datetime.now().month-1)
            midx = list(calendar.month_name).index(mn)

            cal = calendar.monthcalendar(y, midx)

            for week in cal:
                cols = st.columns(7)
                for i, day in enumerate(week):
                    if day != 0:
                        with cols[i]:
                            st.markdown(f"**{day}**")

                            day_data = df[(df['D'].dt.day == day) & (df['D'].dt.month == midx)]

                            for _, row in day_data.iterrows():
                                is_rep = "Bike Reported" in str(row['R'])

                                if st.button(f"{'✅' if is_rep else '⭕'} {row['N']}", key=f"trig_{row['row']}"):
                                    app_sheet.update_cell(row['row'], 8, "Bike Reported" if not is_rep else "")
                                    st.rerun()

    except Exception as e:
        st.error(f"Calendar Error: {e}")
