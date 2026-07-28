import io
from datetime import datetime, timedelta
import pandas as pd
import openpyxl
import requests
import streamlit as st
import gspread
from gspread_formatting import cellFormat, color, format_cell_range
from oauth2client.service_account import ServiceAccountCredentials

st.set_page_config(
    page_title="Kalyx Despatch Terminal", page_icon="🏍️", layout="centered"
)

# Your Live Google Sheet Details
SHEET_ID = "1YZgwm11scCWdiH5-9RxBVe3kEk9obfNtlH4frIYZGSg"
SHEET_EXPORT_URL = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=xlsx"
SHEET_NAME = "Despatch"

# --- 1. STATE MANAGEMENT ---
# This controls the sequential step-by-step screens
if "page" not in st.session_state:
    st.session_state.page = "setup"
if "optimized_route" not in st.session_state:
    st.session_state.optimized_route = []
if "current_stop" not in st.session_state:
    st.session_state.current_stop = 0

# --- 2. DATA LOADING LOGIC ---
@st.cache_data(ttl=30, show_spinner=False)
def load_pending_despatch_tasks():
    response = requests.get(SHEET_EXPORT_URL)
    if response.status_code != 200:
        raise Exception("Failed to download Google Sheet.")

    wb = openpyxl.load_workbook(io.BytesIO(response.content), data_only=True)
    sheet = wb[SHEET_NAME]
    pending_tasks = []
    
    # 7-day date threshold
    one_week_ago = datetime.now() - timedelta(days=7)

    for row_idx in range(2, sheet.max_row + 1):
        # Date filtering (Column A = 1)
        cell_date = sheet.cell(row=row_idx, column=1).value
        task_date = None
        if isinstance(cell_date, datetime):
            task_date = cell_date
        elif isinstance(cell_date, str):
            try:
                task_date = pd.to_datetime(cell_date)
            except:
                pass

        if task_date:
            if hasattr(task_date, "to_pydatetime"):
                task_date = task_date.to_pydatetime()
            if task_date < one_week_ago:
                continue

        # Check Cell Fill Color (Column S = 19). ONLY ORANGE = Complete.
        cell_address = sheet.cell(row=row_idx, column=19)
        fill = cell_address.fill
        is_orange = False
        if fill and fill.start_color and fill.start_color.rgb:
            cell_color = str(fill.start_color.rgb).upper()
            if any(c in cell_color for c in ["A500", "9900", "FFC000", "ED7D31", "F290"]):
                is_orange = True

        # Blue, White, Blank, Green are treated as PENDING
        if not is_orange:
            address = cell_address.value
            company = sheet.cell(row=row_idx, column=17).value
            
            if company and address and str(address).strip() != "nan":
                pending_tasks.append({
                    "id": row_idx,
                    "job_name": str(sheet.cell(row=row_idx, column=5).value or "-"),
                    "task_type": str(sheet.cell(row=row_idx, column=6).value or "Despatch"),
                    "area": str(sheet.cell(row=row_idx, column=14).value or "Unassigned"),
                    "box": float(sheet.cell(row=row_idx, column=16).value or 0),
                    "company": str(company),
                    "address": str(address),
                    "target_date": str(sheet.cell(row=row_idx, column=23).value or "-")[:10],
                    "container": float(sheet.cell(row=row_idx, column=24).value or 0),
                    "bag": float(sheet.cell(row=row_idx, column=25).value or 0),
                    "envelope": float(sheet.cell(row=row_idx, column=26).value or 0),
                    "transport": str(sheet.cell(row=row_idx, column=27).value or "Motorcycle"),
                    "department": str(sheet.cell(row=row_idx, column=37).value or "-"),
                    "client": str(sheet.cell(row=row_idx, column=38).value or "N/A"),
                    "phone": str(sheet.cell(row=row_idx, column=39).value or ""),
                })
    return pending_tasks

# --- 3. GOOGLE SHEETS API WRITEBACK ---
def mark_row_completed_in_sheets(row_idx):
    """Connects to Google Sheets via Service Account and paints the row Orange."""
    if "gcp_service_account" not in st.secrets:
        st.warning("⚠️ No Google Service Account detected in Secrets! (Row completed locally but not on cloud).")
        return False
        
    try:
        credentials = ServiceAccountCredentials.from_json_keyfile_dict(
            dict(st.secrets["gcp_service_account"]),
            scopes=["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
        )
        gc = gspread.authorize(credentials)
        sh = gc.open_by_key(SHEET_ID)
        worksheet = sh.worksheet(SHEET_NAME)
        
        # Mark Columns A to AM (1 to 39) with ORANGE color (#FFA500)
        fmt = cellFormat(backgroundColor=color(1.0, 0.65, 0.0))
        format_cell_range(worksheet, f"A{row_idx}:AM{row_idx}", fmt)
        return True
    except Exception as e:
        st.error(f"Failed to update Google Sheet: {e}")
        return False

# --- 4. OPTIMIZATION LOGIC ---
def smart_directional_sort(stops_list):
    def get_directional_score(address):
        addr = str(address).lower()
        if "perai jaya" in addr: return 10
        elif "pauh" in addr: return 20
        elif "seberang jaya" in addr or "todak" in addr: return 30
        elif "chain ferry" in addr: return 40
        elif "ong yi how" in addr or "teras" in addr: return 50
        elif "selayang" in addr or "sungai dua" in addr: return 60
        elif "jawi" in addr or "valdor" in addr: return 80
        elif "simpang ampat" in addr or "hijauan hills" in addr: return 85
        elif "teguh" in addr or "tinggi" in addr: return 90
        elif "bukit minyak" in addr: return 100
        elif "juru" in addr: return 110
        elif "bukit kecil" in addr: return 120
        elif "maju jaya" in addr: return 130
        else: return 140
    return sorted(stops_list, key=lambda x: get_directional_score(x["address"]))

# --- 5. UI: SETUP PAGE ---
if st.session_state.page == "setup":
    st.title("🏍️ Kalyx Route Setup")
    
    with st.spinner("Fetching pending jobs from Google Sheets..."):
        try:
            tasks = load_pending_despatch_tasks()
        except Exception as e:
            st.error(f"Error loading sheet: {e}")
            st.stop()

    if not tasks:
        st.success("🎉 All recent tasks are completed!")
        st.stop()

    df_tasks = pd.DataFrame(tasks)
    
    col1, col2 = st.columns(2)
    with col1:
        selected_area = st.selectbox("📍 Filter Area:", ["All Areas"] + list(df_tasks["area"].unique()))
    with col2:
        selected_transport = st.selectbox("🚗/🏍️ Transport:", ["All", "Car", "Motorcycle"])
        
    # Input Time in UI overriding sheet!
    shift_time = st.time_input("⏰ Select Target Shift Time (Optional):")
    st.session_state.shift_time_input = shift_time.strftime("%I:%M %p")

    filtered_df = df_tasks
    if selected_area != "All Areas":
        filtered_df = filtered_df[filtered_df["area"] == selected_area]
    if selected_transport != "All":
        filtered_df = filtered_df[filtered_df["transport"].str.contains(selected_transport, case=False)]

    st.markdown(f"### 📋 Select Stops for Current Run ({len(filtered_df)} Available)")
    
    selected_stops = []
    for _, row in filtered_df.iterrows():
        c_check, c_text = st.columns([0.1, 0.9])
        with c_check:
            is_checked = st.checkbox("Select", key=f"chk_{row['id']}")
        with c_text:
            transport_icon = "🚗" if "car" in str(row["transport"]).lower() else "🏍️"
            st.markdown(f"**{row['company']}** {transport_icon} - {row['task_type']} <br><small>📍 {row['address']}</small>", unsafe_allow_html=True)
        if is_checked:
            selected_stops.append(row.to_dict())

    if st.button("🚀 Optimize & Start Loop", type="primary", use_container_width=True):
        if not selected_stops:
            st.warning("Please select at least one stop.")
        else:
            st.session_state.optimized_route = smart_directional_sort(selected_stops)
            st.session_state.current_stop = 0
            st.session_state.page = "route"
            st.rerun()

# --- 6. UI: SINGLE STOP SEQUENTIAL NAVIGATION ---
elif st.session_state.page == "route":
    total_stops = len(st.session_state.optimized_route)
    current_idx = st.session_state.current_stop
    stop = st.session_state.optimized_route[current_idx]

    # Header / Progress
    st.progress((current_idx) / total_stops)
    st.markdown(f"<h3 style='text-align: center; color: #ff6b6b;'>🏁 Stop {current_idx + 1} of {total_stops}</h3>", unsafe_allow_html=True)
    st.markdown("---")

    # Stop Details Card
    transport_icon = "🚗" if "car" in str(stop["transport"]).lower() else "🏍️"
    
    st.title(f"{stop['company']}")
    st.markdown(f"**Task:** {stop['task_type']} {transport_icon} | **Target:** {st.session_state.shift_time_input} | **Dept:** {stop['department']}")
    st.markdown(f"📍 **Address:** {stop['address']}")
    
    # Parcel Logic
    items = []
    if stop["box"] > 0: items.append(f"📦 {int(stop['box'])} Box")
    if stop["container"] > 0: items.append(f"🗃️ {int(stop['container'])} Container")
    if stop["bag"] > 0: items.append(f"🛍️ {int(stop['bag'])} Bag")
    if stop["envelope"] > 0: items.append(f"✉️ {int(stop['envelope'])} Envelope")
    if items:
        st.info(" | ".join(items))
    
    # Contact & Navigation Buttons
    clean_phone = str(stop["phone"]).replace(" ", "").replace("-", "") if stop["phone"] else ""
    col_nav1, col_nav2 = st.columns(2)
    
    with col_nav1:
        if clean_phone and clean_phone != "nan":
            st.markdown(f"<a href='tel:{clean_phone}'><button style='width: 100%; padding:15px; border-radius:8px; border:1px solid #1E90FF; background:transparent; color:#1E90FF;'>📞 Call {stop['client']}</button></a>", unsafe_allow_html=True)
        else:
            st.write(f"📞 Contact: {stop['client']} (No phone)")
            
    with col_nav2:
        maps_url = f"https://www.google.com/maps/dir/?api=1&destination={stop['address'].replace(' ', '+')}"
        st.markdown(f"<a href='{maps_url}' target='_blank'><button style='width: 100%; padding:15px; border-radius:8px; border:1px solid #28a745; background:transparent; color:#28a745;'>🗺️ Google Maps</button></a>", unsafe_allow_html=True)
    
    st.write("")
    st.write("")

    # Mark Complete Button (Advances to next sequence)
    if st.button("✅ Mark Complete & Load Next Job", type="primary", use_container_width=True):
        with st.spinner("Writing to Google Sheets..."):
            # Call API to mark orange
            mark_row_completed_in_sheets(stop["id"])
            
            # Auto-advance
            st.session_state.current_stop += 1
            if st.session_state.current_stop >= total_stops:
                st.session_state.page = "finished"
            st.rerun()

# --- 7. UI: FINISHED SHIFT ---
elif st.session_state.page == "finished":
    st.balloons()
    st.title("🎉 Route Complete!")
    st.success("All selected jobs have been delivered, and Google Sheets is automatically updated.")
    
    if st.button("🔄 Start New Shift", use_container_width=True):
        st.session_state.page = "setup"
        # Reset cache so it pulls the newly orange-painted rows as complete
        st.cache_data.clear() 
        st.rerun()
