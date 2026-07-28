import io
from datetime import datetime, timedelta, time
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

# --- GOOGLE SHEET CONFIGURATION ---
SHEET_ID = "1YZgwm11scCWdiH5-9RxBVe3kEk9obfNtlH4frIYZGSg"
SHEET_EXPORT_URL = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=xlsx"
SHEET_NAME = "Despatch"

# --- 1. SESSION STATE MANAGEMENT ---
if "page" not in st.session_state:
    st.session_state.page = "setup"
if "optimized_route" not in st.session_state:
    st.session_state.optimized_route = []
if "current_stop" not in st.session_state:
    st.session_state.current_stop = 0
if "optimization_time" not in st.session_state:
    st.session_state.optimization_time = None

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
        # Column A: Date check
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

        # Column S (19): Cell Fill Color check. ONLY ORANGE = Complete.
        cell_address = sheet.cell(row=row_idx, column=19)
        fill = cell_address.fill
        is_orange = False
        if fill and fill.start_color and fill.start_color.rgb:
            cell_color = str(fill.start_color.rgb).upper()
            if any(c in cell_color for c in ["A500", "9900", "FFC000", "ED7D31", "F290"]):
                is_orange = True

        # Blue, White, Green, etc., are treated as PENDING
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
                    "sheet_time_slot": str(sheet.cell(row=row_idx, column=28).value or "").strip(),
                    "department": str(sheet.cell(row=row_idx, column=37).value or "-"),
                    "client": str(sheet.cell(row=row_idx, column=38).value or "N/A"),
                    "phone": str(sheet.cell(row=row_idx, column=39).value or ""),
                })
    return pending_tasks

# --- 3. GOOGLE SHEETS WRITEBACK (MARK ORANGE) ---
def mark_row_completed_in_sheets(row_idx):
    if "gcp_service_account" not in st.secrets:
        return False
        
    try:
        credentials = ServiceAccountCredentials.from_json_keyfile_dict(
            dict(st.secrets["gcp_service_account"]),
            scopes=["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
        )
        gc = gspread.authorize(credentials)
        sh = gc.open_by_key(SHEET_ID)
        worksheet = sh.worksheet(SHEET_NAME)
        
        # Paint row Orange (#FFA500)
        fmt = cellFormat(backgroundColor=color(1.0, 0.65, 0.0))
        format_cell_range(worksheet, f"A{row_idx}:AM{row_idx}", fmt)
        return True
    except Exception as e:
        st.error(f"Error updating Google Sheet: {e}")
        return False

# --- 4. ROUTE OPTIMIZATION ALGORITHMS ---
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

def optimize_route(stops_list, start_time):
    # Check if any stop has a custom time input
    timed_stops = [s for s in stops_list if s.get("custom_time") is not None]
    
    # CASE 1: No specific time input provided -> Standard directional optimization
    if not timed_stops:
        return sorted(stops_list, key=lambda x: get_directional_score(x["address"]))
    
    # CASE 2: Build optimization anchored around jobs with target times
    def get_sort_key(item):
        custom_t = item.get("custom_time")
        if custom_t is not None:
            time_val = custom_t.hour * 60 + custom_t.minute
            return (0, time_val, get_directional_score(item["address"]))
        else:
            return (1, 0, get_directional_score(item["address"]))
            
    return sorted(stops_list, key=get_sort_key)


# --- 5. UI PAGE: SETUP & TASK SELECTION ---
if st.session_state.page == "setup":
    st.title("🏍️ Kalyx Despatch Route Planner")
    
    with st.spinner("Fetching live pending tasks..."):
        try:
            tasks = load_pending_despatch_tasks()
        except Exception as e:
            st.error(f"Error loading sheet: {e}")
            st.stop()

    if not tasks:
        st.success("🎉 All recent tasks are completed!")
        st.stop()

    df_tasks = pd.DataFrame(tasks)
    
    # Area & Transport Filters
    col_f1, col_f2 = st.columns(2)
    with col_f1:
        selected_area = st.selectbox("📍 Filter Area:", ["All Areas"] + list(df_tasks["area"].unique()))
    with col_f2:
        selected_transport = st.selectbox("🚗/🏍️ Transport:", ["All", "Car", "Motorcycle"])

    filtered_df = df_tasks
    if selected_area != "All Areas":
        filtered_df = filtered_df[filtered_df["area"] == selected_area]
    if selected_transport != "All":
        filtered_df = filtered_df[filtered_df["transport"].str.contains(selected_transport, case=False)]

    st.markdown(f"### 📋 Pending Tasks ({len(filtered_df)} Available)")
    st.caption("Select tasks for today's run and optionally set specific target time slots per job:")

    selected_stops = []
    
    # Render Task List with Department, PIC Name, Parcel Counts & Time Slot Input
    for _, row in filtered_df.iterrows():
        row_dict = row.to_dict()
        
        # Build document/parcel count string
        items = []
        if row["box"] > 0: items.append(f"📦 {int(row['box'])} Box")
        if row["container"] > 0: items.append(f"🗃️ {int(row['container'])} Container")
        if row["bag"] > 0: items.append(f"🛍️ {int(row['bag'])} Bag")
        if row["envelope"] > 0: items.append(f"✉️ {int(row['envelope'])} Envelope")
        items_str = " | ".join(items) if items else "No document details"

        c_check, c_details, c_time = st.columns([0.08, 0.64, 0.28])
        
        with c_check:
            is_checked = st.checkbox("Select", key=f"chk_{row['id']}")
            
        with c_details:
            transport_icon = "🚗" if "car" in str(row["transport"]).lower() else "🏍️"
            st.markdown(
                f"**{row['company']}** {transport_icon} - *{row['task_type']}*<br>"
                f"<small>🏢 <b>Dept:</b> {row['department']} | 👤 <b>PIC:</b> {row['client']}</small><br>"
                f"<small>📄 <b>Items:</b> {items_str}</small><br>"
                f"<small>📍 {row['address']}</small>", 
                unsafe_allow_html=True
            )
            
        with c_time:
            # Per-job time slot input option
            has_time = st.checkbox("Set Time", key=f"time_chk_{row['id']}")
            job_time = None
            if has_time:
                job_time = st.time_input(
                    "Target Time", 
                    value=time(9, 0), 
                    key=f"t_val_{row['id']}", 
                    label_visibility="collapsed"
                )
            row_dict["custom_time"] = job_time

        if is_checked:
            selected_stops.append(row_dict)
        st.markdown("---")

    # OPTIMIZE BUTTON
    if st.button("🚀 Optimize & Start Loop", type="primary", use_container_width=True):
        if not selected_stops:
            st.warning("Please select at least one task.")
        else:
            # Capture real-time timestamp
            realtime_now = datetime.now()
            st.session_state.optimization_time = realtime_now
            
            # Run optimization
            st.session_state.optimized_route = optimize_route(selected_stops, realtime_now)
            st.session_state.current_stop = 0
            st.session_state.page = "route"
            st.rerun()

# --- 6. UI PAGE: SINGLE ROUTE DISPLAY & SEQUENTIAL EXECUTION ---
elif st.session_state.page == "route":
    total_stops = len(st.session_state.optimized_route)
    current_idx = st.session_state.current_stop
    stop = st.session_state.optimized_route[current_idx]

    # TOP NAV BAR: STOP COUNTER + HOME BUTTON
    col_nav_top, col_home = st.columns([0.75, 0.25])
    with col_nav_top:
        opt_time_str = st.session_state.optimization_time.strftime("%I:%M %p") if st.session_state.optimization_time else ""
        st.markdown(f"### 🏁 Stop {current_idx + 1} of {total_stops}")
        st.caption(f"Optimized at {opt_time_str}")
    with col_home:
        if st.button("🏠 Home", key="btn_home_route", use_container_width=True):
            st.session_state.page = "setup"
            st.session_state.current_stop = 0
            st.session_state.optimized_route = []
            st.rerun()

    st.progress((current_idx) / total_stops)
    st.markdown("---")

    # Current Job Card Details
    transport_icon = "🚗" if "car" in str(stop["transport"]).lower() else "🏍️"
    
    st.title(f"{stop['company']}")
    
    # Time Badge
    time_badge = f"⏰ Slot: {stop['custom_time'].strftime('%I:%M %p')}" if stop.get("custom_time") else "⏰ Flexible / Anytime"
    st.markdown(f"**Task:** {stop['task_type']} {transport_icon} | {time_badge} | **Dept:** {stop['department']}")
    st.markdown(f"📍 **Address:** {stop['address']}")
    
    # Parcel Details
    items = []
    if stop["box"] > 0: items.append(f"📦 {int(stop['box'])} Box")
    if stop["container"] > 0: items.append(f"🗃️ {int(stop['container'])} Container")
    if stop["bag"] > 0: items.append(f"🛍️ {int(stop['bag'])} Bag")
    if stop["envelope"] > 0: items.append(f"✉️ {int(stop['envelope'])} Envelope")
    if items:
        st.info(" | ".join(items))
    
    # Call & Navigation Buttons
    clean_phone = str(stop["phone"]).replace(" ", "").replace("-", "") if stop["phone"] else ""
    col_c1, col_c2 = st.columns(2)
    
    with col_c1:
        if clean_phone and clean_phone != "nan":
            st.markdown(
                f"<a href='tel:{clean_phone}'><button style='width: 100%; padding:14px; border-radius:8px; border:1px solid #1E90FF; background:transparent; color:#1E90FF; font-weight:bold;'>📞 Call {stop['client']}</button></a>", 
                unsafe_allow_html=True
            )
        else:
            st.write(f"📞 Contact: {stop['client']} (No phone)")
            
    with col_c2:
        maps_url = f"https://www.google.com/maps/dir/?api=1&destination={stop['address'].replace(' ', '+')}"
        st.markdown(
            f"<a href='{maps_url}' target='_blank'><button style='width: 100%; padding:14px; border-radius:8px; border:1px solid #28a745; background:transparent; color:#28a745; font-weight:bold;'>🗺️ Navigate</button></a>", 
            unsafe_allow_html=True
        )
    
    st.write("")
    st.write("")

    # COMPLETE BUTTON -> MARKS ORANGE IN SHEETS & AUTO-SHOWS NEXT JOB
    if st.button("✅ Mark Complete & Show Next Job", type="primary", use_container_width=True):
        with st.spinner("Updating Google Sheets..."):
            # Paint row Orange in Google Sheets
            mark_row_completed_in_sheets(stop["id"])
            
            # Step to next job
            st.session_state.current_stop += 1
            if st.session_state.current_stop >= total_stops:
                st.session_state.page = "finished"
            st.rerun()

# --- 7. UI PAGE: SHIFT COMPLETED ---
elif st.session_state.page == "finished":
    st.balloons()
    st.title("🎉 All Tasks Completed!")
    st.success("All selected jobs for this route have been completed and updated live in Google Sheets.")
    
    col_f1, col_f2 = st.columns(2)
    with col_f1:
        if st.button("🔄 Start New Shift", use_container_width=True):
            st.session_state.page = "setup"
            st.cache_data.clear()
            st.rerun()
    with col_f2:
        if st.button("🏠 Home", key="btn_home_finished", use_container_width=True):
            st.session_state.page = "setup"
            st.cache_data.clear()
            st.rerun()
