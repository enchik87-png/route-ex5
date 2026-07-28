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

# --- PRESET START & END LOCATIONS ---
PRESET_LOCATIONS = {
    "Kalyx Consultants Sdn Bhd (Office)": {
        "label": "🏢 Kalyx Consultants Sdn Bhd (Office)",
        "address": "Kalyx Consultants Sdn Bhd, Bukit Mertajam",
        "score": 120
    },
    "Machang Bubok (Home)": {
        "label": "🏠 Machang Bubok (Home)",
        "address": "Machang Bubok, Bukit Mertajam",
        "score": 140
    },
    "Taman Sri Serdang, Bertam (Home)": {
        "label": "🏠 Taman Sri Serdang, Bertam (Home)",
        "address": "Taman Sri Serdang, Bertam, Kepala Batas",
        "score": 0
    }
}

# --- 1. SESSION STATE MANAGEMENT ---
if "page" not in st.session_state:
    st.session_state.page = "setup"
if "optimized_route" not in st.session_state:
    st.session_state.optimized_route = []
if "current_stop" not in st.session_state:
    st.session_state.current_stop = 0
if "optimization_time" not in st.session_state:
    st.session_state.optimization_time = None
if "start_point" not in st.session_state:
    st.session_state.start_point = "Kalyx Consultants Sdn Bhd (Office)"
if "end_point" not in st.session_state:
    st.session_state.end_point = "Kalyx Consultants Sdn Bhd (Office)"


# --- 2. KPI AGING CALCULATION LOGIC ---
def calculate_kpi_status(raw_request_date):
    """
    Calculates effective request date using 9:30 AM cut-off logic:
    - Before 9:30 AM -> Same day request
    - After 9:30 AM -> Next day request
    Returns days elapsed, status label, badge HTML, and red status flag.
    """
    now = datetime.now()
    req_dt = None
    
    if isinstance(raw_request_date, datetime):
        req_dt = raw_request_date
    elif isinstance(raw_request_date, str) and raw_request_date.strip():
        try:
            req_dt = pd.to_datetime(raw_request_date).to_pydatetime()
        except:
            req_dt = now

    if not req_dt:
        req_dt = now

    # 9:30 AM Cut-off logic
    if req_dt.time() > time(9, 30):
        effective_date = req_dt.date() + timedelta(days=1)
    else:
        effective_date = req_dt.date()

    days_elapsed = (now.date() - effective_date).days
    days_elapsed = max(0, days_elapsed)

    if days_elapsed >= 3:
        status_key = "Overdue"
        badge_html = f"<span style='background-color:#FF4D4D; color:white; padding:3px 8px; border-radius:4px; font-weight:bold; font-size:12px;'>🚨 OVERDUE ({days_elapsed} Days)</span>"
        is_red = True
    elif days_elapsed == 2:
        status_key = "Due Today (Day 2)"
        badge_html = f"<span style='background-color:#FFC107; color:black; padding:3px 8px; border-radius:4px; font-weight:bold; font-size:12px;'>⚠️ KPI LIMIT (Day 2)</span>"
        is_red = False
    else:
        status_key = "On Time"
        badge_html = f"<span style='background-color:#28A745; color:white; padding:3px 8px; border-radius:4px; font-weight:bold; font-size:12px;'>✅ ON TIME ({days_elapsed} Day)</span>"
        is_red = False

    formatted_req_str = req_dt.strftime("%d/%m/%Y %I:%M %p")
    return days_elapsed, status_key, badge_html, is_red, formatted_req_str


# --- 3. DATA LOADING LOGIC ---
@st.cache_data(ttl=30, show_spinner=False)
def load_pending_despatch_tasks():
    response = requests.get(SHEET_EXPORT_URL)
    if response.status_code != 200:
        raise Exception("Failed to download Google Sheet.")

    wb = openpyxl.load_workbook(io.BytesIO(response.content), data_only=True)
    sheet = wb[SHEET_NAME]
    pending_tasks = []
    
    # 14-day date threshold for sheet scan
    two_weeks_ago = datetime.now() - timedelta(days=14)

    for row_idx in range(2, sheet.max_row + 1):
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
            if task_date < two_weeks_ago:
                continue

        # Column S (19): Cell Fill Color check. ONLY ORANGE = Complete.
        cell_address = sheet.cell(row=row_idx, column=19)
        fill = cell_address.fill
        is_orange = False
        if fill and fill.start_color and fill.start_color.rgb:
            cell_color = str(fill.start_color.rgb).upper()
            if any(c in cell_color for c in ["A500", "9900", "FFC000", "ED7D31", "F290"]):
                is_orange = True

        # Blue, White, Green, etc. treated as PENDING
        if not is_orange:
            address = cell_address.value
            company = sheet.cell(row=row_idx, column=17).value
            
            if company and address and str(address).strip() != "nan":
                days_elapsed, kpi_status, badge_html, is_red, formatted_req_str = calculate_kpi_status(cell_date)

                pending_tasks.append({
                    "id": row_idx,
                    "requested_date_raw": cell_date,
                    "requested_date_str": formatted_req_str,
                    "kpi_days": days_elapsed,
                    "kpi_status": kpi_status,
                    "kpi_badge_html": badge_html,
                    "is_red_overdue": is_red,
                    "pic_name": str(sheet.cell(row=row_idx, column=5).value or "-"),     # Column E: Our Company PIC
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
                    "client": str(sheet.cell(row=row_idx, column=38).value or "N/A"),     # Column AL: Client Name
                    "phone": str(sheet.cell(row=row_idx, column=39).value or ""),
                })
    return pending_tasks


# --- 4. GOOGLE SHEETS WRITEBACK (MARK ORANGE) ---
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


# --- 5. LOCATION SCORING & ROUTE OPTIMIZATION ---
def get_location_score(address):
    addr = str(address).lower()
    if "bertam" in addr or "serdang" in addr or "kepala batas" in addr: return 0
    elif "perai jaya" in addr: return 10
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
    elif "bukit kecil" in addr or "perda" in addr or "kalyx" in addr: return 120
    elif "maju jaya" in addr: return 130
    elif "machang bubok" in addr or "macang bubok" in addr: return 140
    else: return 100

def optimize_route(stops_list, start_key, end_key):
    start_score = PRESET_LOCATIONS.get(start_key, {}).get("score", 120)
    end_score = PRESET_LOCATIONS.get(end_key, {}).get("score", 120)

    def get_sort_key(item):
        custom_t = item.get("custom_time")
        overdue_priority = 0 if item.get("is_red_overdue") else 1
        item_score = get_location_score(item["address"])
        
        # Determine routing direction based on start & end scores
        if start_score == end_score:
            route_score = abs(item_score - start_score)
        elif start_score < end_score:
            route_score = item_score
        else:
            route_score = -item_score

        if custom_t is not None:
            time_val = custom_t.hour * 60 + custom_t.minute
            return (overdue_priority, 0, time_val, route_score)
        else:
            return (overdue_priority, 1, 0, route_score)
            
    return sorted(stops_list, key=get_sort_key)


# --- 6. UI PAGE: SETUP & TASK SELECTION ---
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
    
    # Area, Transport & KPI Filters
    col_f1, col_f2, col_f3 = st.columns(3)
    with col_f1:
        selected_area = st.selectbox("📍 Filter Area:", ["All Areas"] + list(df_tasks["area"].unique()))
    with col_f2:
        selected_transport = st.selectbox("🚗/🏍️ Transport:", ["All", "Car", "Motorcycle"])
    with col_f3:
        selected_kpi = st.selectbox("⏳ KPI Filter:", ["All Tasks", "🚨 Overdue Only", "⚠️ Due Today (Day 2)", "✅ On Time Only"])

    filtered_df = df_tasks
    if selected_area != "All Areas":
        filtered_df = filtered_df[filtered_df["area"] == selected_area]
    if selected_transport != "All":
        filtered_df = filtered_df[filtered_df["transport"].str.contains(selected_transport, case=False)]
    if selected_kpi == "🚨 Overdue Only":
        filtered_df = filtered_df[filtered_df["is_red_overdue"] == True]
    elif selected_kpi == "⚠️ Due Today (Day 2)":
        filtered_df = filtered_df[filtered_df["kpi_status"] == "Due Today (Day 2)"]
    elif selected_kpi == "✅ On Time Only":
        filtered_df = filtered_df[filtered_df["kpi_status"] == "On Time"]

    st.markdown(f"### 📋 Select Tasks for Today ({len(filtered_df)} Available)")

    selected_stops = []
    
    # Render Task List
    for _, row in filtered_df.iterrows():
        row_dict = row.to_dict()
        
        items = []
        if row["box"] > 0: items.append(f"📦 {int(row['box'])} Box")
        if row["container"] > 0: items.append(f"🗃️ {int(row['container'])} Container")
        if row["bag"] > 0: items.append(f"🛍️ {int(row['bag'])} Bag")
        if row["envelope"] > 0: items.append(f"✉️ {int(row['envelope'])} Envelope")
        items_str = " | ".join(items) if items else "No document details"

        border_style = "border-left: 5px solid #FF4D4D; padding-left: 8px;" if row["is_red_overdue"] else ""

        c_check, c_details, c_time = st.columns([0.08, 0.64, 0.28])
        
        with c_check:
            is_checked = st.checkbox("Select", key=f"chk_{row['id']}")
            
        with c_details:
            transport_icon = "🚗" if "car" in str(row["transport"]).lower() else "🏍️"
            st.markdown(
                f"<div style='{border_style}'>"
                f"<b>{row['company']}</b> {transport_icon} - <i>{row['task_type']}</i> {row['kpi_badge_html']}<br>"
                f"<small>📅 <b>Requested:</b> {row['requested_date_str']}</small><br>"
                f"<small>🏢 <b>Dept:</b> {row['department']} | 👤 <b>PIC:</b> {row['pic_name']} | 🤝 <b>Client:</b> {row['client']}</small><br>"
                f"<small>📄 <b>Items:</b> {items_str}</small><br>"
                f"<small>📍 {row['address']}</small>"
                f"</div>", 
                unsafe_allow_html=True
            )
            
        with c_time:
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

    # START & END POINT SELECTION BEFORE OPTIMIZING
    st.markdown("### 🗺️ Route Endpoints & Optimization")
    st.caption("Select your starting point and final destination for today's run:")
    
    col_sp, col_ep = st.columns(2)
    options_keys = list(PRESET_LOCATIONS.keys())
    
    with col_sp:
        sel_start = st.selectbox("🚩 Start Point:", options_keys, index=0, key="select_start_pt")
    with col_ep:
        sel_end = st.selectbox("🏁 End Point:", options_keys, index=0, key="select_end_pt")

    # OPTIMIZE BUTTON
    if st.button("🚀 Optimize & Start Loop", type="primary", use_container_width=True):
        if not selected_stops:
            st.warning("Please select at least one task.")
        else:
            realtime_now = datetime.now()
            st.session_state.optimization_time = realtime_now
            st.session_state.start_point = sel_start
            st.session_state.end_point = sel_end
            
            # Run optimization based on start and end point
            st.session_state.optimized_route = optimize_route(selected_stops, sel_start, sel_end)
            st.session_state.current_stop = 0
            st.session_state.page = "route"
            st.rerun()

# --- 7. UI PAGE: ROUTE DISPLAY & EXECUTION ---
elif st.session_state.page == "route":
    total_stops = len(st.session_state.optimized_route)
    current_idx = st.session_state.current_stop
    stop = st.session_state.optimized_route[current_idx]

    # TOP NAV BAR: STOP COUNTER + ROUTE SUMMARY
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
    
    # DISPLAY ROUTE ENDPOINTS
    start_label = PRESET_LOCATIONS.get(st.session_state.start_point, {}).get("label", st.session_state.start_point)
    end_label = PRESET_LOCATIONS.get(st.session_state.end_point, {}).get("label", st.session_state.end_point)
    st.info(f"🚩 **Start:** {start_label}  \n🏁 **End:** {end_label}")
    st.markdown("---")

    # Current Job Card Details
    transport_icon = "🚗" if "car" in str(stop["transport"]).lower() else "🏍️"
    
    st.title(f"{stop['company']}")
    st.markdown(f"{stop['kpi_badge_html']}", unsafe_allow_html=True)
    time_badge = f"⏰ Slot: {stop['custom_time'].strftime('%I:%M %p')}" if stop.get("custom_time") else "⏰ Flexible / Anytime"
    st.markdown(f"**Task:** {stop['task_type']} {transport_icon} | {time_badge}")
    st.markdown(f"📅 **Requested Date:** {stop['requested_date_str']}")
    st.markdown(f"🏢 **Dept:** {stop['department']} | 👤 **Kalyx PIC:** {stop['pic_name']}")
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
                f"<a href='tel:{clean_phone}'><button style='width: 100%; padding:14px; border-radius:8px; border:1px solid #1E90FF; background:transparent; color:#1E90FF; font-weight:bold;'>📞 Call Client ({stop['client']})</button></a>", 
                unsafe_allow_html=True
            )
        else:
            st.write(f"📞 Client Contact: {stop['client']} (No phone)")
            
    with col_c2:
        maps_url = f"https://www.google.com/maps/dir/?api=1&destination={stop['address'].replace(' ', '+')}"
        st.markdown(
            f"<a href='{maps_url}' target='_blank'><button style='width: 100%; padding:14px; border-radius:8px; border:1px solid #28a745; background:transparent; color:#28a745; font-weight:bold;'>🗺️ Navigate</button></a>", 
            unsafe_allow_html=True
        )
    
    st.write("")
    st.write("")

    # COMPLETE BUTTON
    if st.button("✅ Mark Complete & Show Next Job", type="primary", use_container_width=True):
        with st.spinner("Updating Google Sheets..."):
            mark_row_completed_in_sheets(stop["id"])
            st.session_state.current_stop += 1
            if st.session_state.current_stop >= total_stops:
                st.session_state.page = "finished"
            st.rerun()

# --- 8. UI PAGE: SHIFT COMPLETED ---
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
