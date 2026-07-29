import io
import time as time_module
import urllib.parse
from datetime import datetime, timedelta, time
import pandas as pd
import openpyxl
import requests
import streamlit as st
import gspread
from gspread_formatting import cellFormat, color, format_cell_range
from oauth2client.service_account import ServiceAccountCredentials

# --- NEW IMPORTS FOR GOOGLE OR-TOOLS ---
from ortools.constraint_solver import routing_enums_pb2
from ortools.constraint_solver import pywrapcp

st.set_page_config(
    page_title="Kalyx Despatch Terminal", page_icon="🏍️", layout="centered"
)

# --- GOOGLE SHEET CONFIGURATION ---
SHEET_ID = "1YZgwm11scCWdiH5-9RxBVe3kEk9obfNtlH4frIYZGSg"
SHEET_EXPORT_URL = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=xlsx"
SHEET_NAME = "Despatch"

# --- PRESET START & END LOCATIONS (WITH GPS COORDINATES) ---
# Coordinates are formatted as (Longitude, Latitude) for OSRM API
PRESET_LOCATIONS = {
    "Kalyx Consultants Sdn Bhd (Office)": {
        "label": "🏢 Kalyx Consultants (Icon City)",
        "coords": (100.4435, 5.3456)  # Approx Icon City, Bukit Mertajam
    },
    "Machang Bubok (Home)": {
        "label": "🏠 Machang Bubok (Home)",
        "coords": (100.5100, 5.3300) 
    },
    "Taman Sri Serdang, Bertam (Home)": {
        "label": "🏠 Taman Sri Serdang, Bertam (Home)",
        "coords": (100.4480, 5.5180) 
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

        cell_address = sheet.cell(row=row_idx, column=19)
        fill = cell_address.fill
        is_orange = False
        if fill and fill.start_color and fill.start_color.rgb:
            cell_color = str(fill.start_color.rgb).upper()
            if any(c in cell_color for c in ["A500", "9900", "FFC000", "ED7D31", "F290"]):
                is_orange = True

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
                    "pic_name": str(sheet.cell(row=row_idx, column=5).value or "-"),
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


# --- 4. GOOGLE SHEETS WRITEBACK ---
def mark_row_completed_in_sheets(row_idx):
    if "gcp_service_account" not in st.secrets:
        return "no_secrets"
        
    try:
        credentials = ServiceAccountCredentials.from_json_keyfile_dict(
            dict(st.secrets["gcp_service_account"]),
            scopes=["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
        )
        gc = gspread.authorize(credentials)
        sh = gc.open_by_key(SHEET_ID)
        worksheet = sh.worksheet(SHEET_NAME)
        
        fmt = cellFormat(backgroundColor=color(1.0, 0.65, 0.0))
        format_cell_range(worksheet, f"A{row_idx}:AM{row_idx}", fmt)
        return True
    except Exception:
        return False


# --- 5. GPS GEOCODING & OR-TOOLS OPTIMIZATION ---
@st.cache_data(ttl=3600, show_spinner=False)
def get_coordinates(address):
    """Converts a text address into GPS (Longitude, Latitude) using free OpenStreetMap"""
    # Append Penang to ensure it stays local
    search_query = f"{address}, Penang, Malaysia"
    url = f"https://nominatim.openstreetmap.org/search?q={urllib.parse.quote(search_query)}&format=json&limit=1"
    headers = {"User-Agent": "KalyxDespatchApp/1.0"}
    
    try:
        resp = requests.get(url, headers=headers, timeout=5)
        if resp.status_code == 200 and len(resp.json()) > 0:
            data = resp.json()[0]
            return (float(data["lon"]), float(data["lat"]))
    except:
        pass
    return None

def get_osrm_matrix(coords_list):
    """Fetches real driving durations (seconds) between all points using OSRM"""
    coords_str = ";".join([f"{lon},{lat}" for lon, lat in coords_list])
    url = f"http://router.project-osrm.org/table/v1/driving/{coords_str}?annotations=duration"
    try:
        resp = requests.get(url, timeout=10)
        data = resp.json()
        if data.get("code") == "Ok":
            return data["durations"]
    except Exception as e:
        print(f"OSRM Error: {e}")
    return None

def optimize_route_osrm(stops_list, start_key, end_key):
    """Uses Google OR-Tools AI to find the fastest physical road route"""
    start_coords = PRESET_LOCATIONS[start_key]["coords"]
    end_coords = PRESET_LOCATIONS[end_key]["coords"]
    
    # 1. Geocode all addresses
    valid_stops = []
    unmapped_stops = []
    
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    for i, stop in enumerate(stops_list):
        status_text.text(f"Geocoding {i+1}/{len(stops_list)}: {stop['company']}...")
        coords = get_coordinates(stop["address"])
        time_module.sleep(1) # Be nice to the free Nominatim API server (1 req/sec)
        
        if coords:
            stop["coords"] = coords
            valid_stops.append(stop)
        else:
            unmapped_stops.append(stop)
            
        progress_bar.progress((i + 1) / len(stops_list))
    
    status_text.empty()
    progress_bar.empty()

    if not valid_stops:
        # If absolutely nothing geocoded, just return the raw list
        return unmapped_stops

    # 2. Build the coordinate list: [START, ...STOPS..., END]
    all_coords = [start_coords] + [s["coords"] for s in valid_stops] + [end_coords]
    
    # 3. Get real road driving matrix
    status_text.text("Calculating Penang road traffic routes...")
    time_matrix = get_osrm_matrix(all_coords)
    status_text.empty()
    
    if not time_matrix:
        st.warning("⚠️ Could not connect to OSRM road network. Falling back to default sorting.")
        return valid_stops + unmapped_stops

    # 4. Google OR-Tools Setup
    num_locations = len(all_coords)
    start_index = 0
    end_index = num_locations - 1
    
    manager = pywrapcp.RoutingIndexManager(num_locations, 1, [start_index], [end_index])
    routing = pywrapcp.RoutingModel(manager)

    def time_callback(from_index, to_index):
        from_node = manager.IndexToNode(from_index)
        to_node = manager.IndexToNode(to_index)
        # OR-Tools requires integers
        return int(time_matrix[from_node][to_node])

    transit_callback_index = routing.RegisterTransitCallback(time_callback)
    routing.SetArcCostEvaluatorOfAllVehicles(transit_callback_index)

    search_parameters = pywrapcp.DefaultRoutingSearchParameters()
    search_parameters.first_solution_strategy = routing_enums_pb2.FirstSolutionStrategy.PATH_CHEAPEST_ARC

    # 5. Solve for the fastest physical route!
    solution = routing.SolveWithParameters(search_parameters)

    optimized_stops = []
    if solution:
        index = routing.Start(0)
        index = solution.Value(routing.NextVar(index)) # Skip Start Point
        
        while not routing.IsEnd(index):
            node_index = manager.IndexToNode(index)
            # Subtract 1 because valid_stops array doesn't include the Start point
            optimized_stops.append(valid_stops[node_index - 1])
            index = solution.Value(routing.NextVar(index))
    else:
        optimized_stops = valid_stops

    # Append any unmapped stops to the end so they aren't lost
    return optimized_stops + unmapped_stops


# --- 6. UI PAGE: SETUP & TASK SELECTION ---
if st.session_state.page == "setup":
    st.title("🏍️ Kalyx AI Despatch Route Planner")
    
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

    st.markdown("### 🗺️ Route Endpoints & Optimization")
    st.caption("Select your starting point and final destination for today's run:")
    
    col_sp, col_ep = st.columns(2)
    options_keys = list(PRESET_LOCATIONS.keys())
    
    with col_sp:
        sel_start = st.selectbox("🚩 Start Point:", options_keys, index=0, key="select_start_pt")
    with col_ep:
        sel_end = st.selectbox("🏁 End Point:", options_keys, index=0, key="select_end_pt")

    if st.button("🚀 Calculate AI Road Route", type="primary", use_container_width=True):
        if not selected_stops:
            st.warning("Please select at least one task.")
        else:
            st.session_state.optimization_time = datetime.now()
            st.session_state.start_point = sel_start
            st.session_state.end_point = sel_end
            
            # Run the new AI Optimization
            with st.spinner("Analyzing maps and traffic constraints..."):
                st.session_state.optimized_route = optimize_route_osrm(selected_stops, sel_start, sel_end)
            
            st.session_state.current_stop = 0
            st.session_state.page = "route"
            st.rerun()


# --- 7. UI PAGE: ROUTE DISPLAY & EXECUTION ---
elif st.session_state.page == "route":
    total_stops = len(st.session_state.optimized_route)
    current_idx = st.session_state.current_stop
    stop = st.session_state.optimized_route[current_idx]

    col_nav_top, col_home = st.columns([0.75, 0.25])
    with col_nav_top:
        opt_time_str = st.session_state.optimization_time.strftime("%I:%M %p") if st.session_state.optimization_time else ""
        st.markdown(f"### 🏁 Stop {current_idx + 1} of {total_stops}")
        st.caption(f"Optimized at {opt_time_str} via OR-Tools")
    with col_home:
        if st.button("🏠 Home", key="btn_home_route", use_container_width=True):
            st.session_state.page = "setup"
            st.session_state.current_stop = 0
            st.session_state.optimized_route = []
            st.rerun()

    st.progress((current_idx) / total_stops)
    
    start_label = PRESET_LOCATIONS.get(st.session_state.start_point, {}).get("label", st.session_state.start_point)
    end_label = PRESET_LOCATIONS.get(st.session_state.end_point, {}).get("label", st.session_state.end_point)
    st.info(f"🚩 **Start:** {start_label}  \n🏁 **End:** {end_label}")
    
    # Alert if the address couldn't be found by GPS
    if "coords" not in stop:
        st.warning("⚠️ **GPS Warning:** This address couldn't be automatically mapped. It has been moved to the end of your route.")

    st.markdown("---")

    transport_icon = "🚗" if "car" in str(stop["transport"]).lower() else "🏍️"
    
    st.title(f"{stop['company']}")
    st.markdown(f"{stop['kpi_badge_html']}", unsafe_allow_html=True)
    time_badge = f"⏰ Slot: {stop['custom_time'].strftime('%I:%M %p')}" if stop.get("custom_time") else "⏰ Flexible / Anytime"
    st.markdown(f"**Task:** {stop['task_type']} {transport_icon} | {time_badge}")
    st.markdown(f"📅 **Requested Date:** {stop['requested_date_str']}")
    st.markdown(f"🏢 **Dept:** {stop['department']} | 👤 **Kalyx PIC:** {stop['pic_name']}")
    st.markdown(f"📍 **Address:** {stop['address']}")
    
    items = []
    if stop["box"] > 0: items.append(f"📦 {int(stop['box'])} Box")
    if stop["container"] > 0: items.append(f"🗃️ {int(stop['container'])} Container")
    if stop["bag"] > 0: items.append(f"🛍️ {int(stop['bag'])} Bag")
    if stop["envelope"] > 0: items.append(f"✉️ {int(stop['envelope'])} Envelope")
    if items:
        st.info(" | ".join(items))
    
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
        maps_url = f"https://www.google.com/maps/dir/?api=1&destination={urllib.parse.quote(stop['address'])}"
        st.markdown(
            f"<a href='{maps_url}' target='_blank'><button style='width: 100%; padding:14px; border-radius:8px; border:1px solid #28a745; background:transparent; color:#28a745; font-weight:bold;'>🗺️ Navigate</button></a>", 
            unsafe_allow_html=True
        )
    
    st.write("")
    st.write("")

    if st.button("✅ Mark Complete & Show Next Job", type="primary", use_container_width=True):
        with st.spinner("Updating Google Sheets..."):
            success = mark_row_completed_in_sheets(stop["id"])
            
            if success == "no_secrets":
                st.warning("⚠️ Google Sheets API not connected yet. Moving to next job in the app only.")
                st.session_state.current_stop += 1
                if st.session_state.current_stop >= total_stops:
                    st.session_state.page = "finished"
                st.rerun()
                
            elif success == True:
                st.session_state.current_stop += 1
                if st.session_state.current_stop >= total_stops:
                    st.session_state.page = "finished"
                st.rerun()
                
            else:
                st.error("❌ Failed to update Google Sheets. Check your internet connection and try again.")


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
