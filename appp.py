from datetime import datetime, timedelta
import io
import openpyxl
import pandas as pd
import requests
import streamlit as st

st.set_page_config(
    page_title="Kalyx Despatch Route Planner", page_icon="🏍️", layout="centered"
)

st.title("🏍️ Kalyx Despatch Route Optimizer")
st.write(
    "Fetching pending tasks (only Orange = Complete) from your Google Sheet"
    " 'Despatch' for the past 1 week with full parcel & transport details!"
)

# Your live Google Sheet export URL
SHEET_EXPORT_URL = "https://docs.google.com/spreadsheets/d/1YZgwm11scCWdiH5-9RxBVe3kEk9obfNtlH4frIYZGSg/export?format=xlsx"


@st.cache_data(ttl=60)
def load_pending_despatch_tasks():
  response = requests.get(SHEET_EXPORT_URL)
  if response.status_code != 200:
    raise Exception("Failed to download Google Sheet. Check sharing permissions.")

  wb = openpyxl.load_workbook(io.BytesIO(response.content), data_only=True)
  if "Despatch" not in wb.sheetnames:
    raise Exception("Sheet named 'Despatch' not found in workbook.")

  sheet = wb["Despatch"]
  pending_tasks = []

  # Define date threshold: exactly 1 week (7 days) ago from today
  today = datetime.now()
  one_week_ago = today - timedelta(days=7)

  # Start from row 2 (skipping header)
  for row_idx in range(2, sheet.max_row + 1):
    # 1. Check Date in Column A
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
        continue  # Skip tasks older than 1 week

    # 2. Check Cell Fill in Column S (Address) for completion status (Only Orange = Complete)
    cell_address = sheet.cell(row=row_idx, column=19)  # Column S
    fill = cell_address.fill

    is_orange = False
    if fill and fill.start_color and fill.start_color.rgb:
      color = str(fill.start_color.rgb).upper()
      if any(c in color for c in ["A500", "9900", "FFC000", "ED7D31", "F290"]):
        is_orange = True

    # If NOT orange, it is pending
    if not is_orange:
      # Mapping columns requested:
      # Col A: Date (handled above)
      # Col E: Job request name (5)
      # Col F: Task type - Collect / Send / Collect Send / List out (6)
      # Col N: Area (14)
      # Col P: Box qty (16)
      # Col Q: Company (17)
      # Col S: Address (19)
      # Col W: Request to be complete date (23)
      # Col X: Container qty (24)
      # Col Y: Bag qty (25)
      # Col Z: Envelope qty (26)
      # Col AA: Transport mode - Car / Motorcycle (27)
      # Col AK: Department (37 -> wait, AK is column 37? Let's check Excel index: A=1... Z=26, AA=27, AB=28... AK = 1+10 = 37? Let's verify: A(1) to Z(26), AA(27) AB(28) AC(29) AD(30) AE(31) AF(32) AG(33) AH(34) AI(35) AJ(36) AK(37).)
      # Col AL: Client Name (38)
      # Col AM: Phone Number (39)

      job_name = sheet.cell(row=row_idx, column=5).value
      task_type = sheet.cell(row=row_idx, column=6).value
      area = sheet.cell(row=row_idx, column=14).value
      box_qty = sheet.cell(row=row_idx, column=16).value
      company = sheet.cell(row=row_idx, column=17).value
      address = cell_address.value
      target_date = sheet.cell(row=row_idx, column=23).value
      container_qty = sheet.cell(row=row_idx, column=24).value
      bag_qty = sheet.cell(row=row_idx, column=25).value
      env_qty = sheet.cell(row=row_idx, column=26).value
      transport_mode = sheet.cell(row=row_idx, column=27).value
      department = sheet.cell(row=row_idx, column=37).value
      client = sheet.cell(row=row_idx, column=38).value
      phone = sheet.cell(row=row_idx, column=39).value

      if company and address and str(address).strip() != "nan":
        pending_tasks.append({
            "id": row_idx,
            "job_name": str(job_name) if job_name else "-",
            "task_type": str(task_type) if task_type else "Despatch",
            "area": str(area) if area else "Unassigned",
            "company": str(company),
            "address": str(address),
            "box": box_qty if box_qty else 0,
            "container": container_qty if container_qty else 0,
            "bag": bag_qty if bag_qty else 0,
            "envelope": env_qty if env_qty else 0,
            "transport": str(transport_mode) if transport_mode else "Motorcycle",
            "target_date": str(target_date)[:10] if target_date else "-",
            "department": str(department) if department else "-",
            "client": str(client) if client else "N/A",
            "phone": str(phone) if phone else "",
            "date": (
                task_date.strftime("%Y-%m-%d") if task_date else "Unknown"
            ),
        })

  return pending_tasks


try:
  tasks = load_pending_despatch_tasks()
except Exception as e:
  st.error(f"Error loading sheet: {e}")
  st.stop()

if not tasks:
  st.success(
      "🎉 All tasks from the past week are completed (only orange rows found)!"
  )
  st.stop()

df_tasks = pd.DataFrame(tasks)

# Filters
col_f1, col_f2 = st.columns(2)
with col_f1:
  selected_area = st.selectbox(
      "📍 Filter by Area:",
      options=["All Areas"] + list(df_tasks["area"].unique()),
  )
with col_f2:
  selected_transport = st.selectbox(
      "🚗/🏍️ Filter Transport Mode:",
      options=["All", "Car", "Motorcycle"],
  )

filtered_df = df_tasks
if selected_area != "All Areas":
  filtered_df = filtered_df[filtered_df["area"] == selected_area]
if selected_transport != "All":
  filtered_df = filtered_df[
      filtered_df["transport"].str.contains(
          selected_transport, case=False, na=False
      )
  ]

st.markdown(
    f"### 📋 Pending Stops for the Past Week ({len(filtered_df)} available)"
)

selected_stops = []
for idx, row in filtered_df.iterrows():
  # Build item summary string
  items = []
  if row["box"] and float(row["box"]) > 0:
    items.append(f"📦 {row['box']} Box(es)")
  if row["container"] and float(row["container"]) > 0:
    items.append(f"🗃️ {row['container']} Container(s)")
  if row["bag"] and float(row["bag"]) > 0:
    items.append(f"🛍️ {row['bag']} Bag(s)")
  if row["envelope"] and float(row["envelope"]) > 0:
    items.append(f"✉️ {row['envelope']} Envelope(s)")
  item_str = " | ".join(items) if items else "No items specified"

  col1, col2 = st.columns([0.1, 0.9])
  with col1:
    is_checked = st.checkbox(
        "Select", key=f"chk_{row['id']}", label_visibility="collapsed"
    )
  with col2:
    transport_icon = (
        "🚗" if "car" in str(row["transport"]).lower() else "🏍️"
    )
    st.markdown(
        f"**{row['company']}** ({row['task_type']}) {transport_icon}<br><small>🏷️"
        f" Job: {row['job_name']} | 🏢 Dept: {row['department']}<br>📦"
        f" {item_str}<br>📅 Target Date: {row['target_date']} | 📍"
        f" {row['address']} | 📞 {row['client']} ({row['phone']})</small>",
        unsafe_allow_html=True,
    )
  if is_checked:
    selected_stops.append(row)

start_point = st.text_input(
    "Start Point Address:", "10-G, Jalan Icon City, 14000 Bukit Mertajam, Penang"
)
end_point = st.text_input(
    "End Point Address:", "Machang Bubok, Bukit Mertajam, Pulau Pinang"
)


def smart_directional_sort(stops_list):
  def get_directional_score(address):
    addr = str(address).lower()
    if "perai jaya" in addr:
      return 10
    elif "pauh" in addr:
      return 20
    elif "todak" in addr or "seberang jaya" in addr:
      return 30
    elif "chain ferry" in addr:
      return 40
    elif "teras jaya" in addr:
      return 50
    elif "ong yi how" in addr or "teratai" in addr:
      return 52
    elif "selayang" in addr:
      return 60
    elif "sungai dua" in addr or "kampung teluk" in addr:
      return 70
    elif "valdor" in addr or "sungai jawi" in addr:
      return 80
    elif "hijauan hills" in addr or "simpang ampat" in addr:
      return 85
    elif "teguh" in addr or "permatang tinggi" in addr:
      return 90
    elif "bukit minyak" in addr:
      return 100
    elif "juru" in addr or "simpang juru" in addr:
      return 110
    elif "bukit kecil" in addr:
      return 120
    elif "maju jaya" in addr:
      return 130
    elif "taman seri maju" in addr or "jalan maju" in addr:
      return 135
    else:
      return 140

  return sorted(stops_list, key=lambda x: get_directional_score(x["address"]))


if st.button("🚀 Optimize Selected Route Loop", type="primary"):
  if not selected_stops:
    st.warning("Please check at least one stop above to generate your route.")
  else:
    optimized_stops = smart_directional_sort(selected_stops)

    st.success(
        f"Route successfully optimized for {len(optimized_stops)} selected"
        " stops!"
    )

    st.markdown("### 🛑 START POINT")
    st.info(f"**Address:** {start_point}")
    start_map = f"https://www.google.com/maps/dir/?api=1&destination={start_point.replace(' ', '+')}"
    st.markdown(f"[🗺️ Navigate to Start Point]({start_map})")
    st.markdown("---")

    st.markdown("### 🛑 OPTIMIZED STOPS")

    for idx, stop in enumerate(optimized_stops, 1):
      transport_icon = "🚗" if "car" in str(stop["transport"]).lower() else "🏍️"
      with st.container():
        st.markdown(
            f"**Stop {idx}: {stop['company']}** — *{stop['task_type']}* "
            f"{transport_icon} <small>(Target: {stop['target_date']})</small>"
        )
        st.write(
            f"🏷️ **Job:** {stop['job_name']} | 🏢 **Dept:**"
            f" {stop['department']}"
        )
        st.write(f"📍 **Address:** {stop['address']}")

        items = []
        if stop["box"] and float(stop["box"]) > 0:
          items.append(f"📦 {stop['box']} Box(es)")
        if stop["container"] and float(stop["container"]) > 0:
          items.append(f"🗃️ {stop['container']} Container(s)")
        if stop["bag"] and float(stop["bag"]) > 0:
          items.append(f"🛍️ {stop['bag']} Bag(s)")
        if stop["envelope"] and float(stop["envelope"]) > 0:
          items.append(f"✉️ {stop['envelope']} Envelope(s)")
        if items:
          st.info(" | ".join(items))

        clean_phone = (
            str(stop["phone"]).replace(" ", "").replace("-", "")
            if stop["phone"]
            else ""
        )
        if clean_phone and clean_phone != "nan":
          st.markdown(
              f"📞 **Contact:** {stop['client']} — [Call"
              f" {stop['phone']}](tel:{clean_phone})"
          )
        else:
          st.markdown(f"📞 **Contact:** {stop['client']} (No phone)")

        maps_url = f"https://www.google.com/maps/dir/?api=1&destination={stop['address'].replace(' ', '+')}"
        st.markdown(f"[🗺️ Navigate From Current Location]({maps_url})")
        st.markdown("---")

    st.markdown("### 🏁 END POINT")
    st.info(f"**Address:** {end_point}")
    end_map = f"https://www.google.com/maps/dir/?api=1&destination={end_point.replace(' ', '+')}"
    st.markdown(f"[🗺️ Navigate to End Point]({end_map})")
