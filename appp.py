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
    "Fetching pending (white) tasks live from your Google Sheet 'Despatch',"
    " sorting traffic flow, and providing click-to-call & navigation!"
)

# Your live Google Sheet export URL
SHEET_EXPORT_URL = "https://docs.google.com/spreadsheets/d/1YZgwm11scCWdiH5-9RxBVe3kEk9obfNtlH4frIYZGSg/export?format=xlsx"


@st.cache_data(ttl=60)
def load_pending_despatch_tasks():
  # Download the workbook directly into memory
  response = requests.get(SHEET_EXPORT_URL)
  if response.status_code != 200:
    raise Exception("Failed to download Google Sheet. Check sharing permissions.")

  wb = openpyxl.load_workbook(io.BytesIO(response.content), data_only=True)
  if "Despatch" not in wb.sheetnames:
    raise Exception("Sheet named 'Despatch' not found in workbook.")

  sheet = wb["Despatch"]
  pending_tasks = []

  # Start from row 2 (skipping header)
  for row_idx in range(2, sheet.max_row + 1):
    cell_address = sheet.cell(row=row_idx, column=19)  # Column S (Address)
    fill = cell_address.fill

    # Check if the address cell has an orange fill (Complete/Done)
    is_orange = False
    if fill and fill.start_color and fill.start_color.rgb:
      color = str(fill.start_color.rgb).upper()
      if any(c in color for c in ["A500", "9900", "FFC000", "ED7D31", "F290"]):
        is_orange = True

    # If it is NOT orange (i.e., White background = Pending/Not Done)
    if not is_orange:
      area = sheet.cell(row=row_idx, column=14).value  # Column N (Area)
      company = sheet.cell(row=row_idx, column=17).value  # Column Q (Company)
      address = cell_address.value  # Column S (Address)
      client = sheet.cell(row=row_idx, column=37).value  # Column AL (Client Name)
      phone = sheet.cell(
          row=row_idx, column=38
      ).value  # Column AM (Phone Number)

      if company and address and str(address).strip() != "nan":
        pending_tasks.append({
            "id": row_idx,
            "area": str(area) if area else "Unassigned",
            "company": str(company),
            "address": str(address),
            "client": str(client) if client else "N/A",
            "phone": str(phone) if phone else "",
        })

  return pending_tasks


try:
  tasks = load_pending_despatch_tasks()
except Exception as e:
  st.error(f"Error loading sheet: {e}")
  st.stop()

if not tasks:
  st.success("🎉 All tasks are completed (no pending white rows found)!")
  st.stop()

df_tasks = pd.DataFrame(tasks)

# Area Filter Dropdown (Column N)
selected_area = st.selectbox(
    "📍 Filter by Area:", options=["All Areas"] + list(df_tasks["area"].unique())
)

if selected_area != "All Areas":
  filtered_df = df_tasks[df_tasks["area"] == selected_area]
else:
  filtered_df = df_tasks

st.markdown(
    f"### 📋 Pending Stops for Today ({len(filtered_df)} available in area)"
)

selected_stops = []
for idx, row in filtered_df.iterrows():
  col1, col2 = st.columns([0.1, 0.9])
  with col1:
    is_checked = st.checkbox(
        "Select", key=f"chk_{row['id']}", label_visibility="collapsed"
    )
  with col2:
    st.markdown(
        f"**{row['company']}**<br><small>📍 {row['address']} | 📞"
        f" {row['client']} ({row['phone']})</small>",
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
  """Sorts stops to ensure smooth directional flow (outbound vs inbound lanes)"""

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
      with st.container():
        st.markdown(f"**Stop {idx}: {stop['company']}**")
        st.write(f"📍 **Address:** {stop['address']}")

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
