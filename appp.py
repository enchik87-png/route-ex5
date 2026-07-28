import streamlit as st

st.set_page_config(
    page_title="EX5 Dynamic Route Optimizer", page_icon="🏍️", layout="centered"
)

st.title("🏍️ EX5 Dynamic Route Optimizer")
st.write(
    "Smart-sequenced for motorcycle traffic flow (outbound primary lane ➔"
    " farthest point ➔ inbound return lane)."
)

raw_data = st.text_area(
    "Paste Raw Details Here:",
    height=200,
    placeholder=(
        "W.G. OOI & ASSOCIATES\t1st Floor, 7004, Jalan Ong Yi How...\tRyan\t04-331"
        " 8332"
    ),
)

start_point = st.text_input(
    "Start Point Address:", "10-G, Jalan Icon City, 14000 Bukit Mertajam, Penang"
)
end_point = st.text_input(
    "End Point Address:", "Machang Bubok, Bukit Mertajam, Pulau Pinang"
)


def smart_directional_sort(stops_list):
  """Sorts stops to ensure smooth directional flow (outbound vs inbound lanes)

  preventing unnecessary mid-road crossings or U-turns.
  """

  def get_directional_score(address):
    addr = address.lower()

    # Phase 1: Outbound Northern / Western Sweep (Left-lane biased)
    if "perai jaya" in addr:
      return 10
    elif "pauh" in addr:
      return 20
    elif "todak" in addr or "seberang jaya" in addr:
      return 30
    elif "chain ferry" in addr:
      return 40

    # Phase 2: Butterworth Loop (Ong Yi How / Teras Jaya outbound vs inbound sides)
    elif "teras jaya" in addr:
      return 50  # Hit first on outbound side
    elif "ong yi how" in addr or "teratai" in addr:
      return (
          52  # Sequenced right after Teras Jaya before turning back inward
      )
    elif "selayang" in addr:
      return 60
    elif "sungai dua" in addr or "kampung teluk" in addr:
      return 70  # Farthest North turnaround point

    # Phase 3: Deep Southern Push (Valdor / Simpang Ampat outbound)
    elif "valdor" in addr or "sungai jawi" in addr:
      return 80  # Farthest South turnaround point
    elif "hijauan hills" in addr or "simpang ampat" in addr:
      return 85

    # Phase 4: Inbound Return Track (Permatang Tinggi, Bukit Minyak, Juru, BM Town)
    elif "teguh" in addr or "permatang tinggi" in addr:
      return 90
    elif "bukit minyak" in addr:
      return 100
    elif "juru" in addr or "simpang juru" in addr:
      return 110
    elif "bukit kecil" in addr:
      return 120
    elif "maju jaya" in addr:
      return 130  # First side of Maju Jaya on inbound pass (e.g. Stop C)
    elif "taman seri maju" in addr or "jalan maju" in addr:
      return (
          135  # Further down Maju stream
      )
    else:
      return 140

  return sorted(stops_list, key=lambda x: get_directional_score(x["address"]))


if st.button("🚀 Generate Side-of-Road Optimized Route", type="primary"):
  if not raw_data.strip():
    st.warning("Please paste some location data first.")
  else:
    lines = raw_data.strip().split("\n")
    parsed_stops = []

    for line in lines:
      if line.strip():
        parts = [p.strip() for p in line.split("\t") if p.strip()]
        if len(parts) >= 2:
          parsed_stops.append({
              "company": parts[0],
              "address": parts[1],
              "contact": parts[2] if len(parts) > 2 else "N/A",
              "phone": parts[3] if len(parts) > 3 else "",
          })
        else:
          parsed_stops.append({
              "company": "Stop Item",
              "address": line.strip(),
              "contact": "See details",
              "phone": "",
          })

    optimized_stops = smart_directional_sort(parsed_stops)

    st.success(
        f"Successfully optimized {len(optimized_stops)} stops keeping traffic"
        " lane flow in mind!"
    )

    st.markdown("### 🛑 START POINT")
    st.info(f"**Address:** {start_point}")
    start_map = f"https://www.google.com/maps/dir/?api=1&destination={start_point.replace(' ', '+')}"
    st.markdown(f"[🗺️ Navigate to Start Point]({start_map})")
    st.markdown("---")

    st.markdown("### 🛑 DIRECTIONAL OPTIMIZED STOPS")

    for idx, stop in enumerate(optimized_stops, 1):
      with st.container():
        st.markdown(f"**Stop {idx}: {stop['company']}**")
        st.write(f"📍 **Address:** {stop['address']}")
        if stop["phone"]:
          st.write(f"📞 **Contact:** {stop['contact']} ({stop['phone']})")
        else:
          st.write(f"📞 **Contact:** {stop['contact']}")

        maps_url = f"https://www.google.com/maps/dir/?api=1&destination={stop['address'].replace(' ', '+')}"
        st.markdown(f"[🗺️ Navigate From Current Location]({maps_url})")
        st.markdown("---")

    st.markdown("### 🏁 END POINT")
    st.info(f"**Address:** {end_point}")
    end_map = f"https://www.google.com/maps/dir/?api=1&destination={end_point.replace(' ', '+')}"
    st.markdown(f"[🗺️ Navigate to End Point]({end_map})")
