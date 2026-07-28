import streamlit as st

st.set_page_config(
    page_title="EX5 Dynamic Route Optimizer", page_icon="🏍️", layout="centered"
)

st.title("🏍️ EX5 Dynamic Route Optimizer")
st.write(
    "Paste your raw list of stops below. The app will break them down into"
    " individual stops with live auto-navigation links!"
)

# Text area for raw paste
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

if st.button("🚀 Process & Generate Route Cards", type="primary"):
  if not raw_data.strip():
    st.warning("Please paste some location data first.")
  else:
    st.success("Route generated successfully!")

    # Display Start Point
    st.markdown("### 🛑 START POINT")
    st.info(f"**Address:** {start_point}")
    start_map = f"https://www.google.com/maps/dir/?api=1&destination={start_point.replace(' ', '+')}"
    st.markdown(f"[🗺️ Navigate to Start Point]({start_map})")
    st.markdown("---")

    # Parse raw text line by line
    lines = raw_data.strip().split("\n")
    valid_stops = []

    for line in lines:
      if line.strip():
        # Split by tabs or multiple spaces if copied from excel/tables
        parts = [p.strip() for p in line.split("\t") if p.strip()]
        if len(parts) >= 2:
          valid_stops.append({
              "company": parts[0],
              "address": parts[1],
              "contact": parts[2] if len(parts) > 2 else "N/A",
              "phone": parts[3] if len(parts) > 3 else "",
          })
        else:
          # Fallback if it's pasted as a single block line
          valid_stops.append(
              {"company": "Stop Item", "address": line.strip(), "contact": "See details", "phone": ""}
          )

    st.markdown(f"### 🛑 STOPS ({len(valid_stops)} Locations)")

    for idx, stop in enumerate(valid_stops, 1):
      with st.container():
        st.markdown(f"**Stop {idx}: {stop['company']}**")
        st.write(f"📍 **Address:** {stop['address']}")
        if stop["phone"]:
          st.write(f"📞 **Contact:** {stop['contact']} ({stop['phone']})")
        else:
          st.write(f"📞 **Contact:** {stop['contact']}")

        # Build clean Google Maps navigation link from current location
        maps_url = f"https://www.google.com/maps/dir/?api=1&destination={stop['address'].replace(' ', '+')}"
        st.markdown(f"[🗺️ Navigate From Current Location]({maps_url})")
        st.markdown("---")

    # Display End Point
    st.markdown("### 🏁 END POINT")
    st.info(f"**Address:** {end_point}")
    end_map = f"https://www.google.com/maps/dir/?api=1&destination={end_point.replace(' ', '+')}"
    st.markdown(f"[🗺️ Navigate to End Point]({end_map})")
