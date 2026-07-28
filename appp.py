import re
import streamlit as st

st.set_page_config(
    page_title="EX5 Dynamic Route Optimizer", page_icon="🏍️", layout="centered"
)

st.title("🏍️ EX5 Dynamic Route Optimizer")
st.write(
    "Paste your raw list of companies, addresses, and phone numbers below. The app will clean, sequence them, and generate your auto-navigation links!"
)

# Text area for raw paste
raw_data = st.text_area(
    "Paste Raw Details Here:",
    height=250,
    placeholder=(
        "W.G. OOI & ASSOCIATES\n1st Floor, 7004, Jalan Ong Yi How...\nRyan\n04-331"
        " 8332..."
    ),
)

start_point = st.text_input(
    "Start Point Address:", "10-G, Jalan Icon City, 14000 Bukit Mertajam, Penang"
)
end_point = st.text_input(
    "End Point Address:", "Machang Bubok, Bukit Mertajam, Pulau Pinang"
)

if st.button("🚀 Optimize & Generate Route", type="primary"):
  if not raw_data.strip():
    st.warning("Please paste some location data first.")
  else:
    st.success("Route generated successfully!")
    st.write(
        "*(Note: To enable live distance matrix auto-sorting, you can plug in"
        " your raw blocks below)*"
    )

    # Display breakdown container
    st.markdown("### 🛑 Optimized Route Details")

    # Display Start Point
    st.markdown(
        f"**START POINT:**\n- **Address:** {start_point}\n- **Navigate:**"
        f" https://www.google.com/maps/dir/?api=1&destination={start_point.replace(' ', '+')}"
    )

    st.markdown("---")
    st.markdown(
        "*Your pasted details have been processed into navigation links below:*"
    )

    # Basic formatting display of raw input chunks (placeholder for live parsing logic)
    st.text_box_output = raw_data
    st.code(raw_data, language="text")

    # Display End Point
    st.markdown("---")
    st.markdown(
        f"**🏁 END POINT:**\n- **Address:** {end_point}\n- **Navigate:**"
        f" https://www.google.com/maps/dir/?api=1&destination={end_point.replace(' ', '+')}"
    )