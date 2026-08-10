import streamlit as st

st.set_page_config(
    page_title="Banjul Port Operations",
    page_icon="🚢",
    layout="wide"
)

st.title("🚢 Banjul Port Operations")

st.write("Welcome to Alport Banjul Operations System")

vessel_name = st.text_input("Vessel Name")

eta = st.text_input("ETA")

berth = st.selectbox(
    "Berth",
    ["Berth 1", "Berth 3A", "Berth 3B", "Anchorage"]
)

discharge = st.number_input(
    "Discharge Containers",
    min_value=0,
    step=1
)

loading = st.number_input(
    "Loading Containers",
    min_value=0,
    step=1
)

if st.button("Calculate Total Moves"):
    total = discharge + loading

    st.success(f"Total Moves: {total}")

    st.write("Vessel:", vessel_name)
    st.write("ETA:", eta)
    st.write("Berth:", berth)
