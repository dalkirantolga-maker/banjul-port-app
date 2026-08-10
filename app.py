import os
import re
from datetime import datetime

import pandas as pd
import streamlit as st


# =========================================================
# PAGE CONFIG
# =========================================================

st.set_page_config(
    page_title="Container Tracking",
    page_icon="📦",
    layout="centered",
    initial_sidebar_state="collapsed"
)

EXCEL_FILE = "containers.xlsx"


# =========================================================
# STYLE
# =========================================================

CSS = """
<style>

    #MainMenu {
        visibility: hidden;
    }

    footer {
        visibility: hidden;
    }

    header {
        visibility: hidden;
    }

    .stApp {
        background-color: #f6f8fb;
    }

    .block-container {
        max-width: 820px;
        padding-top: 2rem;
        padding-bottom: 3rem;
    }

    div[data-testid="stTextInput"] input {
        text-align: center;
        font-size: 24px;
        font-weight: 700;
        letter-spacing: 2px;
        text-transform: uppercase;
        min-height: 58px;
        border-radius: 12px;
    }

    div[data-testid="stSelectbox"] > div > div {
        min-height: 52px;
        border-radius: 12px;
    }

    div.stButton > button {
        min-height: 52px;
        border-radius: 12px;
        font-weight: 700;
        font-size: 16px;
    }

    div[data-testid="stMetric"] {
        background-color: white;
        border: 1px solid #e2e8f0;
        padding: 16px;
        border-radius: 12px;
    }

    div[data-testid="stMetricLabel"] {
        font-size: 13px;
    }

    div[data-testid="stMetricValue"] {
        font-size: 20px;
    }

</style>
"""

st.markdown(CSS, unsafe_allow_html=True)


# =========================================================
# SHIPPING LINE SETTINGS
# =========================================================

LINE_MAP = {
    "MAE": "MAERSK",
    "MAERSK": "MAERSK",

    "CMA": "CMA CGM",
    "CMA CGM": "CMA CGM",
    "CMACGM": "CMA CGM",

    "MSC": "MSC",

    "OBT": "OBT",

    "HAP": "HAPAG-LLOYD",
    "HAPAG": "HAPAG-LLOYD",
    "HAPAG-LLOYD": "HAPAG-LLOYD",

    "ONE": "ONE",

    "COSCO": "COSCO",

    "PIL": "PIL"
}


# =========================================================
# FUNCTIONS
# =========================================================

def normalize_container(value):
    if pd.isna(value):
        return ""

    value = str(value).upper().strip()

    return re.sub(
        r"[^A-Z0-9]",
        "",
        value
    )


def normalize_line(value):
    if pd.isna(value):
        return "-"

    value = str(value).upper().strip()

    if not value:
        return "-"

    return LINE_MAP.get(
        value,
        value
    )


def clean_value(record, column):
    if column not in record.index:
        return "-"

    value = str(record[column]).strip()

    if not value:
        return "-"

    if value.lower() == "nan":
        return "-"

    return value


@st.cache_data(ttl=30)
def load_database(file_name, modified_time):

    df = pd.read_excel(
        file_name,
        dtype=str,
        engine="openpyxl"
    )

    df.columns = [
        str(column).strip().upper()
        for column in df.columns
    ]

    df = df.fillna("")

    if "CONTAINER" not in df.columns:
        raise ValueError(
            "CONTAINER column is missing."
        )

    df["_SEARCH_CONTAINER"] = (
        df["CONTAINER"]
        .apply(normalize_container)
    )

    return df


# =========================================================
# HEADER
# =========================================================

st.caption("ALPORT BANJUL")

st.title("📦 Container Tracking")

st.write(
    "Container identification and shipping line verification"
)


# =========================================================
# LOAD DATABASE
# =========================================================

if not os.path.exists(EXCEL_FILE):

    st.error(
        "Container database is currently unavailable."
    )

    st.write(
        "Please contact the Operations Department."
    )

    st.stop()


try:

    modified_time = os.path.getmtime(
        EXCEL_FILE
    )

    df = load_database(
        EXCEL_FILE,
        modified_time
    )

except Exception:

    st.error(
        "Container database could not be loaded."
    )

    st.write(
        "Please contact the Operations Department."
    )

    st.stop()


update_time = datetime.fromtimestamp(
    modified_time
).strftime("%d/%m/%Y %H:%M")


# =========================================================
# DATABASE STATUS
# =========================================================

status_col1, status_col2 = st.columns(2)

with status_col1:
    st.metric(
        "Containers",
        f"{len(df):,}"
    )

with status_col2:
    st.metric(
        "Last Update",
        update_time
    )


st.write("")


# =========================================================
# SEARCH FORM
# =========================================================

with st.container(border=True):

    st.subheader("Container Search")

    available_lines = sorted(
        {
            normalize_line(agent)
            for agent in df["AGENT"].unique()
            if normalize_line(agent) != "-"
        }
    )

    line_options = [
        "No line selected"
    ] + available_lines

    with st.form(
        "container_search_form",
        clear_on_submit=False
    ):

        selected_line = st.selectbox(
            "Loading Line",
            options=line_options,
            help="Select the shipping line of the vessel you are loading."
        )

        container_input = st.text_input(
            "Container Number",
            placeholder="Example: SEKU6920313",
            max_chars=20
        )

        search_button = st.form_submit_button(
            "🔎 Search Container",
            type="primary",
            use_container_width=True
        )


# =========================================================
# SEARCH
# =========================================================

if search_button:

    search_number = normalize_container(
        container_input
    )

    if not search_number:

        st.warning(
            "Please enter a container number."
        )

        st.stop()


    result = df[
        df["_SEARCH_CONTAINER"] == search_number
    ]


    # =====================================================
    # NOT FOUND
    # =====================================================

    if result.empty:

        st.error("⛔ CONTAINER NOT FOUND")

        st.markdown(
            f"## {search_number}"
        )

        st.error(
            "DO NOT LOAD"
        )

        st.write(
            "This container is not available in the "
            "current database. Please contact Operations "
            "before loading."
        )


    # =====================================================
    # DUPLICATE
    # =====================================================

    elif len(result) > 1:

        st.warning(
            "⚠️ DUPLICATE CONTAINER RECORD"
        )

        st.markdown(
            f"## {search_number}"
        )

        st.write(
            f"{len(result)} records were found for this "
            "container."
        )

        st.error(
            "VERIFY WITH OPERATIONS BEFORE LOADING"
        )


    # =====================================================
    # FOUND
    # =====================================================

    else:

        record = result.iloc[0]

        container = clean_value(
            record,
            "CONTAINER"
        )

        shipping_line = normalize_line(
            clean_value(
                record,
                "AGENT"
            )
        )

        size = clean_value(
            record,
            "SIZE"
        )

        container_type = clean_value(
            record,
            "TYPE"
        )

        status = clean_value(
            record,
            "FULL-MTY"
        )

        area = clean_value(
            record,
            "AREA"
        )

        vessel = clean_value(
            record,
            "VESSEL NAME"
        )

        voyage = clean_value(
            record,
            "VOYAGE NUMBER"
        )

        imo_class = clean_value(
            record,
            "IMO CLS"
        )

        discharge_date = clean_value(
            record,
            "DISCHARGE DATE"
        )


        # =================================================
        # WRONG SHIPPING LINE
        # =================================================

        wrong_line = (
            selected_line != "No line selected"
            and selected_line != shipping_line
        )


        if wrong_line:

            st.error(
                "🛑 STOP — WRONG SHIPPING LINE"
            )

            st.markdown(
                f"# {container}"
            )

            wrong_col1, wrong_col2 = st.columns(2)

            with wrong_col1:
                st.metric(
                    "Container Line",
                    shipping_line
                )

            with wrong_col2:
                st.metric(
                    "Loading Line",
                    selected_line
                )

            st.error(
                "⛔ DO NOT LOAD THIS CONTAINER"
            )

            st.write(
                "The container belongs to a different "
                "shipping line."
            )


        # =================================================
        # CORRECT / VERIFIED
        # =================================================

        else:

            if selected_line != "No line selected":

                st.success(
                    "✅ CORRECT SHIPPING LINE"
                )

            else:

                st.success(
                    "✅ CONTAINER FOUND"
                )


            # CONTAINER NUMBER

            st.markdown(
                f"# {container}"
            )


            # SHIPPING LINE

            with st.container(border=True):

                st.caption(
                    "SHIPPING LINE"
                )

                st.markdown(
                    f"## 🚢 {shipping_line}"
                )


            # =================================================
            # MAIN DETAILS
            # =================================================

            row1_col1, row1_col2 = st.columns(2)

            with row1_col1:

                st.metric(
                    "Size",
                    size
                )

            with row1_col2:

                st.metric(
                    "Type",
                    container_type
                )


            row2_col1, row2_col2 = st.columns(2)

            with row2_col1:

                st.metric(
                    "Status",
                    status
                )

            with row2_col2:

                st.metric(
                    "Area",
                    area
                )


            row3_col1, row3_col2 = st.columns(2)

            with row3_col1:

                st.metric(
                    "Vessel",
                    vessel
                )

            with row3_col2:

                st.metric(
                    "Voyage",
                    voyage
                )


            # =================================================
            # EXTRA INFORMATION
            # =================================================

            if (
                imo_class != "-"
                or discharge_date != "-"
            ):

                with st.expander(
                    "Additional Information"
                ):

                    if imo_class != "-":

                        st.write(
                            "**IMO Class:**",
                            imo_class
                        )

                    if discharge_date != "-":

                        st.write(
                            "**Discharge Date:**",
                            discharge_date
                        )


# =========================================================
# FOOTER
# =========================================================

st.divider()

st.caption(
    "ALPORT BANJUL • Container Verification System • Operations"
)
