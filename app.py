import os
import re
from datetime import datetime

import pandas as pd
import streamlit as st


# =========================================================
# PAGE
# =========================================================

st.set_page_config(
    page_title="ALPORT Container Tracking",
    page_icon="🚢",
    layout="centered",
    initial_sidebar_state="collapsed"
)

EXCEL_FILE = "containers.xlsx"


# =========================================================
# SHIPPING LINE MAP
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
# CSS
# =========================================================

st.html("""
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
    background:
        radial-gradient(circle at top left, rgba(24, 82, 122, 0.10), transparent 30%),
        radial-gradient(circle at top right, rgba(9, 96, 117, 0.08), transparent 28%),
        linear-gradient(180deg, #f8fafc 0%, #eef3f8 100%);
}

.block-container {
    max-width: 850px;
    padding-top: 1.4rem;
    padding-bottom: 3rem;
}

/* HERO */

.hero {
    background:
        linear-gradient(135deg, #071b2d 0%, #0b355b 55%, #14637a 100%);
    border-radius: 24px;
    padding: 38px 25px;
    margin-bottom: 22px;
    text-align: center;
    box-shadow: 0 18px 45px rgba(15, 42, 65, 0.20);
}

.hero-brand {
    color: #9bc3d8;
    font-size: 13px;
    font-weight: 800;
    letter-spacing: 4px;
}

.hero-title {
    color: white;
    font-size: 40px;
    font-weight: 900;
    margin-top: 8px;
}

.hero-sub {
    color: #d7e8f1;
    margin-top: 7px;
    font-size: 14px;
}

/* DATABASE STATUS */

.db-card {
    background: white;
    border: 1px solid #e3e8ef;
    border-radius: 16px;
    padding: 18px 20px;
    box-shadow: 0 6px 20px rgba(15,23,42,0.05);
}

.db-label {
    color: #7b8794;
    font-size: 12px;
    font-weight: 700;
    text-transform: uppercase;
}

.db-value {
    color: #102a43;
    font-size: 22px;
    font-weight: 900;
    margin-top: 5px;
}

/* INPUTS */

div[data-testid="stTextInput"] input {
    height: 62px;
    border-radius: 13px;
    text-align: center;
    font-size: 24px;
    font-weight: 800;
    letter-spacing: 2px;
    text-transform: uppercase;
}

div[data-baseweb="select"] > div {
    min-height: 55px;
    border-radius: 13px !important;
}

div.stButton > button {
    min-height: 54px;
    border-radius: 13px;
    font-size: 16px;
    font-weight: 800;
}

/* RESULT */

.result-success {
    background: linear-gradient(135deg, #157f46, #19a45a);
    color: white;
    border-radius: 16px;
    padding: 17px;
    text-align: center;
    font-size: 20px;
    font-weight: 900;
    box-shadow: 0 10px 25px rgba(25,164,90,0.18);
}

.result-danger {
    background: linear-gradient(135deg, #991b1b, #dc2626);
    color: white;
    border-radius: 16px;
    padding: 18px;
    text-align: center;
    font-size: 23px;
    font-weight: 900;
    box-shadow: 0 10px 25px rgba(220,38,38,0.20);
}

.container-card {
    background:
        linear-gradient(135deg, #081b2d 0%, #0e4268 100%);
    border-radius: 20px;
    padding: 30px;
    text-align: center;
    margin-top: 15px;
    margin-bottom: 18px;
    box-shadow: 0 12px 32px rgba(15,42,65,0.18);
}

.container-label {
    color: #a8c9dd;
    font-size: 11px;
    letter-spacing: 3px;
    font-weight: 700;
}

.container-number {
    color: white;
    font-size: 34px;
    font-weight: 900;
    letter-spacing: 2px;
    margin-top: 5px;
}

.shipping-label {
    color: #a8c9dd;
    font-size: 11px;
    letter-spacing: 3px;
    font-weight: 700;
    margin-top: 22px;
}

.shipping-line {
    color: white;
    font-size: 31px;
    font-weight: 900;
    margin-top: 4px;
}

/* INFO */

div[data-testid="stMetric"] {
    background-color: rgba(255,255,255,0.95);
    border: 1px solid #e3e8ef;
    border-radius: 15px;
    padding: 17px;
    box-shadow: 0 5px 16px rgba(15,23,42,0.04);
}

div[data-testid="stMetricLabel"] {
    color: #758293;
    font-size: 12px;
    font-weight: 700;
}

div[data-testid="stMetricValue"] {
    color: #102a43;
    font-size: 20px;
    font-weight: 800;
}

.footer-custom {
    text-align: center;
    color: #8a98a8;
    font-size: 11px;
    margin-top: 35px;
    letter-spacing: .4px;
}

@media (max-width: 600px) {

    .hero {
        padding: 28px 18px;
        border-radius: 18px;
    }

    .hero-title {
        font-size: 30px;
    }

    .container-number {
        font-size: 27px;
    }

    .shipping-line {
        font-size: 26px;
    }
}

</style>
""")


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

    return LINE_MAP.get(value, value)


def clean_value(record, column):
    if column not in record.index:
        return "-"

    value = str(record[column]).strip()

    if not value or value.lower() == "nan":
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
        raise ValueError("CONTAINER column missing.")

    df["_SEARCH"] = (
        df["CONTAINER"]
        .apply(normalize_container)
    )

    return df


# =========================================================
# HERO
# =========================================================

st.html("""
<div class="hero">
    <div class="hero-brand">
        ALPORT BANJUL
    </div>

    <div class="hero-title">
        CONTAINER TRACKING
    </div>

    <div class="hero-sub">
        Container Identification & Shipping Line Verification
    </div>
</div>
""")


# =========================================================
# DATABASE
# =========================================================

if not os.path.exists(EXCEL_FILE):

    st.error(
        "Container database is unavailable."
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

    st.stop()


update_time = datetime.fromtimestamp(
    modified_time
).strftime("%d/%m/%Y %H:%M")


# =========================================================
# DATABASE CARDS
# =========================================================

db1, db2 = st.columns(2)

with db1:

    st.html(
        f"""
        <div class="db-card">
            <div class="db-label">
                Containers
            </div>

            <div class="db-value">
                {len(df):,}
            </div>
        </div>
        """
    )


with db2:

    st.html(
        f"""
        <div class="db-card">
            <div class="db-label">
                Last Update
            </div>

            <div class="db-value">
                {update_time}
            </div>
        </div>
        """
    )


st.write("")


# =========================================================
# SEARCH
# =========================================================

with st.container(border=True):

    st.subheader(
        "🔎 Container Verification"
    )

    st.caption(
        "Select loading line and enter container number."
    )

    available_lines = sorted(
        {
            normalize_line(value)
            for value in df["AGENT"].unique()
            if normalize_line(value) != "-"
        }
    )

    line_options = [
        "No line selected"
    ] + available_lines


    with st.form(
        "container_search",
        clear_on_submit=False
    ):

        selected_line = st.selectbox(
            "Loading Line",
            line_options
        )

        container_input = st.text_input(
            "Container Number",
            placeholder="SEKU6920313",
            max_chars=20
        )

        search_button = st.form_submit_button(
            "VERIFY CONTAINER",
            type="primary",
            use_container_width=True
        )


# =========================================================
# SEARCH RESULT
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
        df["_SEARCH"] == search_number
    ]


    # =====================================================
    # NOT FOUND
    # =====================================================

    if result.empty:

        st.html(
            """
            <div class="result-danger">
                ⛔ CONTAINER NOT FOUND
            </div>
            """
        )

        st.markdown(
            f"## {search_number}"
        )

        st.error(
            "DO NOT LOAD"
        )

        st.caption(
            "Container is not available in the current database. "
            "Contact Operations before loading."
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

        location = clean_value(
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


        wrong_line = (
            selected_line != "No line selected"
            and selected_line != shipping_line
        )


        # =================================================
        # WRONG LINE
        # =================================================

        if wrong_line:

            st.html(
                """
                <div class="result-danger">
                    🛑 STOP — WRONG SHIPPING LINE
                </div>
                """
            )

            st.html(
                f"""
                <div class="container-card">

                    <div class="container-label">
                        CONTAINER
                    </div>

                    <div class="container-number">
                        {container}
                    </div>

                    <div class="shipping-label">
                        ACTUAL SHIPPING LINE
                    </div>

                    <div class="shipping-line">
                        {shipping_line}
                    </div>

                </div>
                """
            )

            c1, c2 = st.columns(2)

            with c1:
                st.metric(
                    "Container Line",
                    shipping_line
                )

            with c2:
                st.metric(
                    "Selected Line",
                    selected_line
                )

            st.error(
                "⛔ DO NOT LOAD THIS CONTAINER"
            )


        # =================================================
        # VERIFIED
        # =================================================

        else:

            st.html(
                """
                <div class="result-success">
                    ✓ CONTAINER VERIFIED
                </div>
                """
            )

            st.html(
                f"""
                <div class="container-card">

                    <div class="container-label">
                        CONTAINER
                    </div>

                    <div class="container-number">
                        {container}
                    </div>

                    <div class="shipping-label">
                        SHIPPING LINE
                    </div>

                    <div class="shipping-line">
                        {shipping_line}
                    </div>

                </div>
                """
            )


            if selected_line != "No line selected":

                st.success(
                    f"✓ Correct shipping line: {selected_line}"
                )


            # =================================================
            # DETAILS
            # =================================================

            col1, col2 = st.columns(2)

            with col1:

                st.metric(
                    "Size",
                    size
                )

            with col2:

                st.metric(
                    "Type",
                    container_type
                )


            col3, col4 = st.columns(2)

            with col3:

                st.metric(
                    "Status",
                    status
                )

            with col4:

                st.metric(
                    "Location / Area",
                    location
                )


            col5, col6 = st.columns(2)

            with col5:

                st.metric(
                    "Vessel",
                    vessel
                )

            with col6:

                st.metric(
                    "Voyage",
                    voyage
                )


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

st.html("""
<div class="footer-custom">
    ALPORT BANJUL • CONTAINER VERIFICATION SYSTEM • OPERATIONS
</div>
""")
