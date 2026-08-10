import os
import re
from datetime import datetime

import pandas as pd
import streamlit as st


# =========================================================
# PAGE CONFIG
# =========================================================

st.set_page_config(
    page_title="ALPORT Container Tracking",
    page_icon="🚢",
    layout="centered",
    initial_sidebar_state="collapsed"
)

EXCEL_FILE = "containers.xlsx"


# =========================================================
# LINE SETTINGS
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


LINE_EMOJIS = {
    "MAERSK": "🔷",
    "CMA CGM": "🟠",
    "MSC": "🟡",
    "HAPAG-LLOYD": "🔴",
    "ONE": "🟣",
    "COSCO": "🔵",
    "PIL": "🟢",
    "OBT": "⚓"
}


# =========================================================
# CSS
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
        background:
            radial-gradient(
                circle at top left,
                rgba(30, 64, 175, 0.08),
                transparent 32%
            ),
            radial-gradient(
                circle at top right,
                rgba(14, 116, 144, 0.08),
                transparent 30%
            ),
            linear-gradient(
                180deg,
                #f8fafc 0%,
                #eef3f8 100%
            );
    }

    .block-container {
        max-width: 880px;
        padding-top: 1.5rem;
        padding-bottom: 3rem;
    }

    h1, h2, h3 {
        letter-spacing: -0.02em;
    }

    /* HERO */
    .hero-box {
        background:
            linear-gradient(
                135deg,
                #061a2d 0%,
                #0b355b 52%,
                #155e75 100%
            );

        border-radius: 24px;
        padding: 34px 28px;
        margin-bottom: 20px;

        box-shadow:
            0 20px 50px rgba(15, 23, 42, 0.18);

        border: 1px solid rgba(255,255,255,0.08);
    }

    .hero-company {
        color: #9fc8df;
        font-size: 13px;
        font-weight: 800;
        letter-spacing: 4px;
        text-transform: uppercase;
        text-align: center;
    }

    .hero-title {
        color: white;
        text-align: center;
        font-size: 38px;
        font-weight: 900;
        margin-top: 8px;
        margin-bottom: 8px;
    }

    .hero-subtitle {
        color: #d7e7f1;
        text-align: center;
        font-size: 15px;
        line-height: 1.6;
    }

    /* SEARCH CONTAINER */
    div[data-testid="stVerticalBlockBorderWrapper"] {
        border-radius: 18px !important;
        border: 1px solid rgba(148, 163, 184, 0.28) !important;
        background: rgba(255,255,255,0.82) !important;
        box-shadow: 0 8px 28px rgba(15,23,42,0.06);
        backdrop-filter: blur(10px);
    }

    /* TEXT INPUT */
    div[data-testid="stTextInput"] input {
        height: 62px;
        border-radius: 14px;
        text-align: center;
        font-size: 24px;
        font-weight: 800;
        letter-spacing: 2px;
        text-transform: uppercase;
        background-color: #ffffff;
    }

    div[data-testid="stTextInput"] input:focus {
        border-color: #0f4c75;
        box-shadow: 0 0 0 3px rgba(15, 76, 117, 0.10);
    }

    /* SELECT BOX */
    div[data-baseweb="select"] > div {
        min-height: 56px;
        border-radius: 14px !important;
        background-color: white;
    }

    /* BUTTON */
    div.stButton > button {
        min-height: 54px;
        border-radius: 14px;
        font-weight: 800;
        font-size: 16px;
        transition: 0.2s;
    }

    div.stButton > button:hover {
        transform: translateY(-1px);
    }

    /* METRIC CARD */
    div[data-testid="stMetric"] {
        background:
            linear-gradient(
                180deg,
                rgba(255,255,255,0.98),
                rgba(248,250,252,0.98)
            );

        border: 1px solid #e2e8f0;
        border-radius: 15px;
        padding: 18px;

        box-shadow:
            0 5px 18px
            rgba(15,23,42,0.05);
    }

    div[data-testid="stMetricLabel"] {
        color: #64748b;
        font-weight: 700;
        font-size: 12px;
        text-transform: uppercase;
        letter-spacing: 0.6px;
    }

    div[data-testid="stMetricValue"] {
        color: #0f172a;
        font-size: 21px;
        font-weight: 800;
    }

    /* SUCCESS */
    div[data-testid="stAlert"][data-baseweb="notification"] {
        border-radius: 15px;
    }

    /* LINE HERO */
    .line-card {
        padding: 28px;
        border-radius: 19px;
        text-align: center;
        color: white;
        background:
            linear-gradient(
                135deg,
                #0b2239,
                #0f4c75
            );

        box-shadow:
            0 12px 30px rgba(15, 76, 117, 0.18);

        margin-top: 8px;
        margin-bottom: 18px;
    }

    .line-label {
        font-size: 12px;
        letter-spacing: 3px;
        font-weight: 700;
        color: #b9d7e9;
    }

    .line-name {
        font-size: 34px;
        font-weight: 900;
        margin-top: 5px;
    }

    .line-container {
        font-size: 18px;
        color: #e4eff6;
        margin-top: 5px;
        letter-spacing: 1px;
    }

    /* STATUS PILLS */
    .verified-pill {
        background:
            linear-gradient(
                135deg,
                #15803d,
                #16a34a
            );

        color: white;
        text-align: center;
        padding: 15px;
        border-radius: 14px;
        font-size: 19px;
        font-weight: 900;

        box-shadow:
            0 8px 22px
            rgba(22,163,74,0.18);
    }

    .wrong-pill {
        background:
            linear-gradient(
                135deg,
                #991b1b,
                #dc2626
            );

        color: white;
        text-align: center;
        padding: 18px;
        border-radius: 15px;
        font-size: 23px;
        font-weight: 900;

        box-shadow:
            0 10px 28px
            rgba(220,38,38,0.22);
    }

    .stop-card {
        background:
            linear-gradient(
                180deg,
                #fff7f7,
                #fff1f2
            );

        border: 2px solid #ef4444;
        border-radius: 18px;
        padding: 23px;
        margin-top: 16px;

        box-shadow:
            0 10px 30px
            rgba(220,38,38,0.10);
    }

    .footer-text {
        text-align: center;
        color: #94a3b8;
        font-size: 11px;
        letter-spacing: 0.4px;
        margin-top: 35px;
    }

    /* MOBILE */
    @media (max-width: 600px) {

        .block-container {
            padding-left: 1rem;
            padding-right: 1rem;
        }

        .hero-box {
            padding: 27px 18px;
            border-radius: 18px;
        }

        .hero-title {
            font-size: 29px;
        }

        .hero-subtitle {
            font-size: 13px;
        }

        div[data-testid="stTextInput"] input {
            font-size: 21px;
        }

        .line-name {
            font-size: 29px;
        }
    }

</style>
"""

st.markdown(
    CSS,
    unsafe_allow_html=True
)


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


def line_emoji(line):

    return LINE_EMOJIS.get(
        line,
        "🚢"
    )


# =========================================================
# HERO
# =========================================================

st.markdown(
    """
    <div class="hero-box">

        <div class="hero-company">
            ALPORT BANJUL
        </div>

        <div class="hero-title">
            CONTAINER TRACKING
        </div>

        <div class="hero-subtitle">
            Fast container verification for safe and accurate vessel loading
        </div>

    </div>
    """,
    unsafe_allow_html=True
)


# =========================================================
# DATABASE
# =========================================================

if not os.path.exists(EXCEL_FILE):

    st.error(
        "Container database is currently unavailable."
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
# DATABASE STATUS
# =========================================================

status1, status2 = st.columns(2)

with status1:

    st.metric(
        "Live Database",
        f"{len(df):,}",
        help="Total containers in the current database"
    )

with status2:

    st.metric(
        "Last Updated",
        update_time
    )


st.write("")


# =========================================================
# SEARCH PANEL
# =========================================================

with st.container(
    border=True
):

    st.subheader(
        "🔎 Container Verification"
    )

    st.caption(
        "Select loading line and enter the container number."
    )

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
# RESULT
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
        df["_SEARCH_CONTAINER"]
        == search_number
    ]


    # =====================================================
    # NOT FOUND
    # =====================================================

    if result.empty:

        st.markdown(
            """
            <div class="wrong-pill">
                ⛔ CONTAINER NOT FOUND
            </div>
            """,
            unsafe_allow_html=True
        )

        st.markdown(
            f"# {search_number}"
        )

        st.error(
            "DO NOT LOAD"
        )

        st.caption(
            "The container is not available in the current database. "
            "Contact Operations before loading."
        )


    # =====================================================
    # DUPLICATE
    # =====================================================

    elif len(result) > 1:

        st.warning(
            "⚠️ DUPLICATE RECORD"
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


        wrong_line = (
            selected_line != "No line selected"
            and selected_line != shipping_line
        )


        # =================================================
        # WRONG LINE
        # =================================================

        if wrong_line:

            st.markdown(
                """
                <div class="wrong-pill">
                    🛑 STOP — WRONG SHIPPING LINE
                </div>
                """,
                unsafe_allow_html=True
            )

            st.markdown(
                f"""
                <div class="stop-card">
                    <h2 style="text-align:center;margin-bottom:4px;">
                        {container}
                    </h2>
                </div>
                """,
                unsafe_allow_html=True
            )

            wrong1, wrong2 = st.columns(2)

            with wrong1:

                st.metric(
                    "Container Line",
                    f"{line_emoji(shipping_line)} {shipping_line}"
                )

            with wrong2:

                st.metric(
                    "Selected Loading Line",
                    selected_line
                )

            st.error(
                "⛔ DO NOT LOAD THIS CONTAINER"
            )


        # =================================================
        # VERIFIED
        # =================================================

        else:

            st.markdown(
                """
                <div class="verified-pill">
                    ✓ CONTAINER VERIFIED
                </div>
                """,
                unsafe_allow_html=True
            )

            st.write("")

            st.markdown(
                f"""
                <div class="line-card">

                    <div class="line-label">
                        SHIPPING LINE
                    </div>

                    <div class="line-name">
                        {line_emoji(shipping_line)}
                        {shipping_line}
                    </div>

                    <div class="line-container">
                        {container}
                    </div>

                </div>
                """,
                unsafe_allow_html=True
            )


            if selected_line != "No line selected":

                st.success(
                    f"✓ Correct line for {selected_line}"
                )


            # =================================================
            # INFO GRID
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
                    "Location",
                    area
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


            # =================================================
            # ADDITIONAL INFO
            # =================================================

            if (
                imo_class != "-"
                or discharge_date != "-"
            ):

                with st.expander(
                    "More Container Details"
                ):

                    info1, info2 = st.columns(2)

                    with info1:

                        st.caption(
                            "IMO CLASS"
                        )

                        st.write(
                            imo_class
                        )

                    with info2:

                        st.caption(
                            "DISCHARGE DATE"
                        )

                        st.write(
                            discharge_date
                        )


# =========================================================
# FOOTER
# =========================================================

st.markdown(
    """
    <div class="footer-text">
        ALPORT BANJUL • CONTAINER VERIFICATION SYSTEM • OPERATIONS
    </div>
    """,
    unsafe_allow_html=True
)
