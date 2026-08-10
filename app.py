import streamlit as st
import pandas as pd
import os
import re
from datetime import datetime


# =========================================================
# CONFIGURATION
# =========================================================

st.set_page_config(
    page_title="ALPORT Container Tracking",
    page_icon="🚢",
    layout="centered",
    initial_sidebar_state="collapsed"
)

EXCEL_FILE = "containers.xlsx"


# =========================================================
# SHIPPING LINE DEFINITIONS
# =========================================================

LINE_NAMES = {
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


LINE_OPTIONS = [
    "Select Loading Line",
    "MAERSK",
    "CMA CGM",
    "MSC",
    "HAPAG-LLOYD",
    "ONE",
    "COSCO",
    "PIL",
    "OBT"
]


# =========================================================
# PROFESSIONAL DESIGN
# =========================================================

st.markdown(
    """
    <style>

    /* MAIN PAGE */
    .stApp {
        background:
            linear-gradient(
                180deg,
                #f7f9fc 0%,
                #eef2f7 100%
            );
    }

    .block-container {
        max-width: 820px;
        padding-top: 2rem;
        padding-bottom: 3rem;
    }

    /* HIDE STREAMLIT ELEMENTS */
    #MainMenu {
        visibility: hidden;
    }

    footer {
        visibility: hidden;
    }

    header {
        visibility: hidden;
    }

    /* BRAND HEADER */
    .brand {
        background:
            linear-gradient(
                135deg,
                #071b2d 0%,
                #0d3154 55%,
                #174f7a 100%
            );

        border-radius: 22px;
        padding: 32px 25px;
        text-align: center;
        margin-bottom: 25px;

        box-shadow:
            0px 12px 32px
            rgba(15, 42, 65, 0.18);
    }

    .brand-small {
        font-size: 13px;
        font-weight: 700;
        letter-spacing: 4px;
        color: #9fc4df;
        margin-bottom: 7px;
    }

    .brand-title {
        color: white;
        font-size: 34px;
        font-weight: 800;
        line-height: 1.1;
    }

    .brand-subtitle {
        color: #c9ddeb;
        font-size: 14px;
        margin-top: 10px;
    }

    /* SEARCH BOX */
    div[data-testid="stTextInput"] input {
        height: 64px;
        border-radius: 13px;
        font-size: 24px;
        font-weight: 800;
        text-align: center;
        letter-spacing: 2px;
        text-transform: uppercase;
        border: 2px solid #d6dee8;
        background: white;
    }

    div[data-testid="stTextInput"] input:focus {
        border-color: #174f7a;
        box-shadow:
            0 0 0 2px
            rgba(23, 79, 122, 0.12);
    }

    /* SELECTBOX */
    div[data-baseweb="select"] > div {
        min-height: 54px;
        border-radius: 12px;
    }

    /* BUTTON */
    div.stButton > button {
        width: 100%;
        min-height: 55px;
        border-radius: 12px;
        font-size: 17px;
        font-weight: 800;
    }

    /* CONTAINER NUMBER */
    .cntr-number {
        text-align: center;
        font-size: 35px;
        font-weight: 900;
        letter-spacing: 3px;
        color: #102a43;
        margin-top: 22px;
        margin-bottom: 17px;
    }

    /* SUCCESS */
    .verified {
        background:
            linear-gradient(
                135deg,
                #eaf8f0,
                #f3fbf6
            );

        border: 1px solid #9ed9b4;
        border-left: 6px solid #18864b;

        border-radius: 14px;
        padding: 17px;

        text-align: center;
        color: #116534;

        font-size: 19px;
        font-weight: 800;

        margin-top: 22px;
    }

    /* LINE CARD */
    .line-card {
        background:
            linear-gradient(
                135deg,
                #071b2d,
                #123e63
            );

        border-radius: 18px;
        padding: 27px 20px;

        text-align: center;

        box-shadow:
            0px 8px 20px
            rgba(15, 42, 65, 0.16);

        margin-bottom: 20px;
    }

    .line-caption {
        color: #a9c9df;
        font-size: 12px;
        font-weight: 700;
        letter-spacing: 3px;
    }

    .line-value {
        color: white;
        font-size: 37px;
        font-weight: 900;
        margin-top: 5px;
    }

    /* CORRECT LINE */
    .correct-line {
        background:
            linear-gradient(
                135deg,
                #eaf8ef,
                #f5fcf7
            );

        border: 2px solid #22a35a;
        border-radius: 17px;
        padding: 22px;

        text-align: center;
        margin-top: 15px;
    }

    .correct-title {
        color: #13763d;
        font-size: 25px;
        font-weight: 900;
    }

    .correct-sub {
        color: #3b7153;
        margin-top: 5px;
        font-weight: 600;
    }

    /* WRONG LINE */
    .wrong-line {
        background:
            linear-gradient(
                135deg,
                #fff0f0,
                #ffe4e4
            );

        border: 3px solid #d71920;
        border-radius: 20px;

        padding: 30px 20px;

        text-align: center;

        box-shadow:
            0 8px 25px
            rgba(215, 25, 32, 0.14);

        margin-top: 20px;
    }

    .stop-title {
        color: #c1121f;
        font-size: 44px;
        line-height: 1;
        font-weight: 1000;
    }

    .wrong-title {
        color: #c1121f;
        font-size: 24px;
        font-weight: 900;
        margin-top: 13px;
    }

    .wrong-container {
        color: #19212b;
        font-size: 27px;
        font-weight: 900;
        margin-top: 18px;
        letter-spacing: 2px;
    }

    .do-not-load {
        color: white;
        background-color: #c1121f;
        border-radius: 10px;

        font-size: 25px;
        font-weight: 1000;

        padding: 12px;
        margin-top: 22px;
    }

    /* NOT FOUND */
    .not-found {
        background:
            linear-gradient(
                135deg,
                #fff1f1,
                #fff8f8
            );

        border: 2px solid #d71920;
        border-radius: 18px;

        padding: 28px 20px;
        text-align: center;

        margin-top: 22px;
    }

    .not-found-title {
        color: #c1121f;
        font-size: 28px;
        font-weight: 900;
    }

    .not-found-number {
        color: #17212b;
        font-size: 27px;
        font-weight: 900;
        letter-spacing: 2px;
        margin-top: 12px;
    }

    /* INFO GRID */
    .info-box {
        background-color: white;
        border: 1px solid #dde4ec;
        border-radius: 13px;
        padding: 17px;

        min-height: 92px;

        box-shadow:
            0px 3px 10px
            rgba(0,0,0,0.035);

        margin-bottom: 12px;
    }

    .info-label {
        color: #7c8998;
        font-size: 11px;
        font-weight: 800;
        letter-spacing: 1px;
    }

    .info-value {
        color: #142536;
        font-size: 19px;
        font-weight: 800;
        margin-top: 7px;
        word-break: break-word;
    }

    /* DATABASE STATUS */
    .db-status {
        text-align: center;
        font-size: 12px;
        color: #83909e;
        margin-top: 8px;
        margin-bottom: 8px;
    }

    /* FOOTER */
    .custom-footer {
        text-align: center;
        color: #8b98a7;
        font-size: 12px;
        margin-top: 45px;
        line-height: 1.7;
    }

    </style>
    """,
    unsafe_allow_html=True
)


# =========================================================
# FUNCTIONS
# =========================================================

def normalize_container(value):

    if pd.isna(value):
        return ""

    value = str(value).upper().strip()

    value = re.sub(
        r"[^A-Z0-9]",
        "",
        value
    )

    return value


def normalize_line(value):

    if pd.isna(value):
        return "-"

    line = str(value).upper().strip()

    return LINE_NAMES.get(
        line,
        line
    )


def get_value(record, column):

    if column not in record.index:
        return "-"

    value = str(record[column]).strip()

    if value == "":
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
            "CONTAINER column not found."
        )

    df["_SEARCH"] = (
        df["CONTAINER"]
        .apply(normalize_container)
    )

    return df


def info_card(label, value):

    st.markdown(
        f"""
        <div class="info-box">

            <div class="info-label">
                {label}
            </div>

            <div class="info-value">
                {value}
            </div>

        </div>
        """,
        unsafe_allow_html=True
    )


# =========================================================
# HEADER
# =========================================================

st.markdown(
    """
    <div class="brand">

        <div class="brand-small">
            ALPORT BANJUL
        </div>

        <div class="brand-title">
            CONTAINER TRACKING
        </div>

        <div class="brand-subtitle">
            Container Identification & Shipping Line Verification
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
        "Container database could not be loaded. "
        "Please contact Operations."
    )

    st.stop()


update_time = datetime.fromtimestamp(
    modified_time
).strftime("%d/%m/%Y • %H:%M")


st.markdown(
    f"""
    <div class="db-status">
        Database: {len(df):,} containers
        &nbsp;&nbsp;•&nbsp;&nbsp;
        Updated: {update_time}
    </div>
    """,
    unsafe_allow_html=True
)


st.divider()


# =========================================================
# SEARCH FORM
# =========================================================

with st.form(
    "container_search",
    clear_on_submit=False
):

    st.markdown(
        "##### Loading Line"
    )

    selected_line = st.selectbox(
        "Loading Line",
        LINE_OPTIONS,
        label_visibility="collapsed"
    )

    st.markdown(
        "##### Container Number"
    )

    container_input = st.text_input(
        "Container Number",
        placeholder="MRKU2621382",
        label_visibility="collapsed",
        max_chars=20
    )

    search_button = st.form_submit_button(
        "SEARCH CONTAINER",
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

    if search_number == "":

        st.warning(
            "Enter a container number."
        )

        st.stop()


    results = df[
        df["_SEARCH"] == search_number
    ]


    # =====================================================
    # NOT FOUND
    # =====================================================

    if results.empty:

        st.markdown(
            f"""
            <div class="not-found">

                <div class="not-found-title">
                    ✕ CONTAINER NOT FOUND
                </div>

                <div class="not-found-number">
                    {search_number}
                </div>

                <div style="
                    color:#7a2930;
                    margin-top:14px;
                    font-size:15px;
                ">
                    Container is not available
                    in the current database.
                </div>

                <div class="do-not-load">
                    DO NOT LOAD
                </div>

                <div style="
                    color:#7a2930;
                    margin-top:13px;
                    font-size:13px;
                ">
                    Contact Operations for verification.
                </div>

            </div>
            """,
            unsafe_allow_html=True
        )


    # =====================================================
    # DUPLICATE
    # =====================================================

    elif len(results) > 1:

        st.error(
            f"⚠ DUPLICATE RECORD — "
            f"{len(results)} records found for "
            f"{search_number}. "
            f"Verify with Operations before loading."
        )


    # =====================================================
    # CONTAINER FOUND
    # =====================================================

    else:

        record = results.iloc[0]

        container = get_value(
            record,
            "CONTAINER"
        )

        agent_code = get_value(
            record,
            "AGENT"
        )

        shipping_line = normalize_line(
            agent_code
        )

        size = get_value(
            record,
            "SIZE"
        )

        container_type = get_value(
            record,
            "TYPE"
        )

        status = get_value(
            record,
            "FULL-MTY"
        )

        area = get_value(
            record,
            "AREA"
        )

        vessel = get_value(
            record,
            "VESSEL NAME"
        )

        voyage = get_value(
            record,
            "VOYAGE NUMBER"
        )

        imo = get_value(
            record,
            "IMO CLS"
        )

        discharge_date = get_value(
            record,
            "DISCHARGE DATE"
        )


        # =================================================
        # WRONG LINE CHECK
        # =================================================

        wrong_line = False

        if (
            selected_line
            != "Select Loading Line"
        ):

            if (
                shipping_line.upper()
                != selected_line.upper()
            ):

                wrong_line = True


        # =================================================
        # WRONG SHIPPING LINE
        # =================================================

        if wrong_line:

            st.markdown(
                f"""
                <div class="wrong-line">

                    <div class="stop-title">
                        STOP
                    </div>

                    <div class="wrong-title">
                        WRONG SHIPPING LINE
                    </div>

                    <div class="wrong-container">
                        {container}
                    </div>

                    <div style="
                        margin-top:20px;
                        color:#5b3034;
                    ">
                        CONTAINER LINE
                    </div>

                    <div style="
                        font-size:28px;
                        font-weight:900;
                        color:#162635;
                    ">
                        {shipping_line}
                    </div>

                    <div style="
                        margin-top:15px;
                        color:#5b3034;
                    ">
                        SELECTED LOADING LINE
                    </div>

                    <div style="
                        font-size:25px;
                        font-weight:900;
                        color:#162635;
                    ">
                        {selected_line}
                    </div>

                    <div class="do-not-load">
                        DO NOT LOAD
                    </div>

                </div>
                """,
                unsafe_allow_html=True
            )


        # =================================================
        # CORRECT CONTAINER
        # =================================================

        else:

            st.markdown(
                """
                <div class="verified">
                    ✓ CONTAINER VERIFIED
                </div>
                """,
                unsafe_allow_html=True
            )


            st.markdown(
                f"""
                <div class="cntr-number">
                    {container}
                </div>
                """,
                unsafe_allow_html=True
            )


            # LINE

            st.markdown(
                f"""
                <div class="line-card">

                    <div class="line-caption">
                        SHIPPING LINE
                    </div>

                    <div class="line-value">
                        {shipping_line}
                    </div>

                </div>
                """,
                unsafe_allow_html=True
            )


            # CORRECT LINE MESSAGE

            if (
                selected_line
                != "Select Loading Line"
            ):

                st.markdown(
                    f"""
                    <div class="correct-line">

                        <div class="correct-title">
                            ✓ CORRECT SHIPPING LINE
                        </div>

                        <div class="correct-sub">
                            Approved for {selected_line}
                            line verification
                        </div>

                    </div>
                    """,
                    unsafe_allow_html=True
                )


            st.write("")


            # =================================================
            # MAIN INFORMATION
            # =================================================

            col1, col2 = st.columns(2)

            with col1:

                info_card(
                    "SIZE",
                    size
                )

            with col2:

                info_card(
                    "TYPE",
                    container_type
                )


            col3, col4 = st.columns(2)

            with col3:

                info_card(
                    "STATUS",
                    status
                )

            with col4:

                info_card(
                    "LOCATION / AREA",
                    area
                )


            col5, col6 = st.columns(2)

            with col5:

                info_card(
                    "VESSEL",
                    vessel
                )

            with col6:

                info_card(
                    "VOYAGE",
                    voyage
                )


            # =================================================
            # ADDITIONAL INFORMATION
            # =================================================

            with st.expander(
                "Additional Information"
            ):

                col7, col8 = st.columns(2)

                with col7:

                    st.caption(
                        "IMO CLASS"
                    )

                    st.write(
                        imo
                    )

                with col8:

                    st.caption(
                        "DISCHARGE DATE"
                    )

                    st.write(
                        discharge_date
                    )

                st.caption(
                    "AGENT CODE"
                )

                st.write(
                    agent_code
                )


# =========================================================
# FOOTER
# =========================================================

st.markdown(
    """
    <div class="custom-footer">

        <b>ALPORT BANJUL</b><br>

        Container Verification System<br>

        Operations Department

    </div>
    """,
    unsafe_allow_html=True
)
