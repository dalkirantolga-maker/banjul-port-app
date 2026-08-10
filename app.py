import streamlit as st
import pandas as pd
import os
from datetime import datetime

# =========================================================
# PAGE SETTINGS
# =========================================================

st.set_page_config(
    page_title="Container Tracking System",
    page_icon="📦",
    layout="centered"
)

EXCEL_FILE = "containers.xlsx"

# =========================================================
# DESIGN
# =========================================================

st.markdown(
    """
    <style>

    .stApp {
        background-color: #f4f6f8;
    }

    .block-container {
        max-width: 850px;
        padding-top: 2rem;
    }

    .header {
        background: linear-gradient(135deg, #102a43, #1f4e79);
        padding: 28px;
        border-radius: 16px;
        text-align: center;
        margin-bottom: 24px;
        box-shadow: 0 6px 20px rgba(0,0,0,0.10);
    }

    .header-title {
        color: white;
        font-size: 32px;
        font-weight: 800;
        margin: 0;
    }

    .header-subtitle {
        color: #dbeafe;
        font-size: 15px;
        margin-top: 8px;
    }

    .found-box {
        background-color: #e9f8ef;
        border: 2px solid #16a34a;
        color: #166534;
        padding: 15px;
        border-radius: 12px;
        text-align: center;
        font-weight: 800;
        font-size: 19px;
        margin-top: 20px;
    }

    .not-found-box {
        background-color: #fff1f1;
        border: 2px solid #dc2626;
        padding: 24px;
        border-radius: 14px;
        text-align: center;
        margin-top: 20px;
    }

    .not-found-title {
        color: #b91c1c;
        font-size: 27px;
        font-weight: 900;
    }

    .container-number {
        text-align: center;
        font-size: 31px;
        font-weight: 900;
        letter-spacing: 2px;
        color: #102a43;
        margin-top: 18px;
    }

    .agent-box {
        background-color: #102a43;
        border-radius: 14px;
        text-align: center;
        padding: 20px;
        margin-top: 18px;
        margin-bottom: 24px;
    }

    .agent-label {
        color: #bdd7ee;
        font-size: 13px;
        font-weight: 700;
        letter-spacing: 2px;
    }

    .agent-value {
        color: white;
        font-size: 34px;
        font-weight: 900;
        margin-top: 5px;
    }

    .info-card {
        background-color: white;
        border: 1px solid #e2e8f0;
        border-radius: 12px;
        padding: 15px;
        margin-bottom: 12px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.04);
    }

    .info-label {
        color: #64748b;
        font-size: 12px;
        font-weight: 700;
    }

    .info-value {
        color: #0f172a;
        font-size: 18px;
        font-weight: 800;
        margin-top: 4px;
    }

    .footer {
        text-align: center;
        color: #94a3b8;
        font-size: 12px;
        margin-top: 40px;
    }

    div[data-testid="stTextInput"] input {
        height: 58px;
        text-align: center;
        font-size: 23px;
        font-weight: 800;
        letter-spacing: 1px;
        text-transform: uppercase;
    }

    div.stButton > button {
        width: 100%;
        height: 52px;
        border-radius: 10px;
        font-size: 17px;
        font-weight: 800;
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

    return (
        str(value)
        .upper()
        .strip()
        .replace(" ", "")
        .replace("-", "")
    )


@st.cache_data(ttl=60)
def load_database(file_name, modified_time):

    df = pd.read_excel(
        file_name,
        sheet_name="Sayfa1",
        dtype=str,
        engine="openpyxl"
    )

    df.columns = [
        str(column).strip().upper()
        for column in df.columns
    ]

    df = df.fillna("")

    for column in df.columns:
        df[column] = (
            df[column]
            .astype(str)
            .str.strip()
        )

    df["_SEARCH_CONTAINER"] = (
        df["CONTAINER"]
        .apply(normalize_container)
    )

    return df


def safe_value(record, column):

    if column not in record.index:
        return "-"

    value = str(record[column]).strip()

    if value == "" or value.lower() == "nan":
        return "-"

    return value


def format_agent(agent):

    agent = str(agent).strip().upper()

    mappings = {
        "MAE": "MAERSK",
        "CMA": "CMA CGM",
        "MSC": "MSC",
        "OBT": "OBT",
        "ALPORT": "ALPORT",
        "STT": "STT",
        "FERI": "FERRY",
        "NILL": "NILL"
    }

    return mappings.get(agent, agent)


# =========================================================
# HEADER
# =========================================================

st.markdown(
    """
    <div class="header">
        <div class="header-title">
            📦 CONTAINER TRACKING
        </div>

        <div class="header-subtitle">
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

    st.error("Container database is unavailable.")

    st.info(
        "Please contact the Operations Department."
    )

    st.stop()


try:

    modified_time = os.path.getmtime(EXCEL_FILE)

    df = load_database(
        EXCEL_FILE,
        modified_time
    )

except Exception as error:

    st.error(
        "Container database could not be loaded."
    )

    st.exception(error)

    st.stop()


# =========================================================
# DATABASE STATUS
# =========================================================

update_date = datetime.fromtimestamp(
    modified_time
).strftime("%d/%m/%Y %H:%M")

col1, col2 = st.columns(2)

col1.metric(
    "Containers",
    f"{len(df):,}"
)

col2.metric(
    "Database Updated",
    update_date
)

st.divider()

# =========================================================
# SEARCH
# =========================================================

st.markdown(
    "### 🔎 Container Search"
)

container_input = st.text_input(
    "Container Number",
    placeholder="Example: MRKU5878311",
    label_visibility="collapsed"
)

search = st.button(
    "SEARCH",
    type="primary",
    use_container_width=True
)

if search:

    container_number = normalize_container(
        container_input
    )

    if container_number == "":

        st.warning(
            "Please enter a container number."
        )

        st.stop()

    results = df[
        df["_SEARCH_CONTAINER"] == container_number
    ]

    # =====================================================
    # NOT FOUND
    # =====================================================

    if results.empty:

        st.markdown(
            f"""
            <div class="not-found-box">

                <div class="not-found-title">
                    ⛔ CONTAINER NOT FOUND
                </div>

                <div style="
                    font-size:24px;
                    font-weight:900;
                    color:#111827;
                    margin-top:10px;
                ">
                    {container_number}
                </div>

                <div style="
                    margin-top:14px;
                    color:#7f1d1d;
                    font-size:16px;
                ">
                    This container is not available
                    in the current database.
                </div>

                <div style="
                    margin-top:18px;
                    color:#b91c1c;
                    font-size:23px;
                    font-weight:900;
                ">
                    DO NOT LOAD
                </div>

                <div style="
                    margin-top:7px;
                    color:#7f1d1d;
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

        st.warning(
            f"⚠️ {len(results)} records found for {container_number}. "
            "Please verify before loading."
        )

        display_columns = [
            "CONTAINER",
            "AGENT",
            "SIZE",
            "TYPE",
            "FULL-MTY",
            "AREA",
            "VESSEL NAME",
            "VOYAGE NUMBER"
        ]

        available_columns = [
            column
            for column in display_columns
            if column in results.columns
        ]

        st.dataframe(
            results[available_columns],
            hide_index=True,
            use_container_width=True
        )

    # =====================================================
    # FOUND
    # =====================================================

    else:

        record = results.iloc[0]

        container = safe_value(
            record,
            "CONTAINER"
        )

        agent_code = safe_value(
            record,
            "AGENT"
        )

        agent = format_agent(
            agent_code
        )

        size = safe_value(
            record,
            "SIZE"
        )

        container_type = safe_value(
            record,
            "TYPE"
        )

        status = safe_value(
            record,
            "FULL-MTY"
        )

        area = safe_value(
            record,
            "AREA"
        )

        vessel = safe_value(
            record,
            "VESSEL NAME"
        )

        voyage = safe_value(
            record,
            "VOYAGE NUMBER"
        )

        imo = safe_value(
            record,
            "IMO CLS"
        )

        discharge_date = safe_value(
            record,
            "DISCHARGE DATE"
        )

        # FOUND

        st.markdown(
            """
            <div class="found-box">
                ✅ CONTAINER FOUND
            </div>
            """,
            unsafe_allow_html=True
        )

        st.markdown(
            f"""
            <div class="container-number">
                {container}
            </div>
            """,
            unsafe_allow_html=True
        )

        # SHIPPING LINE

        st.markdown(
            f"""
            <div class="agent-box">

                <div class="agent-label">
                    SHIPPING LINE
                </div>

                <div class="agent-value">
                    {agent}
                </div>

            </div>
            """,
            unsafe_allow_html=True
        )

        # =================================================
        # INFORMATION
        # =================================================

        col1, col2 = st.columns(2)

        with col1:

            st.markdown(
                f"""
                <div class="info-card">

                    <div class="info-label">
                        SIZE
                    </div>

                    <div class="info-value">
                        {size}'
                    </div>

                </div>
                """,
                unsafe_allow_html=True
            )

        with col2:

            st.markdown(
                f"""
                <div class="info-card">

                    <div class="info-label">
                        TYPE
                    </div>

                    <div class="info-value">
                        {container_type}
                    </div>

                </div>
                """,
                unsafe_allow_html=True
            )

        col3, col4 = st.columns(2)

        with col3:

            st.markdown(
                f"""
                <div class="info-card">

                    <div class="info-label">
                        STATUS
                    </div>

                    <div class="info-value">
                        {status}
                    </div>

                </div>
                """,
                unsafe_allow_html=True
            )

        with col4:

            st.markdown(
                f"""
                <div class="info-card">

                    <div class="info-label">
                        AREA
                    </div>

                    <div class="info-value">
                        {area}
                    </div>

                </div>
                """,
                unsafe_allow_html=True
            )

        col5, col6 = st.columns(2)

        with col5:

            st.markdown(
                f"""
                <div class="info-card">

                    <div class="info-label">
                        VESSEL
                    </div>

                    <div class="info-value">
                        {vessel}
                    </div>

                </div>
                """,
                unsafe_allow_html=True
            )

        with col6:

            st.markdown(
                f"""
                <div class="info-card">

                    <div class="info-label">
                        VOYAGE
                    </div>

                    <div class="info-value">
                        {voyage}
                    </div>

                </div>
                """,
                unsafe_allow_html=True
            )

        # =================================================
        # ADDITIONAL INFO
        # =================================================

        with st.expander(
            "Additional Information"
        ):

            st.write(
                "**IMO Class:**",
                imo
            )

            st.write(
                "**Discharge Date:**",
                discharge_date
            )

            st.write(
                "**Agent Code:**",
                agent_code
            )

# =========================================================
# FOOTER
# =========================================================

st.markdown(
    """
    <div class="footer">

        ALPORT BANJUL<br>
        Container Tracking System

    </div>
    """,
    unsafe_allow_html=True
)
