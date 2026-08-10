import os
import re
import html
from datetime import datetime

import pandas as pd
import streamlit as st


# =========================================================
# SAYFA AYARLARI
# =========================================================

st.set_page_config(
    page_title="ALPORT Konteyner Takip",
    page_icon="🚢",
    layout="centered",
    initial_sidebar_state="collapsed"
)

EXCEL_FILE = "containers.xlsx"


# =========================================================
# HAT EŞLEŞTİRMELERİ
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


# Her hat için görsel vurgu rengi
LINE_COLORS = {
    "MAERSK": "#42B0D5",
    "CMA CGM": "#E85D2A",
    "MSC": "#F5B800",
    "HAPAG-LLOYD": "#F26B21",
    "ONE": "#D6007F",
    "COSCO": "#1A5CA8",
    "PIL": "#E52329",
    "OBT": "#16A085"
}


# =========================================================
# CSS
# =========================================================

st.html("""
<style>

/* STREAMLIT MENÜLERİNİ GİZLE */

#MainMenu {
    visibility: hidden;
}

footer {
    visibility: hidden;
}

header {
    visibility: hidden;
}


/* SAYFA */

.stApp {
    background:
        radial-gradient(
            circle at 0% 0%,
            rgba(0, 159, 227, 0.12),
            transparent 32%
        ),
        radial-gradient(
            circle at 100% 25%,
            rgba(13, 148, 136, 0.08),
            transparent 28%
        ),
        linear-gradient(
            180deg,
            #f7fafc 0%,
            #edf3f7 100%
        );
}

.block-container {
    max-width: 960px;
    padding-top: 1.4rem;
    padding-bottom: 3rem;
}


/* =====================================================
   HERO
   ===================================================== */

.hero {
    position: relative;
    overflow: hidden;

    min-height: 260px;

    background:
        linear-gradient(
            125deg,
            #061827 0%,
            #0a3555 50%,
            #007c91 100%
        );

    border-radius: 26px;

    padding: 35px 38px;

    margin-bottom: 24px;

    box-shadow:
        0 22px 55px
        rgba(8, 35, 55, 0.22);

    border:
        1px solid rgba(255,255,255,0.10);
}

.hero-glow-one {
    position: absolute;

    width: 280px;
    height: 280px;

    border-radius: 50%;

    background:
        rgba(49, 190, 220, 0.18);

    right: -90px;
    top: -110px;

    filter: blur(2px);
}

.hero-glow-two {
    position: absolute;

    width: 180px;
    height: 180px;

    border-radius: 50%;

    background:
        rgba(255,255,255,0.05);

    right: 140px;
    bottom: -100px;
}

.hero-content {
    position: relative;
    z-index: 3;

    width: 55%;
}

.hero-brand {
    color: #8ed8e4;

    font-size: 12px;
    font-weight: 800;

    letter-spacing: 4px;

    margin-bottom: 9px;
}

.hero-title {
    color: white;

    font-size: 40px;
    font-weight: 900;

    line-height: 1.05;

    letter-spacing: -1px;
}

.hero-subtitle {
    color: #c7dce7;

    font-size: 14px;

    margin-top: 15px;

    max-width: 440px;

    line-height: 1.65;
}

.hero-badge {
    display: inline-block;

    color: #dff8ff;

    background:
        rgba(255,255,255,0.10);

    border:
        1px solid rgba(255,255,255,0.15);

    border-radius: 30px;

    padding: 7px 13px;

    margin-top: 18px;

    font-size: 11px;
}


/* =====================================================
   HERO GEMİ GÖRSELİ
   ===================================================== */

.hero-visual {
    position: absolute;

    right: 25px;
    bottom: 18px;

    width: 38%;

    max-width: 320px;

    opacity: 0.96;
}


/* =====================================================
   ÜST DURUM KARTLARI
   ===================================================== */

.stat-card {
    position: relative;

    overflow: hidden;

    background:
        rgba(255,255,255,0.94);

    border:
        1px solid #dbe5ec;

    border-radius: 17px;

    padding: 20px 22px;

    min-height: 96px;

    box-shadow:
        0 7px 22px
        rgba(15,23,42,0.05);
}

.stat-accent-blue {
    position: absolute;

    top: 0;
    left: 0;

    width: 5px;
    height: 100%;

    background:
        linear-gradient(
            #00a1d5,
            #127ca5
        );
}

.stat-accent-green {
    position: absolute;

    top: 0;
    left: 0;

    width: 5px;
    height: 100%;

    background:
        linear-gradient(
            #18a57b,
            #087b65
        );
}

.stat-icon {
    font-size: 20px;
    margin-bottom: 5px;
}

.stat-label {
    color: #718190;

    font-size: 10px;
    font-weight: 800;

    text-transform: uppercase;

    letter-spacing: 1.3px;
}

.stat-value {
    color: #102b3d;

    font-size: 22px;
    font-weight: 900;

    margin-top: 4px;
}


/* =====================================================
   ARAMA PANELİ
   ===================================================== */

div[data-testid="stVerticalBlockBorderWrapper"] {
    border-radius: 20px !important;

    background:
        rgba(255,255,255,0.92) !important;

    border:
        1px solid #d8e2e9 !important;

    box-shadow:
        0 10px 32px
        rgba(15,23,42,0.06);
}

div[data-testid="stTextInput"] input {
    min-height: 64px;

    background: #ffffff;

    border-radius: 14px;

    border: 1px solid #cbd8e1;

    text-align: center;

    font-size: 24px;
    font-weight: 900;

    letter-spacing: 2px;

    text-transform: uppercase;
}

div[data-testid="stTextInput"] input:focus {
    border-color: #0d80a5;

    box-shadow:
        0 0 0 3px
        rgba(13,128,165,0.10);
}

div[data-baseweb="select"] > div {
    min-height: 56px;

    border-radius: 13px !important;

    background-color: white;
}

div.stButton > button {
    min-height: 56px;

    width: 100%;

    border-radius: 13px;

    border: none;

    background:
        linear-gradient(
            110deg,
            #073b5d,
            #007f9a
        );

    color: white;

    font-size: 16px;
    font-weight: 800;

    letter-spacing: 0.5px;

    box-shadow:
        0 8px 18px
        rgba(0, 105, 140, 0.16);
}

div.stButton > button:hover {
    background:
        linear-gradient(
            110deg,
            #052d49,
            #006d83
        );

    transform: translateY(-1px);
}


/* =====================================================
   DOĞRULANDI
   ===================================================== */

.success-banner {
    position: relative;

    overflow: hidden;

    background:
        linear-gradient(
            115deg,
            #08794f,
            #14a570
        );

    border-radius: 17px;

    padding: 18px 22px;

    color: white;

    margin-top: 24px;

    box-shadow:
        0 12px 28px
        rgba(16, 153, 103, 0.20);
}

.success-title {
    font-size: 19px;
    font-weight: 900;
}

.success-subtitle {
    color: #d9fff0;

    font-size: 12px;

    margin-top: 3px;
}


/* =====================================================
   KONTEYNER SONUÇ KARTI
   ===================================================== */

.container-result {
    position: relative;

    overflow: hidden;

    background:
        linear-gradient(
            135deg,
            #071b2a,
            #103d5a
        );

    border-radius: 21px;

    padding: 28px;

    margin-top: 16px;
    margin-bottom: 18px;

    box-shadow:
        0 15px 35px
        rgba(15,42,65,0.17);
}

.container-accent {
    width: 7px;
    height: 100%;

    position: absolute;

    left: 0;
    top: 0;
}

.result-label {
    color: #8dafc2;

    font-size: 10px;

    letter-spacing: 2px;

    font-weight: 800;

    text-transform: uppercase;
}

.result-number {
    color: white;

    font-size: 34px;
    font-weight: 900;

    letter-spacing: 2px;

    margin-top: 3px;
}

.result-divider {
    height: 1px;

    background:
        rgba(255,255,255,0.10);

    margin:
        22px 0;
}

.result-line {
    color: white;

    font-size: 30px;
    font-weight: 900;

    margin-top: 4px;
}


/* =====================================================
   METRİKLER
   ===================================================== */

div[data-testid="stMetric"] {
    background:
        linear-gradient(
            180deg,
            #ffffff,
            #fbfcfd
        );

    border:
        1px solid #dbe4eb;

    border-radius: 15px;

    padding: 17px;

    box-shadow:
        0 5px 18px
        rgba(15,23,42,0.04);
}

div[data-testid="stMetricLabel"] {
    color: #758493;

    font-size: 10px;
    font-weight: 800;

    letter-spacing: 0.8px;

    text-transform: uppercase;
}

div[data-testid="stMetricValue"] {
    color: #112c3e;

    font-size: 20px;
    font-weight: 900;
}


/* =====================================================
   KRİTİK UYARI
   ===================================================== */

@keyframes alertPulse {

    0% {
        box-shadow:
            0 0 0 0
            rgba(220,38,38,0.30);
    }

    70% {
        box-shadow:
            0 0 0 12px
            rgba(220,38,38,0);
    }

    100% {
        box-shadow:
            0 0 0 0
            rgba(220,38,38,0);
    }
}


.danger-card {
    position: relative;

    overflow: hidden;

    background:
        linear-gradient(
            135deg,
            #7f1515,
            #ce2626
        );

    border-radius: 21px;

    padding: 28px;

    margin-top: 24px;

    color: white;

    text-align: center;

    animation:
        alertPulse 2s infinite;

    box-shadow:
        0 15px 35px
        rgba(200,30,30,0.20);
}

.danger-symbol {
    width: 70px;
    height: 70px;

    margin:
        0 auto 12px auto;

    border-radius: 50%;

    background:
        rgba(255,255,255,0.13);

    border:
        2px solid rgba(255,255,255,0.35);

    display: flex;

    align-items: center;
    justify-content: center;

    font-size: 40px;
    font-weight: 900;
}

.danger-title {
    font-size: 30px;
    font-weight: 1000;

    letter-spacing: 0.4px;
}

.danger-container {
    font-size: 27px;
    font-weight: 900;

    margin-top: 15px;

    letter-spacing: 2px;
}

.danger-info {
    color: #ffe1e1;

    font-size: 13px;

    margin-top: 15px;
}

.danger-line {
    color: white;

    font-size: 25px;
    font-weight: 900;

    margin-top: 4px;
}

.danger-stop {
    margin-top: 22px;

    background: white;

    color: #a91616;

    padding: 13px;

    border-radius: 11px;

    font-size: 21px;
    font-weight: 1000;
}


/* =====================================================
   BULUNAMADI
   ===================================================== */

.not-found {
    background:
        linear-gradient(
            135deg,
            #541414,
            #b32121
        );

    border-radius: 20px;

    padding: 28px;

    text-align: center;

    color: white;

    margin-top: 24px;

    box-shadow:
        0 14px 32px
        rgba(180,30,30,0.18);
}

.not-found-title {
    font-size: 27px;
    font-weight: 900;
}

.not-found-number {
    font-size: 25px;
    font-weight: 900;

    margin-top: 12px;

    letter-spacing: 2px;
}

.not-found-stop {
    background: white;

    color: #9c1d1d;

    border-radius: 10px;

    padding: 12px;

    font-size: 18px;
    font-weight: 900;

    margin-top: 20px;
}


/* =====================================================
   FOOTER
   ===================================================== */

.app-footer {
    text-align: center;

    color: #8a99a6;

    font-size: 10px;

    letter-spacing: 0.8px;

    margin-top: 40px;
}


/* =====================================================
   TELEFON
   ===================================================== */

@media (max-width: 650px) {

    .block-container {
        padding-left: 13px;
        padding-right: 13px;
    }

    .hero {
        min-height: 300px;

        padding:
            25px 22px;

        border-radius: 20px;
    }

    .hero-content {
        width: 100%;
    }

    .hero-title {
        font-size: 31px;
    }

    .hero-visual {
        width: 180px;

        right: 8px;
        bottom: 5px;

        opacity: 0.42;
    }

    .result-number {
        font-size: 28px;
    }

    .result-line {
        font-size: 25px;
    }

    .danger-title {
        font-size: 25px;
    }
}

</style>
""")


# =========================================================
# FONKSİYONLAR
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

    if (
        not value
        or value.lower() == "nan"
    ):
        return "-"

    return value


def safe(value):
    return html.escape(
        str(value)
    )


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
            "CONTAINER sütunu bulunamadı."
        )

    df["_SEARCH"] = (
        df["CONTAINER"]
        .apply(normalize_container)
    )

    return df


# =========================================================
# HERO / GÖRSEL
# =========================================================

st.html("""
<div class="hero">

    <div class="hero-glow-one"></div>
    <div class="hero-glow-two"></div>

    <div class="hero-content">

        <div class="hero-brand">
            ALPORT BANJUL
        </div>

        <div class="hero-title">
            Konteyner<br>
            Takip Sistemi
        </div>

        <div class="hero-subtitle">
            Gemi yüklemelerinde doğru konteyner,
            doğru shipping line ve operasyonel
            güvenlik kontrolü.
        </div>

        <div class="hero-badge">
            OPERASYON • KONTEYNER KONTROL
        </div>

    </div>


    <svg
        class="hero-visual"
        viewBox="0 0 500 330"
        xmlns="http://www.w3.org/2000/svg"
    >

        <!-- SU -->
        <path
            d="M30 285 Q80 270 130 285 T230 285 T330 285 T430 285"
            fill="none"
            stroke="#6ED8E8"
            stroke-width="5"
            opacity="0.55"
        />

        <path
            d="M60 305 Q110 290 160 305 T260 305 T360 305 T460 305"
            fill="none"
            stroke="#6ED8E8"
            stroke-width="3"
            opacity="0.30"
        />


        <!-- GEMİ -->
        <path
            d="
            M85 205
            L430 205
            L395 270
            L140 270
            Z
            "
            fill="#E8F6FA"
            opacity="0.96"
        />


        <!-- KÖPRÜ -->
        <rect
            x="320"
            y="145"
            width="75"
            height="60"
            rx="4"
            fill="#E8F6FA"
        />

        <rect
            x="334"
            y="158"
            width="15"
            height="13"
            fill="#1B6680"
        />

        <rect
            x="357"
            y="158"
            width="15"
            height="13"
            fill="#1B6680"
        />


        <!-- KONTEYNER 1 -->
        <rect
            x="125"
            y="150"
            width="74"
            height="52"
            rx="3"
            fill="#F05A47"
        />

        <line
            x1="143"
            y1="153"
            x2="143"
            y2="199"
            stroke="#FFAA9E"
            stroke-width="2"
        />

        <line
            x1="161"
            y1="153"
            x2="161"
            y2="199"
            stroke="#FFAA9E"
            stroke-width="2"
        />

        <line
            x1="179"
            y1="153"
            x2="179"
            y2="199"
            stroke="#FFAA9E"
            stroke-width="2"
        />


        <!-- KONTEYNER 2 -->
        <rect
            x="203"
            y="150"
            width="74"
            height="52"
            rx="3"
            fill="#F4B83A"
        />

        <line
            x1="221"
            y1="153"
            x2="221"
            y2="199"
            stroke="#FFE2A1"
            stroke-width="2"
        />

        <line
            x1="239"
            y1="153"
            x2="239"
            y2="199"
            stroke="#FFE2A1"
            stroke-width="2"
        />

        <line
            x1="257"
            y1="153"
            x2="257"
            y2="199"
            stroke="#FFE2A1"
            stroke-width="2"
        />


        <!-- KONTEYNER 3 -->
        <rect
            x="164"
            y="94"
            width="74"
            height="52"
            rx="3"
            fill="#26B2AE"
        />

        <line
            x1="182"
            y1="97"
            x2="182"
            y2="143"
            stroke="#8FE7E3"
            stroke-width="2"
        />

        <line
            x1="200"
            y1="97"
            x2="200"
            y2="143"
            stroke="#8FE7E3"
            stroke-width="2"
        />

        <line
            x1="218"
            y1="97"
            x2="218"
            y2="143"
            stroke="#8FE7E3"
            stroke-width="2"
        />


        <!-- VİNÇ -->
        <line
            x1="80"
            y1="70"
            x2="80"
            y2="200"
            stroke="#8FD3E4"
            stroke-width="8"
        />

        <line
            x1="80"
            y1="72"
            x2="270"
            y2="72"
            stroke="#8FD3E4"
            stroke-width="7"
        />

        <line
            x1="240"
            y1="72"
            x2="240"
            y2="120"
            stroke="#8FD3E4"
            stroke-width="3"
        />

        <rect
            x="225"
            y="118"
            width="30"
            height="7"
            rx="2"
            fill="#F4B83A"
        />

    </svg>

</div>
""")


# =========================================================
# VERİTABANI
# =========================================================

if not os.path.exists(EXCEL_FILE):

    st.error(
        "Konteyner veri dosyasına ulaşılamıyor."
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
        "Konteyner veritabanı yüklenemedi."
    )

    st.stop()


update_time = datetime.fromtimestamp(
    modified_time
).strftime("%d.%m.%Y • %H:%M")


# =========================================================
# DURUM KARTLARI
# =========================================================

stat1, stat2 = st.columns(2)


with stat1:

    st.html(
        f"""
        <div class="stat-card">

            <div class="stat-accent-blue"></div>

            <div class="stat-icon">
                ▣
            </div>

            <div class="stat-label">
                Güncel Kayıt
            </div>

            <div class="stat-value">
                {len(df):,} Konteyner
            </div>

        </div>
        """
    )


with stat2:

    st.html(
        f"""
        <div class="stat-card">

            <div class="stat-accent-green"></div>

            <div class="stat-icon">
                ◷
            </div>

            <div class="stat-label">
                Son Güncelleme
            </div>

            <div class="stat-value">
                {update_time}
            </div>

        </div>
        """
    )


st.write("")


# =========================================================
# ARAMA PANELİ
# =========================================================

with st.container(
    border=True
):

    st.subheader(
        "Konteyner Doğrulama"
    )

    st.caption(
        "Yükleme hattını seçin ve konteyner numarasını girin."
    )


    available_lines = sorted(
        {
            normalize_line(value)
            for value in df["AGENT"].unique()
            if normalize_line(value) != "-"
        }
    )


    line_options = [
        "Hat seçilmedi"
    ] + available_lines


    with st.form(
        "container_search",
        clear_on_submit=False
    ):

        selected_line = st.selectbox(
            "Yükleme Hattı",
            line_options
        )

        container_input = st.text_input(
            "Konteyner Numarası",
            placeholder="Örnek: SEKU6920313",
            max_chars=20
        )

        search_button = st.form_submit_button(
            "KONTEYNERİ DOĞRULA",
            type="primary",
            use_container_width=True
        )


# =========================================================
# ARAMA SONUCU
# =========================================================

if search_button:

    search_number = normalize_container(
        container_input
    )


    if not search_number:

        st.warning(
            "Lütfen konteyner numarası girin."
        )

        st.stop()


    result = df[
        df["_SEARCH"] == search_number
    ]


    # =====================================================
    # BULUNAMADI
    # =====================================================

    if result.empty:

        st.html(
            f"""
            <div class="not-found">

                <div style="
                    font-size:45px;
                    margin-bottom:8px;
                ">
                    ⓧ
                </div>

                <div class="not-found-title">
                    KONTEYNER BULUNAMADI
                </div>

                <div class="not-found-number">
                    {safe(search_number)}
                </div>

                <div style="
                    color:#ffdede;
                    margin-top:12px;
                    font-size:13px;
                ">
                    Bu konteyner güncel veritabanında
                    bulunmuyor.
                </div>

                <div class="not-found-stop">
                    YÜKLEME YAPMAYIN
                </div>

            </div>
            """
        )


    # =====================================================
    # DUPLICATE
    # =====================================================

    elif len(result) > 1:

        st.error(
            "Aynı konteyner için birden fazla kayıt bulundu."
        )

        st.warning(
            "Yükleme öncesinde Operasyon Departmanı ile teyit edin."
        )


    # =====================================================
    # KONTEYNER BULUNDU
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


        line_color = LINE_COLORS.get(
            shipping_line,
            "#23A6A8"
        )


        wrong_line = (
            selected_line != "Hat seçilmedi"
            and selected_line != shipping_line
        )


        # =================================================
        # YANLIŞ HAT
        # =================================================

        if wrong_line:

            st.html(
                f"""
                <div class="danger-card">

                    <div class="danger-symbol">
                        !
                    </div>

                    <div class="danger-title">
                        YANLIŞ SHIPPING LINE
                    </div>

                    <div class="danger-container">
                        {safe(container)}
                    </div>

                    <div class="danger-info">
                        KONTEYNERİN KAYITLI HATTI
                    </div>

                    <div class="danger-line">
                        {safe(shipping_line)}
                    </div>

                    <div class="danger-info">
                        YÜKLEME İÇİN SEÇİLEN HAT
                    </div>

                    <div class="danger-line">
                        {safe(selected_line)}
                    </div>

                    <div class="danger-stop">
                        BU KONTEYNERİ YÜKLEMEYİN
                    </div>

                </div>
                """
            )


        # =================================================
        # DOĞRU
        # =================================================

        else:

            if selected_line != "Hat seçilmedi":

                st.html(
                    """
                    <div class="success-banner">

                        <div class="success-title">
                            ✓ Yükleme Kontrolü Başarılı
                        </div>

                        <div class="success-subtitle">
                            Konteyner seçilen shipping line ile eşleşiyor.
                        </div>

                    </div>
                    """
                )

            else:

                st.html(
                    """
                    <div class="success-banner">

                        <div class="success-title">
                            ✓ Konteyner Bulundu
                        </div>

                        <div class="success-subtitle">
                            Konteyner güncel veritabanında kayıtlı.
                        </div>

                    </div>
                    """
                )


            # =================================================
            # ANA SONUÇ
            # =================================================

            st.html(
                f"""
                <div class="container-result">

                    <div
                        class="container-accent"
                        style="
                            background:{line_color};
                        "
                    ></div>

                    <div class="result-label">
                        KONTEYNER NUMARASI
                    </div>

                    <div class="result-number">
                        {safe(container)}
                    </div>

                    <div class="result-divider"></div>

                    <div class="result-label">
                        SHIPPING LINE
                    </div>

                    <div
                        class="result-line"
                        style="
                            color:{line_color};
                        "
                    >
                        {safe(shipping_line)}
                    </div>

                </div>
                """
            )


            # =================================================
            # BİLGİLER
            # =================================================

            c1, c2 = st.columns(2)

            with c1:
                st.metric(
                    "Boyut",
                    size
                )

            with c2:
                st.metric(
                    "Konteyner Tipi",
                    container_type
                )


            c3, c4 = st.columns(2)

            with c3:
                st.metric(
                    "Durum",
                    status
                )

            with c4:
                st.metric(
                    "Saha / Konum",
                    location
                )


            c5, c6 = st.columns(2)

            with c5:
                st.metric(
                    "Gemi",
                    vessel
                )

            with c6:
                st.metric(
                    "Sefer",
                    voyage
                )


            # =================================================
            # DETAY
            # =================================================

            if (
                imo_class != "-"
                or discharge_date != "-"
            ):

                with st.expander(
                    "Operasyon Detayları"
                ):

                    if imo_class != "-":

                        st.write(
                            "**IMO Sınıfı:**",
                            imo_class
                        )


                    if discharge_date != "-":

                        st.write(
                            "**Tahliye Tarihi:**",
                            discharge_date
                        )


# =========================================================
# FOOTER
# =========================================================

st.html("""
<div class="app-footer">
    ALPORT BANJUL • KONTEYNER TAKİP SİSTEMİ • OPERASYON
</div>
""")
