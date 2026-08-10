import os
import re
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


# =========================================================
# TASARIM
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
        linear-gradient(
            180deg,
            #f3f6f9 0%,
            #edf1f5 100%
        );
}

.block-container {
    max-width: 920px;
    padding-top: 2rem;
    padding-bottom: 3rem;
}


/* ======================================================
   ÜST BÖLÜM
   ====================================================== */

.top-panel {
    background:
        linear-gradient(
            135deg,
            #081b2b 0%,
            #0d2f4a 52%,
            #124969 100%
        );

    border-radius: 22px;
    padding: 34px 38px;

    box-shadow:
        0 18px 45px rgba(15, 23, 42, 0.16);

    border:
        1px solid rgba(255,255,255,0.08);

    margin-bottom: 24px;
}

.company-name {
    color: #93b8cf;
    font-size: 12px;
    font-weight: 700;
    letter-spacing: 4px;
    text-transform: uppercase;
}

.system-title {
    color: white;
    font-size: 36px;
    font-weight: 800;
    margin-top: 7px;
    letter-spacing: -0.5px;
}

.system-description {
    color: #ccdde7;
    margin-top: 8px;
    font-size: 15px;
    line-height: 1.6;
}


/* ======================================================
   DURUM KARTLARI
   ====================================================== */

.status-card {
    background: #ffffff;
    border: 1px solid #dce3ea;
    border-radius: 16px;

    padding: 19px 21px;

    box-shadow:
        0 5px 18px rgba(15,23,42,0.045);
}

.status-label {
    color: #758292;
    font-size: 11px;
    letter-spacing: 1px;
    text-transform: uppercase;
    font-weight: 700;
}

.status-value {
    color: #11283b;
    font-size: 22px;
    font-weight: 800;
    margin-top: 5px;
}


/* ======================================================
   ARAMA PANELİ
   ====================================================== */

div[data-testid="stVerticalBlockBorderWrapper"] {
    background: #ffffff !important;
    border: 1px solid #d9e1e8 !important;
    border-radius: 18px !important;

    box-shadow:
        0 8px 25px rgba(15,23,42,0.05);

    padding: 6px;
}

div[data-testid="stTextInput"] input {
    height: 62px;

    border-radius: 12px;

    text-align: center;

    font-size: 24px;
    font-weight: 700;

    letter-spacing: 2px;
    text-transform: uppercase;

    background: #fbfcfd;

    border: 1px solid #cfd8e2;
}

div[data-testid="stTextInput"] input:focus {
    border-color: #1c577e;
    box-shadow:
        0 0 0 3px rgba(28,87,126,0.08);
}

div[data-baseweb="select"] > div {
    min-height: 56px;
    border-radius: 12px !important;
    background: #fbfcfd;
}

div.stButton > button {
    width: 100%;
    min-height: 55px;

    border-radius: 12px;

    font-size: 16px;
    font-weight: 700;

    letter-spacing: 0.3px;

    background:
        linear-gradient(
            135deg,
            #123f60,
            #185d84
        );

    border: none;
}

div.stButton > button:hover {
    background:
        linear-gradient(
            135deg,
            #0f3551,
            #154d70
        );
}


/* ======================================================
   SONUÇ KARTI
   ====================================================== */

.result-card {
    background: #ffffff;

    border: 1px solid #dce3e9;

    border-radius: 18px;

    padding: 28px;

    margin-top: 20px;

    box-shadow:
        0 10px 28px rgba(15,23,42,0.06);
}

.result-small-label {
    color: #7b8896;
    font-size: 11px;
    font-weight: 700;
    letter-spacing: 1.5px;
    text-transform: uppercase;
}

.result-container {
    color: #10293d;
    font-size: 34px;
    font-weight: 800;
    letter-spacing: 2px;
    margin-top: 4px;
}

.result-line {
    color: #143e5e;
    font-size: 29px;
    font-weight: 800;
    margin-top: 4px;
}


/* ======================================================
   BAŞARILI SONUÇ
   ====================================================== */

.success-banner {
    background:
        linear-gradient(
            135deg,
            #eaf6ef,
            #f4fbf7
        );

    border-left: 5px solid #278657;

    color: #20613f;

    border-radius: 13px;

    padding: 16px 18px;

    font-size: 18px;
    font-weight: 700;

    margin-top: 20px;
}


/* ======================================================
   HATALI HAT
   ====================================================== */

.danger-banner {
    background:
        linear-gradient(
            135deg,
            #fff0f0,
            #fff7f7
        );

    border: 2px solid #c93939;

    border-radius: 16px;

    padding: 22px;

    margin-top: 20px;

    text-align: center;
}

.danger-title {
    color: #b42323;
    font-size: 30px;
    font-weight: 900;
}

.danger-text {
    color: #7b2929;
    margin-top: 8px;
    font-size: 15px;
}

.do-not-load {
    background: #b42323;

    color: white;

    border-radius: 10px;

    padding: 12px 15px;

    margin-top: 18px;

    font-size: 20px;
    font-weight: 800;
}


/* ======================================================
   METRİK KARTLARI
   ====================================================== */

div[data-testid="stMetric"] {
    background: #ffffff;

    border: 1px solid #dde4ea;

    border-radius: 14px;

    padding: 17px;

    box-shadow:
        0 4px 15px rgba(15,23,42,0.035);
}

div[data-testid="stMetricLabel"] {
    color: #7b8793;

    font-size: 11px;

    text-transform: uppercase;

    letter-spacing: 0.7px;

    font-weight: 700;
}

div[data-testid="stMetricValue"] {
    color: #12283a;

    font-size: 20px;

    font-weight: 800;
}


/* ======================================================
   ALT BÖLÜM
   ====================================================== */

.footer-text {
    color: #929eaa;

    text-align: center;

    font-size: 11px;

    margin-top: 40px;

    letter-spacing: 0.4px;
}


/* ======================================================
   MOBİL
   ====================================================== */

@media (max-width: 600px) {

    .block-container {
        padding-left: 1rem;
        padding-right: 1rem;
    }

    .top-panel {
        padding: 27px 22px;
        border-radius: 18px;
    }

    .system-title {
        font-size: 29px;
    }

    .system-description {
        font-size: 13px;
    }

    .result-container {
        font-size: 28px;
    }

    .result-line {
        font-size: 25px;
    }

    div[data-testid="stTextInput"] input {
        font-size: 21px;
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
            "Excel dosyasında CONTAINER sütunu bulunamadı."
        )

    df["_SEARCH"] = (
        df["CONTAINER"]
        .apply(normalize_container)
    )

    return df


# =========================================================
# ÜST PANEL
# =========================================================

st.html("""
<div class="top-panel">

    <div class="company-name">
        ALPORT BANJUL
    </div>

    <div class="system-title">
        Konteyner Takip Sistemi
    </div>

    <div class="system-description">
        Gemi yüklemelerinde doğru konteyner ve doğru hat kontrolü için
        operasyon destek sistemi
    </div>

</div>
""")


# =========================================================
# VERİTABANI
# =========================================================

if not os.path.exists(EXCEL_FILE):

    st.error(
        "Konteyner veri dosyasına ulaşılamıyor."
    )

    st.info(
        "Lütfen Operasyon Departmanı ile iletişime geçin."
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

    st.info(
        "Lütfen Operasyon Departmanı ile iletişime geçin."
    )

    st.stop()


update_time = datetime.fromtimestamp(
    modified_time
).strftime("%d.%m.%Y %H:%M")


# =========================================================
# VERİ DURUMU
# =========================================================

status1, status2 = st.columns(2)

with status1:

    st.html(
        f"""
        <div class="status-card">

            <div class="status-label">
                Kayıtlı Konteyner
            </div>

            <div class="status-value">
                {len(df):,}
            </div>

        </div>
        """
    )


with status2:

    st.html(
        f"""
        <div class="status-card">

            <div class="status-label">
                Son Güncelleme
            </div>

            <div class="status-value">
                {update_time}
            </div>

        </div>
        """
    )


st.write("")


# =========================================================
# ARAMA PANELİ
# =========================================================

with st.container(border=True):

    st.subheader(
        "Konteyner Kontrolü"
    )

    st.caption(
        "Yükleme yapılacak hattı seçin ve konteyner numarasını girin."
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
        "container_search_form",
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
            "KONTEYNERİ KONTROL ET",
            type="primary",
            use_container_width=True
        )


# =========================================================
# ARAMA
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
            <div class="danger-banner">

                <div class="danger-title">
                    KONTEYNER BULUNAMADI
                </div>

                <div style="
                    font-size:26px;
                    font-weight:800;
                    color:#172b3b;
                    margin-top:12px;
                    letter-spacing:2px;
                ">
                    {search_number}
                </div>

                <div class="danger-text">
                    Bu konteyner güncel veritabanında bulunmamaktadır.
                </div>

                <div class="do-not-load">
                    YÜKLEME YAPMAYIN
                </div>

                <div class="danger-text">
                    Yükleme öncesinde Operasyon Departmanı ile teyit edin.
                </div>

            </div>
            """
        )


    # =====================================================
    # MÜKERRER KAYIT
    # =====================================================

    elif len(result) > 1:

        st.warning(
            "Aynı konteyner numarası için birden fazla kayıt bulundu."
        )

        st.markdown(
            f"### {search_number}"
        )

        st.error(
            "YÜKLEME ÖNCESİ OPERASYON DEPARTMANI İLE TEYİT EDİN"
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
                <div class="danger-banner">

                    <div class="danger-title">
                        YANLIŞ HAT
                    </div>

                    <div style="
                        font-size:27px;
                        font-weight:800;
                        color:#172b3b;
                        margin-top:14px;
                        letter-spacing:2px;
                    ">
                        {container}
                    </div>

                    <div class="danger-text">
                        Konteynerin kayıtlı hattı:
                    </div>

                    <div style="
                        font-size:26px;
                        font-weight:800;
                        color:#172b3b;
                        margin-top:3px;
                    ">
                        {shipping_line}
                    </div>

                    <div class="danger-text">
                        Seçilen yükleme hattı:
                    </div>

                    <div style="
                        font-size:23px;
                        font-weight:800;
                        color:#172b3b;
                        margin-top:3px;
                    ">
                        {selected_line}
                    </div>

                    <div class="do-not-load">
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
                        ✓ Konteyner ve yükleme hattı doğrulandı
                    </div>
                    """
                )

            else:

                st.html(
                    """
                    <div class="success-banner">
                        ✓ Konteyner bulundu
                    </div>
                    """
                )


            st.html(
                f"""
                <div class="result-card">

                    <div class="result-small-label">
                        KONTEYNER NUMARASI
                    </div>

                    <div class="result-container">
                        {container}
                    </div>

                    <div style="height:20px;"></div>

                    <div class="result-small-label">
                        SHIPPING LINE
                    </div>

                    <div class="result-line">
                        {shipping_line}
                    </div>

                </div>
                """
            )


            # =================================================
            # DETAYLAR
            # =================================================

            col1, col2 = st.columns(2)

            with col1:
                st.metric(
                    "Boyut",
                    size
                )

            with col2:
                st.metric(
                    "Tip",
                    container_type
                )


            col3, col4 = st.columns(2)

            with col3:
                st.metric(
                    "Durum",
                    status
                )

            with col4:
                st.metric(
                    "Saha / Konum",
                    location
                )


            col5, col6 = st.columns(2)

            with col5:
                st.metric(
                    "Gemi",
                    vessel
                )

            with col6:
                st.metric(
                    "Sefer",
                    voyage
                )


            if (
                imo_class != "-"
                or discharge_date != "-"
            ):

                with st.expander(
                    "Diğer Bilgiler"
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
# ALT BÖLÜM
# =========================================================

st.html("""
<div class="footer-text">
    ALPORT BANJUL • KONTEYNER TAKİP SİSTEMİ • OPERASYON DEPARTMANI
</div>
""")
