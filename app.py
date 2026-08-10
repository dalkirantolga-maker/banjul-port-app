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
MAX_HISTORY = 6


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
# SESSION STATE BAŞLANGIÇ DEĞERLERİ
# =========================================================

if "search_history" not in st.session_state:
    st.session_state.search_history = []

if "high_contrast" not in st.session_state:
    st.session_state.high_contrast = False

if "container_query" not in st.session_state:
    st.session_state.container_query = ""


# =========================================================
# CSS
# =========================================================

BASE_CSS = """
<style>

/* STREAMLIT MENÜLERİNİ GİZLE */

#MainMenu { visibility: hidden; }
footer { visibility: hidden; }
header { visibility: hidden; }


/* SAYFA */

.stApp {
    background:
        radial-gradient(circle at 0% 0%, rgba(0, 159, 227, 0.12), transparent 32%),
        radial-gradient(circle at 100% 25%, rgba(13, 148, 136, 0.08), transparent 28%),
        linear-gradient(180deg, #f7fafc 0%, #edf3f7 100%);
}

.block-container {
    max-width: 960px;
    padding-top: 1.4rem;
    padding-bottom: 3rem;
}


/* =====================================================
   ÜST ARAÇ ÇUBUĞU (kontrast anahtarı)
   ===================================================== */

.top-toolbar {
    display: flex;
    justify-content: flex-end;
    margin-bottom: 6px;
}


/* =====================================================
   HERO
   ===================================================== */

.hero {
    position: relative;
    overflow: hidden;
    min-height: 260px;
    background: linear-gradient(125deg, #061827 0%, #0a3555 50%, #007c91 100%);
    border-radius: 26px;
    padding: 35px 38px;
    margin-bottom: 24px;
    box-shadow: 0 22px 55px rgba(8, 35, 55, 0.22);
    border: 1px solid rgba(255,255,255,0.10);
}

.hero-glow-one {
    position: absolute;
    width: 280px; height: 280px;
    border-radius: 50%;
    background: rgba(49, 190, 220, 0.18);
    right: -90px; top: -110px;
    filter: blur(2px);
}

.hero-glow-two {
    position: absolute;
    width: 180px; height: 180px;
    border-radius: 50%;
    background: rgba(255,255,255,0.05);
    right: 140px; bottom: -100px;
}

.hero-content { position: relative; z-index: 3; width: 55%; }

.hero-brand {
    color: #8ed8e4;
    font-size: 12px; font-weight: 800;
    letter-spacing: 4px;
    margin-bottom: 9px;
}

.hero-title {
    color: white;
    font-size: 40px; font-weight: 900;
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
    background: rgba(255,255,255,0.10);
    border: 1px solid rgba(255,255,255,0.15);
    border-radius: 30px;
    padding: 7px 13px;
    margin-top: 18px;
    font-size: 11px;
}

.hero-visual {
    position: absolute;
    right: 25px; bottom: 18px;
    width: 38%; max-width: 320px;
    opacity: 0.96;
}


/* =====================================================
   ÜST DURUM KARTLARI
   ===================================================== */

.stat-card {
    position: relative;
    overflow: hidden;
    background: rgba(255,255,255,0.94);
    border: 1px solid #dbe5ec;
    border-radius: 17px;
    padding: 20px 22px;
    min-height: 96px;
    box-shadow: 0 7px 22px rgba(15,23,42,0.05);
}

.stat-accent-blue {
    position: absolute; top: 0; left: 0;
    width: 5px; height: 100%;
    background: linear-gradient(#00a1d5, #127ca5);
}

.stat-accent-green {
    position: absolute; top: 0; left: 0;
    width: 5px; height: 100%;
    background: linear-gradient(#18a57b, #087b65);
}

.stat-icon { font-size: 20px; margin-bottom: 5px; }

.stat-label {
    color: #718190;
    font-size: 10px; font-weight: 800;
    text-transform: uppercase;
    letter-spacing: 1.3px;
}

.stat-value {
    color: #102b3d;
    font-size: 22px; font-weight: 900;
    margin-top: 4px;
}


/* =====================================================
   ARAMA PANELİ
   ===================================================== */

div[data-testid="stVerticalBlockBorderWrapper"] {
    border-radius: 20px !important;
    background: rgba(255,255,255,0.92) !important;
    border: 1px solid #d8e2e9 !important;
    box-shadow: 0 10px 32px rgba(15,23,42,0.06);
}

div[data-testid="stTextInput"] input {
    min-height: 64px;
    background: #ffffff;
    border-radius: 14px;
    border: 1px solid #cbd8e1;
    text-align: center;
    font-size: 24px; font-weight: 900;
    letter-spacing: 2px;
    text-transform: uppercase;
}

div[data-testid="stTextInput"] input:focus {
    border-color: #0d80a5;
    box-shadow: 0 0 0 3px rgba(13,128,165,0.10);
}

div[data-testid="stTextArea"] textarea {
    background: #ffffff;
    border-radius: 14px;
    border: 1px solid #cbd8e1;
    font-size: 16px; font-weight: 700;
    letter-spacing: 1px;
    text-transform: uppercase;
    font-family: "Courier New", monospace;
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
    background: linear-gradient(110deg, #073b5d, #007f9a);
    color: white;
    font-size: 16px; font-weight: 800;
    letter-spacing: 0.5px;
    box-shadow: 0 8px 18px rgba(0, 105, 140, 0.16);
}

div.stButton > button:hover {
    background: linear-gradient(110deg, #052d49, #006d83);
    transform: translateY(-1px);
}


/* =====================================================
   FORMAT İPUCU (canlı doğrulama)
   ===================================================== */

.format-hint {
    text-align: center;
    font-size: 12px;
    font-weight: 700;
    letter-spacing: 0.5px;
    margin-top: -6px;
    margin-bottom: 10px;
    padding: 6px;
    border-radius: 8px;
}

.format-ok {
    color: #0a7a4d;
    background: rgba(16,165,103,0.10);
}

.format-bad {
    color: #a91616;
    background: rgba(200,30,30,0.08);
}

.format-empty {
    color: #8a99a6;
    background: transparent;
}


/* =====================================================
   ARAMA GEÇMİŞİ (chip'ler)
   ===================================================== */

.history-label {
    color: #718190;
    font-size: 10px; font-weight: 800;
    text-transform: uppercase;
    letter-spacing: 1px;
    margin-bottom: 6px;
}

div[data-testid="column"] div.stButton > button.history-chip,
.history-row div.stButton > button {
    min-height: 34px;
    width: auto;
    padding: 4px 12px;
    font-size: 12px;
    font-weight: 800;
    border-radius: 30px;
    background: #eef4f8;
    color: #10425c;
    box-shadow: none;
    border: 1px solid #d3e2ea;
}

.history-row div.stButton > button:hover {
    background: #dcebf3;
    transform: none;
}

.history-chip-ok div.stButton > button {
    border-color: #b7e6d2;
    background: #eafbf3;
    color: #0a7a4d;
}

.history-chip-bad div.stButton > button {
    border-color: #f3c4c4;
    background: #fdeeee;
    color: #a91616;
}


/* =====================================================
   DOĞRULANDI
   ===================================================== */

.success-banner {
    position: relative;
    overflow: hidden;
    background: linear-gradient(115deg, #08794f, #14a570);
    border-radius: 17px;
    padding: 18px 22px;
    color: white;
    margin-top: 24px;
    box-shadow: 0 12px 28px rgba(16, 153, 103, 0.20);
}

.success-title { font-size: 19px; font-weight: 900; }
.success-subtitle { color: #d9fff0; font-size: 12px; margin-top: 3px; }


/* =====================================================
   KONTEYNER SONUÇ KARTI
   ===================================================== */

.container-result {
    position: relative;
    overflow: hidden;
    background: linear-gradient(135deg, #071b2a, #103d5a);
    border-radius: 21px;
    padding: 28px;
    margin-top: 16px;
    margin-bottom: 4px;
    box-shadow: 0 15px 35px rgba(15,42,65,0.17);
}

.container-accent { width: 7px; height: 100%; position: absolute; left: 0; top: 0; }

.result-label {
    color: #8dafc2;
    font-size: 10px;
    letter-spacing: 2px;
    font-weight: 800;
    text-transform: uppercase;
}

.result-number {
    color: white;
    font-size: 34px; font-weight: 900;
    letter-spacing: 2px;
    margin-top: 3px;
}

.result-divider { height: 1px; background: rgba(255,255,255,0.10); margin: 22px 0; }

.result-line { color: white; font-size: 30px; font-weight: 900; margin-top: 4px; }

.copy-caption {
    font-size: 11px;
    color: #8a99a6;
    margin-top: -6px;
    margin-bottom: 16px;
}


/* =====================================================
   METRİKLER
   ===================================================== */

div[data-testid="stMetric"] {
    background: linear-gradient(180deg, #ffffff, #fbfcfd);
    border: 1px solid #dbe4eb;
    border-radius: 15px;
    padding: 17px;
    box-shadow: 0 5px 18px rgba(15,23,42,0.04);
}

div[data-testid="stMetricLabel"] {
    color: #758493;
    font-size: 10px; font-weight: 800;
    letter-spacing: 0.8px;
    text-transform: uppercase;
}

div[data-testid="stMetricValue"] { color: #112c3e; font-size: 20px; font-weight: 900; }


/* =====================================================
   KRİTİK UYARI
   ===================================================== */

@keyframes alertPulse {
    0% { box-shadow: 0 0 0 0 rgba(220,38,38,0.30); }
    70% { box-shadow: 0 0 0 12px rgba(220,38,38,0); }
    100% { box-shadow: 0 0 0 0 rgba(220,38,38,0); }
}

.danger-card {
    position: relative;
    overflow: hidden;
    background: linear-gradient(135deg, #7f1515, #ce2626);
    border-radius: 21px;
    padding: 28px;
    margin-top: 24px;
    color: white;
    text-align: center;
    animation: alertPulse 2s infinite;
    box-shadow: 0 15px 35px rgba(200,30,30,0.20);
}

.danger-symbol {
    width: 70px; height: 70px;
    margin: 0 auto 12px auto;
    border-radius: 50%;
    background: rgba(255,255,255,0.13);
    border: 2px solid rgba(255,255,255,0.35);
    display: flex; align-items: center; justify-content: center;
    font-size: 40px; font-weight: 900;
}

.danger-title { font-size: 30px; font-weight: 1000; letter-spacing: 0.4px; }
.danger-container { font-size: 27px; font-weight: 900; margin-top: 15px; letter-spacing: 2px; }
.danger-info { color: #ffe1e1; font-size: 13px; margin-top: 15px; }
.danger-line { color: white; font-size: 25px; font-weight: 900; margin-top: 4px; }

.danger-stop {
    margin-top: 22px;
    background: white;
    color: #a91616;
    padding: 13px;
    border-radius: 11px;
    font-size: 21px; font-weight: 1000;
}


/* =====================================================
   BULUNAMADI
   ===================================================== */

.not-found {
    background: linear-gradient(135deg, #541414, #b32121);
    border-radius: 20px;
    padding: 28px;
    text-align: center;
    color: white;
    margin-top: 24px;
    box-shadow: 0 14px 32px rgba(180,30,30,0.18);
}

.not-found-title { font-size: 27px; font-weight: 900; }
.not-found-number { font-size: 25px; font-weight: 900; margin-top: 12px; letter-spacing: 2px; }

.not-found-stop {
    background: white;
    color: #9c1d1d;
    border-radius: 10px;
    padding: 12px;
    font-size: 18px; font-weight: 900;
    margin-top: 20px;
}


/* =====================================================
   TOPLU DOĞRULAMA SATIRLARI
   ===================================================== */

.batch-row {
    display: flex;
    align-items: center;
    gap: 14px;
    padding: 14px 18px;
    border-radius: 13px;
    margin-bottom: 8px;
    font-weight: 800;
}

.batch-row-ok {
    background: rgba(16,165,103,0.09);
    border: 1px solid rgba(16,165,103,0.25);
    color: #0a7a4d;
}

.batch-row-bad {
    background: rgba(200,30,30,0.07);
    border: 1px solid rgba(200,30,30,0.22);
    color: #a91616;
}

.batch-row-warn {
    background: rgba(244,184,58,0.14);
    border: 1px solid rgba(200,150,20,0.28);
    color: #8a6300;
}

.batch-icon { font-size: 18px; min-width: 22px; text-align: center; }
.batch-number { font-size: 15px; letter-spacing: 1px; min-width: 150px; }
.batch-detail { font-size: 12px; font-weight: 700; opacity: 0.85; flex: 1; }

.batch-summary {
    display: flex;
    gap: 10px;
    margin-bottom: 14px;
}

.batch-pill {
    flex: 1;
    text-align: center;
    padding: 12px;
    border-radius: 12px;
    font-size: 12px;
    font-weight: 800;
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
    .block-container { padding-left: 13px; padding-right: 13px; }
    .hero { min-height: 300px; padding: 25px 22px; border-radius: 20px; }
    .hero-content { width: 100%; }
    .hero-title { font-size: 31px; }
    .hero-visual { width: 180px; right: 8px; bottom: 5px; opacity: 0.42; }
    .result-number { font-size: 28px; }
    .result-line { font-size: 25px; }
    .danger-title { font-size: 25px; }
    .batch-row { flex-wrap: wrap; }
}

</style>
"""

# Yüksek kontrast modu için ek/override kurallar
HIGH_CONTRAST_CSS = """
<style>

.stApp {
    background: #ffffff !important;
}

.stat-card, div[data-testid="stVerticalBlockBorderWrapper"], div[data-testid="stMetric"] {
    border: 2px solid #000000 !important;
    box-shadow: none !important;
}

.stat-value, .result-label, div[data-testid="stMetricValue"] {
    color: #000000 !important;
}

div[data-testid="stTextInput"] input, div[data-testid="stTextArea"] textarea {
    border: 2px solid #000000 !important;
    color: #000000 !important;
}

div.stButton > button {
    background: #000000 !important;
    color: #ffffff !important;
    border: 2px solid #000000 !important;
    box-shadow: none !important;
}

.container-result {
    background: #000000 !important;
}

.danger-card {
    background: #b30000 !important;
    animation: none !important;
}

.not-found {
    background: #7a0000 !important;
}

.success-banner {
    background: #005c33 !important;
}

.batch-row-ok { background: #e6f7ee !important; border: 2px solid #0a7a4d !important; }
.batch-row-bad { background: #fbe6e6 !important; border: 2px solid #a91616 !important; }
.batch-row-warn { background: #fff6df !important; border: 2px solid #8a6300 !important; }

</style>
"""

st.html(BASE_CSS)

if st.session_state.high_contrast:
    st.html(HIGH_CONTRAST_CSS)


# =========================================================
# FONKSİYONLAR
# =========================================================

CONTAINER_FORMAT_RE = re.compile(r"^[A-Z]{4}[0-9]{7}$")


def normalize_container(value):
    if pd.isna(value):
        return ""
    value = str(value).upper().strip()
    return re.sub(r"[^A-Z0-9]", "", value)


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


def safe(value):
    return html.escape(str(value))


def is_valid_format(normalized_value):
    """ISO 6346 temel format kontrolü: 4 harf + 7 rakam (check digit dahil).
    Bu, gerçek check-digit doğrulaması değil, sadece biçim kontrolüdür."""
    return bool(CONTAINER_FORMAT_RE.match(normalized_value))


@st.cache_data(ttl=30)
def load_database(file_name, modified_time):
    df = pd.read_excel(file_name, dtype=str, engine="openpyxl")
    df.columns = [str(column).strip().upper() for column in df.columns]
    df = df.fillna("")

    if "CONTAINER" not in df.columns:
        raise ValueError("CONTAINER sütunu bulunamadı.")

    df["_SEARCH"] = df["CONTAINER"].apply(normalize_container)

    return df


def lookup_container(df, raw_number):
    """Tek bir konteyner numarasını veritabanında arar.
    Dönüş: (status, record_or_none) -- status: 'ok' | 'not_found' | 'duplicate' | 'invalid'"""

    normalized = normalize_container(raw_number)

    if not normalized:
        return "invalid", None, normalized

    if not is_valid_format(normalized):
        return "invalid", None, normalized

    result = df[df["_SEARCH"] == normalized]

    if result.empty:
        return "not_found", None, normalized

    if len(result) > 1:
        return "duplicate", result, normalized

    return "ok", result.iloc[0], normalized


def push_history(number, status):
    entry = {"number": number, "status": status}

    st.session_state.search_history = [
        item for item in st.session_state.search_history
        if item["number"] != number
    ]

    st.session_state.search_history.insert(0, entry)
    st.session_state.search_history = st.session_state.search_history[:MAX_HISTORY]


def select_history_number(number):
    st.session_state.container_query = number


# =========================================================
# ÜST ARAÇ ÇUBUĞU (kontrast anahtarı)
# =========================================================

toolbar_col = st.columns([5, 2])[1]

with toolbar_col:
    st.toggle(
        "Yüksek Kontrast",
        key="high_contrast",
        help="Güneş ışığı altında veya düşük görüş koşullarında okunabilirliği artırır."
    )


# =========================================================
# HERO / GÖRSEL
# =========================================================

st.html("""
<div class="hero" role="banner" aria-label="ALPORT Banjul Konteyner Takip Sistemi">

    <div class="hero-glow-one"></div>
    <div class="hero-glow-two"></div>

    <div class="hero-content">

        <div class="hero-brand">ALPORT BANJUL</div>

        <div class="hero-title">
            Konteyner<br>
            Takip Sistemi
        </div>

        <div class="hero-subtitle">
            Gemi yüklemelerinde doğru konteyner,
            doğru shipping line ve operasyonel
            güvenlik kontrolü.
        </div>

        <div class="hero-badge">OPERASYON • KONTEYNER KONTROL</div>

    </div>

    <svg class="hero-visual" viewBox="0 0 500 330" xmlns="http://www.w3.org/2000/svg" aria-hidden="true">
        <path d="M30 285 Q80 270 130 285 T230 285 T330 285 T430 285" fill="none" stroke="#6ED8E8" stroke-width="5" opacity="0.55" />
        <path d="M60 305 Q110 290 160 305 T260 305 T360 305 T460 305" fill="none" stroke="#6ED8E8" stroke-width="3" opacity="0.30" />
        <path d="M85 205 L430 205 L395 270 L140 270 Z" fill="#E8F6FA" opacity="0.96" />
        <rect x="320" y="145" width="75" height="60" rx="4" fill="#E8F6FA" />
        <rect x="334" y="158" width="15" height="13" fill="#1B6680" />
        <rect x="357" y="158" width="15" height="13" fill="#1B6680" />
        <rect x="125" y="150" width="74" height="52" rx="3" fill="#F05A47" />
        <line x1="143" y1="153" x2="143" y2="199" stroke="#FFAA9E" stroke-width="2" />
        <line x1="161" y1="153" x2="161" y2="199" stroke="#FFAA9E" stroke-width="2" />
        <line x1="179" y1="153" x2="179" y2="199" stroke="#FFAA9E" stroke-width="2" />
        <rect x="203" y="150" width="74" height="52" rx="3" fill="#F4B83A" />
        <line x1="221" y1="153" x2="221" y2="199" stroke="#FFE2A1" stroke-width="2" />
        <line x1="239" y1="153" x2="239" y2="199" stroke="#FFE2A1" stroke-width="2" />
        <line x1="257" y1="153" x2="257" y2="199" stroke="#FFE2A1" stroke-width="2" />
        <rect x="164" y="94" width="74" height="52" rx="3" fill="#26B2AE" />
        <line x1="182" y1="97" x2="182" y2="143" stroke="#8FE7E3" stroke-width="2" />
        <line x1="200" y1="97" x2="200" y2="143" stroke="#8FE7E3" stroke-width="2" />
        <line x1="218" y1="97" x2="218" y2="143" stroke="#8FE7E3" stroke-width="2" />
        <line x1="80" y1="70" x2="80" y2="200" stroke="#8FD3E4" stroke-width="8" />
        <line x1="80" y1="72" x2="270" y2="72" stroke="#8FD3E4" stroke-width="7" />
        <line x1="240" y1="72" x2="240" y2="120" stroke="#8FD3E4" stroke-width="3" />
        <rect x="225" y="118" width="30" height="7" rx="2" fill="#F4B83A" />
    </svg>

</div>
""")


# =========================================================
# VERİTABANI
# =========================================================

if not os.path.exists(EXCEL_FILE):
    st.error("Konteyner veri dosyasına ulaşılamıyor.")
    st.stop()

try:
    with st.spinner("Veritabanı yükleniyor..."):
        modified_time = os.path.getmtime(EXCEL_FILE)
        df = load_database(EXCEL_FILE, modified_time)
except Exception:
    st.error("Konteyner veritabanı yüklenemedi.")
    st.stop()

update_time = datetime.fromtimestamp(modified_time).strftime("%d.%m.%Y • %H:%M")


# =========================================================
# DURUM KARTLARI
# =========================================================

stat1, stat2 = st.columns(2)

with stat1:
    st.html(f"""
        <div class="stat-card" role="status" aria-label="Güncel kayıt sayısı">
            <div class="stat-accent-blue"></div>
            <div class="stat-icon">▣</div>
            <div class="stat-label">Güncel Kayıt</div>
            <div class="stat-value">{len(df):,} Konteyner</div>
        </div>
    """)

with stat2:
    st.html(f"""
        <div class="stat-card" role="status" aria-label="Son güncelleme zamanı">
            <div class="stat-accent-green"></div>
            <div class="stat-icon">◷</div>
            <div class="stat-label">Son Güncelleme</div>
            <div class="stat-value">{update_time}</div>
        </div>
    """)

st.write("")


# =========================================================
# HAT SEÇENEKLERİ
# =========================================================

available_lines = sorted({
    normalize_line(value)
    for value in df["AGENT"].unique()
    if normalize_line(value) != "-"
})

line_options = ["Hat seçilmedi"] + available_lines


# =========================================================
# SEKMELER: TEKLİ ARAMA / TOPLU DOĞRULAMA
# =========================================================

tab_single, tab_batch = st.tabs(["🔍 Tekli Arama", "📋 Toplu Doğrulama"])


# ---------------------------------------------------------
# TEKLİ ARAMA
# ---------------------------------------------------------

with tab_single:

    with st.container(border=True):

        st.subheader("Konteyner Doğrulama")
        st.caption("Yükleme hattını seçin ve konteyner numarasını girin.")

        # Arama geçmişi chip'leri
        if st.session_state.search_history:

            st.html('<div class="history-label">SON ARAMALAR</div>')

            history_cols = st.columns(len(st.session_state.search_history))

            for idx, entry in enumerate(st.session_state.search_history):
                with history_cols[idx]:
                    st.button(
                        entry["number"],
                        key=f"history_{idx}_{entry['number']}",
                        on_click=select_history_number,
                        args=(entry["number"],),
                        use_container_width=True
                    )

            st.write("")

        selected_line = st.selectbox("Yükleme Hattı", line_options, key="single_line_select")

        container_input = st.text_input(
            "Konteyner Numarası",
            placeholder="Örnek: SEKU6920313",
            max_chars=20,
            key="container_query"
        )

        # Canlı format ipucu
        live_normalized = normalize_container(container_input)

        if not live_normalized:
            st.html('<div class="format-hint format-empty">Konteyner numarasını girin (harf + rakam)</div>')
        elif is_valid_format(live_normalized):
            st.html(f'<div class="format-hint format-ok">✓ Format geçerli — {safe(live_normalized)}</div>')
        else:
            st.html(
                f'<div class="format-hint format-bad">⚠ Format hatalı olabilir — beklenen: 4 harf + 7 rakam '
                f'(örn. MSKU1234567) — girilen: {safe(live_normalized)}</div>'
            )

        search_clicked = st.button(
            "KONTEYNERİ DOĞRULA",
            type="primary",
            use_container_width=True,
            key="single_search_button"
        )

    if search_clicked:

        if not live_normalized:
            st.warning("Lütfen konteyner numarası girin.")
            st.stop()

        with st.spinner("Kontrol ediliyor..."):
            status, record_or_result, normalized = lookup_container(df, container_input)

        push_history(normalized, status)

        # =================================================
        # GEÇERSİZ FORMAT
        # =================================================

        if status == "invalid":
            st.html(f"""
                <div class="not-found" role="alert">
                    <div style="font-size:45px; margin-bottom:8px;">⚠</div>
                    <div class="not-found-title">GEÇERSİZ FORMAT</div>
                    <div class="not-found-number">{safe(normalized)}</div>
                    <div style="color:#ffdede; margin-top:12px; font-size:13px;">
                        Konteyner numarası standart formata (4 harf + 7 rakam) uymuyor.
                        Lütfen tekrar kontrol edin.
                    </div>
                </div>
            """)

        # =================================================
        # BULUNAMADI
        # =================================================

        elif status == "not_found":
            st.html(f"""
                <div class="not-found" role="alert">
                    <div style="font-size:45px; margin-bottom:8px;">ⓧ</div>
                    <div class="not-found-title">KONTEYNER BULUNAMADI</div>
                    <div class="not-found-number">{safe(normalized)}</div>
                    <div style="color:#ffdede; margin-top:12px; font-size:13px;">
                        Bu konteyner güncel veritabanında bulunmuyor.
                    </div>
                    <div class="not-found-stop">YÜKLEME YAPMAYIN</div>
                </div>
            """)

        # =================================================
        # DUPLICATE
        # =================================================

        elif status == "duplicate":
            st.error("Aynı konteyner için birden fazla kayıt bulundu.")
            st.warning("Yükleme öncesinde Operasyon Departmanı ile teyit edin.")

        # =================================================
        # KONTEYNER BULUNDU
        # =================================================

        else:
            record = record_or_result

            container = clean_value(record, "CONTAINER")
            shipping_line = normalize_line(clean_value(record, "AGENT"))
            size = clean_value(record, "SIZE")
            container_type = clean_value(record, "TYPE")
            status_val = clean_value(record, "FULL-MTY")
            location = clean_value(record, "AREA")
            vessel = clean_value(record, "VESSEL NAME")
            voyage = clean_value(record, "VOYAGE NUMBER")
            imo_class = clean_value(record, "IMO CLS")
            discharge_date = clean_value(record, "DISCHARGE DATE")

            line_color = LINE_COLORS.get(shipping_line, "#23A6A8")

            wrong_line = (
                selected_line != "Hat seçilmedi"
                and selected_line != shipping_line
            )

            if wrong_line:
                st.html(f"""
                    <div class="danger-card" role="alert">
                        <div class="danger-symbol">!</div>
                        <div class="danger-title">YANLIŞ SHIPPING LINE</div>
                        <div class="danger-container">{safe(container)}</div>
                        <div class="danger-info">KONTEYNERİN KAYITLI HATTI</div>
                        <div class="danger-line">{safe(shipping_line)}</div>
                        <div class="danger-info">YÜKLEME İÇİN SEÇİLEN HAT</div>
                        <div class="danger-line">{safe(selected_line)}</div>
                        <div class="danger-stop">BU KONTEYNERİ YÜKLEMEYİN</div>
                    </div>
                """)

            else:
                if selected_line != "Hat seçilmedi":
                    st.html("""
                        <div class="success-banner" role="status">
                            <div class="success-title">✓ Yükleme Kontrolü Başarılı</div>
                            <div class="success-subtitle">Konteyner seçilen shipping line ile eşleşiyor.</div>
                        </div>
                    """)
                else:
                    st.html("""
                        <div class="success-banner" role="status">
                            <div class="success-title">✓ Konteyner Bulundu</div>
                            <div class="success-subtitle">Konteyner güncel veritabanında kayıtlı.</div>
                        </div>
                    """)

                st.html(f"""
                    <div class="container-result">
                        <div class="container-accent" style="background:{line_color};"></div>
                        <div class="result-label">KONTEYNER NUMARASI</div>
                        <div class="result-number">{safe(container)}</div>
                        <div class="result-divider"></div>
                        <div class="result-label">SHIPPING LINE</div>
                        <div class="result-line" style="color:{line_color};">{safe(shipping_line)}</div>
                    </div>
                """)

                st.html('<div class="copy-caption">Numarayı kopyalamak için aşağıdaki kutunun sağ üstündeki simgeye tıklayın</div>')
                st.code(container, language=None)

                c1, c2 = st.columns(2)
                with c1:
                    st.metric("Boyut", size)
                with c2:
                    st.metric("Konteyner Tipi", container_type)

                c3, c4 = st.columns(2)
                with c3:
                    st.metric("Durum", status_val)
                with c4:
                    st.metric("Saha / Konum", location)

                c5, c6 = st.columns(2)
                with c5:
                    st.metric("Gemi", vessel)
                with c6:
                    st.metric("Sefer", voyage)

                if imo_class != "-" or discharge_date != "-":
                    with st.expander("Operasyon Detayları"):
                        if imo_class != "-":
                            st.write("**IMO Sınıfı:**", imo_class)
                        if discharge_date != "-":
                            st.write("**Tahliye Tarihi:**", discharge_date)


# ---------------------------------------------------------
# TOPLU DOĞRULAMA
# ---------------------------------------------------------

with tab_batch:

    with st.container(border=True):

        st.subheader("Toplu Konteyner Doğrulama")
        st.caption("Her satıra bir konteyner numarası gelecek şekilde yapıştırın (yükleme listesinden kopyala-yapıştır yapabilirsiniz).")

        batch_line = st.selectbox("Yükleme Hattı (opsiyonel)", line_options, key="batch_line_select")

        batch_text = st.text_area(
            "Konteyner Numaraları",
            placeholder="MSKU1234567\nTCLU7654321\nCMAU9988776\n...",
            height=180,
            key="batch_query"
        )

        batch_clicked = st.button(
            "HEPSİNİ DOĞRULA",
            type="primary",
            use_container_width=True,
            key="batch_search_button"
        )

    if batch_clicked:

        raw_lines = [line.strip() for line in batch_text.splitlines() if line.strip()]

        if not raw_lines:
            st.warning("Lütfen en az bir konteyner numarası girin.")
            st.stop()

        with st.spinner(f"{len(raw_lines)} konteyner kontrol ediliyor..."):

            seen_in_batch = {}
            rows = []

            for raw in raw_lines:
                status, record_or_result, normalized = lookup_container(df, raw)

                # Yapıştırılan liste içindeki kendi tekrarları
                if normalized in seen_in_batch:
                    rows.append({
                        "number": normalized,
                        "status": "repeat_in_list",
                        "detail": "Bu liste içinde tekrar ediyor"
                    })
                    continue

                seen_in_batch[normalized] = True

                if status == "invalid":
                    rows.append({"number": normalized or raw, "status": "invalid", "detail": "Geçersiz format"})

                elif status == "not_found":
                    rows.append({"number": normalized, "status": "not_found", "detail": "Veritabanında yok — YÜKLEMEYİN"})

                elif status == "duplicate":
                    rows.append({"number": normalized, "status": "duplicate", "detail": "Birden fazla kayıt — teyit gerekli"})

                else:
                    record = record_or_result
                    shipping_line = normalize_line(clean_value(record, "AGENT"))
                    vessel = clean_value(record, "VESSEL NAME")

                    wrong_line = (
                        batch_line != "Hat seçilmedi"
                        and batch_line != shipping_line
                    )

                    if wrong_line:
                        rows.append({
                            "number": normalized,
                            "status": "wrong_line",
                            "detail": f"Kayıtlı hat: {shipping_line} (seçilen: {batch_line}) — YÜKLEMEYİN"
                        })
                    else:
                        rows.append({
                            "number": normalized,
                            "status": "ok",
                            "detail": f"{shipping_line} • {vessel}" if vessel != "-" else shipping_line
                        })

                push_history(normalized, status)

        # Özet sayaçları
        ok_count = sum(1 for r in rows if r["status"] == "ok")
        bad_count = len(rows) - ok_count

        st.html(f"""
            <div class="batch-summary">
                <div class="batch-pill" style="background:rgba(16,165,103,0.12); color:#0a7a4d;">
                    ✓ {ok_count} Uygun
                </div>
                <div class="batch-pill" style="background:rgba(200,30,30,0.09); color:#a91616;">
                    ✗ {bad_count} Sorunlu
                </div>
                <div class="batch-pill" style="background:rgba(15,23,42,0.06); color:#374151;">
                    Toplam {len(rows)}
                </div>
            </div>
        """)

        icon_map = {
            "ok": ("✓", "batch-row-ok"),
            "not_found": ("ⓧ", "batch-row-bad"),
            "invalid": ("⚠", "batch-row-warn"),
            "duplicate": ("⚠", "batch-row-warn"),
            "wrong_line": ("!", "batch-row-bad"),
            "repeat_in_list": ("↻", "batch-row-warn"),
        }

        for row in rows:
            icon, css_class = icon_map.get(row["status"], ("?", "batch-row-warn"))

            st.html(f"""
                <div class="batch-row {css_class}" role="listitem">
                    <div class="batch-icon">{icon}</div>
                    <div class="batch-number">{safe(row['number'])}</div>
                    <div class="batch-detail">{safe(row['detail'])}</div>
                </div>
            """)


# =========================================================
# FOOTER
# =========================================================

st.html("""
<div class="app-footer">
    ALPORT BANJUL • KONTEYNER TAKİP SİSTEMİ • OPERASYON
</div>
""")
