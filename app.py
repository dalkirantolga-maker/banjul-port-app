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
    page_icon="⚓",
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


# Her hat için görsel vurgu rengi (marka rengi, brass zeminle uyumlu tutuldu)
LINE_COLORS = {
    "MAERSK": "#1E7FA6",
    "CMA CGM": "#C1501F",
    "MSC": "#B08D3E",
    "HAPAG-LLOYD": "#C1501F",
    "ONE": "#9C2B5E",
    "COSCO": "#1E4E8C",
    "PIL": "#9C2B2B",
    "OBT": "#1F6E4A"
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
# CSS — "LİMAN KAYIT DEFTERİ" KİMLİĞİ
# Navy + pirinç (brass) + parşömen zemin, deniz haritası dokusu
# =========================================================

BASE_CSS = """
<style>

@import url('https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,500;9..144,600;9..144,700&family=Inter:wght@400;500;600;700;800&family=IBM+Plex+Mono:wght@500;600;700&display=swap');

:root {
    --navy: #0B1F30;
    --navy-deep: #050D16;
    --steel: #163A52;
    --brass: #D9B84A;
    --brass-deep: #C9A227;
    --surface: #0F2033;
    --surface-border: rgba(201,162,39,0.22);
    --parchment: #F3EFE2;
    --parchment-line: rgba(201,162,39,0.09);
    --ink: #E7E4D8;
    --ink-soft: #85A0B3;
    --alert: #C24545;
    --verified: #35A97C;
}

#MainMenu { visibility: hidden; }
footer { visibility: hidden; }
header { visibility: hidden; }


/* =====================================================
   TEMEL TİPOGRAFİ
   ===================================================== */

html, body, [class*="css"] {
    font-family: 'Inter', sans-serif;
    color: var(--ink);
}

h1, h2, h3, .stMarkdown h3 {
    font-family: 'Fraunces', serif !important;
    font-weight: 600 !important;
    color: var(--parchment) !important;
    letter-spacing: -0.3px;
}

p, span, label, .stCaption, div[data-testid="stCaptionContainer"] {
    color: var(--ink) !important;
}

code, .stCode, div[data-testid="stTextInput"] input, div[data-testid="stTextArea"] textarea {
    font-family: 'IBM Plex Mono', monospace !important;
}


/* =====================================================
   SAYFA ZEMİNİ — deniz haritası kağıdı dokusu
   ===================================================== */

.stApp {
    background-color: #060F1A;
    background-image:
        linear-gradient(var(--parchment-line) 1px, transparent 1px),
        linear-gradient(90deg, var(--parchment-line) 1px, transparent 1px),
        radial-gradient(circle at 8% 0%, rgba(201,162,39,0.09), transparent 32%),
        radial-gradient(circle at 100% 100%, rgba(31,110,124,0.08), transparent 35%);
    background-size: 42px 42px, 42px 42px, 100% 100%, 100% 100%;
    background-attachment: fixed;
}

.block-container {
    max-width: 960px;
    padding-top: 1.4rem;
    padding-bottom: 3rem;
}


/* =====================================================
   ÜST ARAÇ ÇUBUĞU
   ===================================================== */

.top-toolbar { display: flex; justify-content: flex-end; margin-bottom: 6px; }

.eyebrow {
    font-family: 'Inter', sans-serif;
    font-size: 10.5px;
    font-weight: 800;
    letter-spacing: 2.6px;
    text-transform: uppercase;
    color: var(--brass-deep);
}


/* =====================================================
   HERO — "kaptan köşkü" paneli
   ===================================================== */

@keyframes heroFadeIn {
    from { opacity: 0; transform: translateY(6px); }
    to { opacity: 1; transform: translateY(0); }
}

.hero {
    position: relative;
    overflow: hidden;
    min-height: 275px;
    background:
        radial-gradient(circle at 88% 15%, rgba(201,162,39,0.16), transparent 45%),
        linear-gradient(155deg, var(--navy-deep) 0%, var(--navy) 55%, #0E2A40 100%);
    border-radius: 6px;
    padding: 40px 42px;
    margin-bottom: 24px;
    box-shadow: 0 24px 55px rgba(7, 22, 36, 0.35);
    border: 1px solid rgba(201,162,39,0.35);
}

.hero::before,
.hero::after {
    content: "";
    position: absolute;
    left: 18px; right: 18px;
    height: 1px;
    background: rgba(201,162,39,0.30);
}
.hero::before { top: 12px; }
.hero::after { bottom: 12px; }

.hero-content {
    position: relative;
    z-index: 3;
    width: 58%;
    animation: heroFadeIn 0.6s ease-out;
}

.hero-brand {
    color: var(--brass);
    font-size: 11.5px;
    font-weight: 800;
    letter-spacing: 4px;
    margin-bottom: 12px;
    text-transform: uppercase;
}

.hero-title {
    font-family: 'Fraunces', serif;
    color: #F7F3E8;
    font-size: 42px;
    font-weight: 600;
    line-height: 1.08;
    letter-spacing: -0.5px;
}

.hero-subtitle {
    color: #B9C7D2;
    font-size: 14px;
    margin-top: 16px;
    max-width: 430px;
    line-height: 1.7;
    font-weight: 400;
}

.hero-badge {
    display: inline-flex;
    align-items: center;
    gap: 8px;
    color: var(--brass);
    background: rgba(201,162,39,0.08);
    border: 1px solid rgba(201,162,39,0.45);
    border-radius: 3px;
    padding: 7px 14px;
    margin-top: 20px;
    font-size: 10.5px;
    font-weight: 700;
    letter-spacing: 1.4px;
}

.hero-visual {
    position: absolute;
    right: 0px;
    top: 0px;
    width: 46%;
    height: 100%;
    opacity: 0.9;
}


/* =====================================================
   ÜST DURUM KARTLARI — "kayıt defteri" fişleri
   ===================================================== */

.stat-card {
    position: relative;
    overflow: hidden;
    background: var(--surface);
    border: 1px solid var(--surface-border);
    border-radius: 4px;
    padding: 20px 22px;
    min-height: 96px;
    box-shadow: 0 4px 18px rgba(0,0,0,0.28);
}

.stat-accent-blue, .stat-accent-green {
    position: absolute; top: 0; left: 0;
    width: 4px; height: 100%;
    background: var(--steel);
}
.stat-accent-green { background: var(--brass); }

.stat-icon { font-size: 18px; margin-bottom: 6px; color: var(--brass); }

.stat-label {
    color: var(--ink-soft);
    font-size: 10px; font-weight: 800;
    text-transform: uppercase;
    letter-spacing: 1.4px;
}

.stat-value {
    font-family: 'Fraunces', serif;
    color: var(--parchment);
    font-size: 23px; font-weight: 600;
    margin-top: 5px;
}


/* =====================================================
   ARAMA PANELİ
   ===================================================== */

div[data-testid="stVerticalBlockBorderWrapper"] {
    border-radius: 5px !important;
    background: var(--surface) !important;
    border: 1px solid var(--surface-border) !important;
    box-shadow: 0 8px 30px rgba(0,0,0,0.30);
}

div[data-testid="stTextInput"] input {
    min-height: 62px;
    background: #0A1826;
    border-radius: 3px;
    border: 1.5px solid rgba(201,162,39,0.35);
    text-align: center;
    font-size: 22px; font-weight: 600;
    letter-spacing: 2px;
    text-transform: uppercase;
    color: var(--parchment);
}

div[data-testid="stTextInput"] input:focus {
    border-color: var(--brass);
    box-shadow: 0 0 0 3px rgba(201,162,39,0.18);
}

div[data-testid="stTextArea"] textarea {
    background: #0A1826;
    border-radius: 3px;
    border: 1.5px solid rgba(201,162,39,0.35);
    font-size: 15px; font-weight: 600;
    letter-spacing: 1px;
    text-transform: uppercase;
    color: var(--parchment);
}

div[data-baseweb="select"] > div {
    min-height: 54px;
    border-radius: 3px !important;
    background-color: #0A1826;
    border-color: rgba(201,162,39,0.35) !important;
    color: var(--parchment) !important;
}

div[data-baseweb="select"] span { color: var(--parchment) !important; }

div.stButton > button {
    min-height: 54px;
    width: 100%;
    border-radius: 3px;
    border: 1px solid var(--brass);
    background: linear-gradient(135deg, var(--brass-deep), var(--brass));
    color: var(--navy-deep);
    font-family: 'Inter', sans-serif;
    font-size: 14.5px; font-weight: 800;
    letter-spacing: 1.6px;
    text-transform: uppercase;
    box-shadow: 0 8px 20px rgba(201,162,39,0.18);
    transition: all 0.15s ease;
}

div.stButton > button:hover {
    background: linear-gradient(135deg, var(--brass), #EBCF6E);
    transform: translateY(-1px);
}


/* =====================================================
   FORMAT İPUCU
   ===================================================== */

.format-hint {
    text-align: center;
    font-family: 'IBM Plex Mono', monospace;
    font-size: 11.5px;
    font-weight: 600;
    letter-spacing: 0.4px;
    margin-top: -6px;
    margin-bottom: 12px;
    padding: 7px;
    border-radius: 3px;
}

.format-ok { color: #6FE3A8; background: rgba(53,169,124,0.12); }
.format-bad { color: #F0A0A0; background: rgba(194,69,69,0.10); }
.format-empty { color: var(--ink-soft); background: transparent; }


/* =====================================================
   ARAMA GEÇMİŞİ
   ===================================================== */

.history-label {
    color: var(--ink-soft);
    font-size: 10px; font-weight: 800;
    text-transform: uppercase;
    letter-spacing: 1.4px;
    margin-bottom: 8px;
}

div.stButton > button[kind="secondary"] {
    min-height: 32px;
    width: auto;
    padding: 4px 13px;
    font-size: 11.5px;
    font-weight: 700;
    letter-spacing: 0.5px;
    text-transform: none;
    border-radius: 30px;
    background: #0A1826;
    color: var(--parchment);
    box-shadow: none;
    border: 1px solid rgba(201,162,39,0.30);
}

div.stButton > button[kind="secondary"]:hover {
    background: rgba(201,162,39,0.12);
    transform: none;
    border-color: var(--brass);
}


/* =====================================================
   DOĞRULANDI BANNER
   ===================================================== */

.success-banner {
    position: relative;
    overflow: hidden;
    background: var(--verified);
    border-radius: 4px;
    padding: 17px 22px;
    color: white;
    margin-top: 24px;
    box-shadow: 0 10px 24px rgba(31,110,74,0.20);
}

.success-title { font-family: 'Fraunces', serif; font-size: 18px; font-weight: 600; }
.success-subtitle { color: #D6EEE1; font-size: 12px; margin-top: 3px; }


/* =====================================================
   KONTEYNER SONUÇ KARTI — resmi manifesto fişi
   ===================================================== */

@keyframes cardReveal {
    from { opacity: 0; transform: translateY(8px); }
    to { opacity: 1; transform: translateY(0); }
}

.container-result {
    position: relative;
    overflow: visible;
    background: var(--navy);
    background-image: radial-gradient(circle at 100% 0%, rgba(201,162,39,0.12), transparent 55%);
    border-radius: 4px;
    padding: 30px 30px 26px 30px;
    margin-top: 18px;
    margin-bottom: 4px;
    box-shadow: 0 18px 38px rgba(11,31,48,0.22);
    border-top: 3px dashed rgba(201,162,39,0.55);
    animation: cardReveal 0.35s ease-out;
}

.container-accent { width: 5px; height: 100%; position: absolute; left: 0; top: 0; }

.manifest-tag {
    position: absolute;
    top: -13px;
    left: 30px;
    background: var(--brass);
    color: var(--navy-deep);
    font-size: 9.5px;
    font-weight: 800;
    letter-spacing: 1.8px;
    text-transform: uppercase;
    padding: 4px 12px;
    border-radius: 2px;
}

.stamp-badge {
    position: absolute;
    top: 22px;
    right: 26px;
    width: 82px;
    height: 82px;
    border-radius: 50%;
    border: 2.5px double currentColor;
    display: flex;
    align-items: center;
    justify-content: center;
    text-align: center;
    font-family: 'Fraunces', serif;
    font-size: 10.5px;
    font-weight: 700;
    letter-spacing: 0.8px;
    text-transform: uppercase;
    line-height: 1.25;
    transform: rotate(-11deg);
    opacity: 0.92;
}

.stamp-verified { color: #6FE3A8; }
.stamp-flagged { color: #F0A0A0; }

.result-label {
    color: #8DA3B5;
    font-size: 10px;
    letter-spacing: 2px;
    font-weight: 800;
    text-transform: uppercase;
}

.result-number {
    font-family: 'IBM Plex Mono', monospace;
    color: #F7F3E8;
    font-size: 30px; font-weight: 700;
    letter-spacing: 2px;
    margin-top: 5px;
}

.result-divider { height: 1px; background: rgba(201,162,39,0.25); margin: 22px 60px 22px 0; }

.result-line { font-family: 'Fraunces', serif; color: #F7F3E8; font-size: 27px; font-weight: 600; margin-top: 4px; }

.copy-caption { font-size: 11px; color: var(--ink-soft); margin-top: 10px; margin-bottom: 16px; }


/* =====================================================
   METRİKLER
   ===================================================== */

div[data-testid="stMetric"] {
    background: var(--surface);
    border: 1px solid var(--surface-border);
    border-radius: 4px;
    padding: 16px;
    box-shadow: 0 4px 16px rgba(0,0,0,0.22);
}

div[data-testid="stMetricLabel"] {
    color: var(--ink-soft);
    font-size: 10px; font-weight: 800;
    letter-spacing: 1px;
    text-transform: uppercase;
}

div[data-testid="stMetricValue"] {
    font-family: 'Fraunces', serif;
    color: var(--parchment); font-size: 19px; font-weight: 600;
}


/* =====================================================
   KRİTİK UYARI
   ===================================================== */

@keyframes alertPulse {
    0% { box-shadow: 0 0 0 0 rgba(156,43,43,0.35); }
    70% { box-shadow: 0 0 0 13px rgba(156,43,43,0); }
    100% { box-shadow: 0 0 0 0 rgba(156,43,43,0); }
}

.danger-card {
    position: relative;
    overflow: visible;
    background: linear-gradient(155deg, #5C1414, var(--alert));
    border-radius: 4px;
    padding: 30px;
    margin-top: 24px;
    color: white;
    text-align: center;
    animation: alertPulse 2.2s infinite, cardReveal 0.35s ease-out;
    box-shadow: 0 18px 38px rgba(156,43,43,0.24);
    border-top: 3px dashed rgba(255,255,255,0.35);
}

.danger-symbol {
    width: 66px; height: 66px;
    margin: 0 auto 14px auto;
    border-radius: 50%;
    background: rgba(255,255,255,0.12);
    border: 2px solid rgba(255,255,255,0.4);
    display: flex; align-items: center; justify-content: center;
    font-family: 'Fraunces', serif;
    font-size: 36px; font-weight: 700;
}

.danger-title { font-family: 'Fraunces', serif; font-size: 27px; font-weight: 700; letter-spacing: 0.2px; }
.danger-container { font-family: 'IBM Plex Mono', monospace; font-size: 24px; font-weight: 700; margin-top: 15px; letter-spacing: 2px; }
.danger-info { color: #FBDADA; font-size: 12.5px; margin-top: 15px; letter-spacing: 0.6px; }
.danger-line { font-family: 'Fraunces', serif; color: white; font-size: 23px; font-weight: 700; margin-top: 4px; }

.danger-stop {
    margin-top: 22px;
    background: white;
    color: var(--alert);
    padding: 13px;
    border-radius: 3px;
    font-family: 'Inter', sans-serif;
    font-size: 15px; font-weight: 800;
    letter-spacing: 1.4px;
    text-transform: uppercase;
}


/* =====================================================
   BULUNAMADI
   ===================================================== */

.not-found {
    background: linear-gradient(155deg, #4A1010, #7A1E1E);
    border-radius: 4px;
    padding: 30px;
    text-align: center;
    color: white;
    margin-top: 24px;
    box-shadow: 0 16px 34px rgba(122,30,30,0.20);
    border-top: 3px dashed rgba(255,255,255,0.3);
    animation: cardReveal 0.35s ease-out;
}

.not-found-title { font-family: 'Fraunces', serif; font-size: 25px; font-weight: 700; }
.not-found-number { font-family: 'IBM Plex Mono', monospace; font-size: 22px; font-weight: 700; margin-top: 12px; letter-spacing: 2px; }

.not-found-stop {
    background: white;
    color: #7A1E1E;
    border-radius: 3px;
    padding: 12px;
    font-size: 15px; font-weight: 800;
    letter-spacing: 1.2px;
    text-transform: uppercase;
    margin-top: 20px;
}


/* =====================================================
   TOPLU DOĞRULAMA SATIRLARI
   ===================================================== */

.batch-row {
    display: flex;
    align-items: center;
    gap: 14px;
    padding: 13px 18px;
    border-radius: 3px;
    margin-bottom: 7px;
    font-weight: 700;
    background: var(--surface);
}

.batch-row-ok { border-left: 4px solid var(--verified); color: #6FE3A8; }
.batch-row-bad { border-left: 4px solid var(--alert); color: #F0A0A0; }
.batch-row-warn { border-left: 4px solid var(--brass); color: #E4C863; }

.batch-icon { font-size: 17px; min-width: 20px; text-align: center; }
.batch-number { font-family: 'IBM Plex Mono', monospace; font-size: 14.5px; letter-spacing: 1px; min-width: 150px; color: var(--parchment); }
.batch-detail { font-size: 12px; font-weight: 600; opacity: 0.85; flex: 1; }

.batch-summary { display: flex; gap: 10px; margin-bottom: 14px; }

.batch-pill {
    flex: 1;
    text-align: center;
    padding: 12px;
    border-radius: 3px;
    font-size: 12px;
    font-weight: 800;
    letter-spacing: 0.4px;
}


/* =====================================================
   FOOTER
   ===================================================== */

.app-footer {
    text-align: center;
    color: #4A6478;
    font-size: 10px;
    letter-spacing: 1.4px;
    text-transform: uppercase;
    margin-top: 44px;
}


/* =====================================================
   HAREKET AZALTMA TERCİHİ
   ===================================================== */

@media (prefers-reduced-motion: reduce) {
    .hero-content, .container-result, .danger-card, .not-found {
        animation: none !important;
    }
    .danger-card { animation: none !important; box-shadow: 0 18px 38px rgba(156,43,43,0.24) !important; }
}


/* =====================================================
   TELEFON
   ===================================================== */

@media (max-width: 650px) {
    .block-container { padding-left: 13px; padding-right: 13px; }
    .hero { min-height: 300px; padding: 28px 24px; }
    .hero-content { width: 100%; }
    .hero-title { font-size: 32px; }
    .hero-visual { width: 150px; height: 150px; top: auto; bottom: 0; right: 0; opacity: 0.35; }
    .result-number { font-size: 25px; }
    .result-line { font-size: 22px; }
    .danger-title { font-size: 23px; }
    .stamp-badge { width: 64px; height: 64px; font-size: 8.5px; top: 16px; right: 16px; }
    .batch-row { flex-wrap: wrap; }
}

</style>
"""

# Yüksek kontrast modu (erişilebilirlik) — palet dışına çıkan, saf siyah/beyaz mod
HIGH_CONTRAST_CSS = """
<style>

.stApp { background: #ffffff !important; background-image: none !important; }

.stat-card, div[data-testid="stVerticalBlockBorderWrapper"], div[data-testid="stMetric"] {
    border: 2px solid #000000 !important;
    box-shadow: none !important;
}

.stat-value, .result-label, div[data-testid="stMetricValue"], h1, h2, h3 { color: #000000 !important; }

div[data-testid="stTextInput"] input, div[data-testid="stTextArea"] textarea {
    border: 2px solid #000000 !important;
    color: #000000 !important;
    background: #ffffff !important;
}

div.stButton > button {
    background: #000000 !important;
    color: #ffffff !important;
    border: 2px solid #000000 !important;
    box-shadow: none !important;
}

.container-result, .hero { background: #000000 !important; background-image: none !important; }
.danger-card { background: #b30000 !important; animation: none !important; }
.not-found { background: #7a0000 !important; }
.success-banner { background: #005c33 !important; }

.batch-row-ok { background: #e6f7ee !important; border-left: 4px solid #0a7a4d !important; }
.batch-row-bad { background: #fbe6e6 !important; border-left: 4px solid #a91616 !important; }
.batch-row-warn { background: #fff6df !important; border-left: 4px solid #8a6300 !important; }

.stamp-badge { color: #ffffff !important; }

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
    Dönüş: (status, record_or_none, normalized) -- status: 'ok' | 'not_found' | 'duplicate' | 'invalid'"""

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
# ÜST ARAÇ ÇUBUĞU
# =========================================================

toolbar_col = st.columns([5, 2])[1]

with toolbar_col:
    st.toggle(
        "Yüksek Kontrast",
        key="high_contrast",
        help="Güneş ışığı altında veya düşük görüş koşullarında okunabilirliği artırır."
    )


# =========================================================
# HERO — pusula gülü ile "kaptan köşkü" paneli
# =========================================================

st.html("""
<div class="hero" role="banner" aria-label="ALPORT Banjul Konteyner Takip Sistemi">

    <div class="hero-content">

        <div class="hero-brand">ALPORT BANJUL &nbsp;·&nbsp; LİMAN OPERASYONLARI</div>

        <div class="hero-title">
            Konteyner<br>
            Kayıt Defteri
        </div>

        <div class="hero-subtitle">
            Gemi yüklemesinden önceki son kontrol noktası:
            konteyner numarası, shipping line ve saha kaydını
            tek ekranda doğrulayın.
        </div>

        <div class="hero-badge">⚓ &nbsp;OPERASYONEL DOĞRULAMA SİSTEMİ</div>

    </div>

    <svg class="hero-visual" viewBox="0 0 400 400" xmlns="http://www.w3.org/2000/svg" aria-hidden="true">
        <circle cx="230" cy="200" r="150" fill="none" stroke="#C9A227" stroke-width="1" opacity="0.35"/>
        <circle cx="230" cy="200" r="112" fill="none" stroke="#C9A227" stroke-width="1" opacity="0.28"/>
        <circle cx="230" cy="200" r="4" fill="#C9A227" opacity="0.85"/>

        <g stroke="#C9A227" stroke-width="1" opacity="0.55">
            <line x1="230" y1="30" x2="230" y2="370" />
            <line x1="60" y1="200" x2="400" y2="200" />
            <line x1="123" y1="93" x2="337" y2="307" opacity="0.35"/>
            <line x1="337" y1="93" x2="123" y2="307" opacity="0.35"/>
        </g>

        <path d="M230 48 L242 190 L230 210 L218 190 Z" fill="#C9A227" opacity="0.9"/>
        <path d="M230 352 L242 210 L230 190 L218 210 Z" fill="#E4D9B0" opacity="0.55"/>
        <path d="M78 200 L220 188 L240 200 L220 212 Z" fill="#E4D9B0" opacity="0.45"/>
        <path d="M382 200 L240 188 L220 200 L240 212 Z" fill="#E4D9B0" opacity="0.45"/>

        <text x="230" y="24" text-anchor="middle" fill="#C9A227" font-family="Fraunces, serif" font-size="15" opacity="0.8">K</text>
        <text x="230" y="392" text-anchor="middle" fill="#C9A227" font-family="Fraunces, serif" font-size="15" opacity="0.8">G</text>
        <text x="42" y="205" text-anchor="middle" fill="#C9A227" font-family="Fraunces, serif" font-size="15" opacity="0.8">B</text>
        <text x="418" y="205" text-anchor="middle" fill="#C9A227" font-family="Fraunces, serif" font-size="15" opacity="0.8">D</text>
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
    with st.spinner("Kayıt defteri yükleniyor..."):
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
# SEKMELER
# =========================================================

tab_single, tab_batch = st.tabs(["⚓ Tekli Arama", "📋 Toplu Doğrulama"])


# ---------------------------------------------------------
# TEKLİ ARAMA
# ---------------------------------------------------------

with tab_single:

    with st.container(border=True):

        st.subheader("Konteyner Doğrulama")
        st.caption("Yükleme hattını seçin ve konteyner numarasını girin.")

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

        if status == "invalid":
            st.html(f"""
                <div class="not-found" role="alert">
                    <div style="font-size:42px; margin-bottom:8px;">⚠</div>
                    <div class="not-found-title">GEÇERSİZ FORMAT</div>
                    <div class="not-found-number">{safe(normalized)}</div>
                    <div style="color:#ffdede; margin-top:12px; font-size:13px;">
                        Konteyner numarası standart formata (4 harf + 7 rakam) uymuyor.
                        Lütfen tekrar kontrol edin.
                    </div>
                </div>
            """)

        elif status == "not_found":
            st.html(f"""
                <div class="not-found" role="alert">
                    <div style="font-size:42px; margin-bottom:8px;">ⓧ</div>
                    <div class="not-found-title">KONTEYNER BULUNAMADI</div>
                    <div class="not-found-number">{safe(normalized)}</div>
                    <div style="color:#ffdede; margin-top:12px; font-size:13px;">
                        Bu konteyner güncel veritabanında bulunmuyor.
                    </div>
                    <div class="not-found-stop">YÜKLEME YAPMAYIN</div>
                </div>
            """)

        elif status == "duplicate":
            st.error("Aynı konteyner için birden fazla kayıt bulundu.")
            st.warning("Yükleme öncesinde Operasyon Departmanı ile teyit edin.")

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

            line_color = LINE_COLORS.get(shipping_line, "#B08D3E")

            wrong_line = (
                selected_line != "Hat seçilmedi"
                and selected_line != shipping_line
            )

            if wrong_line:
                st.html(f"""
                    <div class="danger-card" role="alert">
                        <div class="stamp-badge stamp-flagged">RED<br>FLAG</div>
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
                        <div class="manifest-tag">Manifesto Kaydı</div>
                        <div class="stamp-badge stamp-verified">DOĞRU-<br>LANDI</div>
                        <div class="container-accent" style="background:{line_color};"></div>
                        <div class="result-label">KONTEYNER NUMARASI</div>
                        <div class="result-number">{safe(container)}</div>
                        <div class="result-divider"></div>
                        <div class="result-label">SHIPPING LINE</div>
                        <div class="result-line" style="color:{line_color};">{safe(shipping_line)}</div>
                    </div>
                """)

                st.html('<div class="copy-caption">Numarayı kopyalamak için kutunun sağ üstündeki simgeye tıklayın</div>')
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

        ok_count = sum(1 for r in rows if r["status"] == "ok")
        bad_count = len(rows) - ok_count

        st.html(f"""
            <div class="batch-summary">
                <div class="batch-pill" style="background:rgba(53,169,124,0.14); color:#6FE3A8;">
                    ✓ {ok_count} Uygun
                </div>
                <div class="batch-pill" style="background:rgba(194,69,69,0.12); color:#F0A0A0;">
                    ✗ {bad_count} Sorunlu
                </div>
                <div class="batch-pill" style="background:rgba(201,162,39,0.10); color:#D9B84A;">
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
    ⚓ ALPORT BANJUL &nbsp;·&nbsp; KONTEYNER KAYIT DEFTERİ &nbsp;·&nbsp; OPERASYON
</div>
""")
