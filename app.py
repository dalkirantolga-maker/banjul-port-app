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

# ALPORT Banjul logosu (base64 gömülü — ayrı dosya taşımaya gerek yok)
LOGO_BASE64 = "/9j/4AAQSkZJRgABAQEAYABgAAD/2wBDAAMCAgMCAgMDAwMEAwMEBQgFBQQEBQoHBwYIDAoMDAsKCwsNDhIQDQ4RDgsLEBYQERMUFRUVDA8XGBYUGBIUFRT/2wBDAQMEBAUEBQkFBQkUDQsNFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBT/wAARCACnASEDASIAAhEBAxEB/8QAHwAAAQUBAQEBAQEAAAAAAAAAAAECAwQFBgcICQoL/8QAtRAAAgEDAwIEAwUFBAQAAAF9AQIDAAQRBRIhMUEGE1FhByJxFDKBkaEII0KxwRVS0fAkM2JyggkKFhcYGRolJicoKSo0NTY3ODk6Q0RFRkdISUpTVFVWV1hZWmNkZWZnaGlqc3R1dnd4eXqDhIWGh4iJipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXGx8jJytLT1NXW19jZ2uHi4+Tl5ufo6erx8vP09fb3+Pn6/8QAHwEAAwEBAQEBAQEBAQAAAAAAAAECAwQFBgcICQoL/8QAtREAAgECBAQDBAcFBAQAAQJ3AAECAxEEBSExBhJBUQdhcRMiMoEIFEKRobHBCSMzUvAVYnLRChYkNOEl8RcYGRomJygpKjU2Nzg5OkNERUZHSElKU1RVVldYWVpjZGVmZ2hpanN0dXZ3eHl6goOEhYaHiImKkpOUlZaXmJmaoqOkpaanqKmqsrO0tba3uLm6wsPExcbHyMnK0tPU1dbX2Nna4uPk5ebn6Onq8vP09fb3+Pn6/9oADAMBAAIRAxEAPwD9U6KKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKK+cPGX7ffwo8EeKtV8P311qk97ptw9rO9pZiSLzFOGCtuGcEEfhXs3wy+JOh/FvwXYeKPDs7z6Xe7thlXa6lWKsrLk4IINdNTDVqUFOpBpPqc1PFUas3TpzTa6HU0UV85eLf29vhb4L8Xap4b1F9YGpaddPZziKx3J5ittODu5GR6VNKhVrtqlFu3YqtXpUEnVklfufRtFMhlW4hjlTO11DDPoRT6wNwor5/8AiZ+3B8NfhP431Pwprz6suq6eyLN9nst8eWRXGG3DPDCuY/4eSfB3/nrrn/gvH/xdd8cBipxUo020/I4JY/Cwk4yqJNeZ9T0V5Z8I/wBpr4dfG2U23hfX45tSVS5066QwXG0dSEb7w/3Sa9TrkqU50pclRWfmddOpCrHnptNeQUV5f8Y/2k/APwLhQeKNZWO/kXfFptqvm3Lj12D7o92IFfPq/wDBUnwCbwRnwr4gFvux537jOPXbv/rXVSwOJrx56cG0clXHYahLkqVEmfaVFea/B79onwH8dLSR/CmtJcXcK7ptOuF8q5iHTJQ9R7rke9elVyzpzpScJqz8zrp1IVYqcHdeQUV82eKv+CgHwq8HeJtV0LUZNZF9ptzJaTiOx3LvRirYO7kZB5rLH/BST4OkgGXXAD3On9P/AB+uxZfi2rqm/uON5hhE7OqvvPqeiuB+FPx08EfGuwkufCOuwak8Khp7Ugx3EIPTfG2CB79Peqnxt+P/AIV+AGk6bqPiprxba/ma3hNnB5p3BdxyMjAxXMqFV1PZcr5u3U6XXpKn7XmXL36HpNFfK/8Aw8l+Dv8Az11z/wAF4/8Ai6P+Hkvwd/5665/4Lx/8XXV/Z2L/AOfT+45f7Rwf/P1fefVFFeZfBD9oXwn+0DY6rd+FGvWh02RIp/tkHlHLgkY5OfumvTa4qlOdKThNWaO2nUhVip03dMKKKKzNAooooAKKKKACiiigAooooAKKKKACiiigArz/AOPfxMh+EHwj8S+KndVnsrVhaqx+/cP8sQ9/mI/AGvQK/Pj/AIKhfFjzJ/Dnw8s5uIx/auoKrdzlYVP0G9vxWvQwGH+tYmFPp19EefmGJ+q4adTr09WfEem+FNe8a2PiTXrW3lvoNJiF7qVweSiyShAx9SWbP4E9q+3v+CXvxZw/iP4d3s3Uf2rp6s3cYSZB/wCQ2x/vGvP/ANlb4y/B34bfBXxX4b8ZXF8dW8UGSC+EFgZVjg2FIwGzzjczexNfP/wX+I5+D3xi0HxRZzPLa6dfDzGUFTNbklZBj3QtxX3eJhPG0q1Bwtb4X3/p/gfBYacMFVo14zTv8S7f0vxR+4tfif8AtE/8nKeOv+xguf8A0ca/amwvoNTsbe8tZFmtriNZopFOQ6MMgj6givxW/aJ/5OU8df8AYwXP/o418/w//Gqen6n0PEOtGn6/oftNpn/INtP+uKf+girVVdM/5Btp/wBcU/8AQRVqvlXufVrY/Hb9uzn9qjxt/wBdLb/0mir6V8Ef8E0fBnizwNoWtSeLNdt7nUdPgu2RUhKI8kYYgDbkgE+tfNX7dnH7VHjb/rpbf+k0VaGh/t9/GDw7odhpNjq2nx2djbx2sAOnREqiKFXJI5OAOa/RJU8XUwlH6rLldlf7l5M/OY1cJSxdf63HmV3b735o87+IvhPWf2bfjdqGj2eqk6t4evY5LbUbbKFsqskb4zwSrDIz6jmv1X8cfH5PB37MSfE6WFPtdxo9vdW9s33WuZ0URr9N7gn2Br8xPhp8HviF+1R8SZLsw3l69/dedqmv3SEQwgn5mZsYJA6IPYAAdPvT9u3wNLYfsm2+h6FbSzW2kXVhCsMKFmMSfulGByeWSuLMVTq1sPQqtOV1zfh+Z3Za6tKjiK9JNRs+X8fyR+f3wz+H3jD9rH4xPZG/e61XUHe81HVLsllgiBG6RvYZCqox1UDA6fb0n/BLnwCdCNunifXl1XZxet5Rj3+vlbfu+27PvWf/AME5fgV4z+Gl14k8Q+KNBk0a11W0hisxdkLcHDFjmP7yggj72OnSvuCuLM8zqwr+zw07Rjba39eR25ZllKpQ9piYXlK+9/687n4neNvCXjL9kz4z/Yxemz1zSZUubO/tSRHcRHlXA7qwBBU+jA1+vPwS+Jlv8YPhX4d8XQIsR1K2DTQqciOZSUkX6B1bHtivz2/4KdeJdL1f4z6NptmySX2maWsd6yEHazuzqh9wpBx/tCvrL/gnzpd1pn7MHh5rpWUXNzdXEIb/AJ5mZgPwJUn8a0zP/aMDSxNRWn/w/wDlczyv/Z8fWw1N3h/w3+dj81fjjaLqH7RPjO1diqTeIrqMsvUA3DCvtm+/4JZeEJdOm+w+MtbivGjJhe4iheMNjjcoUEj6EV8QfH2aW2+P3jqWAlZk1+7ZCo5DCdsV0mufta/G17ebTtS8baxbrLGVeNkSFypGOoUMMjuDXuVqWKq06X1aoo2Wv4eR4VGthaVSr9ZpuV3p+PmYXwg8Wax8E/jrpF5p9zi707VBZ3AhY7LiLzNkqH1Vhn9D2Ffq3+0X+zho37SGh6Tpes6rfaVFp1w1yj2IQs5ZduDuB4r4B/Yp/ZT1r4ueMNJ8a6vEsHgzTrsTtNI4L3s0bZESrnONwG4nHHAyTx+rdfP5ziYwxMHRfvxWrX9ep9DkuGlPDTjWj7knon/XoflT+2J+x94d/Zw8IaFq+ja3qeqTahfNavHfCMKqiMtkbVBzkVD+xz+yL4e/aR8OeItR1nWtS0qXTLqO3jSxEZVwyFiTuU88V9Bf8FTv+SYeDf8AsMP/AOiGqr/wSt/5ETx1/wBhK3/9FGu365X/ALL9vze9ff5nF9Tof2r7Dl9223yPoP8AZy/Zo0X9m3TNastG1a/1WPVJo5pGvggKFAwAG0D+8a9hoor4yrVnWm6lR3bPtKVKFGCp01ZIKKKKyNQooooAKKKKACiiigAooooAKKKKACiiigCtqOoW2kafdX15MtvaW0TTTTOcKiKCWY+wAJr8YvEd7qn7U/7SkrQl1m8SauILfIz5FvnavB/uRLk/Q1+z1/YW2q2U9neQR3VpOhjlhlUMkikYKkHqCO1c3oXwm8FeGNTi1HSPCejaZfxAhLq0sY45FyMHDAZGQSK9jL8dDA88uW8mrLyPHzDAzxzhHmtFO78z5M/4dX+Ef+h21r/wGhr5l/a//ZRT9my/8Py6ZqV1rOj6pHIv2m6jVWjnQjKfLxgqykfj6V+vNY/iXwfoXjOzjtNe0ex1m2jfzUhv7dZlV8EbgGBwcEjPvW+HznE06ilVlzR7aGGJybDVKTjSjyy6PU+ff+CfvxX/AOFi/Aey0q6m8zVfDT/2dKGPzGHGYW+m35f+AV+cn7ScjQ/tE/ECVBkprt0w9MiU1+y3hfwB4a8EvcN4f0DTdEa4CiY2FqkPmBc43bQM4yfzNZepfBjwFrF/cX194M0K8vLhzJNcT6fE7yMeSzErkk+tVhsypYbE1Kyg7S6diMTllXE4anRlNXj17n51w/8ABTX4nwQxxLo3hzaihRm1m7DH/PWvpf8AYp/aq8WftF614otPEllplpHplvBLCbCJ0LF2YHdudsj5RXuv/Chvhv8A9CH4d/8ABZD/APE1t+GPh94Y8FS3Enh/w/puiyXChZWsLVITIB0DbQM4yajE4zBVaUo0qPLJ9TTDYPG0qsZ1a/NFdD8lv26vm/ap8b45xJbA4/69oq/Tr4Q+AfC938KfBs8vhvR5ZZNHtHaRrCIliYVySdvJ966LW/hH4I8S6nNqOreEtF1LUJsGW6urCOSR8AAZYjJwAB+FdPaWkFhaw21tClvbwoI44o1CqigYAAHQAVni8wWIoUqMU04L79EjTCZe8PXq1pNNTf3a3FtrWGygSG3hjghQYWOJQqqPYDpXJ/E7x5eeAdCF7YeF9V8VXTkhLXS4g20ju5/hHuAT7V2NFeHNOSaTsz6LDzp0qsZ1Yc8VvG7V/K61+4+Kfhl8YPiV+0J8d7XQ9T87wroOjyNeahpVmrQsBGRtjlc4ZiX2gg4GM8V5/wDtT/t4+NvDfxQ8R+E/At7Z2Wkac32I3wt1lmaZRiUqzZAwxKjj+DNfcPg34W6d4N8Z+MPEsD+bfeI7mOaXKAeUqRhdgPfLbmP+97VZl+EngefUHvpPB2gyXjsWad9MhLsxOSSdvJz3rXKHDBOU8XH2jd9/w/D8z0uL8XSzipShlMVQpQjFWStra8td2+Z2u9Wkj8o/gN+zN45/ab8a/wBp363kGhzXHn6n4hvgx8zJywjLf6yQ+2QOpr9dfDPhzT/CHh7TdD0q3W103T7dLa3hXoiKAAP0q0JrPT0SAPBbKigLFkIFHbA9KkhuYbgExSpKB12MDj8q9HHZhPHSV1aK2R8lgMtjgYtrVvdn4pfGH/k5XxYe3/CTXHP/AG8tX3x+33+zePih4ATxpodpv8T6BblpkiX57u0HLLx1ZOWHtuHcV9CXfwi+H2oavLdXPhLw/canLIZ5JZLGFpmcnJYkjJOec12Dz24fyGkj3sMeUWGSPpXbXzbnnRnSVnD8dv8AI4aGT8kK1Oq7qeu22+v4n5W/sD/tHn4R+Pf+EV1y5MfhXX5VQvIcJaXXRJOeitwrH/dPav1XBBGRyK4ST4H/AA1a4w/gjw55z5bB06HcfU/drtw0NokUIKRLgJGmQOnQAVxZhiqWLq+1px5X1O3LsJWwdL2VSXMuh8Vf8FTRn4YeDfX+2H/9ENVX/glbx4F8dDv/AGjb/wDoo19keIvCfh3x1aRQ63pWna7bQuXjS8gSdEfGCQGBAOKj8LeFvDHg2GWDw9pemaPFcvl49PhjiErAY5CgZIrT6/H6j9U5db7/ADuZ/wBnz+v/AFu+lrW+VjoKKKK8c9gKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigD4Z/anj0yX9pmEazJo0Vh/YEXza99p+zb/MfH+o+fd1x265rL+A/jrQ/hd8UPGesw3FjPoVvoBuNvh5rj7E0odAseJ/m8xiQB2yx96+8bjTbS7k3z2sMz4xukjDHH1Ipo0mxWN4xZ24jfG5BEuDjpkYrx3gJe1dWMle99v+CfqFPjGj/ZscurUJSjyKD9+yaTTulyOz07n52W9/4x0y4g+LL+GvFP8Awlw1RtVub57Vhpzacyf6oHOQNvGSMbT7ZruPGHhd/wBoX44avrfhHU57HU4vDtnrOjTqxQecu392x/hPJGexH1r7kMaGMxlVKEbduOMemKht7C1tW3QW0ULY25jQKcenFJZdpyyndXu/Xua1OOm5/WKWGUKqi4Rad0oNxai01qly2SulZtHwzo/7Qt1qvxh0HxRrehXw1zw1oF3baxp0Ee1nlQtuZAeACDuIPTB9q7Dx78ZvDXxu+JXwZj8KXU19e2+qC7vLQQuHtVOwkOcYyMNnBIwDX1sLC1WdphbRCZs5kCDcc9eaZa6TY2MrSW1lb28jfeeKJVJ+pAq1g6qi4updNpvTXS3n5HFV4ny6VWGIpYNwnCEoRtU91KSmndOOqTm2rNPo2z5D/Zw+OXgz4VeBPEvh7xdrJ0XWU1m7l+y3FvKXKsqgEYU91YfWvLvA2uweDLT4Q+Idbkm0/R18Q6leNcSxuf3P7ob8AEkZBHHvX6FXOh6deymW40+1nlPV5YVZvzIqaXTrW4RFltoZFQYRXjBC/T0qXgajjGLmvd20809dfLyOmPF+DhWr1o4WV67vP31/JOHue5p8bevNtY5L4efGXwf8VnvU8LawuqtZhTOFhkTYGzj76jPQ9K7WoLaxtrMsYLeKAt18tAufyqevVhzJe+7vy/pn5viZUJVW8NFxh0UmpP70orfyCiiirOYKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigD//2Q=="


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
    "MAERSK": "#1E6FA6",
    "CMA CGM": "#C1501F",
    "MSC": "#A6821E",
    "HAPAG-LLOYD": "#C1501F",
    "ONE": "#9C2B5E",
    "COSCO": "#1E4E8C",
    "PIL": "#B0281E",
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
# CSS — kurumsal, açık zeminli operasyon paneli
# =========================================================

BASE_CSS = """
<style>

@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&family=IBM+Plex+Mono:wght@500;600;700&display=swap');

:root {
    --navy: #0F2A44;
    --navy-deep: #0A1E33;
    --steel: #2B5A7D;
    --gold: #B8860B;
    --gold-soft: #D4A72C;
    --bg: #F4F6F8;
    --surface: #FFFFFF;
    --border: #E1E6EB;
    --ink: #1F2937;
    --ink-soft: #6B7280;
    --alert: #B91C1C;
    --alert-bg: #FDECEC;
    --verified: #157347;
    --verified-bg: #E9F7EF;
    --warn: #92660A;
    --warn-bg: #FDF3DC;
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
    font-family: 'Inter', sans-serif !important;
    font-weight: 700 !important;
    color: var(--navy) !important;
    letter-spacing: -0.2px;
}

code, .stCode, div[data-testid="stTextInput"] input, div[data-testid="stTextArea"] textarea {
    font-family: 'IBM Plex Mono', monospace !important;
}


/* =====================================================
   SAYFA ZEMİNİ — düz, kurumsal
   ===================================================== */

.stApp { background: var(--bg); }

.block-container {
    max-width: 960px;
    padding-top: 1.4rem;
    padding-bottom: 3rem;
}


/* =====================================================
   ÜST ARAÇ ÇUBUĞU
   ===================================================== */

.top-toolbar { display: flex; justify-content: flex-end; margin-bottom: 6px; }


/* =====================================================
   HERO
   ===================================================== */

.hero {
    position: relative;
    overflow: hidden;
    background: linear-gradient(135deg, var(--navy-deep) 0%, var(--navy) 60%, #123A5C 100%);
    border-radius: 10px;
    padding: 38px 42px;
    margin-bottom: 24px;
    box-shadow: 0 10px 30px rgba(15,42,68,0.16);
}

.hero-content { position: relative; z-index: 3; width: 62%; }

.hero-logo-badge {
    display: inline-flex;
    align-items: center;
    background: #FFFFFF;
    border-radius: 8px;
    padding: 12px 20px;
    margin-bottom: 20px;
    box-shadow: 0 4px 14px rgba(0,0,0,0.20);
}

.hero-logo-badge img {
    height: 48px;
    display: block;
}

.hero-brand {
    color: var(--gold-soft);
    font-size: 11.5px;
    font-weight: 800;
    letter-spacing: 3px;
    margin-bottom: 10px;
    text-transform: uppercase;
}

.hero-title {
    color: #FFFFFF;
    font-size: 34px;
    font-weight: 800;
    line-height: 1.18;
    letter-spacing: -0.5px;
}

.hero-rule {
    width: 46px;
    height: 3px;
    background: var(--gold-soft);
    border-radius: 2px;
    margin: 16px 0;
}

.hero-subtitle {
    color: #C9D6E2;
    font-size: 14px;
    max-width: 430px;
    line-height: 1.65;
    font-weight: 400;
}

.hero-badge {
    display: inline-flex;
    align-items: center;
    gap: 8px;
    color: var(--navy-deep);
    background: var(--gold-soft);
    border-radius: 4px;
    padding: 7px 14px;
    margin-top: 18px;
    font-size: 10.5px;
    font-weight: 800;
    letter-spacing: 1px;
}

.hero-visual {
    position: absolute;
    right: 22px;
    top: 50%;
    transform: translateY(-50%);
    width: 150px;
    height: 150px;
    opacity: 0.9;
}


/* =====================================================
   ÜST DURUM KARTLARI
   ===================================================== */

.stat-card {
    position: relative;
    overflow: hidden;
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: 8px;
    padding: 18px 20px;
    min-height: 92px;
    box-shadow: 0 2px 8px rgba(15,42,68,0.04);
}

.stat-accent-blue, .stat-accent-green {
    position: absolute; top: 0; left: 0;
    width: 4px; height: 100%;
    background: var(--navy);
}
.stat-accent-green { background: var(--gold); }

.stat-icon { font-size: 17px; margin-bottom: 5px; color: var(--gold); }

.stat-label {
    color: var(--ink-soft);
    font-size: 10px; font-weight: 700;
    text-transform: uppercase;
    letter-spacing: 1.2px;
}

.stat-value {
    color: var(--navy);
    font-size: 21px; font-weight: 800;
    margin-top: 4px;
}


/* =====================================================
   ARAMA PANELİ
   ===================================================== */

div[data-testid="stVerticalBlockBorderWrapper"] {
    border-radius: 10px !important;
    background: var(--surface) !important;
    border: 1px solid var(--border) !important;
    box-shadow: 0 2px 10px rgba(15,42,68,0.05);
}

div[data-testid="stTextInput"] input {
    min-height: 58px;
    background: #FAFBFC;
    border-radius: 6px;
    border: 1.5px solid var(--border);
    text-align: center;
    font-size: 21px; font-weight: 600;
    letter-spacing: 1.5px;
    text-transform: uppercase;
    color: var(--navy);
}

div[data-testid="stTextInput"] input:focus {
    border-color: var(--navy);
    box-shadow: 0 0 0 3px rgba(15,42,68,0.10);
}

div[data-testid="stTextArea"] textarea {
    background: #FAFBFC;
    border-radius: 6px;
    border: 1.5px solid var(--border);
    font-size: 14.5px; font-weight: 600;
    letter-spacing: 0.6px;
    text-transform: uppercase;
    color: var(--navy);
}

div[data-baseweb="select"] > div {
    min-height: 52px;
    border-radius: 6px !important;
    background-color: #FAFBFC;
    border-color: var(--border) !important;
}

div.stButton > button {
    min-height: 52px;
    width: 100%;
    border-radius: 6px;
    border: 1px solid var(--navy);
    background: var(--navy);
    color: #FFFFFF;
    font-family: 'Inter', sans-serif;
    font-size: 14px; font-weight: 700;
    letter-spacing: 1px;
    text-transform: uppercase;
    box-shadow: 0 3px 10px rgba(15,42,68,0.14);
    transition: all 0.15s ease;
}

div.stButton > button:hover {
    background: var(--navy-deep);
    transform: translateY(-1px);
}

/* İkincil butonlar (arama geçmişi chip'leri) */
div.stButton > button[kind="secondary"] {
    min-height: 30px;
    width: auto;
    padding: 4px 12px;
    font-size: 11.5px;
    font-weight: 600;
    letter-spacing: 0.3px;
    text-transform: none;
    border-radius: 30px;
    background: #F0F3F6;
    color: var(--navy);
    box-shadow: none;
    border: 1px solid var(--border);
}

div.stButton > button[kind="secondary"]:hover {
    background: #E4EAF0;
    border-color: var(--navy);
    transform: none;
}


/* =====================================================
   FORMAT İPUCU
   ===================================================== */

.format-hint {
    text-align: center;
    font-family: 'IBM Plex Mono', monospace;
    font-size: 11.5px;
    font-weight: 600;
    letter-spacing: 0.3px;
    margin-top: -6px;
    margin-bottom: 12px;
    padding: 7px;
    border-radius: 5px;
}

.format-ok { color: var(--verified); background: var(--verified-bg); }
.format-bad { color: var(--alert); background: var(--alert-bg); }
.format-empty { color: var(--ink-soft); background: transparent; }


/* =====================================================
   ARAMA GEÇMİŞİ
   ===================================================== */

.history-label {
    color: var(--ink-soft);
    font-size: 10px; font-weight: 700;
    text-transform: uppercase;
    letter-spacing: 1.2px;
    margin-bottom: 8px;
}


/* =====================================================
   DOĞRULANDI BANNER
   ===================================================== */

.success-banner {
    background: var(--verified-bg);
    border: 1px solid #BFE6CC;
    border-radius: 8px;
    padding: 15px 20px;
    margin-top: 22px;
}

.success-title { color: var(--verified); font-size: 16px; font-weight: 800; }
.success-subtitle { color: #3D7856; font-size: 12px; margin-top: 2px; }


/* =====================================================
   KONTEYNER SONUÇ KARTI
   ===================================================== */

.container-result {
    position: relative;
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: 10px;
    padding: 26px 28px;
    margin-top: 16px;
    margin-bottom: 4px;
    box-shadow: 0 4px 16px rgba(15,42,68,0.06);
}

.container-accent { width: 5px; height: 100%; position: absolute; left: 0; top: 0; border-radius: 10px 0 0 10px; }

.status-pill {
    position: absolute;
    top: 22px;
    right: 24px;
    display: inline-flex;
    align-items: center;
    gap: 6px;
    padding: 6px 14px;
    border-radius: 30px;
    font-size: 11px;
    font-weight: 800;
    letter-spacing: 0.6px;
    text-transform: uppercase;
}

.status-pill-ok { background: var(--verified-bg); color: var(--verified); border: 1px solid #BFE6CC; }
.status-pill-bad { background: var(--alert-bg); color: var(--alert); border: 1px solid #F3C6C6; }

.result-label {
    color: var(--ink-soft);
    font-size: 10px;
    letter-spacing: 1.6px;
    font-weight: 700;
    text-transform: uppercase;
}

.result-number {
    font-family: 'IBM Plex Mono', monospace;
    color: var(--navy);
    font-size: 28px; font-weight: 700;
    letter-spacing: 1.5px;
    margin-top: 4px;
}

.result-divider { height: 1px; background: var(--border); margin: 20px 100px 20px 0; }

.result-line { color: var(--navy); font-size: 23px; font-weight: 800; margin-top: 4px; }

.copy-caption { font-size: 11px; color: var(--ink-soft); margin-top: 10px; margin-bottom: 14px; }


/* =====================================================
   METRİKLER
   ===================================================== */

div[data-testid="stMetric"] {
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: 8px;
    padding: 15px;
    box-shadow: 0 2px 8px rgba(15,42,68,0.04);
}

div[data-testid="stMetricLabel"] {
    color: var(--ink-soft);
    font-size: 10px; font-weight: 700;
    letter-spacing: 0.8px;
    text-transform: uppercase;
}

div[data-testid="stMetricValue"] {
    color: var(--navy); font-size: 18px; font-weight: 700;
}


/* =====================================================
   KRİTİK UYARI
   ===================================================== */

.danger-card {
    position: relative;
    background: var(--alert-bg);
    border: 1.5px solid #EAA3A3;
    border-radius: 10px;
    padding: 26px;
    margin-top: 22px;
    color: var(--alert);
    text-align: center;
    box-shadow: 0 4px 16px rgba(185,28,28,0.08);
}

.danger-symbol {
    width: 58px; height: 58px;
    margin: 0 auto 12px auto;
    border-radius: 50%;
    background: #FFFFFF;
    border: 2px solid var(--alert);
    display: flex; align-items: center; justify-content: center;
    font-size: 30px; font-weight: 800;
    color: var(--alert);
}

.danger-title { font-size: 22px; font-weight: 800; letter-spacing: 0.2px; color: var(--alert); }
.danger-container { font-family: 'IBM Plex Mono', monospace; font-size: 20px; font-weight: 700; margin-top: 14px; letter-spacing: 1.5px; color: var(--navy); }
.danger-info { color: #8A3E3E; font-size: 12px; margin-top: 14px; letter-spacing: 0.4px; }
.danger-line { color: var(--alert); font-size: 19px; font-weight: 800; margin-top: 4px; }

.danger-stop {
    margin-top: 20px;
    background: var(--alert);
    color: #FFFFFF;
    padding: 12px;
    border-radius: 6px;
    font-size: 14px; font-weight: 800;
    letter-spacing: 1.2px;
    text-transform: uppercase;
}


/* =====================================================
   BULUNAMADI
   ===================================================== */

.not-found {
    background: var(--alert-bg);
    border: 1.5px solid #EAA3A3;
    border-radius: 10px;
    padding: 26px;
    text-align: center;
    color: var(--alert);
    margin-top: 22px;
    box-shadow: 0 4px 16px rgba(185,28,28,0.08);
}

.not-found-title { font-size: 21px; font-weight: 800; color: var(--alert); }
.not-found-number { font-family: 'IBM Plex Mono', monospace; font-size: 19px; font-weight: 700; margin-top: 10px; letter-spacing: 1.5px; color: var(--navy); }

.not-found-stop {
    background: var(--alert);
    color: #FFFFFF;
    border-radius: 6px;
    padding: 11px;
    font-size: 13.5px; font-weight: 800;
    letter-spacing: 1px;
    text-transform: uppercase;
    margin-top: 18px;
}


/* =====================================================
   TOPLU DOĞRULAMA SATIRLARI
   ===================================================== */

.batch-row {
    display: flex;
    align-items: center;
    gap: 14px;
    padding: 12px 16px;
    border-radius: 6px;
    margin-bottom: 6px;
    font-weight: 700;
    background: var(--surface);
    border: 1px solid var(--border);
}

.batch-row-ok { border-left: 4px solid var(--verified); color: var(--verified); }
.batch-row-bad { border-left: 4px solid var(--alert); color: var(--alert); }
.batch-row-warn { border-left: 4px solid var(--gold); color: var(--warn); }

.batch-icon { font-size: 16px; min-width: 18px; text-align: center; }
.batch-number { font-family: 'IBM Plex Mono', monospace; font-size: 14px; letter-spacing: 0.6px; min-width: 148px; color: var(--navy); }
.batch-detail { font-size: 12px; font-weight: 600; opacity: 0.85; flex: 1; color: var(--ink-soft); }

.batch-summary { display: flex; gap: 10px; margin-bottom: 14px; }

.batch-pill {
    flex: 1;
    text-align: center;
    padding: 11px;
    border-radius: 6px;
    font-size: 12px;
    font-weight: 800;
    letter-spacing: 0.3px;
}


/* =====================================================
   FOOTER
   ===================================================== */

.app-footer {
    text-align: center;
    color: #9AA6B0;
    font-size: 10px;
    letter-spacing: 1.2px;
    text-transform: uppercase;
    margin-top: 40px;
}


/* =====================================================
   TELEFON
   ===================================================== */

@media (max-width: 650px) {
    .block-container { padding-left: 13px; padding-right: 13px; }
    .hero { padding: 26px 24px; }
    .hero-content { width: 100%; }
    .hero-title { font-size: 27px; }
    .hero-visual { display: none; }
    .result-number { font-size: 23px; }
    .result-line { font-size: 19px; }
    .danger-title { font-size: 19px; }
    .status-pill { position: static; display: inline-flex; margin-bottom: 12px; }
    .batch-row { flex-wrap: wrap; }
}

</style>
"""

# Yüksek kontrast modu (erişilebilirlik)
HIGH_CONTRAST_CSS = """
<style>

.stApp { background: #ffffff !important; }

.stat-card, div[data-testid="stVerticalBlockBorderWrapper"], div[data-testid="stMetric"],
.container-result, .batch-row {
    border: 2px solid #000000 !important;
    box-shadow: none !important;
}

.stat-value, .result-label, .result-number, .result-line, div[data-testid="stMetricValue"], h1, h2, h3 {
    color: #000000 !important;
}

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

.hero { background: #000000 !important; }

.danger-card, .not-found { background: #ffffff !important; border: 2px solid #B91C1C !important; }
.danger-title, .not-found-title, .danger-line { color: #B91C1C !important; }

.success-banner { background: #ffffff !important; border: 2px solid #157347 !important; }
.success-title { color: #157347 !important; }

.batch-row-ok { background: #ffffff !important; border: 2px solid #157347 !important; }
.batch-row-bad { background: #ffffff !important; border: 2px solid #B91C1C !important; }
.batch-row-warn { background: #ffffff !important; border: 2px solid #92660A !important; }

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
# HERO
# =========================================================

st.html(f"""
<div class="hero" role="banner" aria-label="ALPORT Banjul Konteyner Takip Sistemi">

    <div class="hero-content">

        <div class="hero-logo-badge">
            <img src="data:image/jpeg;base64,{LOGO_BASE64}" alt="ALPORT Banjul logosu" />
        </div>

        <div class="hero-brand">LİMAN OPERASYONLARI</div>

        <div class="hero-title">Konteyner Takip ve Doğrulama Sistemi</div>

        <div class="hero-rule"></div>

        <div class="hero-subtitle">
            Gemi yüklemesinden önceki son kontrol noktası:
            konteyner numarası, shipping line ve saha kaydını
            tek ekranda doğrulayın.
        </div>

        <div class="hero-badge">⚓ &nbsp;OPERASYONEL DOĞRULAMA</div>

    </div>

    <svg class="hero-visual" viewBox="0 0 200 200" xmlns="http://www.w3.org/2000/svg" aria-hidden="true">
        <circle cx="100" cy="100" r="88" fill="none" stroke="#D4A72C" stroke-width="1.2" opacity="0.35"/>
        <path d="M100 40 C100 40 130 70 130 105 C130 128 117 145 100 145 C83 145 70 128 70 105 C70 70 100 40 100 40 Z"
              fill="none" stroke="#D4A72C" stroke-width="2.2" opacity="0.75"/>
        <line x1="100" y1="40" x2="100" y2="18" stroke="#D4A72C" stroke-width="2.2" opacity="0.75"/>
        <line x1="86" y1="24" x2="114" y2="24" stroke="#D4A72C" stroke-width="2.2" opacity="0.75"/>
        <line x1="70" y1="105" x2="45" y2="118" stroke="#D4A72C" stroke-width="2" opacity="0.6"/>
        <line x1="130" y1="105" x2="155" y2="118" stroke="#D4A72C" stroke-width="2" opacity="0.6"/>
        <circle cx="100" cy="98" r="10" fill="none" stroke="#D4A72C" stroke-width="2" opacity="0.75"/>
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
                    <div style="font-size:38px; margin-bottom:8px;">⚠</div>
                    <div class="not-found-title">GEÇERSİZ FORMAT</div>
                    <div class="not-found-number">{safe(normalized)}</div>
                    <div style="color:#8A3E3E; margin-top:10px; font-size:13px;">
                        Konteyner numarası standart formata (4 harf + 7 rakam) uymuyor.
                        Lütfen tekrar kontrol edin.
                    </div>
                </div>
            """)

        elif status == "not_found":
            st.html(f"""
                <div class="not-found" role="alert">
                    <div style="font-size:38px; margin-bottom:8px;">ⓧ</div>
                    <div class="not-found-title">KONTEYNER BULUNAMADI</div>
                    <div class="not-found-number">{safe(normalized)}</div>
                    <div style="color:#8A3E3E; margin-top:10px; font-size:13px;">
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

            line_color = LINE_COLORS.get(shipping_line, "#A6821E")

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
                        <div class="status-pill status-pill-ok">✓ Doğrulandı</div>
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
                <div class="batch-pill" style="background:#E9F7EF; color:#157347;">
                    ✓ {ok_count} Uygun
                </div>
                <div class="batch-pill" style="background:#FDECEC; color:#B91C1C;">
                    ✗ {bad_count} Sorunlu
                </div>
                <div class="batch-pill" style="background:#EEF2F6; color:#0F2A44;">
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
    ⚓ ALPORT BANJUL &nbsp;·&nbsp; KONTEYNER TAKİP SİSTEMİ &nbsp;·&nbsp; OPERASYON
</div>
""")
