# ============================================================
# AYÇA Business Market V1.2 Stable Executive Dashboard
# Market / Perakende İşletme Yönetim Paneli
# ------------------------------------------------------------
# Beklenen Excel sayfası: AYCA_BUSINESS_DATA
# Çalıştırma:
#   pip install streamlit pandas numpy openpyxl plotly
#   streamlit run app.py
# ============================================================

from __future__ import annotations

import re
from datetime import datetime
from io import BytesIO
from html import escape
from typing import Any

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

st.set_page_config(
    page_title="AYÇA Business Market",
    page_icon="🛒",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ============================================================
# CSS / TEMA - V1.1
# Bu sürümde büyük rakamlarda kart/buton taşması düzeltildi.
# ============================================================
st.markdown("""
<style>
:root{
  --bg:#F8FAFC;--panel:#FFFFFF;--border:#E2E8F0;--text:#0F172A;--muted:#64748B;
  --blue:#2563EB;--green:#10B981;--orange:#F59E0B;--red:#EF4444;--purple:#8B5CF6;--cyan:#06B6D4;
}
html,body,[data-testid="stAppViewContainer"]{background:var(--bg);color:var(--text)}
[data-testid="stHeader"]{background:rgba(248,250,252,.86);backdrop-filter:blur(10px)}
[data-testid="stSidebar"]{background:linear-gradient(180deg,#FFFFFF 0%,#F1F5F9 100%);border-right:1px solid var(--border)}
.block-container{padding-top:1.2rem;max-width:1580px}
.ayca-header{display:flex;justify-content:space-between;align-items:center;gap:16px;margin-bottom:16px;flex-wrap:wrap}
.ayca-title h1{margin:0;font-size:clamp(24px,2.2vw,32px);font-weight:950;letter-spacing:-.7px}
.ayca-title p{margin:6px 0 0;color:var(--muted);font-size:14px}
.header-pill{background:#fff;border:1px solid var(--border);border-radius:16px;padding:12px 16px;box-shadow:0 8px 24px rgba(15,23,42,.05);font-weight:900;white-space:nowrap}
.metric-card,.mini-card,.soft-panel,.exec-card,.task-card{background:#fff;border:1px solid var(--border);border-radius:20px;padding:16px;box-shadow:0 12px 30px rgba(15,23,42,.055);box-sizing:border-box;min-width:0}
.metric-card{min-height:142px;background:linear-gradient(180deg,#fff 0%,#F8FBFF 100%);position:relative;overflow:hidden;display:flex;flex-direction:column;justify-content:space-between;gap:6px}
.metric-label{color:var(--muted);font-size:11px;font-weight:900;letter-spacing:.35px;text-transform:uppercase;margin-bottom:4px;line-height:1.25;white-space:normal;overflow-wrap:anywhere}
.metric-value{font-size:clamp(18px,1.65vw,25px);font-weight:950;margin-bottom:4px;letter-spacing:-.55px;line-height:1.08;white-space:normal;overflow-wrap:anywhere;word-break:break-word;font-variant-numeric:tabular-nums}
.metric-sub{color:var(--muted);font-size:12px;line-height:1.25;white-space:normal;overflow-wrap:anywhere;word-break:break-word}
.metric-up{color:var(--green);font-size:12px;font-weight:900;white-space:nowrap}.metric-down{color:var(--red);font-size:12px;font-weight:900;white-space:nowrap}
.section-title{font-size:21px;font-weight:950;margin:20px 0 12px;letter-spacing:-.3px}
.exec-grid{display:grid;grid-template-columns:1.25fr .75fr;gap:16px;margin:14px 0 18px}.exec-card{background:linear-gradient(135deg,#FFFFFF 0%,#EFF6FF 100%);border-color:#BFDBFE}
.exec-title{font-size:22px;font-weight:950;letter-spacing:-.4px;margin-bottom:8px}.exec-sub{color:var(--muted);font-size:14px;line-height:1.55;margin-bottom:14px}.exec-list-item{background:rgba(255,255,255,.82);border:1px solid rgba(226,232,240,.95);border-radius:16px;padding:12px 13px;margin:9px 0;font-size:14px;line-height:1.45;font-weight:760;overflow-wrap:anywhere}
.score-big{font-size:clamp(38px,4vw,54px);font-weight:950;letter-spacing:-2px;color:var(--blue);line-height:1;margin:4px 0 8px}.score-label{color:var(--muted);font-size:13px;font-weight:850;text-transform:uppercase;letter-spacing:.35px}
.health-row{margin:12px 0}.health-head{display:flex;justify-content:space-between;color:#334155;font-weight:900;font-size:13px;margin-bottom:6px;gap:8px}.health-bar-bg{height:10px;background:#E2E8F0;border-radius:999px;overflow:hidden}.health-bar-fill{height:10px;border-radius:999px;background:linear-gradient(90deg,#2563EB,#10B981)}
.radar-grid{display:grid;grid-template-columns:repeat(6,minmax(0,1fr));gap:12px;margin:12px 0 18px}.radar-card{background:#fff;border:1px solid #E2E8F0;border-radius:18px;padding:15px;box-shadow:0 10px 26px rgba(15,23,42,.05);min-height:116px;min-width:0;overflow:hidden}.radar-title{color:#64748B;font-size:11px;font-weight:900;text-transform:uppercase;letter-spacing:.35px;margin-bottom:9px;line-height:1.25}.radar-value{font-size:clamp(16px,1.35vw,21px);font-weight:950;margin-bottom:5px;line-height:1.1;overflow-wrap:anywhere;word-break:break-word;font-variant-numeric:tabular-nums}.radar-note{color:#64748B;font-size:12px;line-height:1.35;overflow-wrap:anywhere}
.task-grid{display:grid;grid-template-columns:1fr 1fr;gap:16px;margin:14px 0 18px}.task-item{border-bottom:1px solid #E2E8F0;padding:10px 0;font-size:14px;font-weight:760;line-height:1.45;overflow-wrap:anywhere}.task-item:last-child{border-bottom:0}
.chip{display:inline-block;border-radius:999px;padding:5px 10px;font-size:12px;font-weight:900}.chip-red{background:#FEE2E2;color:#B91C1C}.chip-orange{background:#FEF3C7;color:#B45309}.chip-green{background:#DCFCE7;color:#047857}.chip-blue{background:#DBEAFE;color:#1D4ED8}.chip-purple{background:#EDE9FE;color:#6D28D9}.small-muted{color:#64748B;font-size:13px}
.stTabs [data-baseweb="tab"]{background:#fff;border:1px solid var(--border);border-radius:13px;padding:10px 14px;font-weight:800;white-space:normal;min-height:42px}.stTabs [aria-selected="true"]{background:#DBEAFE;color:#2563EB;border-color:#BFDBFE}div[data-testid="stDataFrame"]{border-radius:16px;overflow:hidden;border:1px solid var(--border)}
.stButton>button,.stDownloadButton>button{border-radius:13px;border:1px solid #BFDBFE;background:#fff;color:#2563EB;font-weight:900;white-space:normal;min-height:44px;line-height:1.2;overflow-wrap:anywhere}.stButton>button:hover,.stDownloadButton>button:hover{border-color:#2563EB;background:#DBEAFE;color:#2563EB}
@media(max-width:1300px){.radar-grid{grid-template-columns:repeat(3,minmax(0,1fr))}.metric-value{font-size:20px}}
@media(max-width:1100px){.exec-grid,.task-grid{grid-template-columns:1fr}.radar-grid{grid-template-columns:1fr 1fr}}
@media(max-width:760px){.radar-grid{grid-template-columns:1fr}.metric-card{min-height:120px}.header-pill{white-space:normal}.stTabs [data-baseweb="tab"]{font-size:12px;padding:8px 10px}}
</style>
""", unsafe_allow_html=True)

# ============================================================
# DEMO AUTH
# ============================================================
DEMO_USERS = {
    "basic": {"password": "basic2026", "name": "Basic Demo", "store": "AYÇA Market", "membership": "Basic"},
    "premium": {"password": "premium2026", "name": "Premium Demo", "store": "AYÇA Market", "membership": "Premium"},
}


def safe_rerun():
    try:
        st.rerun()
    except Exception:
        st.experimental_rerun()


def is_premium() -> bool:
    return st.session_state.get("membership", "Basic").lower() == "premium"


def show_auth():
    st.markdown(
        """
        <style>
        [data-testid="stSidebar"]{display:none}
        .login-card{background:#fff;border:1px solid #E2E8F0;border-radius:28px;padding:30px;box-shadow:0 18px 45px rgba(15,23,42,.08)}
        .login-hero{background:linear-gradient(135deg,#fff 0%,#EFF6FF 55%,#DCFCE7 100%);border:1px solid #BFDBFE;border-radius:28px;padding:34px;box-shadow:0 18px 45px rgba(15,23,42,.08);min-height:420px}
        .login-title{font-size:38px;line-height:1.08;font-weight:950;letter-spacing:-1px;color:#0F172A;margin-bottom:12px}
        .login-sub{color:#64748B;font-size:15px;line-height:1.6;margin-bottom:24px}
        .feature-row{background:rgba(255,255,255,.75);border:1px solid rgba(226,232,240,.8);border-radius:16px;padding:11px 13px;margin:10px 0;font-weight:750;font-size:14px}
        </style>
        """,
        unsafe_allow_html=True,
    )
    left, right = st.columns([1.1, .9])
    with left:
        st.markdown(
            """
            <div class="login-hero">
                <div style="font-size:44px">🛒</div>
                <div class="login-title">AYÇA Business Market</div>
                <div class="login-sub">Market sahibinin sabah ilk bakacağı dijital yönetim paneli. Satış, stok, kârlılık, raf verimliliği, sermaye kilidi ve tahsilat açığını tek ekranda yorumlar.</div>
                <div class="feature-row">💰 Hangi kategori en kârlı?</div>
                <div class="feature-row">📦 Param hangi stoklarda kilitli?</div>
                <div class="feature-row">💀 Ölü stoklar ne kadar sermaye bağlıyor?</div>
                <div class="feature-row">🤖 Bugün ne yapmalıyım?</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with right:
        st.markdown('<div class="login-card">', unsafe_allow_html=True)
        st.markdown("### Giriş Yap")
        username = st.text_input("Kullanıcı adı", value="premium")
        password = st.text_input("Şifre", type="password")
        if st.button("🚀 Dashboard'a Giriş Yap", use_container_width=True):
            user = DEMO_USERS.get(username.strip().lower())
            if user and password == user["password"]:
                st.session_state.authenticated = True
                st.session_state.auth_user = user["name"]
                st.session_state.auth_store = user["store"]
                st.session_state.membership = user["membership"]
                safe_rerun()
            else:
                st.error("Premium: premium / premium2026 · Basic: basic / basic2026")
        st.info("Premium demo: premium / premium2026\n\nBasic demo: basic / basic2026")
        st.markdown('</div>', unsafe_allow_html=True)


if "authenticated" not in st.session_state:
    st.session_state.authenticated = False
if not st.session_state.authenticated:
    show_auth()
    st.stop()

# Basic lock wrappers
_orig_dataframe = st.dataframe
_orig_plotly = st.plotly_chart
_orig_download = st.download_button


def locked_info(msg="Basic demo bu bölümde sınırlı önizleme gösterir. Detaylar Premium'da açılır."):
    st.warning("🔒 " + msg)


def df_limited(data=None, *args, **kwargs):
    if is_premium():
        return _orig_dataframe(data, *args, **kwargs)
    locked_info("Tablo en fazla 3 satır gösterilir.")
    return _orig_dataframe(data.head(3) if hasattr(data, "head") else data, *args, **kwargs)


def chart_limited(*args, **kwargs):
    if is_premium():
        return _orig_plotly(*args, **kwargs)
    locked_info("Grafikler Premium üyelikte görünür.")


def download_limited(*args, **kwargs):
    if is_premium():
        return _orig_download(*args, **kwargs)
    locked_info("Excel raporu Premium üyelikte indirilir.")
    return False


st.dataframe = df_limited
st.plotly_chart = chart_limited
st.download_button = download_limited

# ============================================================
# HELPERS
# ============================================================

def normalize(value: Any) -> str:
    tr = str.maketrans("çğıöşüİı", "cgiosuii")
    return re.sub(r"_+", "_", re.sub(r"[^a-z0-9]+", "_", str(value).strip().lower().translate(tr))).strip("_")


def find_col(cols, candidates):
    norm = {c: normalize(c) for c in cols}
    for candidate in candidates:
        cn = normalize(candidate)
        for c, n in norm.items():
            if cn == n:
                return c
    for candidate in candidates:
        cn = normalize(candidate)
        for c, n in norm.items():
            if cn in n:
                return c
    return None


def compact_money(x):
    """Kartlarda taşma olmaması için büyük rakamları kısaltır."""
    try:
        val = float(x)
    except Exception:
        return "₺0"
    sign = "-" if val < 0 else ""
    val = abs(val)
    if val >= 1_000_000_000:
        return f"{sign}₺{val/1_000_000_000:.1f} Mr".replace(".", ",")
    if val >= 1_000_000:
        return f"{sign}₺{val/1_000_000:.1f} Mn".replace(".", ",")
    if val >= 100_000:
        return f"{sign}₺{val/1_000:.0f} Bin"
    return f"{sign}₺{val:,.0f}".replace(",", ".")


def money(x):
    try:
        return f"₺{float(x):,.0f}".replace(",", ".")
    except Exception:
        return "₺0"


def num(x, d=1):
    try:
        return f"{float(x):,.{d}f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return "0"


def pct(x):
    try:
        return "%" + num(float(x) * 100, 1)
    except Exception:
        return "%0,0"


def trend(cur, prev):
    try:
        if float(prev) == 0:
            return ("▲ yeni", "metric-up") if float(cur) > 0 else ("-", "metric-up")
        ratio = (float(cur) - float(prev)) / float(prev)
        return (("▲ " + pct(ratio), "metric-up") if ratio >= 0 else ("▼ " + pct(abs(ratio)), "metric-down"))
    except Exception:
        return ("-", "metric-up")


def date_parse(s):
    return pd.to_datetime(s, errors="coerce", dayfirst=True)


def metric_card(label, value, sub="", tr=None, tc="metric-up"):
    th = f" <span class='{tc}'>{escape(str(tr))}</span>" if tr else ""
    st.markdown(
        f"""
        <div class="metric-card">
          <div>
            <div class="metric-label">{escape(str(label))}</div>
            <div class="metric-value">{escape(str(value))}</div>
          </div>
          <div class="metric-sub">{escape(str(sub))}{th}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def apply_theme(fig, h=360):
    fig.update_layout(
        height=h,
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="#FFFFFF",
        font=dict(color="#0F172A"),
        margin=dict(l=10, r=10, t=48, b=10),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
    )
    fig.update_xaxes(showgrid=False, color="#64748B")
    fig.update_yaxes(gridcolor="#E2E8F0", color="#64748B")
    return fig

# ============================================================
# DATA READER
# ============================================================

def read_excel_smart(file):
    xls = pd.ExcelFile(file)
    sheet = None
    for s in xls.sheet_names:
        if normalize(s) == "ayca_business_data":
            sheet = s
            break
    if sheet is None:
        for s in xls.sheet_names:
            if any(k in normalize(s) for k in ["data", "veri", "satis", "market"]):
                sheet = s
                break
    if sheet is None:
        sheet = xls.sheet_names[0]
    return pd.read_excel(file, sheet_name=sheet), sheet, xls.sheet_names


def standardize(raw):
    cols = list(raw.columns)
    mapping = {
        "tarih": find_col(cols, ["Tarih"]),
        "fis": find_col(cols, ["Fiş No", "Fis No", "Fatura No"]),
        "kanal": find_col(cols, ["Kanal", "Kaynak"]),
        "barkod": find_col(cols, ["Barkod"]),
        "urun": find_col(cols, ["Ürün Adı", "Urun Adi", "Ürün"]),
        "marka": find_col(cols, ["Marka"]),
        "kategori": find_col(cols, ["Kategori", "Grup"]),
        "alt": find_col(cols, ["Alt Kategori", "Alt Grup"]),
        "adet": find_col(cols, ["Adet", "Miktar"]),
        "alis": find_col(cols, ["Alış Birim TL", "Alis Birim TL", "Alış Fiyatı", "Alis Fiyati"]),
        "satis": find_col(cols, ["Satış Birim TL", "Satis Birim TL", "Satış Fiyatı", "Satis Fiyati"]),
        "ciro": find_col(cols, ["Ciro TL", "Ciro", "Satış Tutarı"]),
        "maliyet": find_col(cols, ["Maliyet TL", "Maliyet"]),
        "kar": find_col(cols, ["Brüt Kar TL", "Brut Kar TL", "Kar TL", "Kâr TL"]),
        "marj": find_col(cols, ["Brüt Kar %", "Brut Kar %", "Kar Marjı"]),
        "odeme": find_col(cols, ["Ödeme Tipi", "Odeme Tipi"]),
        "tahsil_durum": find_col(cols, ["Tahsilat Durumu"]),
        "vade": find_col(cols, ["Tahsilat Vade Tarihi", "Vade Tarihi"]),
        "tahsil_edilen": find_col(cols, ["Tahsil Edilen TL"]),
        "acik": find_col(cols, ["Açık Tahsilat TL", "Acik Tahsilat TL", "Veresiye", "Alacak"]),
        "stok": find_col(cols, ["Mevcut Stok", "Stok", "Stok Adedi"]),
        "skt": find_col(cols, ["SKT", "Son Kullanma Tarihi", "Miad Tarihi"]),
        "tedarikci": find_col(cols, ["Tedarikçi", "Tedarikci"]),
        "raf_kodu": find_col(cols, ["Raf Kodu", "Raf Lokasyonu"]),
        "raf_metre": find_col(cols, ["Raf Metre", "Raf Alanı", "Raf Alani"]),
        "son90": find_col(cols, ["Son 90 Gün Satış", "Son 90 Gun Satis", "son_90", "son90"]),
        "son180": find_col(cols, ["Son 180 Gün Satış", "Son 180 Gun Satis", "son_180", "son180"]),
        "kampanya": find_col(cols, ["Kampanya"]),
        "iade": find_col(cols, ["İade Edilebilir", "Iade Edilebilir"]),
        "min": find_col(cols, ["Min Stok"]),
        "max": find_col(cols, ["Max Stok"]),
    }
    required = ["tarih", "urun", "kategori", "adet", "ciro", "kar", "stok", "son90"]
    missing = [x for x in required if mapping[x] is None]
    if missing:
        raise ValueError("Eksik zorunlu kolonlar: " + ", ".join(missing))

    df = pd.DataFrame()
    df["tarih"] = date_parse(raw[mapping["tarih"]])
    df["fis"] = raw[mapping["fis"]].astype(str) if mapping["fis"] else np.arange(len(raw)).astype(str)
    df["kanal"] = raw[mapping["kanal"]].astype(str) if mapping["kanal"] else "Mağaza"
    df["barkod"] = raw[mapping["barkod"]].astype(str) if mapping["barkod"] else ""
    df["urun"] = raw[mapping["urun"]].astype(str)
    df["marka"] = raw[mapping["marka"]].astype(str) if mapping["marka"] else "Bilinmiyor"
    df["kategori"] = raw[mapping["kategori"]].astype(str)
    df["alt"] = raw[mapping["alt"]].astype(str) if mapping["alt"] else "Genel"

    numeric_defaults = [
        ("adet", "adet", 0), ("alis", "alis", 0), ("satis", "satis", 0), ("ciro", "ciro", 0),
        ("maliyet", "maliyet", np.nan), ("kar", "kar", 0), ("stok", "stok", 0), ("son90", "son90", 0),
        ("son180", "son180", 0), ("raf_metre", "raf_metre", 0.5), ("min_stok", "min", 0), ("max_stok", "max", 0),
        ("tahsil_edilen", "tahsil_edilen", 0), ("acik_tahsilat", "acik", 0),
    ]
    for col, key, default in numeric_defaults:
        if mapping.get(key):
            df[col] = pd.to_numeric(raw[mapping[key]], errors="coerce").fillna(default)
        else:
            df[col] = default

    df["maliyet"] = df["maliyet"].fillna(df["adet"] * df["alis"])
    # Marj kolonu güvenli hesaplama
    # Not: fillna() içine ndarray verilirse Streamlit Cloud'da
    # "value parameter must be a scalar, dict or Series, but you passed a ndarray" hatası oluşabilir.
    calculated_margin = pd.Series(
        np.where(df["ciro"] > 0, df["kar"] / df["ciro"], 0),
        index=df.index
    )

    if mapping["marj"]:
        margin_series = pd.to_numeric(raw[mapping["marj"]], errors="coerce")
        df["marj"] = margin_series.combine_first(calculated_margin).fillna(0)
    else:
        df["marj"] = calculated_margin.fillna(0)

    df["odeme"] = raw[mapping["odeme"]].astype(str) if mapping["odeme"] else "Bilinmiyor"
    df["tahsil_durum"] = raw[mapping["tahsil_durum"]].astype(str) if mapping["tahsil_durum"] else "Tahsil Edildi"
    df["vade"] = date_parse(raw[mapping["vade"]]) if mapping["vade"] else df["tarih"]
    df["skt"] = date_parse(raw[mapping["skt"]]) if mapping["skt"] else pd.NaT
    df["tedarikci"] = raw[mapping["tedarikci"]].astype(str) if mapping["tedarikci"] else "Bilinmiyor"
    df["raf_kodu"] = raw[mapping["raf_kodu"]].astype(str) if mapping["raf_kodu"] else "Bilinmiyor"
    df["kampanya"] = raw[mapping["kampanya"]].astype(str) if mapping["kampanya"] else "Hayır"
    df["iade"] = raw[mapping["iade"]].astype(str) if mapping["iade"] else "Hayır"
    return df.dropna(subset=["tarih"])


def product_table(df, today):
    s = df.sort_values("tarih")
    p = s.groupby(["barkod", "urun"], dropna=False).agg(
        marka=("marka", "last"), kategori=("kategori", "last"), alt=("alt", "last"), toplam_adet=("adet", "sum"),
        toplam_ciro=("ciro", "sum"), toplam_maliyet=("maliyet", "sum"), toplam_kar=("kar", "sum"),
        ort_satis=("satis", "mean"), ort_alis=("alis", "mean"), stok=("stok", "last"), skt=("skt", "last"),
        tedarikci=("tedarikci", "last"), raf_kodu=("raf_kodu", "last"), raf_metre=("raf_metre", "last"),
        son90=("son90", "last"), son180=("son180", "last"), min_stok=("min_stok", "last"), max_stok=("max_stok", "last"),
        iade=("iade", "last"), son_satis=("tarih", "max"), islem=("fis", "nunique"),
    ).reset_index()
    p["kar_marji"] = np.where(p["toplam_ciro"] > 0, p["toplam_kar"] / p["toplam_ciro"], 0)
    p["gunluk_tuketim"] = p["son90"] / 90
    p["tahmini_bitis_gunu"] = np.where(p["gunluk_tuketim"] > 0, p["stok"] / p["gunluk_tuketim"], np.inf)
    p["stok_degeri"] = p["stok"] * p["ort_alis"]
    p["skt_kalan_gun"] = (p["skt"] - today).dt.days
    p["son_satis_kac_gun"] = (today - p["son_satis"]).dt.days
    p["raf_verimlilik"] = np.where(p["raf_metre"] > 0, p["toplam_kar"] / p["raf_metre"], 0)
    p["stok_devir"] = np.where(p["stok_degeri"] > 0, p["toplam_ciro"] / p["stok_degeri"], 0)
    return p


def classify(p, critical, warning, dead_days, skt_days):
    p = p.copy()
    p["olu_stok_mu"] = (p["stok"] > 0) & ((p["son90"] <= 0) | (p["son_satis_kac_gun"] >= dead_days))
    p["yavas_stok_mu"] = (p["stok"] > 0) & (p["stok_devir"] < 0.60) & (~p["olu_stok_mu"])
    p["kritik_stok_mu"] = (p["gunluk_tuketim"] > 0) & (p["tahmini_bitis_gunu"] <= critical)
    p["uyari_stok_mu"] = (p["gunluk_tuketim"] > 0) & (p["tahmini_bitis_gunu"] <= warning)
    p["skt_risk_mu"] = p["skt_kalan_gun"].notna() & (p["skt_kalan_gun"] <= skt_days)
    p["siparis_adedi"] = np.ceil(np.maximum(0, (warning * 2 * p["gunluk_tuketim"]) - p["stok"]))
    p["siparis_maliyeti"] = p["siparis_adedi"] * p["ort_alis"]
    p["siparis_oncelik"] = np.select([p["kritik_stok_mu"], p["uyari_stok_mu"]], ["Acil", "Bu Hafta"], default="Takip")
    p["aksiyon"] = np.select(
        [p["olu_stok_mu"], p["skt_risk_mu"], p["kritik_stok_mu"], p["kar_marji"] < 0.10],
        ["Kampanya / iade / raf azalt", "Öne çıkar / kampanya", "Acil sipariş", "Fiyat-maliyet kontrolü"],
        default="Takip et",
    )
    return p


def business_score(p, df, period, prev):
    total = max(1, len(p))
    crit = len(p[p["kritik_stok_mu"]]) / total
    dead_ratio = len(p[p["olu_stok_mu"]]) / total
    slow = len(p[p["yavas_stok_mu"]]) / total
    margin = period["kar"].sum() / period["ciro"].sum() if period["ciro"].sum() > 0 else 0
    prev_ciro = prev["ciro"].sum()
    cur_ciro = period["ciro"].sum()
    sales_trend = (cur_ciro - prev_ciro) / prev_ciro if prev_ciro > 0 else 0
    open_rate = df["acik_tahsilat"].sum() / df["ciro"].sum() if df["ciro"].sum() > 0 else 0
    score = 78 + sales_trend * 18 + margin * 45 - crit * 22 - dead_ratio * 28 - slow * 14 - open_rate * 40
    return int(max(0, min(100, round(score))))


def health_scores(p, period, prev):
    cur = period["ciro"].sum()
    pr = prev["ciro"].sum()
    tr = (cur - pr) / pr if pr > 0 else 0
    margin = period["kar"].sum() / cur if cur > 0 else 0
    total = max(1, len(p))
    crit = len(p[p["kritik_stok_mu"]]) / total
    dead = len(p[p["olu_stok_mu"]]) / total
    slow = len(p[p["yavas_stok_mu"]]) / total
    return {
        "Satış Performansı": int(max(0, min(100, 70 + tr * 50))),
        "Kârlılık": int(max(0, min(100, 55 + margin * 150))),
        "Stok Sağlığı": int(max(0, min(100, 100 - crit * 70 - dead * 45))),
        "Sermaye Verimliliği": int(max(0, min(100, 100 - slow * 60 - dead * 65))),
        "Kategori Performansı": int(max(0, min(100, 70 + len(p["kategori"].unique()) * 2))),
        "Nakit Akışı": int(max(0, min(100, 100 - (period["acik_tahsilat"].sum() / cur * 220 if cur > 0 else 0)))),
    }


def ai_comment(p, period, prev, category):
    msgs = []
    cur = period["ciro"].sum()
    pr = prev["ciro"].sum()
    if pr > 0:
        ratio = (cur - pr) / pr
        msgs.append(f"Güncel dönem cirosu önceki döneme göre {pct(abs(ratio))} {'arttı' if ratio >= 0 else 'düştü'}.")
    if not category.empty:
        best = category.sort_values("kar", ascending=False).iloc[0]
        msgs.append(f"En kârlı ürün grubu {best['Kategori']} görünüyor; dönem kârı {money(best['kar'])}.")
    dead = p[p["olu_stok_mu"]]
    slow = p[p["yavas_stok_mu"]]
    if not dead.empty:
        msgs.append(f"Ölü stoklarda {money(dead['stok_degeri'].sum())} sermaye kilitli; kampanya/iade aksiyonu önerilir.")
    if not slow.empty:
        msgs.append(f"Yavaş dönen ürünlerde {money(slow['stok_degeri'].sum())} stok değeri var; raf alanı gözden geçirilmeli.")
    crit = p[p["kritik_stok_mu"]]
    if not crit.empty:
        row = crit.sort_values("tahmini_bitis_gunu").iloc[0]
        msgs.append(f"{row['urun']} yaklaşık {num(row['tahmini_bitis_gunu'], 0)} gün içinde bitebilir; acil sipariş kontrolü yapın.")
    low = p[(p["toplam_ciro"] > 0) & (p["kar_marji"] < 0.10)]
    if not low.empty:
        msgs.append(f"{len(low)} üründe marj %10 altında; fiyat/maliyet kontrolü önerilir.")
    return " ".join(msgs) if msgs else "Genel tablo dengeli görünüyor; satış, stok ve kârlılık takibi sürdürülmeli."


def excel_report(df, p, cat, order, dead, slow, collection):
    """Boş tablo gelse bile Excel raporunu güvenli üretir."""
    out = BytesIO()

    def safe_df(x):
        if x is None or not isinstance(x, pd.DataFrame) or x.empty:
            return pd.DataFrame({"Bilgi": ["Bu bölüm için veri bulunamadı."]})
        return x.replace([np.inf, -np.inf], np.nan)

    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        pd.DataFrame({
            "Rapor": ["AYÇA Business Market"],
            "Sürüm": ["V1.2 Stable"],
            "Durum": ["Başarıyla oluşturuldu"],
            "Tarih": [pd.Timestamp.now()],
        }).to_excel(writer, "Ozet", index=False)
        safe_df(df).to_excel(writer, "Ham Veri", index=False)
        safe_df(p).to_excel(writer, "Ürün Analizi", index=False)
        safe_df(cat).to_excel(writer, "Kategori Karlılık", index=False)
        safe_df(order).to_excel(writer, "Sipariş", index=False)
        safe_df(dead).to_excel(writer, "Ölü Stok", index=False)
        safe_df(slow).to_excel(writer, "Yavaş Stok", index=False)
        safe_df(collection).to_excel(writer, "Tahsilat", index=False)
    out.seek(0)
    return out.getvalue()

# ============================================================
# SIDEBAR
# ============================================================
st.sidebar.success(f"Giriş: {st.session_state.get('auth_user', 'Demo')} · {st.session_state.get('membership', 'Basic')}")
if st.sidebar.button("Çıkış Yap", use_container_width=True):
    for key in ["authenticated", "auth_user", "auth_store", "membership"]:
        st.session_state.pop(key, None)
    safe_rerun()

st.sidebar.title("🛒 AYÇA Business")
st.sidebar.caption("Market V1.2 Stable Demo")
store_name = st.sidebar.text_input("Market / İşletme Adı", value="AYÇA Market")
owner_name = st.sidebar.text_input("Kullanıcı", value="Abdullah Bey")
file = st.sidebar.file_uploader("Market Excel dosyasını yükle", type=["xlsx", "xls"])
st.sidebar.markdown("---")
critical_days = st.sidebar.slider("Kritik stok günü", 1, 15, 5)
warning_days = st.sidebar.slider("Sipariş uyarı günü", 7, 45, 21)
dead_days = st.sidebar.slider("Ölü stok günü", 30, 180, 90)
skt_days = st.sidebar.slider("SKT uyarı günü", 15, 180, 60)
period_choice = st.sidebar.selectbox("Dashboard dönemi", ["Son 7 gün", "Son 14 gün", "Son 30 gün", "Son 90 gün", "Tüm veri"], index=2)

if file is None:
    st.markdown(
        """
        <div class="ayca-header">
          <div class="ayca-title"><h1>AYÇA Business Market V1.2</h1><p>Market verinizi yükleyerek satış, stok, kârlılık ve sermaye analizini başlatın.</p></div>
          <div class="header-pill">Dosya bekleniyor</div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.info("Sol menüden AYCA_BUSINESS_DATA sayfalı demo Excel dosyasını yükleyin.")
    st.stop()

try:
    raw, sheet, sheets = read_excel_smart(file)
    df = standardize(raw)
except Exception as e:
    st.error(f"Dosya okunurken hata oluştu: {e}")
    st.stop()

# ============================================================
# ANALYSIS
# ============================================================
today = pd.Timestamp.today().normalize()
p = classify(product_table(df, today), critical_days, warning_days, dead_days, skt_days)
max_date = df["tarih"].max()
if period_choice == "Son 7 gün":
    start = max_date - pd.Timedelta(days=7)
elif period_choice == "Son 14 gün":
    start = max_date - pd.Timedelta(days=14)
elif period_choice == "Son 30 gün":
    start = max_date - pd.Timedelta(days=30)
elif period_choice == "Son 90 gün":
    start = max_date - pd.Timedelta(days=90)
else:
    start = df["tarih"].min()

period = df[df["tarih"] >= start].copy()
days_count = max(1, (max_date - start).days)
prev = df[(df["tarih"] >= start - pd.Timedelta(days=days_count)) & (df["tarih"] < start)].copy()

cur_ciro = period["ciro"].sum()
prev_ciro = prev["ciro"].sum()
cur_kar = period["kar"].sum()
prev_kar = prev["kar"].sum()
margin = cur_kar / cur_ciro if cur_ciro > 0 else 0
prev_margin = prev_kar / prev_ciro if prev_ciro > 0 else 0
transactions = period["fis"].nunique()
avg_basket = cur_ciro / transactions if transactions else 0
total_stock = p["stok_degeri"].sum()
open_collection = df["acik_tahsilat"].sum()
score = business_score(p, df, period, prev)

category = period.groupby("kategori").agg(
    ciro=("ciro", "sum"), kar=("kar", "sum"), adet=("adet", "sum"), islem=("fis", "nunique"), acik=("acik_tahsilat", "sum")
).reset_index().rename(columns={"kategori": "Kategori"})
category["marj"] = np.where(category["ciro"] > 0, category["kar"] / category["ciro"], 0)
cat_stock = p.groupby("kategori").agg(stok_degeri=("stok_degeri", "sum"), raf_metre=("raf_metre", "sum")).reset_index().rename(columns={"kategori": "Kategori"})
category = category.merge(cat_stock, on="Kategori", how="left")
category["raf_verimlilik"] = np.where(category["raf_metre"] > 0, category["kar"] / category["raf_metre"], 0)
category["stok_devir"] = np.where(category["stok_degeri"] > 0, category["ciro"] / category["stok_degeri"], 0)

critical = p[p["kritik_stok_mu"]].copy()
order = p[(p["uyari_stok_mu"]) & (p["siparis_adedi"] > 0)].copy()
dead = p[p["olu_stok_mu"]].copy()
slow = p[p["yavas_stok_mu"]].copy()
skt = p[p["skt_risk_mu"]].copy()
low_margin = p[(p["toplam_ciro"] > 0) & (p["kar_marji"] < 0.10)].copy()
collection = df[df["acik_tahsilat"] > 0].copy()

# ============================================================
# UI
# ============================================================
st.markdown(
    f"""
    <div class="ayca-header">
      <div class="ayca-title"><h1>AYÇA Business Market V1.2</h1><p>{escape(store_name)} · {period_choice} · Sayfa: {escape(str(sheet))} · {datetime.now().strftime('%d.%m.%Y')}</p></div>
      <div class="header-pill">İşletme Skoru: {score}/100</div>
    </div>
    """,
    unsafe_allow_html=True,
)

cols = st.columns(6)
metric_args = [
    ("Güncel Ciro", compact_money(cur_ciro), period_choice, *trend(cur_ciro, prev_ciro)),
    ("Brüt Kâr", compact_money(cur_kar), "Dönem kârı", *trend(cur_kar, prev_kar)),
    ("Kâr Marjı", pct(margin), "Brüt kâr / ciro", *trend(margin, prev_margin)),
    ("Ortalama Sepet", compact_money(avg_basket), "Ciro / fiş", None, "metric-up"),
    ("Stok Değeri", compact_money(total_stock), "Mevcut ürünlerin maliyet değeri", None, "metric-up"),
    ("Tahsilat Açığı", compact_money(open_collection), "Açık veresiye / alacak", None, "metric-down"),
]
for col, args in zip(cols, metric_args):
    with col:
        metric_card(*args)

health = health_scores(p, period, prev)
health_html = "".join([
    f'<div class="health-row"><div class="health-head"><span>{escape(str(k))}</span><span>{v}/100</span></div><div class="health-bar-bg"><div class="health-bar-fill" style="width:{v}%"></div></div></div>'
    for k, v in health.items()
])
ai = ai_comment(p, period, prev, category)

tasks = []
if not order.empty:
    row = order.sort_values(["siparis_oncelik", "tahmini_bitis_gunu"]).iloc[0]
    tasks.append(f"☐ {escape(str(row['urun']))} için sipariş kontrolü yap; tahmini bitiş {num(row['tahmini_bitis_gunu'], 0)} gün.")
if not dead.empty:
    tasks.append(f"☐ Ölü stok listesindeki {len(dead)} ürün için kampanya/iade planı çıkar.")
if not slow.empty:
    tasks.append(f"☐ Yavaş dönen stoklarda {compact_money(slow['stok_degeri'].sum())} sermayeyi gözden geçir.")
if not low_margin.empty:
    tasks.append(f"☐ Marjı düşük {len(low_margin)} üründe fiyat/maliyet kontrolü yap.")
if not collection.empty:
    tasks.append(f"☐ Açık tahsilat toplamı {compact_money(collection['acik_tahsilat'].sum())}; vade kontrolü yap.")
if not skt.empty:
    tasks.append(f"☐ SKT riski olan {len(skt)} ürünü öne çıkar veya kampanyaya al.")
if not tasks:
    tasks = ["☐ Genel tablo dengeli; haftalık satış ve stok kontrolünü sürdür."]

tasks_html = "".join([f'<div class="task-item">{t}</div>' for t in tasks[:6]])

left, right = st.columns([1.25, .75])
with left:
    st.markdown(
        f"""
        <div class="exec-card">
          <div class="exec-title">🤖 AYÇA Ticaret Asistanı</div>
          <div class="exec-sub">{escape(ai)}</div>
          <div class="exec-list-item">Bu ekran; ciroyu, kârı, stokta kilitli sermayeyi, tahsilat açığını ve sipariş ihtiyacını birlikte yorumlar.</div>
        </div>
        """,
        unsafe_allow_html=True,
    )
with right:
    st.markdown(
        f"""
        <div class="exec-card">
          <div class="score-label">İşletme Sağlık Skoru</div>
          <div class="score-big">{score}</div>
          {health_html}
        </div>
        """,
        unsafe_allow_html=True,
    )

st.markdown('<div class="section-title">Bugünün Yönetim Paneli</div>', unsafe_allow_html=True)
radar_items = [
    ("Acil Sipariş", len(order), "Ürün"),
    ("Kritik Stok", len(critical), "Ürün"),
    ("Ölü Stok", compact_money(dead["stok_degeri"].sum()) if not dead.empty else "₺0", "Sermaye"),
    ("Yavaş Stok", compact_money(slow["stok_degeri"].sum()) if not slow.empty else "₺0", "Sermaye"),
    ("SKT Riski", len(skt), "Ürün"),
    ("Düşük Marj", len(low_margin), "Ürün"),
]
radar_html = "<div class='radar-grid'>" + "".join([
    f"<div class='radar-card'><div class='radar-title'>{escape(str(t))}</div><div class='radar-value'>{escape(str(v))}</div><div class='radar-note'>{escape(str(n))}</div></div>"
    for t, v, n in radar_items
]) + "</div>"
st.markdown(radar_html, unsafe_allow_html=True)

st.markdown('<div class="task-grid"><div class="task-card"><div class="exec-title">✅ Bugün Ne Yapmalıyım?</div>' + tasks_html + '</div><div class="task-card"><div class="exec-title">💡 Kısa Yorum</div><div class="exec-sub">Büyük rakamlar kartlarda artık kısaltılmış gösterilir; detaylı tutarlar tablo ve Excel raporunda tam değer olarak korunur.</div></div></div>', unsafe_allow_html=True)

# ============================================================
# TABS
# ============================================================
t_overview, t_cat, t_order, t_dead, t_shelf, t_collect, t_report = st.tabs([
    "Genel Bakış", "Kategori Kârlılık", "Sipariş", "Ölü/Yavaş Stok", "Raf Verimliliği", "Tahsilat", "Rapor"
])

with t_overview:
    st.markdown("#### Ciro ve kâr trendi")
    daily = period.groupby("tarih").agg(ciro=("ciro", "sum"), kar=("kar", "sum"), acik=("acik_tahsilat", "sum")).reset_index()
    if not daily.empty:
        fig = go.Figure()
        fig.add_trace(go.Scatter(x=daily["tarih"], y=daily["ciro"], mode="lines+markers", name="Ciro"))
        fig.add_trace(go.Scatter(x=daily["tarih"], y=daily["kar"], mode="lines+markers", name="Brüt Kâr"))
        st.plotly_chart(apply_theme(fig, 390), use_container_width=True)
    st.dataframe(period.sort_values("tarih", ascending=False).head(250), use_container_width=True)

with t_cat:
    st.markdown("#### Hangi ürün grubu daha kârlı?")
    view = category.sort_values("kar", ascending=False).copy()
    st.dataframe(view, use_container_width=True)
    if not view.empty:
        st.plotly_chart(apply_theme(px.bar(view.head(20), x="Kategori", y="kar", title="Kategori Bazlı Brüt Kâr"), 390), use_container_width=True)
        st.plotly_chart(apply_theme(px.bar(view.head(20), x="Kategori", y="raf_verimlilik", title="Kategori Raf Verimliliği"), 390), use_container_width=True)

with t_order:
    st.markdown("#### Sipariş asistanı")
    st.info(f"Toplam önerilen sipariş maliyeti: {money(order['siparis_maliyeti'].sum()) if not order.empty else '₺0'}")
    columns = ["urun", "kategori", "stok", "son90", "gunluk_tuketim", "tahmini_bitis_gunu", "siparis_adedi", "siparis_maliyeti", "siparis_oncelik", "tedarikci"]
    st.dataframe(order.sort_values("tahmini_bitis_gunu")[columns].replace([np.inf, -np.inf], np.nan), use_container_width=True)
    if not order.empty:
        st.plotly_chart(apply_theme(px.bar(order.sort_values("siparis_maliyeti", ascending=False).head(20), x="urun", y="siparis_maliyeti", color="siparis_oncelik", title="Sipariş Maliyeti Öncelik Listesi"), 420), use_container_width=True)

with t_dead:
    st.markdown("#### Ölü stoklar ve yavaş dönen stoklar")
    st.warning(f"Ölü stoklarda toplam {money(dead['stok_degeri'].sum()) if not dead.empty else '₺0'} değerinde sermaye görünüyor.")
    dead_cols = ["urun", "kategori", "stok", "stok_degeri", "son90", "son_satis_kac_gun", "iade", "aksiyon"]
    slow_cols = ["urun", "kategori", "stok", "stok_degeri", "stok_devir", "son90", "aksiyon"]
    st.markdown("##### Ölü stoklar")
    st.dataframe(dead.sort_values("stok_degeri", ascending=False)[dead_cols], use_container_width=True)
    st.markdown("##### Yavaş dönen stoklar")
    st.dataframe(slow.sort_values("stok_degeri", ascending=False)[slow_cols], use_container_width=True)

with t_shelf:
    st.markdown("#### Raf verimliliği = kâr / raf metre")
    shelf = p.sort_values("raf_verimlilik", ascending=False)
    st.dataframe(shelf[["urun", "kategori", "raf_kodu", "raf_metre", "toplam_kar", "raf_verimlilik", "stok_degeri", "stok_devir"]], use_container_width=True)
    if not category.empty:
        st.plotly_chart(apply_theme(px.bar(category.sort_values("raf_verimlilik", ascending=False), x="Kategori", y="raf_verimlilik", title="Kategori Raf Verimliliği"), 390), use_container_width=True)

with t_collect:
    st.markdown("#### Tahsilat açığı / veresiye riski")
    st.dataframe(collection[["tarih", "fis", "urun", "kategori", "odeme", "tahsil_durum", "vade", "ciro", "acik_tahsilat"]], use_container_width=True)
    by_day = collection.groupby("vade").agg(acik=("acik_tahsilat", "sum"), fis=("fis", "nunique")).reset_index() if not collection.empty else pd.DataFrame()
    if not by_day.empty:
        st.plotly_chart(apply_theme(px.bar(by_day, x="vade", y="acik", title="Vadeye Göre Açık Tahsilat"), 360), use_container_width=True)

with t_report:
    st.markdown("#### Excel raporu indir")
    data = excel_report(df, p, category, order, dead, slow, collection)
    st.download_button(
        "📥 AYÇA Business Market Raporu İndir",
        data=data,
        file_name="ayca_business_market_rapor.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
    st.caption("Rapor içinde Ozet, Ham Veri, Ürün Analizi, Kategori Kârlılık, Sipariş, Ölü Stok, Yavaş Stok ve Tahsilat sayfaları bulunur.")

st.divider()
st.caption("AYÇA Business Market V1.2 · Veriler kalıcı olarak saklanmaz. Excel dosyası anlık analiz edilir.")
