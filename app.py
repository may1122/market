# AYÇA Business Market - Streamlit App
# Gerekli paketler: streamlit pandas openpyxl plotly
# Çalıştırma: streamlit run app.py

import io
from datetime import datetime

import pandas as pd
import plotly.express as px
import streamlit as st

st.set_page_config(page_title="AYÇA Business Market", page_icon="🛒", layout="wide")

# -----------------------------
# Yardımcı Fonksiyonlar
# -----------------------------

def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip().lower().replace(" ", "_").replace("ı", "i") for c in df.columns]
    return df


def read_excel_sheets(uploaded_file):
    xl = pd.ExcelFile(uploaded_file)
    sheets = {name: normalize_columns(pd.read_excel(uploaded_file, sheet_name=name)) for name in xl.sheet_names}
    return sheets


def find_sheet(sheets, candidates):
    for key, df in sheets.items():
        k = key.lower().replace("ı", "i")
        if any(c in k for c in candidates):
            return df
    return None


def money(x):
    try:
        return f"₺{float(x):,.0f}".replace(",", ".")
    except Exception:
        return "₺0"


def safe_div(a, b):
    return a / b if b not in [0, None] else 0


def prepare_sales(df):
    df = df.copy()
    required = ["tarih", "barkod", "urun_adi", "kategori", "adet", "satis_fiyati", "alis_fiyati"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        st.error(f"Satış sayfasında eksik kolonlar: {missing}")
        st.stop()

    df["tarih"] = pd.to_datetime(df["tarih"], errors="coerce")
    for c in ["adet", "satis_fiyati", "alis_fiyati"]:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)
    df["ciro"] = df["adet"] * df["satis_fiyati"]
    df["maliyet"] = df["adet"] * df["alis_fiyati"]
    df["brut_kar"] = df["ciro"] - df["maliyet"]
    df["brut_marj_%"] = df.apply(lambda r: safe_div(r["brut_kar"], r["ciro"]) * 100, axis=1)
    df["gun"] = df["tarih"].dt.date
    return df


def prepare_stock(df):
    df = df.copy()
    required = ["barkod", "urun_adi", "kategori", "stok_adedi", "alis_fiyati", "satis_fiyati", "min_stok", "miad_tarihi"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        st.error(f"Stok sayfasında eksik kolonlar: {missing}")
        st.stop()

    df["miad_tarihi"] = pd.to_datetime(df["miad_tarihi"], errors="coerce")
    for c in ["stok_adedi", "alis_fiyati", "satis_fiyati", "min_stok"]:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)
    df["stok_degeri"] = df["stok_adedi"] * df["alis_fiyati"]
    df["potansiyel_satis_degeri"] = df["stok_adedi"] * df["satis_fiyati"]
    df["stok_marji_%"] = df.apply(lambda r: safe_div(r["satis_fiyati"] - r["alis_fiyati"], r["satis_fiyati"]) * 100, axis=1)
    return df


def prepare_receivables(df):
    if df is None:
        return pd.DataFrame(columns=["musteri", "tutar", "vade_tarihi", "durum"])
    df = df.copy()
    if "tutar" in df.columns:
        df["tutar"] = pd.to_numeric(df["tutar"], errors="coerce").fillna(0)
    if "vade_tarihi" in df.columns:
        df["vade_tarihi"] = pd.to_datetime(df["vade_tarihi"], errors="coerce")
    return df


def risk_label(row):
    if row.get("stok_adedi", 0) <= 0:
        return "Stok Yok"
    if row.get("stok_adedi", 0) <= row.get("min_stok", 0):
        return "Kritik Stok"
    if row.get("gunluk_satis_hizi", 0) == 0:
        return "Ölü Stok"
    if row.get("tahmini_bitis_gunu", 999) <= 7:
        return "7 Gün İçinde Biter"
    if row.get("miada_kalan_gun", 9999) <= 60:
        return "Miad Riski"
    return "Normal"

# -----------------------------
# Arayüz
# -----------------------------

st.title("🛒 AYÇA Business Market")
st.caption("Marketler için satış, stok, kârlılık, sipariş ve tahsilat karar destek ekranı")

with st.sidebar:
    st.header("Veri Yükleme")
    uploaded_file = st.file_uploader("Excel yükle", type=["xlsx"])
    st.info("Beklenen sayfalar: Satislar, Stoklar, Tahsilat")
    st.divider()
    analysis_days = st.slider("Satış hızı hesaplama günü", 7, 90, 30)
    order_cover_days = st.slider("Sipariş hedef stok günü", 7, 60, 21)

if uploaded_file is None:
    st.warning("Başlamak için örnek Excel dosyasını veya kendi market verinizi yükleyin.")
    st.stop()

sheets = read_excel_sheets(uploaded_file)
sales_raw = find_sheet(sheets, ["satis", "sales"])
stock_raw = find_sheet(sheets, ["stok", "stock", "envanter"])
receivables_raw = find_sheet(sheets, ["tahsilat", "alacak", "receivable"])

if sales_raw is None or stock_raw is None:
    st.error("Excel içinde en az Satislar ve Stoklar sayfaları olmalı.")
    st.stop()

sales = prepare_sales(sales_raw)
stock = prepare_stock(stock_raw)
receivables = prepare_receivables(receivables_raw)

max_date = sales["tarih"].max()
min_speed_date = max_date - pd.Timedelta(days=analysis_days)
recent_sales = sales[sales["tarih"] >= min_speed_date].copy()

product_sales = recent_sales.groupby(["barkod", "urun_adi", "kategori"], as_index=False).agg(
    satilan_adet=("adet", "sum"),
    ciro=("ciro", "sum"),
    brut_kar=("brut_kar", "sum")
)
product_sales["gunluk_satis_hizi"] = product_sales["satilan_adet"] / analysis_days
product_sales["marj_%"] = product_sales.apply(lambda r: safe_div(r["brut_kar"], r["ciro"]) * 100, axis=1)

merged = stock.merge(product_sales[["barkod", "satilan_adet", "ciro", "brut_kar", "gunluk_satis_hizi", "marj_%"]], on="barkod", how="left")
for c in ["satilan_adet", "ciro", "brut_kar", "gunluk_satis_hizi", "marj_%"]:
    merged[c] = merged[c].fillna(0)

merged["tahmini_bitis_gunu"] = merged.apply(lambda r: safe_div(r["stok_adedi"], r["gunluk_satis_hizi"]) if r["gunluk_satis_hizi"] > 0 else 9999, axis=1)
merged["onerilen_siparis"] = ((merged["gunluk_satis_hizi"] * order_cover_days) - merged["stok_adedi"]).clip(lower=0).round(0)
merged["miada_kalan_gun"] = (merged["miad_tarihi"] - pd.Timestamp.today()).dt.days
merged["risk"] = merged.apply(risk_label, axis=1)

# Tahsilat
if not receivables.empty and "tutar" in receivables.columns:
    total_receivable = receivables["tutar"].sum()
    if "vade_tarihi" in receivables.columns:
        overdue_receivable = receivables.loc[receivables["vade_tarihi"] < pd.Timestamp.today(), "tutar"].sum()
    else:
        overdue_receivable = 0
else:
    total_receivable = 0
    overdue_receivable = 0

# KPI
k1, k2, k3, k4, k5 = st.columns(5)
k1.metric("Toplam Ciro", money(sales["ciro"].sum()))
k2.metric("Brüt Kâr", money(sales["brut_kar"].sum()))
k3.metric("Brüt Marj", f"{safe_div(sales['brut_kar'].sum(), sales['ciro'].sum()) * 100:.1f}%")
k4.metric("Stok Maliyeti", money(stock["stok_degeri"].sum()))
k5.metric("Tahsilat Açığı", money(overdue_receivable))

st.divider()

# Günlük öneriler
st.subheader("🤖 Bugünün İlk 5 Aksiyonu")
actions = []
critical = merged[merged["risk"].isin(["Stok Yok", "Kritik Stok", "7 Gün İçinde Biter"])].sort_values("onerilen_siparis", ascending=False)
dead = merged[(merged["gunluk_satis_hizi"] == 0) & (merged["stok_adedi"] > 0)].sort_values("stok_degeri", ascending=False)
expiry = merged[merged["miada_kalan_gun"].between(0, 60)].sort_values("miada_kalan_gun")
low_margin = merged[(merged["marj_%"] > 0) & (merged["marj_%"] < 12)].sort_values("marj_%")

if not critical.empty:
    r = critical.iloc[0]
    actions.append(f"**Sipariş:** {r['urun_adi']} için yaklaşık **{int(r['onerilen_siparis'])} adet** sipariş öneriliyor.")
if not dead.empty:
    r = dead.iloc[0]
    actions.append(f"**Ölü stok:** {r['urun_adi']} ürününde **{money(r['stok_degeri'])}** sermaye rafta bekliyor.")
if not expiry.empty:
    r = expiry.iloc[0]
    actions.append(f"**Miad:** {r['urun_adi']} ürününün miadına **{int(r['miada_kalan_gun'])} gün** kaldı.")
if not low_margin.empty:
    r = low_margin.iloc[0]
    actions.append(f"**Düşük marj:** {r['urun_adi']} marjı **%{r['marj_%']:.1f}**. Fiyat kontrol edilmeli.")
if overdue_receivable > 0:
    actions.append(f"**Tahsilat:** Vadesi geçmiş toplam **{money(overdue_receivable)}** alacak var.")

if not actions:
    st.success("Bugün için kritik aksiyon görünmüyor.")
else:
    for a in actions[:5]:
        st.markdown("- " + a)

# Sekmeler
summary_tab, sales_tab, stock_tab, order_tab, receivable_tab = st.tabs([
    "📊 Özet", "🔥 Satış", "📦 Stok & Risk", "🧾 Sipariş Tavsiyesi", "💰 Tahsilat"
])

with summary_tab:
    c1, c2 = st.columns(2)
    daily = sales.groupby("gun", as_index=False).agg(ciro=("ciro", "sum"), brut_kar=("brut_kar", "sum"))
    c1.plotly_chart(px.line(daily, x="gun", y="ciro", title="Günlük Ciro"), use_container_width=True)
    cat = sales.groupby("kategori", as_index=False).agg(ciro=("ciro", "sum"), brut_kar=("brut_kar", "sum")).sort_values("ciro", ascending=False)
    c2.plotly_chart(px.bar(cat, x="kategori", y="ciro", title="Kategori Bazlı Ciro"), use_container_width=True)

with sales_tab:
    top_products = product_sales.sort_values("ciro", ascending=False)
    st.dataframe(top_products, use_container_width=True)
    st.plotly_chart(px.bar(top_products.head(20), x="urun_adi", y="ciro", title="En Çok Ciro Yapan 20 Ürün"), use_container_width=True)

with stock_tab:
    st.dataframe(merged.sort_values(["risk", "stok_degeri"], ascending=[True, False]), use_container_width=True)
    risk_count = merged.groupby("risk", as_index=False).agg(urun_sayisi=("barkod", "count"), stok_degeri=("stok_degeri", "sum"))
    st.plotly_chart(px.bar(risk_count, x="risk", y="stok_degeri", title="Risklere Göre Stok Değeri"), use_container_width=True)

with order_tab:
    orders = merged[merged["onerilen_siparis"] > 0].sort_values("onerilen_siparis", ascending=False)
    st.dataframe(orders[["barkod", "urun_adi", "kategori", "stok_adedi", "gunluk_satis_hizi", "tahmini_bitis_gunu", "onerilen_siparis"]], use_container_width=True)
    csv = orders.to_csv(index=False).encode("utf-8-sig")
    st.download_button("Sipariş listesini indir", data=csv, file_name="ayca_business_siparis.csv", mime="text/csv")

with receivable_tab:
    if receivables.empty:
        st.info("Tahsilat sayfası bulunamadı veya boş.")
    else:
        st.dataframe(receivables, use_container_width=True)
        st.metric("Toplam Alacak", money(total_receivable))
        st.metric("Vadesi Geçmiş Alacak", money(overdue_receivable))
