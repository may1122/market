ChatGPT


Ayça business marketin son çalışmanın txt ve exceli paylaş benimle

Son çalışmada oluşturduğumuz AYÇA Business dosyalarını buldum.


ayca_business_market_v1_app.py.txt
Belge

ayca_business_market_buyuk_demo_veri.xlsx
Hesap Tablosu
TXT Dosyası
ayca_business_market_v1_app.py.txt
AYÇA Business Market Executive sürümü. İçeriklerinde:

İşletme sağlık skoru

En kârlı ürün grubu analizi

Ölü stok sermaye kilidi

Yavaş dönen stoklar

Sipariş asistanı

Tahsilat/veresiye açığı

Raf verimliliği

İçerideki toplam mal değeri

AI Ticaret Asistanı

Excel rapor çıktısı

bulunuyor. 


Excel Dosyası
ayca_business_market_buyuk_demo_veri.xlsx

Özellikleri:

270 gün veri

12.805 satış satırı

1.500 SKU

Ortalama günlük ciro ≈ 109.465 TL

Satış + stok + kârlılık + tahsilat + raf verimliliği analizleri için uygun demo veri seti 


Bu arada son konuştuğumuz noktada sistemi daha da büyütmüştük:

✅ En kârlı ürün grubu
✅ Ölü stok sermaye kilidi
✅ Yavaş dönen stoklar
✅ Tahsilat açığı
✅ Toplam stok değeri
✅ Sipariş önceliklendirme
✅ Raf verimliliği
✅ Yönetici özeti (AI yorumları)

Bir sonraki aşama olarak konuştuğumuz AYÇA Business V2.0'da şunları ekleyebiliriz:

ABC ürün analizi

Pareto (%80 ciroyu getiren ürünler)

Saatlik satış analizi

Kampanya öneri motoru

Kasiyer performansı

Şube karşılaştırması

Nakit akış tahmini

Otomatik satın alma bütçesi

"Bugün ne yapmalıyım?" ekranı

Bence market ürünü burada eczane ürününden bile daha büyük bir ticari potansiyele sahip.


============================================================
AYÇA Business Market V1.0 Executive
Market / bakkal / perakende işletme karar destek paneli
------------------------------------------------------------
Beklenen Excel sayfası: AYCA_BUSINESS_DATA
Örnek veri: ayca_business_market_ornek_veri_v1.xlsx
Çalıştırma:
pip install streamlit pandas numpy openpyxl plotly
streamlit run app.py
============================================================
from future import annotations

import re
from datetime import datetime
from io import BytesIO
from typing import Optional

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

============================================================
CSS
============================================================
st.markdown("""

""", unsafe_allow_html=True)

============================================================
Helpers
============================================================
def normalize_col_name(name: str) -> str:
tr = str.maketrans("çğıöşüİı", "cgiosuii")
name = str(name).strip().lower().translate(tr)
name = re.sub(r"[^a-z0-9]+", "", name)
return re.sub(r"+", "", name).strip("")

def find_col(columns, candidates) -> Optional[str]:
normalized = {c: normalize_col_name(c) for c in columns}
for cand in candidates:
cn = normalize_col_name(cand)
for original, norm in normalized.items():
if cn == norm:
return original
for cand in candidates:
cn = normalize_col_name(cand)
for original, norm in normalized.items():
if cn in norm:
return original
return None

def money_fmt(x) -> str:
try:
return f"₺{float(x):,.0f}".replace(",", ".")
except Exception:
return "₺0"

def num_fmt(x, d=1) -> str:
try:
if pd.isna(x) or np.isinf(float(x)):
return "0"
return f"{float(x):,.{d}f}".replace(",", "X").replace(".", ",").replace("X", ".")
except Exception:
return "0"

def pct_fmt(x) -> str:
return "%" + num_fmt(float(x) * 100 if pd.notna(x) else 0, 1)

def trend(current, previous):
try:
if previous == 0:
return "▲ yeni", "up"
r = (current - previous) / previous
return ("▲ " + pct_fmt(r), "up") if r >= 0 else ("▼ " + pct_fmt(abs(r)), "down")
except Exception:
return "", "up"

def metric_card(label, value, sub="", trend_text="", trend_class="up"):
t = f" {trend_text}" if trend_text else ""
st.markdown(f"""
{label}{value}{sub}{t}
""", unsafe_allow_html=True)

def plot_theme(fig, height=360):
fig.update_layout(height=height, paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="#FFFFFF", font=dict(color="#0F172A"), margin=dict(l=10,r=10,t=48,b=10), legend=dict(orientation="h", y=1.02, x=1, xanchor="right", yanchor="bottom"))
fig.update_xaxes(showgrid=False, color="#64748B")
fig.update_yaxes(gridcolor="#E2E8F0", color="#64748B")
return fig

def read_excel_smart(uploaded_file):
xls = pd.ExcelFile(uploaded_file)
preferred = None
for s in xls.sheet_names:
if normalize_col_name(s) == "ayca_business_data":
preferred = s; break
if preferred is None:
for s in xls.sheet_names:
if any(k in normalize_col_name(s) for k in ["data", "veri", "satis", "market"]):
preferred = s; break
if preferred is None:
preferred = xls.sheet_names[0]
return pd.read_excel(uploaded_file, sheet_name=preferred), preferred

def standardize(raw: pd.DataFrame) -> pd.DataFrame:
cols = list(raw.columns)
m = {
"tarih": find_col(cols, ["Tarih", "Satış Tarihi"]),
"fis": find_col(cols, ["Fiş No", "Fis No", "Belge No"]),
"barkod": find_col(cols, ["Barkod", "Barcode"]),
"urun": find_col(cols, ["Ürün Adı", "Urun Adi", "Mal Adı"]),
"marka": find_col(cols, ["Marka", "Brand"]),
"kategori": find_col(cols, ["Kategori", "Category"]),
"alt": find_col(cols, ["Alt Kategori", "Alt Grup"]),
"adet": find_col(cols, ["Adet", "Miktar", "Satış Adedi"]),
"alis": find_col(cols, ["Alış Birim TL", "Alis Birim TL", "Alış Fiyatı"]),
"satis": find_col(cols, ["Satış Birim TL", "Satis Birim TL", "Satış Fiyatı"]),
"ciro": find_col(cols, ["Ciro TL", "Satış Tutarı", "Ciro"]),
"maliyet": find_col(cols, ["Maliyet TL", "Maliyet"]),
"kar": find_col(cols, ["Brüt Kar TL", "Brut Kar TL", "Kar TL"]),
"odeme": find_col(cols, ["Ödeme Tipi", "Odeme Tipi", "Tahsilat Tipi"]),
"veresiye": find_col(cols, ["Veresiye Tutar TL", "Tahsilat Açığı", "Tahsilat Acigi"]),
"stok": find_col(cols, ["Mevcut Stok", "Stok", "Kalan Stok"]),
"stok_degeri": find_col(cols, ["Stok Değeri TL", "Stok Degeri TL"]),
"son90": find_col(cols, ["Son 90 Gün Satış", "Son 90 Gun Satis", "Son 90"]),
"bitis": find_col(cols, ["Tahmini Bitiş Günü", "Tahmini Bitis Gunu"]),
"skt": find_col(cols, ["SKT", "Miad", "Son Kullanma Tarihi"]),
"tedarikci": find_col(cols, ["Tedarikçi", "Tedarikci"]),
"raf": find_col(cols, ["Raf Kodu", "Raf Lokasyonu", "Raf"]),
"raf_cm": find_col(cols, ["Raf Payı CM", "Raf Payi CM", "Raf Cm"]),
"son_satis": find_col(cols, ["Son Satış Tarihi", "Son Satis Tarihi"]),
}
required = ["tarih", "urun", "kategori", "adet", "ciro", "kar", "stok", "son90"]
missing = [x for x in required if m.get(x) is None]
if missing:
raise ValueError("Eksik zorunlu kolonlar: " + ", ".join(missing))

df = pd.DataFrame()
df["tarih"] = pd.to_datetime(raw[m["tarih"]], errors="coerce", dayfirst=True)
df["fis"] = raw[m["fis"]].astype(str) if m["fis"] else ""
df["barkod"] = raw[m["barkod"]].astype(str) if m["barkod"] else ""
df["urun"] = raw[m["urun"]].astype(str)
df["marka"] = raw[m["marka"]].astype(str) if m["marka"] else "Bilinmiyor"
df["kategori"] = raw[m["kategori"]].astype(str)
df["alt_kategori"] = raw[m["alt"]].astype(str) if m["alt"] else "Genel"
df["adet"] = pd.to_numeric(raw[m["adet"]], errors="coerce").fillna(0)
df["alis_birim"] = pd.to_numeric(raw[m["alis"]], errors="coerce").fillna(0) if m["alis"] else 0
df["satis_birim"] = pd.to_numeric(raw[m["satis"]], errors="coerce").fillna(0) if m["satis"] else 0
df["ciro"] = pd.to_numeric(raw[m["ciro"]], errors="coerce").fillna(0)
df["maliyet"] = pd.to_numeric(raw[m["maliyet"]], errors="coerce").fillna(0) if m["maliyet"] else df["adet"] * df["alis_birim"]
df["brut_kar"] = pd.to_numeric(raw[m["kar"]], errors="coerce").fillna(0)
df["odeme"] = raw[m["odeme"]].astype(str) if m["odeme"] else "Bilinmiyor"
df["veresiye"] = pd.to_numeric(raw[m["veresiye"]], errors="coerce").fillna(0) if m["veresiye"] else np.where(df["odeme"].str.contains("veresiye", case=False, na=False), df["ciro"], 0)
df["stok"] = pd.to_numeric(raw[m["stok"]], errors="coerce").fillna(0)
df["stok_degeri_raw"] = pd.to_numeric(raw[m["stok_degeri"]], errors="coerce").fillna(0) if m["stok_degeri"] else 0
df["son90"] = pd.to_numeric(raw[m["son90"]], errors="coerce").fillna(0)
df["bitis_raw"] = pd.to_numeric(raw[m["bitis"]], errors="coerce").fillna(np.nan) if m["bitis"] else np.nan
df["skt"] = pd.to_datetime(raw[m["skt"]], errors="coerce", dayfirst=True) if m["skt"] else pd.NaT
df["tedarikci"] = raw[m["tedarikci"]].astype(str) if m["tedarikci"] else "Bilinmiyor"
df["raf"] = raw[m["raf"]].astype(str) if m["raf"] else "Bilinmiyor"
df["raf_cm"] = pd.to_numeric(raw[m["raf_cm"]], errors="coerce").fillna(20) if m["raf_cm"] else 20
df["son_satis"] = pd.to_datetime(raw[m["son_satis"]], errors="coerce", dayfirst=True) if m["son_satis"] else df["tarih"]
df = df.dropna(subset=["tarih"])
return df
def product_table(df: pd.DataFrame, today: pd.Timestamp) -> pd.DataFrame:
d = df.sort_values("tarih")
p = d.groupby(["barkod", "urun"], dropna=False).agg(
marka=("marka", "last"), kategori=("kategori", "last"), alt_kategori=("alt_kategori", "last"),
toplam_adet=("adet", "sum"), toplam_ciro=("ciro", "sum"), toplam_maliyet=("maliyet", "sum"), toplam_kar=("brut_kar", "sum"),
ort_alis=("alis_birim", "mean"), ort_satis=("satis_birim", "mean"), stok=("stok", "last"), stok_degeri_raw=("stok_degeri_raw", "last"),
son90=("son90", "last"), bitis_raw=("bitis_raw", "last"), skt=("skt", "last"), tedarikci=("tedarikci", "last"), raf=("raf", "last"), raf_cm=("raf_cm", "last"), son_satis=("son_satis", "max"),
islem=("fis", pd.Series.nunique)
).reset_index()
p["kar_marji"] = np.where(p["toplam_ciro"] > 0, p["toplam_kar"] / p["toplam_ciro"], 0)
p["gunluk_satis"] = p["son90"] / 90
p["tahmini_bitis_gunu"] = np.where(p["gunluk_satis"] > 0, p["stok"] / p["gunluk_satis"], np.inf)
p["stok_degeri"] = np.where(p["stok_degeri_raw"] > 0, p["stok_degeri_raw"], p["stok"] * p["ort_alis"])
p["stok_devir"] = np.where(p["stok_degeri"] > 0, p["toplam_ciro"] / p["stok_degeri"], np.inf)
p["raf_verim_ciro"] = np.where(p["raf_cm"] > 0, p["toplam_ciro"] / p["raf_cm"], 0)
p["raf_verim_kar"] = np.where(p["raf_cm"] > 0, p["toplam_kar"] / p["raf_cm"], 0)
p["skt_kalan_gun"] = (p["skt"] - today).dt.days
p["son_satis_kac_gun"] = (today - p["son_satis"]).dt.days
return p

def classify(p: pd.DataFrame, critical_days, warning_days, dead_days):
q = p.copy()
q["hiz_sinifi"] = np.select(
[q["son90"] <= 0, q["tahmini_bitis_gunu"] <= critical_days, q["tahmini_bitis_gunu"] <= warning_days, q["tahmini_bitis_gunu"] > 90],
["Hareketsiz", "Acil", "Bu Hafta", "Yavaş"], default="Normal")
q["olu_stok_mu"] = (q["stok"] > 0) & ((q["son90"] <= 0) | (q["son_satis_kac_gun"] >= dead_days))
q["sermaye_sinifi"] = np.select([q["olu_stok_mu"], q["tahmini_bitis_gunu"] > 90, q["tahmini_bitis_gunu"] <= warning_days], ["Ölü/Yavaş", "Yavaş Dönen", "Hızlı Dönen"], default="Normal")
q["siparis_adedi"] = np.ceil((q["gunluk_satis"] * warning_days * 1.7) - q["stok"]).clip(lower=0)
q["siparis_maliyeti"] = q["siparis_adedi"] * q["ort_alis"]
q["kampanya_onerisi"] = np.select(
[q["olu_stok_mu"], (q["son_satis_kac_gun"] > 60) & (q["stok_degeri"] > 5000), q["kar_marji"] < 0.10],
["İndirim/Kampanya", "Raf azalt + kampanya", "Fiyat/maliyet kontrol"], default="Takip")
return q

def score_business(p, period_df, prev_df, current_margin, previous_margin, veresiye_total):
total = max(1, len(p))
critical = len(p[(p["gunluk_satis"] > 0) & (p["tahmini_bitis_gunu"] <= 7)]) / total
dead = len(p[p["olu_stok_mu"]]) / total
slow_value_ratio = p[p["sermaye_sinifi"].isin(["Ölü/Yavaş", "Yavaş Dönen"])] ["stok_degeri"].sum() / max(1, p["stok_degeri"].sum())
low_margin = len(p[(p["toplam_ciro"] > 0) & (p["kar_marji"] < 0.10)]) / total
c = period_df["ciro"].sum(); pr = prev_df["ciro"].sum()
sales = 0 if pr == 0 else (c-pr)/pr
score = 76 + sales22 + current_margin35 - critical25 - dead20 - slow_value_ratio22 - low_margin10 - min(.12, veresiye_total/max(c,1))*40
return int(max(0, min(100, round(score))))

def ai_advisor(period_df, prev_df, p, critical_df, dead_df, order_df, low_margin_df, veresiye_total):
messages = []
c = period_df["ciro"].sum(); pr = prev_df["ciro"].sum()
if pr > 0:
r = (c-pr)/pr
messages.append(f"Seçilen dönemde ciro önceki döneme göre {pct_fmt(abs(r))} {'arttı' if r>=0 else 'düştü'}.")
cat_now = period_df.groupby("kategori")["ciro"].sum().sort_values(ascending=False)
cat_prev = prev_df.groupby("kategori")["ciro"].sum() if not prev_df.empty else pd.Series(dtype=float)
if not cat_now.empty:
best_cat = cat_now.index[0]
messages.append(f"Ciro lideri {best_cat}; bu kategori raf ve stok odağında tutulmalı.")
profit_cat = p.groupby("kategori")["toplam_kar"].sum().sort_values(ascending=False)
if not profit_cat.empty:
messages.append(f"En çok para kazandıran ürün grubu {profit_cat.index[0]}; sadece ciroya değil bu gruba da alan açılmalı.")
if not critical_df.empty:
x = critical_df.sort_values("tahmini_bitis_gunu").iloc[0]
messages.append(f"{x['urun']} yaklaşık {num_fmt(x['tahmini_bitis_gunu'],0)} gün içinde bitebilir; sipariş önceliği yüksek.")
if not dead_df.empty:
messages.append(f"Ölü/yavaş stoklarda {money_fmt(dead_df['stok_degeri'].sum())} sermaye kilitli; kampanya ve raf azaltma aksiyonu önerilir.")
if not order_df.empty:
messages.append(f"Sipariş motoru {len(order_df)} ürün için {money_fmt(order_df['siparis_maliyeti'].sum())} yaklaşık sipariş maliyeti hesapladı.")
if not low_margin_df.empty:
low = low_margin_df.sort_values("toplam_ciro", ascending=False).iloc[0]
messages.append(f"{low['urun']} yüksek satışa rağmen düşük marjlı; alış fiyatı veya satış fiyatı kontrol edilmeli.")
if veresiye_total > 0:
messages.append(f"Tahsilat/veresiye açığı {money_fmt(veresiye_total)}; nakit akışını bozabilir.")
return " ".join(messages) if messages else "Genel tablo dengeli. Satış, stok ve kârlılık riskleri düşük görünüyor."

def excel_report(df, p, cat, order_df, dead_df, low_margin_df, cash_df):
out = BytesIO()
with pd.ExcelWriter(out, engine="openpyxl") as writer:
df.to_excel(writer, "Ham_Veri", index=False)
p.replace([np.inf, -np.inf], np.nan).to_excel(writer, "Urun_Analizi", index=False)
cat.to_excel(writer, "Kategori_Analizi", index=False)
order_df.replace([np.inf, -np.inf], np.nan).to_excel(writer, "Siparis", index=False)
dead_df.replace([np.inf, -np.inf], np.nan).to_excel(writer, "Olu_Stok", index=False)
low_margin_df.replace([np.inf, -np.inf], np.nan).to_excel(writer, "Dusuk_Marj", index=False)
cash_df.to_excel(writer, "Tahsilat", index=False)
return out.getvalue()

============================================================
Sidebar
============================================================
st.sidebar.title("🛒 AYÇA Business")
st.sidebar.caption("Market V1.0 Executive")
market_adi = st.sidebar.text_input("Market Adı", "Demo Market")
kullanici = st.sidebar.text_input("Kullanıcı", "Abdullah Bey")
uploaded = st.sidebar.file_uploader("Market Excel dosyasını yükle", type=["xlsx", "xls"])
st.sidebar.markdown("---")
period = st.sidebar.selectbox("Analiz dönemi", ["Son 7 gün", "Son 14 gün", "Son 30 gün", "Son 90 gün", "Tüm veri"], index=2)
critical_days = st.sidebar.slider("Acil stok günü", 1, 15, 5)
warning_days = st.sidebar.slider("Sipariş plan günü", 7, 45, 21)
dead_days = st.sidebar.slider("Ölü stok günü", 30, 240, 90)
skt_days = st.sidebar.slider("SKT uyarı günü", 30, 180, 90)

if uploaded is None:
st.markdown("""
AYÇA Business MarketMarket sahibinin dijital genel müdür yardımcısı.Excel bekleniyor
""", unsafe_allow_html=True)
st.info("Sol menüden örnek Excel dosyasını yükle: ayca_business_market_ornek_veri_v1.xlsx")
st.stop()

try:
raw, sheet_name = read_excel_smart(uploaded)
df = standardize(raw)
except Exception as exc:
st.error(f"Dosya okunurken hata oluştu: {exc}")
st.stop()

today = pd.Timestamp.today().normalize()
p = classify(product_table(df, today), critical_days, warning_days, dead_days)
max_date = df["tarih"].max()
if period == "Son 7 gün": start = max_date - pd.Timedelta(days=7)
elif period == "Son 14 gün": start = max_date - pd.Timedelta(days=14)
elif period == "Son 30 gün": start = max_date - pd.Timedelta(days=30)
elif period == "Son 90 gün": start = max_date - pd.Timedelta(days=90)
else: start = df["tarih"].min()
period_df = df[df["tarih"] >= start].copy()
days_len = max(1, (max_date - start).days)
prev_df = df[(df["tarih"] >= start - pd.Timedelta(days=days_len)) & (df["tarih"] < start)].copy()

ciro = period_df["ciro"].sum(); prev_ciro = prev_df["ciro"].sum()
kar = period_df["brut_kar"].sum(); prev_kar = prev_df["brut_kar"].sum()
marj = kar / ciro if ciro else 0; prev_marj = prev_kar / prev_ciro if prev_ciro else 0
fis = period_df["fis"].nunique(); avg_basket = ciro / fis if fis else 0
stock_value = p["stok_degeri"].sum()
veresiye_total = period_df["veresiye"].sum()
critical_df = p[(p["gunluk_satis"] > 0) & (p["tahmini_bitis_gunu"] <= critical_days)].copy()
order_df = p[(p["gunluk_satis"] > 0) & (p["siparis_adedi"] > 0)].copy().sort_values("siparis_maliyeti", ascending=False)
dead_df = p[p["olu_stok_mu"]].copy().sort_values("stok_degeri", ascending=False)
low_margin_df = p[(p["toplam_ciro"] > 0) & (p["kar_marji"] < 0.10)].copy().sort_values("toplam_ciro", ascending=False)
skt_df = p[p["skt_kalan_gun"].notna() & (p["skt_kalan_gun"] <= skt_days)].copy().sort_values("skt_kalan_gun")
score = score_business(p, period_df, prev_df, marj, prev_marj, veresiye_total)

Category table
cat = p.groupby("kategori").agg(ciro=("toplam_ciro","sum"), kar=("toplam_kar","sum"), stok_degeri=("stok_degeri","sum"), sku=("urun","count"), raf_cm=("raf_cm","sum"), son90=("son90","sum")).reset_index()
cat["marj"] = np.where(cat["ciro"]>0, cat["kar"]/cat["ciro"], 0)
cat["raf_kar_verimi"] = np.where(cat["raf_cm"]>0, cat["kar"]/cat["raf_cm"], 0)
cat["stok_devir"] = np.where(cat["stok_degeri"]>0, cat["ciro"]/cat["stok_degeri"], np.inf)
cat = cat.sort_values("kar", ascending=False)

cash_df = period_df.groupby("odeme").agg(ciro=("ciro","sum"), veresiye=("veresiye","sum"), islem=("fis", pd.Series.nunique)).reset_index().sort_values("ciro", ascending=False)

st.markdown(f"""

KPI
ct, cc = trend(ciro, prev_ciro); kt, kc = trend(kar, prev_kar); mt, mc = trend(marj, prev_marj)
cols = st.columns(6)
with cols[0]: metric_card("Ciro", money_fmt(ciro), period, ct, cc)
with cols[1]: metric_card("Brüt Kâr", money_fmt(kar), "Dönem kârı", kt, kc)
with cols[2]: metric_card("Kâr Marjı", pct_fmt(marj), "Brüt kar / ciro", mt, mc)
with cols[3]: metric_card("Toplam Stok Değeri", money_fmt(stock_value), "İçerideki mal değeri")
with cols[4]: metric_card("Veresiye Açığı", money_fmt(veresiye_total), "Tahsilat riski")
with cols[5]: metric_card("Ortalama Sepet", money_fmt(avg_basket), f"{fis} işlem")

Executive AI block
ai_text = ai_advisor(period_df, prev_df, p, critical_df, dead_df, order_df, low_margin_df, veresiye_total)
st.markdown(f"""

Radar + daily tasks
fast_value = p[p["sermaye_sinifi"]=="Hızlı Dönen"]["stok_degeri"].sum()
slow_value = p[p["sermaye_sinifi"].isin(["Yavaş Dönen", "Ölü/Yavaş"])] ["stok_degeri"].sum()
best_profit_cat = cat.iloc[0]["kategori"] if not cat.empty else "-"
best_margin_cat = cat.sort_values("marj", ascending=False).iloc[0]["kategori"] if not cat.empty else "-"
radar_html = f"""

st.markdown('Yönetici Görev Listesi', unsafe_allow_html=True)
tasks = []
if not critical_df.empty:
x = critical_df.sort_values("tahmini_bitis_gunu").iloc[0]
tasks.append(f"☐ {x['urun']} için sipariş aç; tahmini bitiş {num_fmt(x['tahmini_bitis_gunu'],0)} gün.")
if not dead_df.empty:
x = dead_df.iloc[0]
tasks.append(f"☐ {x['urun']} için kampanya/raf azaltma aksiyonu al; bağlı para {money_fmt(x['stok_degeri'])}.")
if not cat.empty:
tasks.append(f"☐ {best_profit_cat} kategorisine stok ve raf önceliği ver; en yüksek kâr buradan geliyor.")
if not low_margin_df.empty:
x = low_margin_df.iloc[0]
tasks.append(f"☐ {x['urun']} düşük marjlı; alış veya satış fiyatını kontrol et.")
if veresiye_total > 0:
tasks.append(f"☐ Veresiye/tahsilat listesini kontrol et; açık tutar {money_fmt(veresiye_total)}.")
while len(tasks) < 5:
tasks.append("☐ Gün sonunda satış-stok dengesini tekrar kontrol et.")
st.markdown('' + ''.join([f'{t}' for t in tasks[:5]]) + '', unsafe_allow_html=True)

Tabs
st.markdown('Detaylı Analiz Merkezleri', unsafe_allow_html=True)
tabs = st.tabs(["📊 Kategori Gücü", "💰 Kârlılık", "🔒 Sermaye Kilidi", "💀 Ölü Stok", "🛒 Sipariş", "📦 Stok", "⏳ SKT", "💳 Tahsilat", "📥 Rapor"])

with tabs[0]:
c1, c2 = st.columns([1,1])
with c1:
fig = px.bar(cat.sort_values("kar", ascending=True), x="kar", y="kategori", orientation="h", title="Kategori Bazlı Brüt Kâr")
st.plotly_chart(plot_theme(fig, 420), use_container_width=True)
with c2:
fig = px.scatter(cat, x="ciro", y="marj", size="stok_degeri", hover_name="kategori", title="Ciro / Marj / Stok Değeri Haritası")
st.plotly_chart(plot_theme(fig, 420), use_container_width=True)
st.dataframe(cat.assign(ciro=cat.ciro.round(0), kar=cat.kar.round(0), stok_degeri=cat.stok_degeri.round(0), marj=cat.marj.round(3)).replace([np.inf,-np.inf], np.nan), use_container_width=True)

with tabs[1]:
st.subheader("En kârlı ürün grupları ve düşük marj uyarıları")
top_products = p.sort_values("toplam_kar", ascending=False).head(30)
fig = px.bar(top_products.sort_values("toplam_kar", ascending=True), x="toplam_kar", y="urun", orientation="h", title="En Çok Kâr Getiren Ürünler")
st.plotly_chart(plot_theme(fig, 520), use_container_width=True)
st.markdown("#### Düşük marjlı ama satış yapan ürünler")
st.dataframe(low_margin_df[["urun","kategori","marka","toplam_ciro","toplam_kar","kar_marji","stok","stok_degeri"]].head(100), use_container_width=True)

with tabs[2]:
st.subheader("İçerideki mal değeri ve sermaye kilidi")
capital = p.groupby("sermaye_sinifi")["stok_degeri"].sum().reset_index().sort_values("stok_degeri", ascending=False)
fig = px.pie(capital, names="sermaye_sinifi", values="stok_degeri", title="Stok Sermayesi Dağılımı")
st.plotly_chart(plot_theme(fig, 420), use_container_width=True)
st.dataframe(p.sort_values("stok_degeri", ascending=False)[["urun","kategori","marka","stok","stok_degeri","son90","tahmini_bitis_gunu","sermaye_sinifi","stok_devir"]].head(150).replace([np.inf,-np.inf], np.nan), use_container_width=True)

with tabs[3]:
st.subheader("Ölü / yavaş stok merkezi")
st.warning(f"Ölü/yavaş stoklarda toplam {money_fmt(dead_df['stok_degeri'].sum())} değerinde sermaye görünüyor.")
st.dataframe(dead_df[["urun","kategori","marka","stok","stok_degeri","son90","son_satis_kac_gun","raf","raf_cm","kampanya_onerisi"]].head(200), use_container_width=True)

with tabs[4]:
st.subheader("Sipariş asistanı")
st.info(f"Toplam önerilen sipariş maliyeti: {money_fmt(order_df['siparis_maliyeti'].sum())}")
st.dataframe(order_df[["urun","kategori","marka","stok","son90","gunluk_satis","tahmini_bitis_gunu","siparis_adedi","siparis_maliyeti","tedarikci"]].head(200).replace([np.inf,-np.inf], np.nan), use_container_width=True)

with tabs[5]:
st.subheader("Stok bitiş tahmini")
stock_view = p.sort_values("tahmini_bitis_gunu")[["urun","kategori","marka","stok","son90","gunluk_satis","tahmini_bitis_gunu","stok_degeri","hiz_sinifi"]]
st.dataframe(stock_view.replace([np.inf,-np.inf], np.nan).head(250), use_container_width=True)

with tabs[6]:
st.subheader("SKT / son kullanma riski")
st.dataframe(skt_df[["urun","kategori","marka","stok","stok_degeri","skt","skt_kalan_gun","son90","kampanya_onerisi"]].head(200), use_container_width=True)

with tabs[7]:
st.subheader("Tahsilat ve ödeme tipi analizi")
fig = px.bar(cash_df, x="odeme", y="ciro", title="Ödeme Tipine Göre Ciro")
st.plotly_chart(plot_theme(fig, 380), use_container_width=True)
st.dataframe(cash_df, use_container_width=True)

with tabs[8]:
st.subheader("Excel raporu indir")
st.download_button("📥 AYÇA Business analiz raporunu indir", data=excel_report(df, p, cat, order_df, dead_df, low_margin_df, cash_df), file_name="ayca_business_market_analiz_raporu.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
st.caption("Rapor; ham veri, ürün analizi, kategori analizi, sipariş, ölü stok, düşük marj ve tahsilat sayfalarını içerir.")
