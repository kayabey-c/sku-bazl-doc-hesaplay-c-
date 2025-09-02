# app.py
# -*- coding: utf-8 -*-
# NOT: Sadece I/O Streamlit'e uyarlandı. Hesap mantığı ve blok içerikleri korunmuştur.

import io
import re
import unicodedata
import calendar
from datetime import datetime
import datetime as _dt

import numpy as np
import pandas as pd
import streamlit as st

# ---------------- Streamlit başlıkları (minimal) ----------------
st.set_page_config(page_title="DOC Hesaplayıcı (SKU Bazlı)", layout="wide")
st.title("📦 Days of Coverage (DOC) — SKU Bazlı")

uploaded_file = st.file_uploader("Excel dosyasını yükleyin (.xlsx)", type=["xlsx"])
run = st.button("Hesapla", type="primary", disabled=uploaded_file is None)

if not run:
    st.stop()

if uploaded_file is None:
    st.error("Lütfen bir .xlsx dosyası yükleyin.")
    st.stop()

# ==============================
# Colab'daki 2. BLOK BAŞLANGIÇ
# ==============================

# Colab: uploaded = files.upload() ... df = pd.read_excel(...)
# Streamlit eşleniği:
df = pd.read_excel(uploaded_file)

# 2.BLOK
# Plant ve Key Figure değişken ataması
plant_col = "Plant"
kf_col = "Key Figure"

import datetime as dt_for_block2

# Month columns yakalar (datetime tipindeki)
months_columns = [c for c in df.columns if isinstance(c, (pd.Timestamp, dt_for_block2.datetime))]
months_columns.sort()
print("Months columns:", months_columns[:6], "... toplam:", len(months_columns))

# Temizlenmiş Metin
def norm_text(s: str) -> str:
    s = str(s).strip()
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    s = s.lower()
    s = re.sub(r"\s+", " ", s)
    return s

# Key Figure türleri
KF_PATTERNS = {
    "consensus": [
        "kisit siz consensus",
        "consensus",
        "kisit siz consensus sell in forecast / malzeme tuketim mik",
        "kisit siz consensus forecast / malzeme tuketim mik",
        "kisit siz consensus sell in forecast / malzeme tuketim mik.",
        "kısıtsız consensus sell-in forecast / malzeme tüketim mik",
        "kısıtsız consensus sell-in forecast / malzeme tüketim mik."
    ],
    "beginning_stock": ["baslangic stok", "beginning stock"],
    "transport_receipt": ["transport receipt"],
    "recommended_order": ["recommended order"],
    "projected_stock": [
        "unconstrained projected stock",
        "projected stock",
        "unconstrainded projected stock"
    ],
    "doc": ["unconstrained days of coverage", "days of coverage"]
}

# key figure'leri tek tip kategori etiketlerine çevirmek
def classify_kf(val):
    v = norm_text(val)
    for key, patterns in KF_PATTERNS.items():
        for p in patterns:
            if p in v:
                return key
    return None

# bitmiş ürün tespitinde kullanılacak kalıplar
FG_PATTERNS = [
    "bitmis urun", "finished good", "finished goods", "fg"
]

def is_finished_good(val) -> bool:
    v = norm_text(val)
    return any(p in v for p in FG_PATTERNS)

# total/grand total gibi satırları ayıklamak için
TOTAL_TOKENS = [
    "toplam", "genel toplam", "grand total", "total", "sum", "subtotal", "ara toplam"
]

def is_total_like(val) -> bool:
    v = norm_text(val)
    return any(tok in v for tok in TOTAL_TOKENS)

# key figure kategorilerini yazdırmak
df["_kf_class"] = df[kf_col].map(classify_kf)

# ⚠️ ÖNEMLİ: consensus satırlarını EIP'e çeviren satır kaldırıldı!
# df.loc[df["_kf_class"] == "consensus", "Plant"] = "EIP"  # ← SİLİNDİ / YORUMA ALINDI

print("\nKey Figure eşleştirme sonucu (unique):")
print(df[["_kf_class", kf_col]].drop_duplicates())

# ✅ EŞLEŞME TESTİ – Gözle kontrol için
df["_key_figure_normalized"] = df[kf_col].map(norm_text)
print("\n--- İçinde 'consensus' geçen normalized satırlar ---")
print(df[df["_key_figure_normalized"].str.contains("consensus", na=False)][[kf_col, "_key_figure_normalized"]])

# İsteğe bağlı: consensus satırlarında Plant dağılımını göster (EIP/GP kontrolü)
_consensus_plant_counts = (
    df.loc[df["_kf_class"].eq("consensus")]
      .assign(_plant=df.get(plant_col, pd.Series(index=df.index)).astype(str).str.upper())
      .groupby("_plant")
      .size()
      .sort_values(ascending=False)
)
print("\nConsensus satırlarının Plant dağılımı (kontrol amaçlı):")
print(_consensus_plant_counts)

# =================================
# Colab'daki 3. BLOK BAŞLANGIÇ
# =================================
import re as re_block3, unicodedata as u_block3, datetime as _dt_block3
import pandas as pd as pd_block3
import numpy as np as np_block3

plant_col = "Plant"
kf_col    = "Key Figure"

# --- Yardımcılar ---
def norm_text_3(s: str) -> str:
    s = str(s).strip()
    s = u_block3.normalize("NFKD", s)
    s = "".join(ch for ch in s if not u_block3.combining(ch))
    s = s.lower()
    s = re_block3.sub(r"\s+", " ", s)
    return s

KF_PATTERNS_3 = {
    "consensus": [
        "kisit siz consensus","consensus",
        "kisit siz consensus sell in forecast / malzeme tuketim mik",
        "kisit siz consensus forecast / malzeme tuketim mik",
        "kisit siz consensus sell in forecast / malzeme tuketim mik.",
        "kısıtsız consensus sell-in forecast / malzeme tüketim mik",
        "kısıtsız consensus sell-in forecast / malzeme tüketim mik."
    ],
    "beginning_stock": ["baslangic stok","beginning stock"],
    "transport_receipt": ["transport receipt"],
    "recommended_order": ["recommended order"],
    "projected_stock": [
        "unconstrained projected stock","projected stock","unconstrainded projected stock"
    ],
    "doc": ["unconstrained days of coverage","days of coverage"]
}

def classify_kf_3(val):
    v = norm_text_3(val)
    for key, pats in KF_PATTERNS_3.items():
        for p in pats:
            if p in v:
                return key
    return None

# --- Ay kolonlarını esnek yakala ---
def detect_month_columns_flexible(df_in: pd.DataFrame):
    month_cols = []
    for c in df_in.columns:
        if isinstance(c, (pd.Timestamp, _dt_block3.datetime)):
            ts = pd.Timestamp(c)
            month_cols.append((c, pd.Timestamp(ts.year, ts.month, 1)))
            continue
        s = str(c).strip()
        m = re.match(r"^(\d{4}[-/]\d{2}[-/]\d{2})", s)
        if m:
            ts = pd.to_datetime(m.group(1), errors="coerce")
            if pd.notna(ts):
                month_cols.append((c, pd.Timestamp(ts.year, ts.month, 1)))
    month_cols = list(dict.fromkeys(month_cols))
    month_cols.sort(key=lambda x: x[1])
    return month_cols

month_cols = detect_month_columns_flexible(df)
assert month_cols, "Ay kolonları bulunamadı. Başlıklar datetime olmalı veya 'YYYY-MM-DD ...' ile başlamalı."

# --- Ürün adı & kod kolonları ---
name_candidates = ["Product (Text-TR)","Product Name","Product","Ürün","Urun",
                   "Product (Text-EN)","Description","Malzeme Açıklaması","Aciklama"]
name_col = next((c for c in name_candidates if c in df.columns), None)
assert name_col is not None, "Ürün adı kolonu bulunamadı (ör. 'Product (Text-TR]')."

mat_candidates = ["Bileşen","Bilesen","Malzeme","Malzeme Kodu","Material","Material Code","Component"]
mat_col = next((c for c in mat_candidates if c in df.columns), None)
if mat_col:
    df[mat_col] = df[mat_col].astype(str).str.strip()

# --- Normalize ürün key ---
def norm_strict(s: str) -> str:
    s = str(s).strip()
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    return s.lower()

df["PRODUCT_KEY"] = df[name_col].map(norm_strict)

# --- KF sınıfı ve Plant normalizasyonu (Plant'ı zorla EIP yapmıyoruz) ---
df["_kf_class"] = df[kf_col].map(classify_kf_3)

def safe_series(df_in, colname):
    if colname in df_in.columns:
        return df_in[colname].astype(str).fillna("")
    return pd.Series([""] * len(df_in), index=df_in.index)

plant_series  = safe_series(df, "Plant")
plant1_series = safe_series(df, "Plant-1")

df["_plant_norm"] = (plant_series + " " + plant1_series).str.upper()
df["_plant_norm"] = df["_plant_norm"].str.extract(r"(EIP|GP)", expand=False)

# --- Long form ---
month_names = [c for c, _ in month_cols]
col_to_ts   = dict(month_cols)

id_keep = [c for c in df.columns if c not in month_names]
for extra in ["PRODUCT_KEY", name_col, "_plant_norm"]:
    if extra not in id_keep:
        id_keep.append(extra)
if mat_col and mat_col not in id_keep:
    id_keep.append(mat_col)

df_long = df.melt(id_vars=id_keep, value_vars=month_names,
                  var_name="month_col", value_name="value")
df_long["month_ts"]  = df_long["month_col"].map(col_to_ts)
df_long["_kf_class"] = df_long[kf_col].map(classify_kf_3)
df_long["value"]     = pd.to_numeric(df_long["value"], errors="coerce")

# --- Kodun EIP'e ait olup olmadığını tespit (herhangi bir satırda EIP geçmesi yeter) ---
if mat_col:
    eip_code_flag = (
        df_long
        .dropna(subset=["PRODUCT_KEY", name_col, mat_col])
        .assign(_is_eip_code=lambda x: x["_plant_norm"].eq("EIP"))
        .groupby(["PRODUCT_KEY", name_col, mat_col], as_index=False)["_is_eip_code"]
        .max()
        .rename(columns={"_is_eip_code": "_belongs_to_EIP"})
    )
    eip_code_flag["_belongs_to_EIP"] = eip_code_flag["_belongs_to_EIP"].astype(bool)
else:
    eip_code_flag = None

# --- Filtreler ---
mask_projected = df_long["_kf_class"].eq("projected_stock")
mask_consensus_anyplant = df_long["_kf_class"].eq("consensus")

# --- Proj. stok: tüm plantlerden (EIP + GP), ürün bazında ---
proj_name_month = (
    df_long.loc[mask_projected]
           .dropna(subset=["PRODUCT_KEY", name_col, "month_ts"])
           .groupby(["PRODUCT_KEY", name_col, "month_ts"])["value"]
           .sum()
           .rename("monthly_projected_all_plants")
)

# --- Konsensus: SADECE EIP'e ait kod(lar)dan, ürün bazında ---
if mat_col:
    cons_df = (
        df_long.loc[mask_consensus_anyplant]
              .dropna(subset=["PRODUCT_KEY", name_col, mat_col, "month_ts"])
              .merge(eip_code_flag, on=["PRODUCT_KEY", name_col, mat_col], how="left")
    )
    cons_df = cons_df[cons_df["_belongs_to_EIP"] == True]
    cons_name_month = (
        cons_df.groupby(["PRODUCT_KEY", name_col, "month_ts"])["value"]
               .sum()
               .rename("monthly_consensus_eip_only_from_eip_codes")
    )
else:
    # Malzeme kodu yoksa, satır bazında Plant filtresiyle sadece EIP consensus al
    cons_name_month = (
        df_long.loc[mask_consensus_anyplant & df_long["_plant_norm"].eq("EIP")]
               .dropna(subset=["PRODUCT_KEY", name_col, "month_ts"])
               .groupby(["PRODUCT_KEY", name_col, "month_ts"])["value"]
               .sum()
               .rename("monthly_consensus_eip_only_from_eip_codes")
    )

# --- Birleşim & sıralama ---
sku_df = pd.concat([proj_name_month, cons_name_month], axis=1).reset_index()
sku_df = sku_df.sort_values(["PRODUCT_KEY", "month_ts"])

# --- DOC fonksiyonları ---
MAX_DOC_IF_NO_RUNOUT = 600
DAYS_PER_MONTH       = 30
CONSENSUS_UNIT_MULTIPLIER = 1.0

def doc_days_from_stock(stock_val, future_monthly_demand):
    if pd.isna(stock_val) or float(stock_val) <= 0:
        return 0.0
    stock_val = float(stock_val)
    cum = 0.0
    full_months = 0
    for dm in pd.Series(future_monthly_demand).fillna(0).astype(float):
        dm = max(0.0, dm)
        if dm == 0:
            full_months += 1
            continue
        if cum + dm < stock_val:
            cum += dm
            full_months += 1
        else:
            remaining = stock_val - cum
            frac = max(0.0, remaining) / dm
            return full_months * DAYS_PER_MONTH + frac * DAYS_PER_MONTH
    return MAX_DOC_IF_NO_RUNOUT

def compute_doc_per_product(sdf):
    sdf = sdf.sort_values("month_ts").copy()
    stock = sdf["monthly_projected_all_plants"].fillna(0).astype(float).values
    dem   = np.clip(
        sdf["monthly_consensus_eip_only_from_eip_codes"].fillna(0).astype(float) * CONSENSUS_UNIT_MULTIPLIER,
        0, None
    )
    docs = []
    for i in range(len(sdf)):
        future = dem[i+1:]
        docs.append(doc_days_from_stock(stock[i], future))
    sdf["DOC_days"] = docs
    return sdf

# --- Ürün adı (SKU) bazında DOC ---
sku_doc_res = (
    sku_df.groupby(["PRODUCT_KEY", name_col], group_keys=False)
          .apply(compute_doc_per_product)
          .reset_index(drop=True)
)

# --- ÇIKTI İÇİN material_code'u yine birleştir (3'lü vs. kalabilir) ---
if mat_col:
    code_map = (
        df[[mat_col, "PRODUCT_KEY", name_col]]
          .drop_duplicates()
          .groupby(["PRODUCT_KEY", name_col])[mat_col]
          .agg(lambda s: " / ".join(sorted(set(map(str, s)))))
          .reset_index()
          .rename(columns={mat_col: "material_code"})
    )
    sku_doc_res = sku_doc_res.merge(code_map, on=["PRODUCT_KEY", name_col], how="left")
else:
    sku_doc_res["material_code"] = np.nan

# --- Toplam görünüm (eski mantık) ---
sku_doc_res_core = sku_doc_res.copy()
total_monthly = (
    sku_doc_res_core.groupby("month_ts")[["monthly_projected_all_plants","monthly_consensus_eip_only_from_eip_codes"]]
                   .sum()
                   .rename(columns={
                       "monthly_projected_all_plants": "monthly_projected_eip_gp",
                       "monthly_consensus_eip_only_from_eip_codes": "monthly_consensus_eip"
                   })
                   .sort_index()
)

months      = total_monthly.index.to_list()
stock_total = total_monthly["monthly_projected_eip_gp"].reindex(months).fillna(0).astype(float)
dem_total   = total_monthly["monthly_consensus_eip"].reindex(months).fillna(0).astype(float).clip(lower=0)

doc_vals_total = []
for i in range(len(months)):
    doc_vals_total.append(doc_days_from_stock(stock_total.iloc[i], dem_total.iloc[i+1:]))
total_monthly["DOC_days"] = doc_vals_total

print("Hazır: sku_doc_res (ürün bazlı DOC, konsensus sadece EIP'e ait kodlardan) ve total_monthly (toplam).")

# =================================
# Colab'daki 4. BLOK BAŞLANGIÇ
# =================================
# Colab: Excel'e yazıp files.download(...)
# Streamlit: Bellekte Excel yazıp download_button ile indir.

# Ürün bazlı çıktı
cols_order = [
    "PRODUCT_KEY", "product_name", "material_code",
    "month", "proj_stock_all_plants", "consensus_eip_only", "DOC_days"
]

out_df = sku_doc_res.copy()

# Gerekli yeniden adlandırmalar
rename_map = {
    "month_ts": "month",
    "monthly_projected_all_plants": "proj_stock_all_plants",
    "monthly_consensus_eip_only_from_eip_codes": "consensus_eip_only",
    name_col: "product_name",
}
out_df = out_df.rename(columns=rename_map)

# material_code yoksa ve orijinal kod kolonu duruyorsa onu material_code yap
if "material_code" not in out_df.columns and 'mat_col' in globals() and mat_col in out_df.columns:
    out_df = out_df.rename(columns={mat_col: "material_code"})

# Sadece var olan kolonları yaz
out_df = out_df[[c for c in cols_order if c in out_df.columns]]

# Bellekte Excel yaz (ürün bazlı)
prod_buffer = io.BytesIO()
with pd.ExcelWriter(prod_buffer, engine="xlsxwriter") as writer:
    out_df.to_excel(writer, index=False, sheet_name="product_monthly_doc")
prod_buffer.seek(0)

# Toplam çıktı
total_buffer = io.BytesIO()
total_monthly_reset = total_monthly.reset_index(names=["month"])
with pd.ExcelWriter(total_buffer, engine="xlsxwriter") as writer:
    total_monthly_reset.to_excel(writer, index=False, sheet_name="DOC_summary")
total_buffer.seek(0)

# İndirme butonları (yalnızca gerekli iki dosya)
c1, c2 = st.columns(2)
with c1:
    st.download_button(
        label="⬇️ DOC_by_PRODUCT.xlsx",
        data=prod_buffer,
        file_name="DOC_by_PRODUCT.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
with c2:
    st.download_button(
        label="⬇️ DOC_summary.xlsx",
        data=total_buffer,
        file_name="DOC_summary.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
