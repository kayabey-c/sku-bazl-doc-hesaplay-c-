# -*- coding: utf-8 -*-
import io
import re
import unicodedata
from datetime import datetime as _dt

import numpy as np
import pandas as pd
import streamlit as st

st.set_page_config(page_title="DOC Hesaplayıcı (SKU Bazlı)", layout="wide")
st.title("📦 Days of Coverage (DOC) — SKU Bazlı")
st.caption("Excel yükleyin → *projected stock* ve *consensus (EIP)* üzerinden ürün bazlı DOC hesapları.")

# ===================== Yardımcılar / Ortak Ayarlar =====================
plant_col = "Plant"
kf_col    = "Key Figure"

def norm_text(s: str) -> str:
    s = str(s).strip()
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    s = s.lower()
    s = re.sub(r"\s+", " ", s)
    return s

KF_PATTERNS = {
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

def classify_kf(val):
    v = norm_text(val)
    for key, pats in KF_PATTERNS.items():
        for p in pats:
            if p in v:
                return key
    return None

def detect_month_columns_flexible(df: pd.DataFrame):
    """
    Datetime başlıklı kolonları veya başı 'YYYY-MM-DD' olan kolonları ay olarak yakalar.
    Dönüş: [(orijinal_kolon_adı, month_ts), ...] month_ts: Timestamp(YYYY,MM,1)
    """
    month_cols = []
    for c in df.columns:
        if isinstance(c, (pd.Timestamp, _dt)):
            ts = pd.Timestamp(c)
            month_cols.append((c, pd.Timestamp(ts.year, ts.month, 1)))
            continue
        s = str(c).strip()
        m = re.match(r"^(\d{4}[-/]\d{2}[-/]\d{2})", s)
        if m:
            ts = pd.to_datetime(m.group(1), errors="coerce")
            if pd.notna(ts):
                month_cols.append((c, pd.Timestamp(ts.year, ts.month, 1)))
    # uniq + sırala
    month_cols = list(dict.fromkeys(month_cols))
    month_cols.sort(key=lambda x: x[1])
    return month_cols

def norm_strict(s: str) -> str:
    s = str(s).strip()
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    return s.lower()

# ---- DOC yardımcıları (orijinal mantık) ----
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
    dem   = np.clip(sdf["monthly_consensus_eip"].fillna(0).astype(float).values * CONSENSUS_UNIT_MULTIPLIER, 0, None)
    docs = []
    for i in range(len(sdf)):
        future = dem[i+1:]
        docs.append(doc_days_from_stock(stock[i], future))
    sdf["DOC_days"] = docs
    return sdf

# ===================== UI — Girdi =====================
with st.sidebar:
    st.header("⚙️ Ayarlar")
    st.markdown(
        "- Excel dosyanızda ay kolonları **datetime** ya da **'YYYY-MM-DD ...' ile başlayan metin** olmalı.\n"
        "- `Key Figure` ve `Plant` alanları bekleniyor."
    )
    show_debug = st.checkbox("Gelişmiş/Debug bilgileri göster", value=False)

uploaded = st.file_uploader("Excel dosyanızı yükleyin", type=["xls", "xlsx"])
sheet_to_read = None

if uploaded is not None:
    try:
        # Çok sayfa varsa kullanıcı seçsin
        xls = pd.ExcelFile(uploaded)
        if len(xls.sheet_names) > 1:
            sheet_to_read = st.selectbox("Sayfa seçin", xls.sheet_names, index=0)
        else:
            sheet_to_read = 0
        df = pd.read_excel(xls, sheet_name=sheet_to_read)
    except Exception as e:
        st.error(f"Excel okunamadı: {e}")
        st.stop()
else:
    st.info("👆 Bir Excel dosyası yükleyin.")
    st.stop()

# ===================== İşlem =====================
# KF sınıflandır / EIP işaretle
if kf_col not in df.columns:
    st.error(f"Beklenen kolon bulunamadı: `{kf_col}`")
    st.stop()
if plant_col not in df.columns:
    st.warning(f"`{plant_col}` kolonu yok; yine de devam edilecek (Plant-1 ile birlikte EIP/GP ayıklanırsa kullanılır).")

df["_kf_class"] = df[kf_col].map(classify_kf)
df.loc[df["_kf_class"] == "consensus", plant_col] = "EIP"

# Ay kolonları
month_cols = detect_month_columns_flexible(df)
if not month_cols:
    st.error("Ay kolonları bulunamadı. Başlıklar datetime olmalı veya 'YYYY-MM-DD ...' ile başlamalı.")
    st.stop()

# Ürün adı & kod kolonları
name_candidates = ["Product (Text-TR)","Product Name","Product","Ürün","Urun",
                   "Product (Text-EN)","Description","Malzeme Açıklaması","Aciklama"]
name_col = next((c for c in name_candidates if c in df.columns), None)
if name_col is None:
    st.error("Ürün adı kolonu bulunamadı (ör. 'Product (Text-TR)')")
    st.stop()

mat_candidates = ["Bileşen","Bilesen","Malzeme","Malzeme Kodu","Material","Material Code","Component"]
mat_col = next((c for c in mat_candidates if c in df.columns), None)
if mat_col:
    df[mat_col] = df[mat_col].astype(str).str.strip()

# Normalize ürün anahtarı
df["PRODUCT_KEY"] = df[name_col].map(norm_strict)

# Plant normalizasyonu
plant_series  = df["Plant"].astype(str).fillna("")   if "Plant"   in df.columns else pd.Series("", index=df.index)
plant1_series = df["Plant-1"].astype(str).fillna("") if "Plant-1" in df.columns else pd.Series("", index=df.index)
df["_plant_norm"] = (plant_series + " " + plant1_series).str.upper()
df["_plant_norm"] = df["_plant_norm"].str.extract(r"(EIP|GP)", expand=False)

# Long form
month_names = [c for c, _ in month_cols]
col_to_ts   = dict(month_cols)

id_keep = [c for c in df.columns if c not in month_names]
for extra in ["PRODUCT_KEY", name_col, "_plant_norm"]:
    if extra not in id_keep:
        id_keep.append(extra)

df_long = df.melt(id_vars=id_keep, value_vars=month_names,
                  var_name="month_col", value_name="value")
df_long["month_ts"]  = df_long["month_col"].map(col_to_ts)
df_long["_kf_class"] = df_long[kf_col].map(classify_kf)
df_long["value"]     = pd.to_numeric(df_long["value"], errors="coerce")

# Filtreler
is_eip         = df_long["_plant_norm"].eq("EIP")
mask_consensus = (df_long["_kf_class"] == "consensus") & is_eip
mask_projected = (df_long["_kf_class"] == "projected_stock")

# Ürün × Ay toplamları
cons_name_month = (
    df_long.loc[mask_consensus]
          .dropna(subset=["PRODUCT_KEY","month_ts"])
          .groupby(["PRODUCT_KEY", name_col, "month_ts"])["value"]
          .sum()
          .rename("monthly_consensus_eip")
)
proj_name_month = (
    df_long.loc[mask_projected]
          .dropna(subset=["PRODUCT_KEY","month_ts"])
          .groupby(["PRODUCT_KEY", name_col, "month_ts"])["value"]
          .sum()
          .rename("monthly_projected_all_plants")
)

sku_df = pd.concat([proj_name_month, cons_name_month], axis=1).reset_index()
sku_df = sku_df.sort_values(["PRODUCT_KEY","month_ts"])

# Ürün bazında DOC
sku_doc_res = (
    sku_df.groupby(["PRODUCT_KEY", name_col], group_keys=False)
          .apply(compute_doc_per_product)
          .reset_index(drop=True)
)

# Toplam görünüm için çekirdek (merge öncesi)
sku_doc_res_core = sku_doc_res.copy()

# Ürün kodu ekle (çoğaltma yapmadan)
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

# Toplam görünüm (eski mantık)
total_monthly = (
    sku_doc_res_core.groupby("month_ts")[["monthly_projected_all_plants","monthly_consensus_eip"]]
                   .sum()
                   .rename(columns={
                       "monthly_projected_all_plants": "monthly_projected_eip_gp",
                       "monthly_consensus_eip": "monthly_consensus_eip"
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

# ===================== Önizleme =====================
st.subheader("✅ Ürün Bazlı DOC (Önizleme)")
show_cols = {
    "month_ts": "month",
    "monthly_projected_all_plants": "proj_stock_all_plants",
    "monthly_consensus_eip": "consensus_eip_only",
    name_col: "product_name",
}
preview = sku_doc_res.rename(columns=show_cols).copy()
st.dataframe(preview.head(50), use_container_width=True)

st.subheader("📊 Toplam Özet (Önizleme)")
st.dataframe(total_monthly.reset_index(names=["month"]).head(36), use_container_width=True)

if show_debug:
    st.divider()
    st.write("**Tespit edilen ay kolonları:**", month_names[:6], "… (toplam:", len(month_names), ")")
    st.write("**KF eşleşmeleri (örnek):**")
    st.dataframe(df[["_kf_class", kf_col]].drop_duplicates().head(20), use_container_width=True)

# ===================== İndirilebilir Çıktılar =====================
# 1) SKU bazlı dosya
prod_buffer = io.BytesIO()
with pd.ExcelWriter(prod_buffer, engine="xlsxwriter") as writer:
    out_df = sku_doc_res.rename(columns={
        "month_ts": "month",
        "monthly_projected_all_plants": "proj_stock_all_plants",
        "monthly_consensus_eip": "consensus_eip_only",
        name_col: "product_name",
    })
    cols_order = [
        "PRODUCT_KEY", "product_name", "material_code",
        "month", "proj_stock_all_plants", "consensus_eip_only", "DOC_days"
    ]
    out_df = out_df[[c for c in cols_order if c in out_df.columns]]
    out_df.to_excel(writer, index=False, sheet_name="product_monthly_doc")
prod_buffer.seek(0)

# 2) Toplam özet dosyası
sum_buffer = io.BytesIO()
total_monthly.reset_index(names=["month"]).to_excel(sum_buffer, index=False)
sum_buffer.seek(0)

col1, col2 = st.columns(2)
with col1:
    st.download_button(
        "⬇️ DOC_by_PRODUCT.xlsx",
        data=prod_buffer,
        file_name="DOC_by_PRODUCT.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
with col2:
    st.download_button(
        "⬇️ DOC_summary.xlsx",
        data=sum_buffer,
        file_name="DOC_summary.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

st.success("Bitti! Dosyaları indirebilirsiniz.")
