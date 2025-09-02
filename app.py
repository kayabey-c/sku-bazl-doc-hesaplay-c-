# app.py
# -*- coding: utf-8 -*-

import io
import re
import unicodedata
import datetime as dt

import numpy as np
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Days of Coverage (DOC) Hesaplayıcı", layout="wide")
st.title("📦 Days of Coverage (DOC) Hesaplayıcı")
st.caption("Excel yükleyin → projected stock ve consensus demand üzerinden DOC hesaplayın.")

uploaded_file = st.file_uploader("Excel'i sürükleyip bırakın", type=["xlsx"])
if uploaded_file is None:
    st.stop()

# ===== 2. BLOK (Colab ile aynı mantık) =====
df = pd.read_excel(uploaded_file)

plant_col = "Plant"
kf_col = "Key Figure"

months_columns = [c for c in df.columns if isinstance(c, (pd.Timestamp, dt.datetime))]
months_columns.sort()

def norm_text(s: str) -> str:
    s = str(s).strip()
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    s = s.lower()
    s = re.sub(r"\s+", " ", s)
    return s

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

def classify_kf(val):
    v = norm_text(val)
    for key, patterns in KF_PATTERNS.items():
        for p in patterns:
            if p in v:
                return key
    return None

df["_kf_class"] = df[kf_col].map(classify_kf)
df["_key_figure_normalized"] = df[kf_col].map(norm_text)

# ===== 3. BLOK (Colab ile aynı mantık) =====
import datetime as _dt

def norm_text_3(s: str) -> str:
    s = str(s).strip()
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    s = s.lower()
    s = re.sub(r"\s+", " ", s)
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

def detect_month_columns_flexible(df_in: pd.DataFrame):
    month_cols = []
    for c in df_in.columns:
        if isinstance(c, (pd.Timestamp, _dt.datetime)):
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

name_candidates = ["Product (Text-TR)","Product Name","Product","Ürün","Urun",
                   "Product (Text-EN)","Description","Malzeme Açıklaması","Aciklama"]
name_col = next((c for c in name_candidates if c in df.columns), None)
assert name_col is not None, "Ürün adı kolonu bulunamadı (ör. 'Product (Text-TR]')."

mat_candidates = ["Bileşen","Bilesen","Malzeme","Malzeme Kodu","Material","Material Code","Component"]
mat_col = next((c for c in mat_candidates if c in df.columns), None)
if mat_col:
    df[mat_col] = df[mat_col].astype(str).str.strip()

def norm_strict(s: str) -> str:
    s = str(s).strip()
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    return s.lower()

df["PRODUCT_KEY"] = df[name_col].map(norm_strict)
df["_kf_class"] = df[kf_col].map(classify_kf_3)

def safe_series(df_in, colname):
    if colname in df_in.columns:
        return df_in[colname].astype(str).fillna("")
    return pd.Series([""] * len(df_in), index=df_in.index)

plant_series  = safe_series(df, "Plant")
plant1_series = safe_series(df, "Plant-1")
df["_plant_norm"] = (plant_series + " " + plant1_series).str.upper()
df["_plant_norm"] = df["_plant_norm"].str.extract(r"(EIP|GP)", expand=False)

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

mask_projected = df_long["_kf_class"].eq("projected_stock")
mask_consensus_anyplant = df_long["_kf_class"].eq("consensus")

proj_name_month = (
    df_long.loc[mask_projected]
           .dropna(subset=["PRODUCT_KEY", name_col, "month_ts"])
           .groupby(["PRODUCT_KEY", name_col, "month_ts"])["value"]
           .sum()
           .rename("monthly_projected_all_plants")
)

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
    cons_name_month = (
        df_long.loc[mask_consensus_anyplant & df_long["_plant_norm"].eq("EIP")]
               .dropna(subset=["PRODUCT_KEY", name_col, "month_ts"])
               .groupby(["PRODUCT_KEY", name_col, "month_ts"])["value"]
               .sum()
               .rename("monthly_consensus_eip_only_from_eip_codes")
    )

sku_df = pd.concat([proj_name_month, cons_name_month], axis=1).reset_index()
sku_df = sku_df.sort_values(["PRODUCT_KEY", "month_ts"])

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

sku_doc_res = (
    sku_df.groupby(["PRODUCT_KEY", name_col], group_keys=False)
          .apply(compute_doc_per_product)
          .reset_index(drop=True)
)

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

# ===== Ön izleme + İndirmeler =====
st.subheader("📁 Yüklenen Veri")
tab1, tab2 = st.tabs(["Tümü", "Sadece 'consensus' & 'projected stock'"])
with tab1:
    st.dataframe(df, use_container_width=True)
with tab2:
    st.dataframe(df[df["_kf_class"].isin(["consensus", "projected_stock"])], use_container_width=True)

st.subheader("📊 DOC Sonuç Tablosu (Toplam)")
total_monthly_reset = total_monthly.reset_index(names=["month"])
st.dataframe(total_monthly_reset, use_container_width=True)

# Excel çıktıları
cols_order = [
    "PRODUCT_KEY", "product_name", "material_code",
    "month", "proj_stock_all_plants", "consensus_eip_only", "DOC_days"
]

out_df = sku_doc_res.copy().rename(columns={
    "month_ts": "month",
    "monthly_projected_all_plants": "proj_stock_all_plants",
    "monthly_consensus_eip_only_from_eip_codes": "consensus_eip_only",
    name_col: "product_name",
})
if "material_code" not in out_df.columns and 'mat_col' in globals() and mat_col in out_df.columns:
    out_df = out_df.rename(columns={mat_col: "material_code"})
out_df = out_df[[c for c in cols_order if c in out_df.columns]]

prod_buffer = io.BytesIO()
with pd.ExcelWriter(prod_buffer, engine="xlsxwriter") as writer:
    out_df.to_excel(writer, index=False, sheet_name="product_monthly_doc")
prod_buffer.seek(0)

total_buffer = io.BytesIO()
with pd.ExcelWriter(total_buffer, engine="xlsxwriter") as writer:
    total_monthly_reset.to_excel(writer, index=False, sheet_name="DOC_summary")
total_buffer.seek(0)

c1, c2 = st.columns(2)
with c1:
    st.download_button("⬇️ Excel'i indir (DOC_by_PRODUCT.xlsx)", prod_buffer,
                       file_name="DOC_by_PRODUCT.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       use_container_width=True)
with c2:
    st.download_button("⬇️ Excel'i indir (DOC_summary.xlsx)", total_buffer,
                       file_name="DOC_summary.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       use_container_width=True)

