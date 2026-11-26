import io
import re
import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st

st.set_page_config(page_title="vivo NJ MIS Portal", layout="wide")

# --- লগইন সিস্টেম ---
if "auth" not in st.session_state:
  st.session_state.auth = False

if not st.session_state.auth:
  st.markdown("## 🔐 vivo Narayanganj Region - MIS Portal")
  with st.form("login"):
    user = st.text_input("Username")
    pwd = st.text_input("Password", type="password")
    if st.form_submit_button("Login"):
      if user.strip().lower() == "naowaz.cse@gmail.com" and pwd == "sumoncse11":
        st.session_state.auth = True
        st.rerun()
      else:
        st.error("Username অথবা Password সঠিক নয়।")
  st.stop()

st.title("📊 vivo NJ Region — Complete 30-Report MIS Generator")

with st.sidebar:
  st.header("📁 ফাইল আপলোড")
  sales_up = st.file_uploader("Sales File (.xlsx)", type=["xlsx", "xls"])
  inv_up = st.file_uploader("Inventory File (.xlsx)", type=["xlsx", "xls"])
  shop_up = st.file_uploader(
      "Shop Information (.xlsx)", type=["xlsx", "xls"]
  )
  target_up = st.file_uploader(
      "Target File - Optional (.xlsx)", type=["xlsx", "xls"]
  )
  st.divider()
  low_vba_threshold = st.number_input(
      "Low Sales VBA Threshold (Qty)", min_value=1, value=20, step=1
  )
  st.divider()
  pbi_link = st.text_input("Power BI Embed URL")


# --- হেল্পার ফাংশন ---
def clean_df(df):
  if df is None or df.empty:
    return None
  out = df.copy()
  out.columns = [str(c).strip() for c in out.columns]
  seen = {}
  cols = []
  for c in out.columns:
    if c in seen:
      seen[c] += 1
      cols.append(f"{c}_{seen[c]}")
    else:
      seen[c] = 0
      cols.append(c)
  out.columns = cols
  return out


def nkey(s):
  return (
      pd.Series(s, dtype="object")
      .fillna("")
      .astype(str)
      .str.replace(r"\.0$", "", regex=True)
      .str.strip()
      .str.lower()
      .str.replace(r"[^a-z0-9\u0980-\u09ff\u4e00-\u9fff]+", "", regex=True)
  )


def ndisp(s):
  return (
      pd.Series(s, dtype="object")
      .fillna("")
      .astype(str)
      .str.replace(r"\.0$", "", regex=True)
      .str.strip()
  )


def nnum(s):
  cleaned = (
      pd.Series(s, dtype="object")
      .fillna(0)
      .astype(str)
      .str.replace(",", "", regex=False)
      .str.replace("%", "", regex=False)
      .str.strip()
  )
  return pd.to_numeric(cleaned, errors="coerce").fillna(0.0)


def find_col(df, kws):
  if df is None or df.empty:
    return None
  for kw in kws:
    for c in df.columns:
      if kw.lower() in str(c).lower():
        return c
  return None


@st.cache_data(show_spinner=False)
def read_xlsx_any(raw_bytes, sheet=None, header=0):
  bio = io.BytesIO(raw_bytes)
  try:
    if sheet is not None:
      return pd.read_excel(bio, sheet_name=sheet, header=header)
    return pd.read_excel(bio, header=header)
  except Exception:
    bio.seek(0)
    return pd.read_excel(bio, header=header)


def prep_shop(file):
  if not file:
    return None
  raw = file.getvalue()
  df = clean_df(read_xlsx_any(raw, sheet="NJ Shop Information", header=1))
  if df is None:
    df = clean_df(read_xlsx_any(raw, header=0))
  if df is None:
    return None
  sc = find_col(df, ["store code", "商店代码", "store_code"])
  if not sc:
    return None
  out = df.copy()
  out["_sk"] = nkey(out[sc])
  out = out[out["_sk"] != ""].drop_duplicates(subset=["_sk"], keep="first")
  return out


def detect_dim_tag(text):
  tl = str(text).lower()
  if any(k in tl for k in ["vba", "ভিবিএ", "promoter", "ba name", "ba_name"]):
    return "vba"
  if any(k in tl for k in ["asm", "এএসএম"]):
    return "asm"
  if any(
      k in tl
      for k in ["se 名称", "se name", "se_name", "sales executive", "se"]
  ):
    return "se"
  if any(k in tl for k in ["store code", "商店代码", "store_code", "shop name"]):
    return "shop"
  if any(k in tl for k in ["model name", "model_name", "model"]):
    return "model"
  if any(k in tl for k in ["md name", "md code", "md"]):
    return "md"
  return "other"


def prep_target(file):
  if not file:
    return []
  raw = file.getvalue()
  xls = pd.ExcelFile(io.BytesIO(raw))
  records = []
  for sh in xls.sheet_names:
    try:
      grid = pd.read_excel(io.BytesIO(raw), sheet_name=sh, header=None)
    except Exception:
      continue
    if grid is None or grid.empty:
      continue
    for h in range(min(20, len(grid))):
      sub = grid.iloc[h:].copy()
      hdr = [
          str(x).strip()
          if str(x).strip() and not str(x).lower().startswith("unnamed")
          else f"C{i}"
          for i, x in enumerate(sub.iloc[0])
      ]
      body = clean_df(sub.iloc[1:].reset_index(drop=True))
      if body is None or body.empty:
        continue
      if len(hdr) == len(body.columns):
        body.columns = hdr
        body = clean_df(body)

      num_cols = []
      for c in body.columns:
        vals = nnum(body[c])
        if float(vals.abs().sum()) > 0:
          num_cols.append((c, vals))

      text_cols = [
          c for c in body.columns if all(c != nc for nc, _ in num_cols)
      ]
      for tc in text_cols:
        keys = nkey(body[tc])
        if (keys != "").sum() == 0:
          continue
        d_tag = detect_dim_tag(f"{sh} {tc}")
        for nc, nvals in num_cols:
          pair = pd.DataFrame({"_k": keys, "_v": nvals})
          pair = pair[pair["_k"] != ""]
          if not pair.empty:
            grp = pair.groupby("_k", as_index=False)["_v"].sum()
            records.append({
                "sheet": sh,
                "dim": d_tag,
                "key_col": tc,
                "val_col": nc,
                "map": dict(zip(grp["_k"], grp["_v"])),
            })
  return records


def enrich_df(d, shop_df):
  if d is None:
    return None
  qc = find_col(d, ["quantity", "qty", "stock"])
  d["_qty"] = nnum(d[qc]) if qc else 1.0
  ac = find_col(d, ["sales amount", "amount", "value"])
  d["_amt"] = nnum(d[ac]) if ac else 0.0

  sc = find_col(d, ["store code", "商店代码", "store_code"])
  if sc:
    d["_sk"] = nkey(d[sc])
    if shop_df is not None:
      extra = [c for c in shop_df.columns if c != "_sk" and c not in d.columns]
      if extra:
        d = pd.merge(d, shop_df[["_sk"] + extra], on="_sk", how="left")
  else:
    d["_sk"] = ""

  # Model ও Version কলাম সমন্বয়
  m_col = find_col(d, ["model name", "model_name", "model"])
  v_col = find_col(d, ["version", "ver", "color"])
  if m_col and v_col and m_col != v_col:
    d["_model_ver"] = (
        ndisp(d[m_col]).replace("", "Unmapped")
        + " ("
        + ndisp(d[v_col]).replace("", "NoVer")
        + ")"
    )
  elif m_col:
    d["_model_ver"] = ndisp(d[m_col]).replace("", "Unmapped")
  else:
    d["_model_ver"] = "All Models"

  # VS (vivo Store) শপ শনাক্তকরণ
  shop_c = find_col(d, ["shop name", "商店名称", "store name", "channel", "type"])
  vs_c = find_col(d, ["vs", "vivo store", "shop type", "store type"])
  if vs_c:
    d["_is_vs"] = (
        ndisp(d[vs_c])
        .str.lower()
        .str.contains(r"\bvs\b|vivo store|yes|1|true", regex=True)
    )
  elif shop_c:
    d["_is_vs"] = (
        ndisp(d[shop_c])
        .str.lower()
        .str.contains(r"\(vs\)|\bvs\b|vivo store", regex=True)
    )
  else:
    d["_is_vs"] = False

  # Date / Month / Year কলাম
  dt_c = find_col(d, ["sales date", "date", "billing date", "create time"])
  if dt_c:
    dt_parsed = pd.to_datetime(d[dt_c], errors="coerce")
    d["_dt"] = dt_parsed
    d["_ym"] = dt_parsed.dt.to_period("M").astype(str).replace("NaT", "Unknown")
    d["_yr"] = dt_parsed.dt.year.fillna(0).astype(int).astype(str).replace("0", "Unknown")
  else:
    d["_dt"] = pd.NaT
    d["_ym"] = "Current"
    d["_yr"] = "Current"

  return d


df_sales = enrich_df(
    clean_df(read_xlsx_any(sales_up.getvalue())) if sales_up else None,
    prep_shop(shop_up) if shop_up else None,
)
df_inv = enrich_df(
    clean_df(read_xlsx_any(inv_up.getvalue())) if inv_up else None,
    prep_shop(shop_up) if shop_up else None,
)
target_recs = prep_target(target_up) if target_up else []


def get_days(d):
  if d is not None and "_dt" in d.columns:
    dt = d["_dt"].dropna()
    if not dt.empty:
      span = (dt.max() - dt.min()).days + 1
      if span > 0:
        return int(span)
  return 30


def col_lookup(df, kws):
  return find_col(df, kws)


def agg_by_dims(df, dims_spec, q_col="_qty", a_col="_amt", sk_col="_sk"):
  """dims_spec: list of (display_name, series_or_colname)"""
  tmp = df.copy()
  k_cols = []
  d_cols = []
  for idx, (d_name, src) in enumerate(dims_spec):
    ser = tmp[src] if isinstance(src, str) and src in tmp.columns else src
    kc = f"_dk_{idx}"
    dc = d_name if d_name not in d_cols else f"{d_name}_{idx}"
    tmp[kc] = nkey(ser)
    tmp[dc] = ndisp(ser).replace("", "Unmapped")
    k_cols.append(kc)
    d_cols.append(dc)

  spec = {}
  if q_col and q_col in tmp.columns:
    spec["Qty"] = (q_col, "sum")
  if a_col and a_col in tmp.columns and float(nnum(tmp[a_col]).abs().sum()) > 0:
    spec["Amount"] = (a_col, "sum")
  if sk_col and sk_col in tmp.columns:
    spec["Shops"] = (sk_col, "nunique")
  for dc in d_cols:
    spec[dc] = (dc, "first")

  grp = tmp.groupby(k_cols, as_index=False).agg(**spec)
  return grp, k_cols, d_cols


def match_target_series(disp_series, recs, q, dim_hint=""):
  if not recs:
    return pd.Series(0.0, index=disp_series.index)
  keys = nkey(disp_series)
  wanted_dim = detect_dim_tag(f"{q} {dim_hint}")
  q_words = [
      w
      for w in re.findall(r"[a-z0-9\u0980-\u09ff\u4e00-\u9fff]+", q)
      if len(w) > 1
  ]
  scored = []
  for r in recs:
    overlap = sum(1 for k in keys.unique() if k and k in r["map"])
    if overlap <= 0:
      continue
    meta = f"{r['sheet']} {r['key_col']} {r['val_col']}".lower()
    hits = sum(1 for w in q_words if w in meta)
    dim_bonus = 10 if (wanted_dim != "other" and r["dim"] == wanted_dim) else 0
    scored.append((dim_bonus + hits, overlap, r["map"]))
  if not scored:
    return pd.Series(0.0, index=disp_series.index)
  scored.sort(key=lambda x: (x[0], x[1]), reverse=True)
  best = scored[0][2]
  return keys.map(lambda k: float(best.get(k, 0.0)))


def build_30_report(q, s_df, i_df, t_recs, days_p, low_vba_lim):
  has_s = s_df is not None and not s_df.empty
  has_i = i_df is not None and not i_df.empty
  if not has_s and not has_i:
    return None, None, None, "Report"

  shop_c_s = col_lookup(s_df, ["shop name", "商店名称", "store code", "商店代码"])
  shop_c_i = col_lookup(i_df, ["shop name", "商店名称", "store code", "商店代码"])
  md_c_s = col_lookup(s_df, ["md name", "md code", "md"])
  md_c_i = col_lookup(i_df, ["md name", "md code", "md"])
  se_c_s = col_lookup(s_df, ["se 名称", "se name", "se_name", "se"])
  asm_c_s = col_lookup(s_df, ["asm name", "asm"])
  vba_c_s = col_lookup(s_df, ["vba name", "vba_name", "vba", "promoter", "ba name"])
  mod_c_s = col_lookup(s_df, ["model name", "model_name", "model"])
  mod_c_i = col_lookup(i_df, ["model name", "model_name", "model"])
  area_c_s = col_lookup(s_df, ["business area", "area", "zone", "region"])
  mkt_c_s = col_lookup(s_df, ["market", "upazola", "thana", "district", "bazar"])

  is_vs_q = bool(re.search(r"\bvs\b", q)) or "vivo store" in q
  is_ver_q = "version" in q or "ver" in q
  is_model_q = "model" in q or "মডেল" in q
  is_turn_q = any(k in q for k in ["turnover", "turn over", "টার্নওভার"])
  is_growth_q = any(k in q for k in ["growth", "degrowth", "de-growth", "two month", "2 month"])

  # --- ২২. Zero Sales Shop List ---
  if "zero sales" in q and ("shop" in q or "store" in q):
    base_shops = pd.DataFrame()
    if has_i and "_sk" in i_df.columns:
      base_shops = i_df[["_sk"] + ([shop_c_i] if shop_c_i else [])].drop_duplicates(subset=["_sk"])
      base_shops = base_shops.rename(columns={shop_c_i: "Shop_Name"}) if shop_c_i else base_shops
    elif has_s and "_sk" in s_df.columns:
      base_shops = s_df[["_sk"] + ([shop_c_s] if shop_c_s else [])].drop_duplicates(subset=["_sk"])
      base_shops = base_shops.rename(columns={shop_c_s: "Shop_Name"}) if shop_c_s else base_shops
    if has_s and not base_shops.empty:
      sold_map = s_df.groupby("_sk", as_index=False)["_qty"].sum().rename(columns={"_qty": "Sales_Qty"})
      m = pd.merge(base_shops, sold_map, on="_sk", how="left")
      m["Sales_Qty"] = nnum(m["Sales_Qty"])
      zero_s = m[m["Sales_Qty"] <= 0].drop(columns=["_sk"], errors="ignore")
      lbl = "Shop_Name" if "Shop_Name" in zero_s.columns else zero_s.columns[0]
      return zero_s, lbl, "Sales_Qty", "Zero Sales Shop List"

  # --- ২৩. Zero Stock Shop List ---
  if "zero stock" in q and ("shop" in q or "store" in q):
    base_shops = pd.DataFrame()
    if has_s and "_sk" in s_df.columns:
      base_shops = s_df[["_sk"] + ([shop_c_s] if shop_c_s else [])].drop_duplicates(subset=["_sk"])
      base_shops = base_shops.rename(columns={shop_c_s: "Shop_Name"}) if shop_c_s else base_shops
    elif has_i and "_sk" in i_df.columns:
      base_shops = i_df[["_sk"] + ([shop_c_i] if shop_c_i else [])].drop_duplicates(subset=["_sk"])
      base_shops = base_shops.rename(columns={shop_c_i: "Shop_Name"}) if shop_c_i else base_shops
    if has_i and not base_shops.empty:
      stk_map = i_df.groupby("_sk", as_index=False)["_qty"].sum().rename(columns={"_qty": "Stock_Qty"})
      m = pd.merge(base_shops, stk_map, on="_sk", how="left")
      m["Stock_Qty"] = nnum(m["Stock_Qty"])
      zero_i = m[m["Stock_Qty"] <= 0].drop(columns=["_sk"], errors="ignore")
      lbl = "Shop_Name" if "Shop_Name" in zero_i.columns else zero_i.columns[0]
      return zero_i, lbl, "Stock_Qty", "Zero Stock Shop List"

  # --- ২১. SE Wise Hire Purchase Sales ---
  if "hire" in q or "hp" in q or "installment" in q:
    if has_s:
      hp_col = find_col(s_df, ["hire", "hp", "payment", "pay type", "order type", "channel"])
      sub_s = s_df.copy()
      if hp_col:
        mask = ndisp(sub_s[hp_col]).str.lower().str.contains("hire|hp|installment", regex=True)
        if mask.any():
          sub_s = sub_s[mask]
      g_col = se_c_s or shop_c_s or "_sk"
      sg, sk, sd = agg_by_dims(sub_s, [(g_col, g_col)])
      rep = sg.drop(columns=sk, errors="ignore").sort_values(by="Qty", ascending=False)
      return rep, sd[0], "Qty", "SE Wise Hire Purchase Sales"

  # --- ১০. SE Wise Shop Coverage Model Version Wise ---
  if "coverage" in q and "se" in q:
    if has_s:
      dims = []
      if se_c_s:
        dims.append(("SE_Name", se_c_s))
      if is_ver_q or is_model_q:
        dims.append(("Model_Version", "_model_ver"))
      if not dims:
        dims = [("SE_Name", se_c_s or "_sk"), ("Model_Version", "_model_ver")]
      sg, sk, sd = agg_by_dims(s_df, dims)
      sg = sg.rename(columns={"Shops": "Covered_Shops", "Qty": "Sales_Qty"})
      if len(sd) > 1:
        sg["Label"] = sg[sd].astype(str).agg(" | ".join, axis=1)
        x_c = "Label"
      else:
        x_c = sd[0]
      rep = sg.drop(columns=sk, errors="ignore").sort_values(by="Covered_Shops", ascending=False)
      return rep, x_c, "Covered_Shops", "SE Wise Shop Coverage (Model Version Wise)"

  # --- ১৭, ১৮, ১৯, ২০. Growth / Degrowth / Two Month ---
  if is_growth_q and has_s:
    if "vba" in q and vba_c_s:
      g_dim = ("VBA_Name", vba_c_s)
      title = "VBA Sales Growth / Degrowth"
    elif "se" in q and se_c_s:
      g_dim = ("SE_Name", se_c_s)
      title = "Two Month SE Wise Sales Growth / Degrowth" if "two month" in q or "2 month" in q else "SE Wise Sales Growth / Degrowth"
    else:
      g_dim = ("Shop_Name", shop_c_s or "_sk")
      title = "Shop Wise Sales Growth / Degrowth"

    months = sorted([m for m in s_df["_ym"].unique() if m not in ("Unknown", "Current")])
    if len(months) >= 2:
      m1, m2 = months[-2], months[-1]
      d1 = s_df[s_df["_ym"] == m1]
      d2 = s_df[s_df["_ym"] == m2]
      g1, k1, _ = agg_by_dims(d1, [g_dim])
      g2, k2, d2_cols = agg_by_dims(d2, [g_dim])
      g1 = g1.rename(columns={"Qty": f"Qty_{m1}"})[[k1[0], f"Qty_{m1}"]]
      g2 = g2.rename(columns={"Qty": f"Qty_{m2}"})
      m_df = pd.merge(g2, g1, left_on=k2[0], right_on=k1[0], how="outer")
      m_df[f"Qty_{m1}"] = nnum(m_df[f"Qty_{m1}"])
      m_df[f"Qty_{m2}"] = nnum(m_df[f"Qty_{m2}"])
      m_df["Growth_Qty"] = m_df[f"Qty_{m2}"] - m_df[f"Qty_{m1}"]
      m_df["Growth_%"] = np.where(
          m_df[f"Qty_{m1}"] > 0,
          np.round((m_df["Growth_Qty"] / m_df[f"Qty_{m1}"]) * 100, 2),
          0.0,
      )
      m_df["Status"] = np.where(m_df["Growth_Qty"] >= 0, "Growth", "Degrowth")
      lbl = d2_cols[0]
      m_df[lbl] = ndisp(m_df[lbl]).replace("", "Unmapped")
      rep = m_df.drop(columns=[k1[0], k2[0]], errors="ignore").sort_values(by="Growth_Qty", ascending=False)
      return rep, lbl, "Growth_Qty", f"{title} ({m1} vs {m2})"
    else:
      sg, sk, sd = agg_by_dims(s_df, [g_dim])
      rep = sg.drop(columns=sk, errors="ignore").sort_values(by="Qty", ascending=False)
      return rep, sd[0], "Qty", title

  # --- ২৯. Yearly Sales Analysis ---
  if ("year" in q or "yearly" in q) and has_s:
    dims = [("Year", "_yr")]
    if "model" in q and mod_c_s:
      dims.append(("Model", mod_c_s))
    elif "se" in q and se_c_s:
      dims.append(("SE", se_c_s))
    elif "shop" in q and shop_c_s:
      dims.append(("Shop", shop_c_s))
    sg, sk, sd = agg_by_dims(s_df, dims)
    x_c = "Label" if len(sd) > 1 else sd[0]
    if len(sd) > 1:
      sg["Label"] = sg[sd].astype(str).agg(" | ".join, axis=1)
    rep = sg.drop(columns=sk, errors="ignore").sort_values(by="Qty", ascending=False)
    return rep, x_c, "Qty", "Yearly Sales Analysis"

  # --- ৬, ৭, ৮. Turnover Days (MD / Shop / VS Shop / Model Version Wise) ---
  if is_turn_q and has_s and has_i:
    s_use = s_df[s_df["_is_vs"]] if is_vs_q and s_df["_is_vs"].any() else s_df
    i_use = i_df[i_df["_is_vs"]] if is_vs_q and i_df["_is_vs"].any() else i_df

    if "md" in q and (md_c_s or md_c_i):
      dims_s = [("MD_Name", md_c_s or md_c_i)]
      dims_i = [("MD_Name", md_c_i or md_c_s)]
      if is_ver_q or is_model_q:
        dims_s.append(("Model_Version", "_model_ver"))
        dims_i.append(("Model_Version", "_model_ver"))
    else:
      dims_s = [("Shop_Name", shop_c_s or "_sk")]
      dims_i = [("Shop_Name", shop_c_i or "_sk")]
      if is_ver_q or is_model_q:
        dims_s.append(("Model_Version", "_model_ver"))
        dims_i.append(("Model_Version", "_model_ver"))

    sg, sk, sd = agg_by_dims(s_use, dims_s)
    ig, ik, id_ = agg_by_dims(i_use, dims_i)
    sg = sg.rename(columns={"Qty": "Sales_Qty"})
    ig = ig.rename(columns={"Qty": "Stock_Qty"})
    m = pd.merge(ig, sg, left_on=ik, right_on=sk, how="outer", suffixes=("_i", "_s"))
    m["Stock_Qty"] = nnum(m["Stock_Qty"])
    m["Sales_Qty"] = nnum(m["Sales_Qty"])
    for dc in id_:
      ci = dc if dc in m.columns else f"{dc}_i"
      cs = f"{dc}_s"
      if ci in m.columns and cs in m.columns:
        m[dc] = ndisp(m[ci]).where(nkey(m[ci]) != "", ndisp(m[cs])).replace("", "Unmapped")
        for drop_c in (ci, cs):
          if drop_c != dc and drop_c in m.columns:
            m = m.drop(columns=[drop_c])
    m["Daily_Sales"] = (m["Sales_Qty"] / days_p).round(2)
    m["Turnover_Days"] = np.where(
        m["Sales_Qty"] > 0, np.round((m["Stock_Qty"] / m["Sales_Qty"]) * days_p, 1), 0.0
    )
    x_c = "Label" if len(id_) > 1 else id_[0]
    if len(id_) > 1:
      m["Label"] = m[id_].astype(str).agg(" | ".join, axis=1)
    rep = m.drop(columns=ik + sk, errors="ignore").sort_values(by="Turnover_Days", ascending=False)
    prefix = "(VS) " if is_vs_q else ""
    return rep, x_c, "Turnover_Days", f"{prefix}{' + '.join(id_)} Turn Over Days ({days_p}d)"

  # --- ৯. Sales Vs Stock Analysis (MD / VS Shop / Shop) ---
  if ("vs" in q and ("stock" in q or "sales" in q)) or ("sales vs stock" in q) or ("stock vs sales" in q):
    if has_s and has_i:
      s_use = s_df[s_df["_is_vs"]] if is_vs_q and s_df["_is_vs"].any() else s_df
      i_use = i_df[i_df["_is_vs"]] if is_vs_q and i_df["_is_vs"].any() else i_df
      if "md" in q and (md_c_s or md_c_i):
        dims_s = [("MD_Name", md_c_s or md_c_i)]
        dims_i = [("MD_Name", md_c_i or md_c_s)]
      else:
        dims_s = [("Shop_Name", shop_c_s or "_sk")]
        dims_i = [("Shop_Name", shop_c_i or "_sk")]
      sg, sk, sd = agg_by_dims(s_use, dims_s)
      ig, ik, id_ = agg_by_dims(i_use, dims_i)
      sg = sg.rename(columns={"Qty": "Sales_Qty", "Amount": "Sales_Amount"})
      ig = ig.rename(columns={"Qty": "Stock_Qty"})
      m = pd.merge(ig, sg, left_on=ik, right_on=sk, how="outer", suffixes=("_i", "_s"))
      m["Stock_Qty"] = nnum(m["Stock_Qty"])
      m["Sales_Qty"] = nnum(m["Sales_Qty"])
      m["Stock_Gap"] = m["Stock_Qty"] - m["Sales_Qty"]
      m["Stock_Need"] = np.maximum(m["Sales_Qty"] - m["Stock_Qty"], 0.0)
      lbl = id_[0]
      if lbl not in m.columns:
        lbl = f"{lbl}_i" if f"{lbl}_i" in m.columns else f"{lbl}_s"
      rep = m.drop(columns=ik + sk, errors="ignore").sort_values(by="Stock_Qty", ascending=False)
      return rep, lbl, "Stock_Qty", f"{'(VS) ' if is_vs_q else ''}{lbl} Sales Vs Stock Analysis"

  # --- ১, ২, ৩, ৪, ৫. Stock Reports (MD Stock, (VS) Shop Stock, Shop Stock, Model Wise Shop Stock, Model Version Wise (VS) Shop Stock) ---
  if any(k in q for k in ["stock", "inventory", "স্টক", "ইনভেন্টরি"]) and has_i:
    i_use = i_df[i_df["_is_vs"]] if is_vs_q and i_df["_is_vs"].any() else i_df
    dims = []
    if "md" in q and md_c_i:
      dims.append(("MD_Name", md_c_i))
    elif "shop" in q or "store" in q or is_vs_q or not is_model_q:
      dims.append(("Shop_Name", shop_c_i or "_sk"))

    if is_ver_q:
      dims.append(("Model_Version", "_model_ver"))
    elif is_model_q and mod_c_i:
      dims.append(("Model_Name", mod_c_i))

    if not dims:
      dims = [("Shop_Name", shop_c_i or "_sk")]

    ig, ik, id_ = agg_by_dims(i_use, dims)
    ig = ig.rename(columns={"Qty": "Stock_Qty"})
    x_c = "Label" if len(id_) > 1 else id_[0]
    if len(id_) > 1:
      ig["Label"] = ig[id_].astype(str).agg(" | ".join, axis=1)
    rep = ig.drop(columns=ik, errors="ignore").sort_values(by="Stock_Qty", ascending=False)
    prefix = "(VS) " if is_vs_q else ""
    return rep, x_c, "Stock_Qty", f"{prefix}{' + '.join(id_)} Stock"

  # --- ১১ থেকে ১৬, ২৪ থেকে ২৮, ৩০. Sales & Target vs Achievement Reports ---
  if has_s:
    s_use = s_df.copy()
    dims = []
    dim_hint = q

    if "vba" in q and vba_c_s:
      dims.append(("VBA_Name", vba_c_s))
      dim_hint = "vba"
    elif "asm" in q and asm_c_s:
      dims.append(("ASM_Name", asm_c_s))
      dim_hint = "asm"
    elif "se" in q and se_c_s:
      dims.append(("SE_Name", se_c_s))
      dim_hint = "se"
    elif "area" in q and area_c_s:
      dims.append(("Area", area_c_s))
      dim_hint = "area"
    elif ("market" in q or "bazar" in q or "thana" in q or "upazila" in q) and mkt_c_s:
      dims.append(("Market", mkt_c_s))
      dim_hint = "market"
    elif ("shop" in q or "store" in q) and shop_c_s:
      dims.append(("Shop_Name", shop_c_s))
      dim_hint = "shop"
    elif is_model_q and mod_c_s:
      dims.append(("Model_Name", mod_c_s))
      dim_hint = "model"

    if is_ver_q:
      dims.append(("Model_Version", "_model_ver"))
    elif is_model_q and mod_c_s and not any("Model" in d[0] for d in dims):
      dims.append(("Model_Name", mod_c_s))

    if not dims:
      dims = [("SE_Name", se_c_s or shop_c_s or mod_c_s or "_sk")]

    sg, sk, sd = agg_by_dims(s_use, dims)
    sg = sg.rename(columns={"Qty": "Sales_Qty", "Amount": "Sales_Amount"})

    if "target" in q or "achievement" in q or t_recs:
      sg["Target_Qty"] = match_target_series(sg[sd[0]], t_recs, q, dim_hint)
      sg["Achievement_%"] = np.where(
          sg["Target_Qty"] > 0,
          np.round((sg["Sales_Qty"] / sg["Target_Qty"]) * 100, 2),
          0.0,
      )

    if "low" in q and "vba" in q:
      sg = sg[sg["Sales_Qty"] <= low_vba_lim]

    x_c = "Label" if len(sd) > 1 else sd[0]
    if len(sd) > 1:
      sg["Label"] = sg[sd].astype(str).agg(" | ".join, axis=1)
    rep = sg.drop(columns=sk, errors="ignore").sort_values(by="Sales_Qty", ascending=False)
    return rep, x_c, "Sales_Qty", f"{' + '.join(sd)} Sales Report"

  return None, None, None, "Report"


# --- কমান্ড বক্স ---
st.markdown("### 💬 কমান্ড বক্স (৩০টি রিপোর্টের যেকোনোটি লিখুন)")
cmd = st.text_input(
    "কমান্ড লিখুন",
    placeholder=(
        "যেমন: MD Stock / (VS) Shop Stock / Model Version Wise (VS) Shop Stock"
        " / Shop Turn Over Days / SE Wise Shop Coverage Model Version Wise /"
        " ASM, SE and VBA Sales Target Vs Achievement / Zero Sales Shop List"
    ),
)

q = (cmd or "").strip().lower()
days_p = get_days(df_sales)
report, x_col, y_col, c_title = None, None, None, "Report"

try:
  if q:
    report, x_col, y_col, c_title = build_30_report(
        q, df_sales, df_inv, target_recs, days_p, low_vba_threshold
    )
  elif df_sales is not None or df_inv is not None:
    report, x_col, y_col, c_title = build_30_report(
        "se wise sales", df_sales, df_inv, target_recs, days_p, low_vba_threshold
    )
except Exception as e:
  st.warning(f"রিপোর্ট নোট: {e}")

if report is not None and not report.empty:
  c1, c2 = st.columns(2)
  with c1:
    st.subheader("📋 রিপোর্ট টেবিল")
    st.dataframe(report, use_container_width=True)
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
      report.to_excel(w, index=False, sheet_name="Report")
    st.download_button(
        "⬇️ Download Excel (.xlsx)",
        data=buf.getvalue(),
        file_name="vivo_nj_report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
  with c2:
    st.subheader("🖼️ চার্ট")
    if x_col and y_col and x_col in report.columns and y_col in report.columns:
      fig = px.bar(report.head(25), x=x_col, y=y_col, text=y_col, title=c_title)
      st.plotly_chart(fig, use_container_width=True)

with st.expander("🔍 Target File ও কলাম ম্যাপিং যাচাই"):
  if target_recs:
    st.write([
        {
            "Sheet": r["sheet"],
            "Dim": r["dim"],
            "Key_Col": r["key_col"],
            "Target_Col": r["val_col"],
            "Keys": len(r["map"]),
        }
        for r in target_recs
    ])
  else:
    st.caption("Target ফাইল আপলোড করলে শিট ও কলামগুলো এখানে দেখাবে।")

if pbi_link:
  st.divider()
  st.components.v1.iframe(pbi_link, height=600, scrolling=True)
