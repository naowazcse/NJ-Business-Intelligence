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

st.title("📊 vivo NJ Region — 30 Reports + Virtual Shop & Extended MIS")

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
  report_mode = st.selectbox(
      "🎛️ রিপোর্ট মোড (Mode)",
      [
          "Auto (Command Driven)",
          "Stock & Inventory Mode",
          "Turnover Days Mode",
          "Sales & Target Mode",
          "Growth / Degrowth Mode",
          "Shop Audit (Zero / Coverage / HP)",
          "Custom / Any Column Mode",
      ],
  )
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


def drop_total_rows(df, col_name):
  if df is None or df.empty or col_name not in df.columns:
    return df
  s_low = ndisp(df[col_name]).str.lower()
  mask = ~s_low.isin(["grand total", "total", "subtotal", "sub total", ""])
  return df[mask].copy()


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
  out = drop_total_rows(out, sc)
  name_c = find_col(out, ["store name", "shop name", "商店名称"])
  if name_c:
    out = drop_total_rows(out, name_c)
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
  if any(k in tl for k in ["store code", "商店代码", "store_code", "shop name", "store name"]):
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
        body_clean = drop_total_rows(body, tc)
        keys = nkey(body_clean[tc])
        if (keys != "").sum() == 0:
          continue
        d_tag = detect_dim_tag(f"{sh} {tc}")
        for nc, _ in num_cols:
          nvals = nnum(body_clean[nc])
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
    d = drop_total_rows(d, sc)
    d["_sk"] = nkey(d[sc])
    if shop_df is not None:
      extra = [c for c in shop_df.columns if c != "_sk" and c not in d.columns]
      if extra:
        d = pd.merge(d, shop_df[["_sk"] + extra], on="_sk", how="left")
  else:
    d["_sk"] = ""

  shop_c = find_col(d, ["store name", "shop name", "商店名称", "store name"])
  if shop_c:
    d = drop_total_rows(d, shop_c)
    d["_shop_disp"] = ndisp(d[shop_c]).replace("", "Unmapped")
  elif sc:
    d["_shop_disp"] = ndisp(d[sc]).replace("", "Unmapped")
  else:
    d["_shop_disp"] = "Unknown Shop"

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

  vs_c = find_col(d, ["vs", "vivo store", "shop type", "store type"])
  if vs_c:
    d["_is_vs"] = (
        ndisp(d[vs_c])
        .str.lower()
        .str.contains(r"\bvs\b|vivo store|virtual|yes|1|true", regex=True)
    )
  else:
    d["_is_vs"] = (
        d["_shop_disp"]
        .str.lower()
        .str.contains(r"\(vs\)|\bvs\b|vivo store|virtual", regex=True)
    )

  dt_c = find_col(d, ["sales date", "date", "billing date", "create time"])
  if dt_c:
    dt_parsed = pd.to_datetime(d[dt_c], errors="coerce")
    d["_dt"] = dt_parsed
    d["_ym"] = dt_parsed.dt.to_period("M").astype(str).replace("NaT", "Unknown")
    d["_yr"] = (
        dt_parsed.dt.year.fillna(0)
        .astype(int)
        .astype(str)
        .replace("0", "Unknown")
    )
  else:
    d["_dt"] = pd.NaT
    d["_ym"] = "Current"
    d["_yr"] = "Current"

  return d


shop_master = prep_shop(shop_up) if shop_up else None
df_sales = enrich_df(
    clean_df(read_xlsx_any(sales_up.getvalue())) if sales_up else None,
    shop_master,
)
df_inv = enrich_df(
    clean_df(read_xlsx_any(inv_up.getvalue())) if inv_up else None,
    shop_master,
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


def apply_text_filters(df, q):
  if df is None or df.empty or not q:
    return df
  out = df.copy()
  for c in out.columns:
    if str(c).startswith("_"):
      continue
    if out[c].dtype == object or str(out[c].dtype).startswith("string"):
      vals = out[c].dropna().astype(str).str.strip().unique()
      if 1 < len(vals) <= 250:
        for v in vals:
          vl = v.lower().strip()
          if len(vl) >= 3 and re.search(rf"\b{re.escape(vl)}\b", q):
            out = out[out[c].astype(str).str.strip().str.lower() == vl]
            break
  return out


def agg_by_dims(df, dims_spec, q_col="_qty", a_col="_amt", sk_col="_sk"):
  tmp = df.copy()
  k_cols, d_cols = [], []
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
  if (
      a_col
      and a_col in tmp.columns
      and float(nnum(tmp[a_col]).abs().sum()) > 0
  ):
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


def make_label_and_sort(df, k_cols, d_cols, sort_col, asc=False, priority_cols=None):
  out = df.copy()
  if len(d_cols) > 1:
    out["Label"] = out[d_cols].astype(str).agg(" | ".join, axis=1)
    x_c = "Label"
  else:
    x_c = d_cols[0]
  out = out.drop(columns=k_cols, errors="ignore")
  if sort_col in out.columns:
    out = out.sort_values(by=sort_col, ascending=asc)
  if priority_cols:
    front = [c for c in priority_cols if c in out.columns]
    rest = [c for c in out.columns if c not in front]
    out = out[front + rest]
  return out, x_c


def calc_stock_sales_merge(i_df, s_df, dims_i, dims_s, days_p, mode_type="turnover"):
  sg, sk, sd = agg_by_dims(s_df, dims_s)
  ig, ik, id_ = agg_by_dims(i_df, dims_i)
  sg = sg.rename(columns={"Qty": "Sales_Qty", "Amount": "Sales_Amount"})
  ig = ig.rename(columns={"Qty": "Stock_Qty"})
  m = pd.merge(ig, sg, left_on=ik, right_on=sk, how="outer", suffixes=("_i", "_s"))
  m["Stock_Qty"] = nnum(m["Stock_Qty"])
  m["Sales_Qty"] = nnum(m["Sales_Qty"])

  for dc in id_:
    ci = dc if dc in m.columns else f"{dc}_i"
    cs = f"{dc}_s"
    if ci in m.columns and cs in m.columns:
      m[dc] = (
          ndisp(m[ci])
          .where(nkey(m[ci]) != "", ndisp(m[cs]))
          .replace("", "Unmapped")
      )
      for drop_c in (ci, cs):
        if drop_c != dc and drop_c in m.columns:
          m = m.drop(columns=[drop_c])
    elif cs in m.columns and dc not in m.columns:
      m[dc] = ndisp(m[cs]).replace("", "Unmapped")
      m = m.drop(columns=[cs], errors="ignore")

  if mode_type == "turnover":
    m["Daily_Sales"] = (m["Sales_Qty"] / days_p).round(2)
    m["Turnover_Days"] = np.where(
        m["Sales_Qty"] > 0,
        np.round((m["Stock_Qty"] / m["Sales_Qty"]) * days_p, 1),
        0.0,
    )
    return make_label_and_sort(
        m,
        ik + sk,
        id_,
        "Turnover_Days",
        False,
        priority_cols=id_ + ["Turnover_Days", "Stock_Qty", "Sales_Qty", "Daily_Sales"],
    )
  else:
    m["Stock_Gap"] = m["Stock_Qty"] - m["Sales_Qty"]
    m["Stock_Need"] = np.maximum(m["Sales_Qty"] - m["Stock_Qty"], 0.0)
    sort_c = "Stock_Need" if "need" in mode_type else "Stock_Qty"
    return make_label_and_sort(
        m,
        ik + sk,
        id_,
        sort_c,
        False,
        priority_cols=id_ + ["Stock_Qty", "Sales_Qty", "Stock_Need", "Stock_Gap"],
    )


def run_30_plus_engine(q, mode_sel, s_raw, i_raw, t_recs, days_p, low_vba_lim):
  s_df = apply_text_filters(s_raw, q)
  i_df = apply_text_filters(i_raw, q)
  has_s = s_df is not None and not s_df.empty
  has_i = i_df is not None and not i_df.empty
  if not has_s and not has_i:
    return None, None, None, "Report", "No Data"

  shop_c_s = "_shop_disp" if (has_s and "_shop_disp" in s_df.columns) else find_col(s_df, ["store name", "shop name", "商店名称", "store code", "商店代码"])
  shop_c_i = "_shop_disp" if (has_i and "_shop_disp" in i_df.columns) else find_col(i_df, ["store name", "shop name", "商店名称", "store code", "商店代码"])
  md_c_s = find_col(s_df, ["md name", "md code", "md"])
  md_c_i = find_col(i_df, ["md name", "md code", "md"])
  se_c_s = find_col(s_df, ["se 名称", "se name", "se_name", "se"])
  asm_c_s = find_col(s_df, ["asm name", "asm"])
  vba_c_s = find_col(s_df, ["vba name", "vba_name", "vba", "promoter", "ba name"])
  mod_c_s = find_col(s_df, ["model name", "model_name", "model"])
  mod_c_i = find_col(i_df, ["model name", "model_name", "model"])
  area_c_s = find_col(s_df, ["business area", "area", "zone", "region"])
  mkt_c_s = find_col(s_df, ["market", "upazola", "thana", "district", "bazar"])

  is_virtual = "virtual" in q or bool(re.search(r"\bvs\b", q)) or "vivo store" in q
  is_ver = "version" in q or "ver" in q
  is_mod = "model" in q or "মডেল" in q
  is_turn = any(k in q for k in ["turnover", "turn over", "টার্নওভার"])
  is_growth = any(k in q for k in ["growth", "degrowth", "de-growth", "two month", "2 month"])

  if mode_sel == "Stock & Inventory Mode" and not is_turn:
    q = f"{q} stock"
  elif mode_sel == "Turnover Days Mode":
    is_turn = True
  elif mode_sel == "Growth / Degrowth Mode":
    is_growth = True

  # Extra: MD Stock Need
  if "need" in q and ("md" in q or "stock" in q) and has_s and has_i:
    d_s = [("MD_Name", md_c_s or md_c_i or "_sk")]
    d_i = [("MD_Name", md_c_i or md_c_s or "_sk")]
    res, x_c = calc_stock_sales_merge(i_df, s_df, d_i, d_s, days_p, mode_type="need")
    return res, x_c, "Stock_Need", "MD Stock Need Analysis", "Extra: MD Stock Need"

  # R22: Zero Sales Shop List
  if "zero sales" in q:
    base = pd.DataFrame()
    if has_i and "_sk" in i_df.columns:
      base = i_df[["_sk", shop_c_i]].drop_duplicates(subset=["_sk"]).rename(columns={shop_c_i: "Store_Name"})
    elif has_s and "_sk" in s_df.columns:
      base = s_df[["_sk", shop_c_s]].drop_duplicates(subset=["_sk"]).rename(columns={shop_c_s: "Store_Name"})
    if has_s and not base.empty:
      sold = s_df.groupby("_sk", as_index=False)["_qty"].sum().rename(columns={"_qty": "Sales_Qty"})
      m = pd.merge(base, sold, on="_sk", how="left")
      m["Sales_Qty"] = nnum(m["Sales_Qty"])
      res = m[m["Sales_Qty"] <= 0].drop(columns=["_sk"], errors="ignore")
      return res, "Store_Name", "Sales_Qty", "Zero Sales Shop List", "R22: Zero Sales Shop List"

  # R23: Zero Stock Shop List
  if "zero stock" in q:
    base = pd.DataFrame()
    if has_s and "_sk" in s_df.columns:
      base = s_df[["_sk", shop_c_s]].drop_duplicates(subset=["_sk"]).rename(columns={shop_c_s: "Store_Name"})
    elif has_i and "_sk" in i_df.columns:
      base = i_df[["_sk", shop_c_i]].drop_duplicates(subset=["_sk"]).rename(columns={shop_c_i: "Store_Name"})
    if has_i and not base.empty:
      stk = i_df.groupby("_sk", as_index=False)["_qty"].sum().rename(columns={"_qty": "Stock_Qty"})
      m = pd.merge(base, stk, on="_sk", how="left")
      m["Stock_Qty"] = nnum(m["Stock_Qty"])
      res = m[m["Stock_Qty"] <= 0].drop(columns=["_sk"], errors="ignore")
      return res, "Store_Name", "Stock_Qty", "Zero Stock Shop List", "R23: Zero Stock Shop List"

  # R21: SE Wise Hire Purchase Sales
  if ("hire" in q or "hp" in q or "installment" in q) and has_s:
    hp_c = find_col(s_df, ["hire", "hp", "payment", "pay type", "order type", "channel"])
    sub = s_df.copy()
    if hp_c:
      mask = ndisp(sub[hp_c]).str.lower().str.contains("hire|hp|installment", regex=True)
      if mask.any():
        sub = sub[mask]
    sg, sk, sd = agg_by_dims(sub, [("SE_Name", se_c_s or shop_c_s or "_sk")])
    sg = sg.rename(columns={"Qty": "HP_Sales_Qty", "Amount": "HP_Sales_Amount"})
    res, x_c = make_label_and_sort(sg, sk, sd, "HP_Sales_Qty", False, priority_cols=sd + ["HP_Sales_Qty", "HP_Sales_Amount"])
    return res, x_c, "HP_Sales_Qty", "SE Wise Hire Purchase Sales", "R21: SE Wise Hire Purchase Sales"

  # R10: SE Wise Shop Coverage Model Version Wise
  if "coverage" in
