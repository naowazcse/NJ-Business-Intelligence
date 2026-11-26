import io
import re
import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st

st.set_page_config(page_title="vivo NJ MIS Portal", layout="wide")

if "auth" not in st.session_state:
  st.session_state.auth = False

if not st.session_state.auth:
  st.markdown("## vivo Narayanganj Region - MIS Portal")
  with st.form("login"):
    u = st.text_input("Username")
    p = st.text_input("Password", type="password")
    if st.form_submit_button("Login"):
      if u.strip().lower() == "naowaz.cse@gmail.com" and p == "sumoncse11":
        st.session_state.auth = True
        st.rerun()
      else:
        st.error("Invalid credentials")
  st.stop()

st.title("vivo NJ Region - 30 Reports + Virtual Shop MIS")

with st.sidebar:
  st.header("Upload Files")
  sales_up = st.file_uploader("Sales File (.xlsx)", type=["xlsx", "xls"])
  inv_up = st.file_uploader("Inventory File (.xlsx)", type=["xlsx", "xls"])
  shop_up = st.file_uploader("Shop Information (.xlsx)", type=["xlsx", "xls"])
  target_up = st.file_uploader(
      "Target File - Optional (.xlsx)", type=["xlsx", "xls"]
  )
  st.divider()
  report_mode = st.selectbox(
      "Report Mode",
      [
          "Auto",
          "Stock & Inventory",
          "Turnover Days",
          "Sales & Target",
          "Growth / Degrowth",
          "Shop Audit",
          "Custom Column",
      ],
  )
  low_vba_lim = st.number_input(
      "Low Sales VBA Threshold", min_value=1, value=20, step=1
  )
  st.divider()
  pbi_link = st.text_input("Power BI Embed URL")


def clean_df(df):
  if df is None or df.empty:
    return None
  out = df.copy()
  out.columns = [str(c).strip() for c in out.columns]
  seen, cols = {}, []
  for c in out.columns:
    seen[c] = seen.get(c, 0) + 1
    cols.append(c if seen[c] == 1 else f"{c}_{seen[c]}")
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
  return pd.to_numeric(
      pd.Series(s, dtype="object")
      .fillna(0)
      .astype(str)
      .str.replace(",", "", regex=False)
      .str.replace("%", "", regex=False)
      .str.strip(),
      errors="coerce",
  ).fillna(0.0)


def find_col(df, kws):
  if df is None or df.empty:
    return None
  for kw in kws:
    for c in df.columns:
      if kw.lower() in str(c).lower():
        return c
  return None


def drop_totals(df, col):
  if df is None or df.empty or col not in df.columns:
    return df
  bad = {"grand total", "total", "subtotal", "sub total", ""}
  return df[~ndisp(df[col]).str.lower().isin(bad)].copy()


@st.cache_data(show_spinner=False)
def read_xlsx_any(raw, sheet=None, header=0):
  bio = io.BytesIO(raw)
  try:
    return pd.read_excel(bio, sheet_name=sheet, header=header)
  except Exception:
    bio.seek(0)
    return pd.read_excel(bio, header=header)


def prep_shop(f):
  if not f:
    return None
  raw = f.getvalue()
  df = clean_df(read_xlsx_any(raw, sheet="NJ Shop Information", header=1))
  if df is None:
    df = clean_df(read_xlsx_any(raw, header=0))
  if df is None:
    return None
  sc = find_col(df, ["store code", "商店代码", "store_code"])
  if not sc:
    return None
  out = drop_totals(df, sc)
  nc = find_col(out, ["store name", "shop name", "商店名称"])
  if nc:
    out = drop_totals(out, nc)
  out["_sk"] = nkey(out[sc])
  return out[out["_sk"] != ""].drop_duplicates(subset=["_sk"], keep="first")


def detect_dim(txt):
  t = str(txt).lower()
  if any(k in t for k in ["vba", "promoter", "ba name", "ba_name"]):
    return "vba"
  if "asm" in t:
    return "asm"
  if any(k in t for k in ["se 名称", "se name", "se_name", "sales executive", "se"]):
    return "se"
  if any(k in t for k in ["store code", "商店代码", "store_code", "shop name", "store name"]):
    return "shop"
  if any(k in t for k in ["model name", "model_name", "model"]):
    return "model"
  if any(k in t for k in ["md name", "md code", "md"]):
    return "md"
  return "other"


def prep_target(f):
  if not f:
    return []
  raw = f.getvalue()
  xls = pd.ExcelFile(io.BytesIO(raw))
  recs = []
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
      num_cols = [c for c in body.columns if float(nnum(body[c]).abs().sum()) > 0]
      txt_cols = [c for c in body.columns if c not in num_cols]
      for tc in txt_cols:
        bc = drop_totals(body, tc)
        ks = nkey(bc[tc])
        if (ks != "").sum() == 0:
          continue
        dtag = detect_dim(f"{sh} {tc}")
        for nc in num_cols:
          pair = pd.DataFrame({"_k": ks, "_v": nnum(bc[nc])})
          pair = pair[pair["_k"] != ""]
          if not pair.empty:
            grp = pair.groupby("_k", as_index=False)["_v"].sum()
            recs.append({
                "sheet": sh,
                "dim": dtag,
                "key_col": tc,
                "val_col": nc,
                "map": dict(zip(grp["_k"], grp["_v"])),
            })
  return recs


def enrich_df(d, shop_df):
  if d is None:
    return None
  qc = find_col(d, ["quantity", "qty", "stock"])
  d["_qty"] = nnum(d[qc]) if qc else 1.0
  ac = find_col(d, ["sales amount", "amount", "value"])
  d["_amt"] = nnum(d[ac]) if ac else 0.0

  sc = find_col(d, ["store code", "商店代码", "store_code"])
  if sc:
    d = drop_totals(d, sc)
    d["_sk"] = nkey(d[sc])
    if shop_df is not None:
      extra = [c for c in shop_df.columns if c != "_sk" and c not in d.columns]
      if extra:
        d = pd.merge(d, shop_df[["_sk"] + extra], on="_sk", how="left")
  else:
    d["_sk"] = ""

  sh_c = find_col(d, ["store name", "shop name", "商店名称"])
  if sh_c:
    d = drop_totals(d, sh_c)
    d["_shop_disp"] = ndisp(d[sh_c]).replace("", "Unmapped")
  elif sc:
    d["_shop_disp"] = ndisp(d[sc]).replace("", "Unmapped")
  else:
    d["_shop_disp"] = "Unknown Shop"

  mc = find_col(d, ["model name", "model_name", "model"])
  vc = find_col(d, ["version", "ver", "color"])
  if mc and vc and mc != vc:
    d["_model_ver"] = (
        ndisp(d[mc]).replace("", "Unmapped")
        + " ("
        + ndisp(d[vc]).replace("", "NoVer")
        + ")"
    )
  elif mc:
    d["_model_ver"] = ndisp(d[mc]).replace("", "Unmapped")
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

  dc = find_col(d, ["sales date", "date", "billing date", "create time"])
  if dc:
    dt_p = pd.to_datetime(d[dc], errors="coerce")
    d["_dt"] = dt_p
    d["_ym"] = dt_p.dt.to_period("M").astype(str).replace("NaT", "Unknown")
    d["_yr"] = (
        dt_p.dt.year.fillna(0).astype(int).astype(str).replace("0", "Unknown")
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


def apply_filters(df, q):
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


def agg_dims(df, dims_spec):
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

  spec = {"Qty": ("_qty", "sum")}
  if "_amt" in tmp.columns and float(nnum(tmp["_amt"]).abs().sum()) > 0:
    spec["Amount"] = ("_amt", "sum")
  if "_sk" in tmp.columns:
    spec["Shops"] = ("_sk", "nunique")
  for dc in d_cols:
    spec[dc] = (dc, "first")
  return tmp.groupby(k_cols, as_index=False).agg(**spec), k_cols, d_cols


def match_target(disp_ser, recs, q, hint=""):
  if not recs:
    return pd.Series(0.0, index=disp_ser.index)
  ks = nkey(disp_ser)
  wdim = detect_dim(f"{q} {hint}")
  qwords = [
      w for w in re.findall(r"[a-z0-9\u0980-\u09ff\u4e00-\u9fff]+", q) if len(w) > 1
  ]
  scored = []
  for r in recs:
    ov = sum(1 for k in ks.unique() if k and k in r["map"])
    if ov <= 0:
      continue
    meta = f"{r['sheet']} {r['key_col']} {r['val_col']}".lower()
    hits = sum(1 for w in qwords if w in meta)
    bonus = 10 if (wdim != "other" and r["dim"] == wdim) else 0
    scored.append((bonus + hits, ov, r["map"]))
  if not scored:
    return pd.Series(0.0, index=disp_ser.index)
  scored.sort(key=lambda x: (x[0], x[1]), reverse=True)
  best = scored[0][2]
  return ks.map(lambda k: float(best.get(k, 0.0)))


def finalize(df, k_cols, d_cols, sort_col, asc=False, front_cols=None):
  out = df.copy()
  if len(d_cols) > 1:
    out["Label"] = out[d_cols].astype(str).agg(" | ".join, axis=1)
    x_c = "Label"
  else:
    x_c = d_cols[0]
  out = out.drop(columns=k_cols, errors="ignore")
  if sort_col in out.columns:
    out = out.sort_values(by=sort_col, ascending=asc)
  if front_cols:
    f = [c for c in front_cols if c in out.columns]
    r = [c for c in out.columns if c not in f]
    out = out[f + r]
  return out, x_c


def merge_stock_sales(i_df, s_df, dims_i, dims_s, days_p, mtype="turnover"):
  sg, sk, _ = agg_dims(s_df, dims_s)
  ig, ik, id_ = agg_dims(i_df, dims_i)
  sg = sg.rename(columns={"Qty": "Sales_Qty", "Amount": "Sales_Amount"})
  ig = ig.rename(columns={"Qty": "Stock_Qty"})
  m = pd.merge(ig, sg, left_on=ik, right_on=sk, how="outer", suffixes=("_i", "_s"))
  m["Stock_Qty"] = nnum(m["Stock_Qty"])
  m["Sales_Qty"] = nnum(m["Sales_Qty"])

  for dc in id_:
    ci = dc if dc in m.columns else f"{dc}_i"
    cs = f"{dc}_s"
    if ci in m.columns and cs in m.columns:
      m[dc] = ndisp(m[ci]).where(nkey(m[ci]) != "", ndisp(m[cs])).replace("", "Unmapped")
      for dc_drop in (ci, cs):
        if dc_drop != dc and dc_drop in m.columns:
          m = m.drop(columns=[dc_drop])
    elif cs in m.columns and dc not in m.columns:
      m[dc] = ndisp(m[cs]).replace("", "Unmapped")
      m = m.drop(columns=[cs], errors="ignore")

  if mtype == "turnover":
    m["Daily_Sales"] = (m["Sales_Qty"] / days_p).round(2)
    m["Turnover_Days"] = np.where(
        m["Sales_Qty"] > 0, np.round((m["Stock_Qty"] / m["Sales_Qty"]) * days_p, 1), 0.0
    )
    return finalize(
        m,
        ik + sk,
        id_,
        "Turnover_Days",
        False,
        id_ + ["Turnover_Days", "Stock_Qty", "Sales_Qty", "Daily_Sales"],
    )

  m["Stock_Gap"] = m["Stock_Qty"] - m["Sales_Qty"]
  m["Stock_Need"] = np.maximum(m["Sales_Qty"] - m["Stock_Qty"], 0.0)
  scol = "Stock_Need" if "need" in mtype else "Stock_Qty"
  return finalize(
      m,
      ik + sk,
      id_,
      scol,
      False,
      id_ + ["Stock_Qty", "Sales_Qty", "Stock_Need", "Stock_Gap"],
  )


def run_engine(q, mode_sel, s_raw, i_raw, t_recs, days_p, vba_lim):
  s_df = apply_filters(s_raw, q)
  i_df = apply_filters(i_raw, q)
  has_s = s_df is not None and not s_df.empty
  has_i = i_df is not None and not i_df.empty
  if not has_s and not has_i:
    return None, None, None, "Report", "No Data"

  sh_s = "_shop_disp" if (has_s and "_shop_disp" in s_df.columns) else find_col(s_df, ["store name", "shop name", "store code"])
  sh_i = "_shop_disp" if (has_i and "_shop_disp" in i_df.columns) else find_col(i_df, ["store name", "shop name", "store code"])
  md_s = find_col(s_df, ["md name", "md code", "md"])
  md_i = find_col(i_df, ["md name", "md code", "md"])
  se_s = find_col(s_df, ["se 名称", "se name", "se_name", "se"])
  asm_s = find_col(s_df, ["asm name", "asm"])
  vba_s = find_col(s_df, ["vba name", "vba_name", "vba", "promoter", "ba name"])
  mod_s = find_col(s_df, ["model name", "model_name", "model"])
  mod_i = find_col(i_df, ["model name", "model_name", "model"])
  area_s = find_col(s_df, ["business area", "area", "zone", "region"])
  mkt_s = find_col(s_df, ["market", "upazola", "thana", "district", "bazar"])

  is_virt = "virtual" in q or bool(re.search(r"\bvs\b", q)) or "vivo store" in q
  is_ver = "version" in q or "ver" in q
  is_mod = "model" in q
  is_turn = any(k in q for k in ["turnover", "turn over"])
  is_gr = any(k in q for k in ["growth", "degrowth", "de-growth", "two month", "2 month"])

  if mode_sel == "Stock & Inventory" and not is_turn:
    q = f"{q} stock"
  elif mode_sel == "Turnover Days":
    is_turn = True
  elif mode_sel == "Growth / Degrowth":
    is_gr = True

  if "need" in q and ("md" in q or "stock" in q) and has_s and has_i:
    res, xc = merge_stock_sales(i_df, s_df, [("MD_Name", md_i or md_s or "_sk")], [("MD_Name", md_s or md_i or "_sk")], days_p, "need")
    return res, xc, "Stock_Need", "MD Stock Need", "MD Stock Need"

  if "zero sales" in q and has_s:
    base = i_df[["_sk", sh_i]].drop_duplicates(subset=["_sk"]).rename(columns={sh_i: "Store_Name"}) if (has_i and "_sk" in i_df.columns) else s_df[["_sk", sh_s]].drop_duplicates(subset=["_sk"]).rename(columns={sh_s: "Store_Name"})
    sold = s_df.groupby("_sk", as_index=False)["_qty"].sum().rename(columns={"_qty": "Sales_Qty"})
    m = pd.merge(base, sold, on="_sk", how="left")
    m["Sales_Qty"] = nnum(m["Sales_Qty"])
    res = m[m["Sales_Qty"] <= 0].drop(columns=["_sk"], errors="ignore")
    return res, "Store_Name", "Sales_Qty", "Zero Sales Shop List", "R22: Zero Sales Shop List"

  if "zero stock" in q and has_i:
    base = s_df[["_sk", sh_s]].drop_duplicates(subset=["_sk"]).rename(columns={sh_s: "Store_Name"}) if (has_s and "_sk" in s_df.columns) else i_df[["_sk", sh_i]].drop_duplicates(subset=["_sk"]).rename(columns={sh_i: "Store_Name"})
    stk = i_df.groupby("_sk", as_index=False)["_qty"].sum().rename(columns={"_qty": "Stock_Qty"})
    m = pd.merge(base, stk, on="_sk", how="left")
    m["Stock
