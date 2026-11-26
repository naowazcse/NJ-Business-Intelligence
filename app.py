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

VIRTUAL_SHOP_LIST = [
    "(VS) NJ - Jahanara Enterprise (Gopalganj)",
    "(VS) NJ - JSP Mobile Shop (Munshiganj City)",
    "(VS) NJ - Kushtia Distribution (Kushtia)",
    "(VS) NJ - M/s Khan Distribution (Faridpur)",
    "(VS) NJ - M/S Mridha traders(Gopalganj Thana 2)",
    "(VS) NJ - Maa Telecom (Madaripur)",
    "(VS) NJ - MD Mahamudul Hasan Enterprise (Rajbari 1)",
    "(VS) NJ - Mobile link (Narayanganj)",
    "(VS) NJ - Modern Telecom (Shariatpur)",
    "(VS) NJ - Pigeon International (Kushtia City)",
    "(VS) NJ - Safa Enterprise (Narayanganj County-02)",
    "(VS) NJ - Sajid Telecom Thana (Munshiganj)",
    "(VS) NJ - SK International(Gopalganj)",
    "(VS) NJ - Vai Vai Telecom (Narayanganj County-01)",
]

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


VIRTUAL_SHOP_KEYS = set(nkey(VIRTUAL_SHOP_LIST).tolist())


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
  bad = {"grand total", "total", "subtotal", "sub total", "sum", ""}
  return df[~ndisp(df[col]).str.lower().isin(bad)].copy()


@st.cache_data(show_spinner=False)
def read_smart_xlsx(raw, sheet=None):
  bio = io.BytesIO(raw)
  try:
    grid = pd.read_excel(bio, sheet_name=sheet if sheet else 0, header=None)
  except Exception:
    bio.seek(0)
    return pd.read_excel(bio)
  if grid is None or grid.empty:
    return None
  markers = (
      "store code",
      "store name",
      "shop name",
      "model name",
      "model",
      "version",
      "quantity",
      "qty",
      "stock",
      "商店代码",
      "商店名称",
  )
  best_h = 0
  best_hits = -1
  for h in range(min(15, len(grid))):
    row_txt = " ".join(str(x).lower() for x in grid.iloc[h].tolist())
    hits = sum(1 for m in markers if m in row_txt)
    if hits > best_hits:
      best_hits = hits
      best_h = h
  sub = grid.iloc[best_h:].copy()
  hdr = [
      str(x).strip()
      if str(x).strip() and not str(x).lower().startswith("unnamed")
      else f"Col_{i}"
      for i, x in enumerate(sub.iloc[0])
  ]
  body = sub.iloc[1:].reset_index(drop=True)
  if len(hdr) == len(body.columns):
    body.columns = hdr
  return clean_df(body)


def prep_shop(f):
  if not f:
    return None
  raw = f.getvalue()
  df = read_smart_xlsx(raw, sheet="NJ Shop Information")
  if df is None:
    df = read_smart_xlsx(raw)
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
  if any(
      k in t for k in ["se 名称", "se name", "se_name", "sales executive", "se"]
  ):
    return "se"
  if any(
      k in t
      for k in ["store code", "商店代码", "store_code", "shop name", "store name"]
  ):
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
      num_cols = [
          c for c in body.columns if float(nnum(body[c]).abs().sum()) > 0
      ]
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
  else:
    d["_sk"] = ""

  sh_c = find_col(d, ["store name", "shop name", "商店名称", "store_name"])
  if sh_c:
    d = drop_totals(d, sh_c)
    d["_shop_disp"] = ndisp(d[sh_c])
  else:
    d["_shop_disp"] = ""

  if shop_df is not None and "_sk" in d.columns:
    m_sh_c = find_col(shop_df, ["store name", "shop name", "商店名称"])
    if m_sh_c:
      sh_map = dict(zip(shop_df["_sk"], ndisp(shop_df[m_sh_c])))
      mapped = d["_sk"].map(lambda k: sh_map.get(k, ""))
      d["_shop_disp"] = d["_shop_disp"].where(d["_shop_disp"] != "", mapped)
    extra = [c for c in shop_df.columns if c != "_sk" and c not in d.columns]
    if extra:
      d = pd.merge(d, shop_df[["_sk"] + extra], on="_sk", how="left")

  if sc:
    d["_shop_disp"] = d["_shop_disp"].where(
        d["_shop_disp"] != "", ndisp(d[sc])
    )
  d["_shop_disp"] = d["_shop_disp"].replace("", "Unmapped")

  mc = find_col(d, ["model name", "model_name", "model"])
  vc = find_col(d, ["version", "ver", "color"])
  d["_model_only"] = ndisp(d[mc]).replace("", "Unmapped") if mc else "All Models"
  d["_ver_only"] = ndisp(d[vc]).replace("", "NoVer") if vc else "All Versions"

  if mc and vc and mc != vc:
    d["_model_ver"] = d["_model_only"] + " (" + d["_ver_only"] + ")"
  else:
    d["_model_ver"] = d["_model_only"]

  in_vs_list = nkey(d["_shop_disp"]).isin(VIRTUAL_SHOP_KEYS)
  vs_c = find_col(d, ["vs", "vivo store", "shop type", "store type"])
  if vs_c:
    vs_flag = (
        ndisp(d[vs_c])
        .str.lower()
        .str.contains(
            r"(?:^|[^a-z0-9])vs(?:[^a-z0-9]|$)|vivo store|virtual|yes|1|true",
            regex=True,
        )
    )
    d["_is_vs"] = in_vs_list | vs_flag
  else:
    vs_name_flag = (
        d["_shop_disp"]
        .str.lower()
        .str.contains(
            r"(?:^|[^a-z0-9])vs(?:[^a-z0-9]|$)|vivo store|virtual", regex=True
        )
    )
    d["_is_vs"] = in_vs_list | vs_name_flag

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
    read_smart_xlsx(sales_up.getvalue()) if sales_up else None, shop_master
)
df_inv = enrich_df(
    read_smart_xlsx(inv_up.getvalue()) if inv_up else None, shop_master
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
      w
      for w in re.findall(r"[a-z0-9\u0980-\u09ff\u4e00-\u9fff]+", q)
      if len(w) > 1
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
  m = pd.merge(
      ig, sg, left_on=ik, right_on=sk, how="outer", suffixes=("_i", "_s")
  )
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
      for dc_drop in (ci, cs):
        if dc_drop != dc and dc_drop in m.columns:
          m = m.drop(columns=[dc_drop])
    elif cs in m.columns and dc not in m.columns:
      m[dc] = ndisp(m[cs]).replace("", "Unmapped")
      m = m.drop(columns=[cs], errors="ignore")

  if mtype == "turnover":
    m["Daily_Sales"] = (m["Sales_Qty"] / days_p).round(2)
    m["Turnover_Days"] = np.where(
        m["Sales_Qty"] > 0,
        np.round((m["Stock_Qty"] / m["Sales_Qty"]) * days_p, 1),
        0.0,
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

  sh_s = (
      "_shop_disp"
      if (has_s and "_shop_disp" in s_df.columns)
      else find_col(s_df, ["store name", "shop name", "store code"])
  )
  sh_i = (
      "_shop_disp"
      if (has_i and "_shop_disp" in i_df.columns)
      else find_col(i_df, ["store name", "shop name", "store code"])
  )
  md_s = find_col(s_df, ["md name", "md code", "md"])
  md_i = find_col(i_df, ["md name", "md code", "md"])
  se_s = find_col(s_df, ["se 名称", "se name", "se_name", "se"])
  asm_s = find_col(s_df, ["asm name", "asm"])
  vba_s = find_col(
      s_df, ["vba name", "vba_name", "vba", "promoter", "ba name"]
  )
  area_s = find_col(s_df, ["business area", "area", "zone", "region"])
  mkt_s = find_col(s_df, ["market", "upazola", "thana", "district", "bazar"])

  is_virt = (
      "virtual" in q
      or bool(re.search(r"(?:^|[^a-z0-9])vs(?:[^a-z0-9]|$)", q))
      or "vivo store" in q
  )
  is_ver = "version" in q or "ver" in q
  is_mod = "model" in q
  is_turn = any(k in q for k in ["turnover", "turn over"])
  is_gr = any(
      k in q
      for k in ["growth", "degrowth", "de-growth", "two month", "2 month"]
  )

  if mode_sel == "Stock & Inventory" and not is_turn:
    q = f"{q} stock"
  elif mode_sel == "Turnover Days":
    is_turn = True
  elif mode_sel == "Growth / Degrowth":
    is_gr = True

  if "need" in q and ("md" in q or "stock" in q) and has_s and has_i:
    res, xc = merge_stock_sales(
        i_df,
        s_df,
        [("MD_Name", md_i or md_s or "_sk")],
        [("MD_Name", md_s or md_i or "_sk")],
        days_p,
        "need",
    )
    return res, xc, "Stock_Need", "MD Stock Need", "MD Stock Need"

  if "zero sales" in q and has_s:
    base = (
        i_df[["_sk", sh_i]]
        .drop_duplicates(subset=["_sk"])
        .rename(columns={sh_i: "Virtual_Shop_Name"})
        if (has_i and "_sk" in i_df.columns)
        else s_df[["_sk", sh_s]]
        .drop_duplicates(subset=["_sk"])
        .rename(columns={sh_s: "Virtual_Shop_Name"})
    )
    sold = (
        s_df.groupby("_sk", as_index=False)["_qty"]
        .sum()
        .rename(columns={"_qty": "Sales_Qty"})
    )
    m = pd.merge(base, sold, on="_sk", how="left")
    m["Sales_Qty"] = nnum(m["Sales_Qty"])
    res = m[m["Sales_Qty"] <= 0].drop(columns=["_sk"], errors="ignore")
    return (
        res,
        "Virtual_Shop_Name",
        "Sales_Qty",
        "Zero Sales Shop List",
        "R22: Zero Sales Shop List",
    )

  if "zero stock" in q and has_i:
    base = (
        s_df[["_sk", sh_s]]
        .drop_duplicates(subset=["_sk"])
        .rename(columns={sh_s: "Virtual_Shop_Name"})
        if (has_s and "_sk" in s_df.columns)
        else i_df[["_sk", sh_i]]
        .drop_duplicates(subset=["_sk"])
        .rename(columns={sh_i: "Virtual_Shop_Name"})
    )
    stk = (
        i_df.groupby("_sk", as_index=False)["_qty"]
        .sum()
        .rename(columns={"_qty": "Stock_Qty"})
    )
    m = pd.merge(base, stk, on="_sk", how="left")
    m["Stock_Qty"] = nnum(m["Stock_Qty"])
    res = m[m["Stock_Qty"] <= 0].drop(columns=["_sk"], errors="ignore")
    return (
        res,
        "Virtual_Shop_Name",
        "Stock_Qty",
        "Zero Stock Shop List",
        "R23: Zero Stock Shop List",
    )

  if any(k in q for k in ["hire", "hp", "installment"]) and has_s:
    hpc = find_col(
        s_df, ["hire", "hp", "payment", "pay type", "order type", "channel"]
    )
    sub = s_df.copy()
    if hpc:
      mk = (
          ndisp(sub[hpc])
          .str.lower()
          .str.contains("hire|hp|installment", regex=True)
      )
      if mk.any():
        sub = sub[mk]
    sg, sk, sd = agg_dims(sub, [("SE_Name", se_s or sh_s or "_sk")])
    sg = sg.rename(columns={"Qty": "HP_Sales_Qty", "Amount": "HP_Sales_Amount"})
    res, xc = finalize(
        sg, sk, sd, "HP_Sales_Qty", False, sd + ["HP_Sales_Qty", "HP_Sales_Amount"]
    )
    return (
        res,
        xc,
        "HP_Sales_Qty",
        "SE Wise Hire Purchase Sales",
        "R21: SE Wise Hire Purchase Sales",
    )

  if "coverage" in q and has_s:
    sg, sk, sd = agg_dims(
        s_df, [("SE_Name", se_s or "_sk"), ("Model_Version", "_model_ver")]
    )
    sg = sg.rename(columns={"Shops": "Covered_Shops", "Qty": "Sales_Qty"})
    res, xc = finalize(
        sg, sk, sd, "Covered_Shops", False, sd + ["Covered_Shops", "Sales_Qty"]
    )
    return (
        res,
        xc,
        "Covered_Shops",
        "SE Wise Shop Coverage",
        "R10: Shop Coverage",
    )

  if is_gr and has_s:
    if "vba" in q:
      gdim, rid, tlbl = (
          ("VBA_Name", vba_s or se_s or "_sk"),
          "R17: VBA Growth Degrowth",
          "VBA Growth Degrowth",
      )
    elif "se" in q:
      gdim, rid, tlbl = (
          ("SE_Name", se_s or "_sk"),
          "R19/R20: Two Month SE Growth Degrowth",
          "Two Month SE Growth Degrowth",
      )
    else:
      gdim, rid, tlbl = (
          ("Virtual_Shop_Name", sh_s or "_sk"),
          "R18: Shop Growth Degrowth",
          "Shop Growth Degrowth",
      )
    ms = sorted(
        [m for m in s_df["_ym"].unique() if m not in ("Unknown", "Current")]
    )
    if len(ms) >= 2:
      m1, m2 = ms[-2], ms[-1]
      g1, k1, _ = agg_dims(s_df[s_df["_ym"] == m1], [gdim])
      g2, k2, d2 = agg_dims(s_df[s_df["_ym"] == m2], [gdim])
      g1 = g1.rename(columns={"Qty": f"Qty_{m1}"})[[k1[0], f"Qty_{m1}"]]
      g2 = g2.rename(columns={"Qty": f"Qty_{m2}"})
      mdf = pd.merge(g2, g1, left_on=k2[0], right_on=k1[0], how="outer")
      mdf[f"Qty_{m1}"] = nnum(mdf[f"Qty_{m1}"])
      mdf[f"Qty_{m2}"] = nnum(mdf[f"Qty_{m2}"])
      mdf["Growth_Qty"] = mdf[f"Qty_{m2}"] - mdf[f"Qty_{m1}"]
      mdf["Growth_%"] = np.where(
          mdf[f"Qty_{m1}"] > 0,
          np.round((mdf["Growth_Qty"] / mdf[f"Qty_{m1}"]) * 100, 2),
          0.0,
      )
      mdf["Status"] = np.where(mdf["Growth_Qty"] >= 0, "Growth", "Degrowth")
      lbl = d2[0]
      mdf[lbl] = ndisp(mdf[lbl]).replace("", "Unmapped")
      res = mdf.drop(columns=[k1[0], k2[0]], errors="ignore").sort_values(
          by="Growth_Qty", ascending=False
      )
      return res, lbl, "Growth_Qty", f"{tlbl} ({m1} vs {m2})", rid
    sg, sk, sd = agg_dims(s_df, [gdim])
    res, xc = finalize(sg, sk, sd, "Qty", False, sd + ["Qty"])
    return res, xc, "Qty", tlbl, rid

  if ("year" in q or "yearly" in q) and has_s:
    sg, sk, sd = agg_dims(s_df, [("Year", "_yr")])
    res, xc = finalize(sg, sk, sd, "Qty", False, sd + ["Qty"])
    return res, xc, "Qty", "Yearly Sales Analysis", "R29: Yearly Sales Analysis"

  if is_turn and has_s and has_i:
    su = s_df[s_df["_is_vs"]] if is_virt else s_df
    iu = i_df[i_df["_is_vs"]] if is_virt else i_df
    if "md" in q and (md_s or md_i):
      ds, di, rid = (
          [("MD_Name", md_s or md_i)],
          [("MD_Name", md_i or md_s)],
          "R6/R8: MD Turn Over Days",
      )
    else:
      ds, di = (
          [("Virtual_Shop_Name", sh_s or "_sk")],
          [("Virtual_Shop_Name", sh_i or "_sk")],
      )
      rid = (
          "R8: (VS) Model Version Turn Over Days"
          if (is_virt and (is_ver or is_mod))
          else "R7: Shop Turn Over Days"
      )
    if is_ver or is_mod:
      ds.append(("Model_Version", "_model_ver"))
      di.append(("Model_Version", "_model_ver"))
    res, xc = merge_stock_sales(iu, su, di, ds, days_p, "turnover")
    return res, xc, "Turnover_Days", f"{di[0][0]} Turn Over Days", rid

  if (
      (
          ("vs" in q and ("stock" in q or "sales" in q))
          or "sales vs stock" in q
          or "stock vs sales" in q
      )
      and has_s
      and has_i
  ):
    su = s_df[s_df["_is_vs"]] if is_virt else s_df
    iu = i_df[i_df["_is_vs"]] if is_virt else i_df
    if "md" in q and (md_s or md_i):
      ds, di, rid = (
          [("MD_Name", md_s or md_i)],
          [("MD_Name", md_i or md_s)],
          "R9: MD Sales Vs Stock",
      )
    else:
      ds, di = (
          [("Virtual_Shop_Name", sh_s or "_sk")],
          [("Virtual_Shop_Name", sh_i or "_sk")],
      )
      rid = (
          "R9: Virtual/(VS) Shop Sales Vs Stock"
          if is_virt
          else "R9: Shop Sales Vs Stock"
      )
    res, xc = merge_stock_sales(iu, su, di, ds, days_p, "vs_stock")
    return res, xc, "Stock_Qty", f"{di[0][0]} Sales Vs Stock", rid

  if any(k in q for k in ["stock", "inventory", "virtual", "vs"]) and has_i:
    iu = i_df[i_df["_is_vs"]] if is_virt else i_df
    if "md" in q and not ("shop" in q or "store" in q or is_virt):
      dims, rid = [("MD_Name", md_i or "_sk")], "R1: MD Stock"
    else:
      sh_col_name = "Virtual_Shop_Name" if is_virt else "Shop_Name"
      dims = [(sh_col_name, sh_i or "_sk")]
      if is_ver and is_mod:
        dims.append(("Model_Name", "_model_only"))
        dims.append(("Version", "_ver_only"))
        rid = (
            "R5: (VS) Shop Model Version Wise Stock"
            if is_virt
            else "R5: Shop Model Version Wise Stock"
        )
      elif is_ver:
        dims.append(("Version", "_ver_only"))
        rid = (
            "R5: (VS) Shop Version Wise Stock"
            if is_virt
            else "R5: Shop Version Wise Stock"
        )
      elif is_mod:
        dims.append(("Model_Name", "_model_only"))
        rid = (
            "R4: (VS) Shop Model Wise Stock"
            if is_virt
            else "R4: Shop Model Wise Stock"
        )
      else:
        rid = "(VS) Shop Stock" if is_virt else "Shop Stock"
    ig, ik, id_ = agg_dims(iu, dims)
    ig = ig.rename(columns={"Qty": "Stock_Qty", "Amount": "Stock_Value"})
    res, xc = finalize(
        ig, ik, id_, "Stock_Qty", False, id_ + ["Stock_Qty", "Stock_Value"]
    )
    return res, xc, "Stock_Qty", f"{' + '.join(id_)} Stock", rid

  if mode_sel == "Custom Column" and has_s:
    acols = [c for c in s_df.columns if not str(c).startswith("_")]
    mcols = [c for c in acols if len(str(c)) > 1 and str(c).lower() in q]
    if mcols:
      dims = [(c, c) for c in mcols[:3]]
      sg, sk, sd = agg_dims(s_df, dims)
      sg = sg.rename(columns={"Qty": "Sales_Qty", "Amount": "Sales_Amount"})
      res, xc = finalize(
          sg, sk, sd, "Sales_Qty", False, sd + ["Sales_Qty", "Sales_Amount"]
      )
      return (
          res,
          xc,
          "Sales_Qty",
          f"Custom ({' + '.join(sd)})",
          "Custom Column",
      )

  if has_s:
    dims, hint, rid = [], "se", "R12/R25: SE Wise Sales"
    if "low" in q and "vba" in q:
      dims, hint, rid = (
          [("VBA_Name", vba_s or se_s or "_sk")],
          "vba",
          "R30: Low Sales VBA",
      )
    elif "vba" in q:
      dims, hint, rid = (
          [("VBA_Name", vba_s or se_s or "_sk")],
          "vba",
          "R14/R16/R24: VBA Sales",
      )
    elif "asm" in q:
      dims, hint, rid = (
          [("ASM_Name", asm_s or se_s or "_sk")],
          "asm",
          "R11/R14/R15: ASM Sales",
      )
    elif "area" in q:
      dims, hint, rid = (
          [("Area", area_s or mkt_s or "_sk")],
          "area",
          "R26: Area Sales",
      )
    elif "market" in q:
      dims, hint, rid = (
          [("Market", mkt_s or area_s or "_sk")],
          "market",
          "R27: Market Sales",
      )
    elif "shop" in q or "store" in q:
      dims, hint, rid = (
          [("Shop_Name", sh_s or "_sk")],
          "shop",
          "R13: Shop Sales",
      )
    elif is_mod and not ("se" in q or "asm" in q or "vba" in q):
      dims, hint, rid = (
          [("Model_Name", "_model_only")],
          "model",
          "R28: Model Sales",
      )
    else:
      dims.append(("SE_Name", se_s or "_sk"))

    if is_ver:
      dims.append(("Version", "_ver_only"))
    elif is_mod and not any("Model" in d[0] for d in dims):
      dims.append(("Model_Name", "_model_only"))

    sg, sk, sd = agg_dims(s_df, dims)
    sg = sg.rename(columns={"Qty": "Sales_Qty", "Amount": "Sales_Amount"})
    if "target" in q or "achievement" in q or t_recs:
      sg["Target_Qty"] = match_target(sg[sd[0]], t_recs, q, hint)
      sg["Achievement_%"] = np.where(
          sg["Target_Qty"] > 0,
          np.round((sg["Sales_Qty"] / sg["Target_Qty"]) * 100, 2),
          0
