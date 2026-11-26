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

st.title("📊 vivo NJ Region — Smart Report Generator")

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
  pbi_link = st.text_input("Power BI Embed URL")


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


def prep_target(file):
  """Target ফাইলের সব শিট ও সব রো থেকে যেকোনো সম্ভাব্য কী ও টার্গেট কলাম বের করে"""
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
          str(x).strip() if str(x).strip() and not str(x).lower().startswith("unnamed") else f"C{i}"
          for i, x in enumerate(sub.iloc[0])
      ]
      body = clean_df(sub.iloc[1:].reset_index(drop=True))
      if body is None or body.empty:
        continue
      # কলাম সংখ্যা না মিললে সেইফটি
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
        for nc, nvals in num_cols:
          pair = pd.DataFrame({"_k": keys, "_v": nvals})
          pair = pair[pair["_k"] != ""]
          if not pair.empty:
            grp = pair.groupby("_k", as_index=False)["_v"].sum()
            records.append({
                "sheet": sh,
                "key_col": tc,
                "val_col": nc,
                "map": dict(zip(grp["_k"], grp["_v"])),
            })
  return records


df_sales = clean_df(read_xlsx_any(sales_up.getvalue())) if sales_up else None
df_inv = clean_df(read_xlsx_any(inv_up.getvalue())) if inv_up else None
df_shop = prep_shop(shop_up) if shop_up else None
target_recs = prep_target(target_up) if target_up else []

for tag, d in [("s", df_sales), ("i", df_inv)]:
  if d is not None:
    qc = find_col(d, ["quantity", "qty", "stock"])
    d["_qty"] = nnum(d[qc]) if qc else 1.0
    ac = find_col(d, ["sales amount", "amount", "value"])
    d["_amt"] = nnum(d[ac]) if ac else 0.0
    sc = find_col(d, ["store code", "商店代码", "store_code"])
    if sc:
      d["_sk"] = nkey(d[sc])
      if df_shop is not None:
        extra = [c for c in df_shop.columns if c != "_sk" and c not in d.columns]
        if extra:
          d = pd.merge(d, df_shop[["_sk"] + extra], on="_sk", how="left")
    if tag == "s":
      df_sales = d
    else:
      df_inv = d


def get_days(d):
  if d is not None:
    dc = find_col(d, ["sales date", "date"])
    if dc:
      dt = pd.to_datetime(d[dc], errors="coerce").dropna()
      if not dt.empty:
        span = (dt.max() - dt.min()).days + 1
        if span > 0:
          return int(span)
  return 30


def pick_cols_from_query(q, dfs):
  avail = []
  for d in dfs:
    if d is not None:
      for c in d.columns:
        if not str(c).startswith("_") and c not in avail:
          avail.append(c)

  picks = []
  for c in avail:
    cl = str(c).lower().strip()
    if len(cl) > 1 and cl in q and c not in picks:
      picks.append(c)

  rules = [
      (["shop", "store", "দোকান", "শপ"], ["shop name", "商店名称", "store code", "商店代码"]),
      (["model", "মডেল", "product"], ["model name", "model"]),
      (["md", "এমডি"], ["md name", "md code"]),
      (["se", "এসই"], ["se 名称", "se name", "se_name", "se"]),
      (["asm", "এএসএম"], ["asm name", "asm"]),
      (["rsm", "আরএসএম"], ["rsm name", "rsm"]),
      (["area", "এরিয়া", "thana", "upazila"], ["business area", "upazola"]),
  ]
  for kws, c_kws in rules:
    if any(re.search(rf"\b{re.escape(k)}\b", q) or (not k.isascii() and k in q) for k in kws):
      m = next((c for kw in c_kws for c in avail if kw in str(c).lower()), None)
      if m and m not in picks:
        picks.append(m)

  if not picks:
    dflt = next(
        (
            c
            for kw in ["se 名称", "se name", "model name", "md name", "shop name"]
            for c in avail
            if kw in str(c).lower()
        ),
        avail[0] if avail else None,
    )
    if dflt:
      picks.append(dflt)
  return picks


def match_target_for_series(disp_series, recs, q):
  if not recs:
    return pd.Series(0.0, index=disp_series.index)
  keys = nkey(disp_series)
  # কমান্ডের শব্দের সাথে মিলে এমন record আগে খোঁজা
  q_words = [w for w in re.findall(r"[a-z0-9\u0980-\u09ff\u4e00-\u9fff]+", q) if len(w) > 1]
  scored = []
  for r in recs:
    meta = f"{r['sheet']} {r['key_col']} {r['val_col']}".lower()
    hits = sum(1 for w in q_words if w in meta)
    overlap = sum(1 for k in keys.unique() if k and k in r["map"])
    if overlap > 0:
      scored.append((hits, overlap, r["map"]))

  if not scored:
    return pd.Series(0.0, index=disp_series.index)

  scored.sort(key=lambda x: (x[0], x[1]), reverse=True)
  best_map = scored[0][2]
  return keys.map(lambda k: float(best_map.get(k, 0.0)))


st.markdown("### 💬 কমান্ড বক্স (যেকোনো রিপোর্ট বা কলামের নাম লিখুন)")
cmd = st.text_input(
    "কমান্ড লিখুন",
    placeholder="যেমন: SE wise target vs achievement / ASM wise sales / Shop model wise turn over days / MD stock need",
)

q = (cmd or "").strip().lower()
days_p = get_days(df_sales)
report = None
x_col, y_col, c_title = None, None, "Report"

try:
  if df_sales is not None or df_inv is not None:
    g_cols = pick_cols_from_query(q, [df_sales, df_inv])
    has_s = df_sales is not None and not df_sales.empty
    has_i = df_inv is not None and not df_inv.empty
    is_turn = any(k in q for k in ["turnover", "turn over", "টার্নওভার"])
    is_stk = any(k in q for k in ["stock", "inventory", "need", "স্টক", "ইনভেন্টরি"])

    def agg_one(df, cols, q_alias, a_alias=None, s_alias=None):
      tmp = df.copy()
      kcols, dcols = [], []
      for i, c in enumerate(cols):
        use_c = c if c in tmp.columns else next(
            (x for x in tmp.columns if not str(x).startswith("_")), None
        )
        if use_c is None:
          use_c = f"_dummy_{i}"
          tmp[use_c] = "All"
        kc = f"_k{i}"
        dc = c if c not in dcols else f"{c}_{i}"
        tmp[kc] = nkey(tmp[use_c])
        tmp[dc] = ndisp(tmp[use_c]).replace("", "Unmapped")
        kcols.append(kc)
        dcols.append(dc)
      spec = {q_alias: ("_qty", "sum")}
      if a_alias and "_amt" in tmp.columns:
        spec[a_alias] = ("_amt", "sum")
      if s_alias and "_sk" in tmp.columns:
        spec[s_alias] = ("_sk", "nunique")
      for dc in dcols:
        spec[dc] = (dc, "first")
      return tmp.groupby(kcols, as_index=False).agg(**spec), kcols, dcols

    if (is_turn or ("need" in q and has_s and has_i)) and has_s and has_i:
      sg, sk, sd = agg_one(df_sales, g_cols, "Sales_Qty", "Sales_Amount", "Active_Shops")
      ig, ik, id_ = agg_one(df_inv, g_cols, "Stock_Qty")
      m = pd.merge(ig, sg, on=sk, how="outer", suffixes=("_i", "_s"))
      m["Stock_Qty"] = nnum(m["Stock_Qty"])
      m["Sales_Qty"] = nnum(m["Sales_Qty"])
      for dc in sd:
        cs = dc if dc in m.columns else f"{dc}_s"
        ci = f"{dc}_i"
        if cs in m.columns and ci in m.columns:
          m[dc] = ndisp(m[cs]).where(nkey(m[cs]) != "", ndisp(m[ci])).replace("", "Unmapped")
          for drop_c in (cs, ci):
            if drop_c != dc and drop_c in m.columns:
              m = m.drop(columns=[drop_c])
      if is_turn:
        m["Daily_Sales"] = (m["Sales_Qty"] / days_p).round(2)
        m["Turnover_Days"] = np.where(
            m["Sales_Qty"] > 0, np.round((m["Stock_Qty"] / m["Sales_Qty"]) * days_p, 1), 0.0
        )
        sort_by, y_col, c_title = "Turnover_Days", "Turnover_Days", f"Turnover Days ({days_p}d)"
      else:
        m["Stock_Need"] = np.maximum(m["Sales_Qty"] - m["Stock_Qty"], 0.0)
        sort_by, y_col, c_title = "Stock_Need", "Stock_Need", "Stock Need"
      x_col = "Label" if len(sd) > 1 else sd[0]
      if len(sd) > 1:
        m["Label"] = m[sd].astype(str).agg(" | ".join, axis=1)
      report = m.drop(columns=sk, errors="ignore").sort_values(by=sort_by, ascending=False)

    elif is_stk and has_i:
      ig, ik, id_ = agg_one(df_inv, g_cols, "Stock_Qty")
      x_col = "Label" if len(id_) > 1 else id_[0]
      if len(id_) > 1:
        ig["Label"] = ig[id_].astype(str).agg(" | ".join, axis=1)
      y_col, c_title = "Stock_Qty", f"Stock by {' + '.join(id_)}"
      report = ig.drop(columns=ik, errors="ignore").sort_values(by="Stock_Qty", ascending=False)

    elif has_s:
      sg, sk, sd = agg_one(df_sales, g_cols, "Sales_Qty", "Sales_Amount", "Active_Shops")
      if target_recs:
        sg["Target_Qty"] = match_target_for_series(sg[sd[0]], target_recs, q)
        sg["Achievement_%"] = np.where(
            sg["Target_Qty"] > 0, np.round((sg["Sales_Qty"] / sg["Target_Qty"]) * 100, 2), 0.0
        )
      x_col = "Label" if len(sd) > 1 else sd[0]
      if len(sd) > 1:
        sg["Label"] = sg[sd].astype(str).agg(" | ".join, axis=1)
      y_col, c_title = "Sales_Qty", f"Sales by {' + '.join(sd_)}" if False else f"Sales by {' + '.join(sd)}"
      report = sg.drop(columns=sk, errors="ignore").sort_values(by="Sales_Qty", ascending=False)

    elif has_i:
      ig, ik, id_ = agg_one(df_inv, g_cols, "Stock_Qty")
      x_col = "Label" if len(id_) > 1 else id_[0]
      if len(id_) > 1:
        ig["Label"] = ig[id_].astype(str).agg(" | ".join, axis=1)
      y_col, c_title = "Stock_Qty", f"Stock by {' + '.join(id_)}"
      report = ig.drop(columns=ik, errors="ignore").sort_values(by="Stock_Qty", ascending=False)

except Exception as e:
  st.warning(f"নোট: {e}")

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
        file_name="vivo_report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
  with c2:
    st.subheader("🖼️ চার্ট")
    if x_col and y_col and x_col in report.columns and y_col in report.columns:
      fig = px.bar(report.head(25), x=x_col, y=y_col, text=y_col, title=c_title)
      st.plotly_chart(fig, use_container_width=True)

with st.expander("🔍 Target File কলাম ও শিট যাচাই (Target 0 আসলে এখান থেকে কলামের নাম দেখুন)"):
  if target_recs:
    st.write(
        [{"Sheet": r["sheet"], "Key_Col": r["key_col"], "Target_Col": r["val_col"], "Matched_Keys": len(r["map"])} for r in target_recs]
    )
  else:
    st.caption("কোনো Target ফাইল আপলোড করা হয়নি বা নিউমেরিক কলাম পাওয়া যায়নি।")

if pbi_link:
  st.divider()
  st.components.v1.iframe(pbi_link, height=600, scrolling=True)
