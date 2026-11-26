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
    ok = st.form_submit_button("Login")
    if ok:
      if user.strip().lower() == "naowaz.cse@gmail.com" and pwd == "sumoncse11":
        st.session_state.auth = True
        st.rerun()
      else:
        st.error("Username অথবা Password সঠিক নয়।")
  st.stop()

st.title("📊 vivo NJ Region — Dynamic Excel Report Generator")

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
  st.header("📈 Power BI Embed")
  pbi_link = st.text_input("Power BI Embed URL")


def clean_cols(df):
  if df is None:
    return None
  df = df.copy()
  df.columns = [str(c).strip() for c in df.columns]
  seen = {}
  out = []
  for c in df.columns:
    if c in seen:
      seen[c] += 1
      out.append(f"{c}_{seen[c]}")
    else:
      seen[c] = 0
      out.append(c)
  df.columns = out
  return df


def norm_key(series):
  return (
      series.fillna("")
      .astype(str)
      .str.replace(r"\.0$", "", regex=True)
      .str.strip()
      .str.lower()
  )


def norm_disp(series):
  return (
      series.fillna("")
      .astype(str)
      .str.replace(r"\.0$", "", regex=True)
      .str.strip()
  )


def to_num(series):
  return pd.to_numeric(series, errors="coerce").fillna(0.0)


def find_col(df, keywords):
  if df is None or df.empty:
    return None
  for kw in keywords:
    for c in df.columns:
      if kw.lower() in str(c).lower():
        return c
  return None


@st.cache_data(show_spinner=False)
def read_excel_safe(file_bytes, sheet_name=None, header=0):
  bio = io.BytesIO(file_bytes)
  if sheet_name is not None:
    try:
      return pd.read_excel(bio, sheet_name=sheet_name, header=header)
    except Exception:
      bio.seek(0)
      return pd.read_excel(bio, header=header)
  return pd.read_excel(bio, header=header)


def load_shop_df(file):
  if file is None:
    return None
  raw = file.getvalue()
  df = clean_cols(
      read_excel_safe(raw, sheet_name="NJ Shop Information", header=1)
  )
  store_col = find_col(df, ["store code", "商店代码", "store_code"])
  if not store_col:
    return None
  keep = [store_col]
  for c in df.columns:
    if c not in keep:
      keep.append(c)
  out = df[keep].copy()
  out["_store_key"] = norm_key(out[store_col])
  out[store_col] = norm_disp(out[store_col])
  return out[out["_store_key"] != ""].drop_duplicates(
      subset=["_store_key"], keep="first"
  )


def load_target_df(file):
  if file is None:
    return None
  raw = file.getvalue()
  xls = pd.ExcelFile(io.BytesIO(raw))
  for sh in xls.sheet_names:
    tmp = clean_cols(pd.read_excel(io.BytesIO(raw), sheet_name=sh))
    if tmp is None or tmp.empty:
      continue
    if all(str(c).startswith("Unnamed") for c in tmp.columns) and len(tmp) > 0:
      tmp.columns = [str(x).strip() for x in tmp.iloc[0]]
      tmp = clean_cols(tmp.iloc[1:].reset_index(drop=True))
    key_col = find_col(
        tmp,
        [
            "se 名称",
            "se name",
            "se_name",
            "sales executive",
            "store code",
            "商店代码",
            "md name",
            "model",
            "name",
        ],
    )
    tgt_col = find_col(
        tmp, ["target", "টার্গেট", "tgt", "achievement", "qty"]
    )
    if key_col and tgt_col:
      out = pd.DataFrame({
          "_tgt_key": norm_key(tmp[key_col]),
          "Target_Qty": to_num(tmp[tgt_col]),
      })
      out = out[out["_tgt_key"] != ""]
      return out.groupby("_tgt_key", as_index=False)["Target_Qty"].sum()
  return None


df_sales = (
    clean_cols(read_excel_safe(sales_up.getvalue())) if sales_up else None
)
df_inv = clean_cols(read_excel_safe(inv_up.getvalue())) if inv_up else None
df_shop = load_shop_df(shop_up) if shop_up else None
df_target = load_target_df(target_up) if target_up else None

# Shop তথ্য Sales এবং Inventory-তে যুক্ত করা
for tag, d in [("sales", df_sales), ("inv", df_inv)]:
  if d is not None:
    q_col = find_col(d, ["quantity", "qty", "stock"])
    d["_qty"] = to_num(d[q_col]) if q_col else 1.0
    a_col = find_col(d, ["sales amount", "amount", "value", "revenue"])
    d["_amt"] = to_num(d[a_col]) if a_col else 0.0
    s_col = find_col(d, ["store code", "商店代码", "store_code"])
    if s_col:
      d["_store_key"] = norm_key(d[s_col])
      if df_shop is not None:
        add_cols = [
            c
            for c in df_shop.columns
            if c != "_store_key" and c not in d.columns
        ]
        if add_cols:
          d = pd.merge(
              d,
              df_shop[["_store_key"] + add_cols],
              on="_store_key",
              how="left",
          )
    if tag == "sales":
      df_sales = d
    else:
      df_inv = d


def get_days(df_s):
  if df_s is not None:
    dt_col = find_col(df_s, ["sales date", "date"])
    if dt_col:
      dts = pd.to_datetime(df_s[dt_col], errors="coerce").dropna()
      if not dts.empty:
        span = (dts.max() - dts.min()).days + 1
        if span > 0:
          return int(span)
  return 30


def extract_group_cols(q, df_list):
  """কমান্ড থেকে এক বা একাধিক গ্রুপ কলাম খুঁজে বের করে (যেমন shop + model, md, se, asm ইত্যাদি)"""
  pool = []
  for d in df_list:
    if d is not None:
      for c in d.columns:
        if not str(c).startswith("_") and c not in pool:
          pool.append(c)

  alias_map = [
      (["shop", "store", "দোকান", "শপ"], ["shop name", "商店名称", "store code", "商店代码"]),
      (["model", "মডেল", "product", "পণ্য"], ["model name", "model"]),
      (["md", "এমডি", "distributor"], ["md name", "md code"]),
      (["se", "এসই", "sales executive"], ["se 名称", "se name", "se_name", "se"]),
      (["asm", "এএসএম"], ["asm name", "asm"]),
      (["rsm", "আরএসএম"], ["rsm name", "rsm"]),
      (["area", "এরিয়া", "thana", "upazila"], ["business area", "upazola"]),
  ]

  found = []
  # ১. সরাসরি কলামের নাম কমান্ডে থাকলে
  for c in pool:
    cl = str(c).lower().strip()
    if len(cl) > 1 and cl in q and c not in found:
      found.append(c)

  # ২. Alias শব্দ কমান্ডে থাকলে
  for kws, col_kws in alias_map:
    if any(re.search(rf"\b{re.escape(k)}\b", q) or (not k.isascii() and k in q) for k in kws):
      for kw in col_kws:
        matched = next((c for c in pool if kw in str(c).lower()), None)
        if matched and matched not in found:
          found.append(matched)
          break

  return found


def apply_filters(df, q):
  """কমান্ডে কোনো কলামের মান উল্লেখ থাকলে সে অনুযায়ী ফিল্টার করে"""
  if df is None or df.empty or not q:
    return df
  out = df.copy()
  for c in out.columns:
    if str(c).startswith("_"):
      continue
    if out[c].dtype == object or str(out[c].dtype).startswith("string"):
      vals = out[c].dropna().astype(str).str.strip().unique()
      # খুব বেশি ইউনিক ভ্যালু থাকলে স্ক্যান বাদ দিয়ে গতি ঠিক রাখা
      if 1 < len(vals) <= 300:
        for v in vals:
          vl = v.lower().strip()
          if len(vl) >= 3 and re.search(rf"\b{re.escape(vl)}\b", q):
            out = out[out[c].astype(str).str.strip().str.lower() == vl]
            break
  return out


def top_n_from_cmd(q, default_n=None):
  m = re.search(r"\b(?:top|first)\s+(\d+)\b", q)
  if m:
    return int(m.group(1))
  return default_n


def build_dynamic_report(q, df_s, df_i, df_t, days_p):
  s_sub = apply_filters(df_s, q)
  i_sub = apply_filters(df_i, q)

  grp_cols = extract_group_cols(q, [s_sub, i_sub])
  has_sales = s_sub is not None and not s_sub.empty
  has_inv = i_sub is not None and not i_sub.empty

  want_turnover = any(
      k in q for k in ["turnover", "turn over", "turnaround", "টার্নওভার"]
  )
  want_stock = any(
      k in q for k in ["stock", "inventory", "need", "স্টক", "ইনভেন্টরি"]
  )
  want_target = any(
      k in q for k in ["target", "achievement", "টার্গেট", "অ্যাচিভমেন্ট"]
  )

  # কমান্ডে গ্রুপিং কলাম না বলা থাকলে ডিফল্ট কলাম নির্ধারণ
  if not grp_cols:
    if want_stock and not has_sales and has_inv:
      fallback = find_col(
          i_sub, ["model name", "se 名称", "se name", "md name", "shop name"]
      )
    elif has_sales:
      fallback = find_col(
          s_sub, ["se 名称", "se name", "model name", "md name", "shop name"]
      )
    elif has_inv:
      fallback = find_col(
          i_sub, ["model name", "se 名称", "se name", "md name", "shop name"]
      )
    else:
      fallback = None
    if fallback:
      grp_cols = [fallback]

  if not grp_cols:
    return None, None, None, "Report"

  # যে কলামগুলো গ্রুপিং হবে সেগুলোর নরম্যালাইজড কী তৈরি
  def prep_grouped(df, g_cols, qty_alias, amt_alias=None, shop_alias=None):
    tmp = df.copy()
    key_cols = []
    disp_cols = []
    for idx, gc in enumerate(g_cols):
      if gc not in tmp.columns:
        # একই অর্থের কলাম অন্য নামে থাকলে খুঁজে নেওয়া
        alt = None
        gl = str(gc).lower()
        if "shop" in gl or "store" in gl or "商店" in gl:
          alt = find_col(tmp, ["shop name", "store code", "商店代码"])
        elif "model" in gl:
          alt = find_col(tmp, ["model name", "model"])
        elif "md" in gl:
          alt = find_col(tmp, ["md name", "md code"])
        elif "se" in gl:
          alt = find_col(tmp, ["se 名称", "se name", "se"])
        if alt:
          gc_use = alt
        else:
          gc_use = f"_g_{idx}"
          tmp[gc_use] = "All"
      else:
        gc_use = gc

      k_name = f"_gk_{idx}"
      d_name = gc if gc not in key_cols else f"{gc}_{idx}"
      tmp[k_name] = norm_key(tmp[gc_use])
      tmp[d_name] = norm_disp(tmp[gc_use]).replace("", "Unmapped")
      key_cols.append(k_name)
      disp_cols.append(d_name)

    agg_dict = {qty_alias: ("_qty", "sum")}
    if amt_alias and "_amt" in tmp.columns:
      agg_dict[amt_alias] = ("_amt", "sum")
    if shop_alias and "_store_key" in tmp.columns:
      agg_dict[shop_alias] = ("_store_key", "nunique")

    for d_name in disp_cols:
      agg_dict[d_name] = (d_name, "first")

    grouped = tmp.groupby(key_cols, as_index=False).agg(**agg_dict)
    return grouped, key_cols, disp_cols

  # Sales ও Inventory একসাথে দরকার এমন কমান্ড (Turnover / Stock Need / উভয় ডেটা থাকলে)
  need_both = want_turnover or ("need" in q) or (has_sales and has_inv and want_stock and not want_target)

  if need_both and has_sales and has_inv:
    s_grp, s_keys, s_disps = prep_grouped(
        s_sub, grp_cols, "Sales_Qty", "Sales_Amount", "Active_Shops"
    )
    i_grp, i_keys, i_disps = prep_grouped(i_sub, grp_cols, "Stock_Qty", None, None)

    merged = pd.merge(i_grp, s_grp, on=s_keys, how="outer", suffixes=("_inv", "_sls"))
    merged["Stock_Qty"] = to_num(merged["Stock_Qty"])
    merged["Sales_Qty"] = to_num(merged["Sales_Qty"])

    for idx, dk in enumerate(s_disps):
      c_sls = dk if dk in merged.columns else f"{dk}_sls"
      c_inv = f"{dk}_inv"
      if c_sls in merged.columns and c_inv in merged.columns:
        merged[dk] = (
            norm_disp(merged[c_sls])
            .where(norm_key(merged[c_sls]) != "", norm_disp(merged[c_inv]))
            .replace("", "Unmapped")
        )
        for drop_c in [c_sls, c_inv]:
          if drop_c != dk and drop_c in merged.columns:
            merged = merged.drop(columns=[drop_c])
      elif c_inv in merged.columns and dk not in merged.columns:
        merged[dk] = norm_disp(merged[c_inv]).replace("", "Unmapped")
        merged = merged.drop(columns=[c_inv], errors="ignore")

    if want_turnover:
      merged["Daily_Sales"] = (merged["Sales_Qty"] / days_p).round(2)
      merged["Turnover_Days"] = np.where(
          merged["Sales_Qty"] > 0,
          np.round((merged["Stock_Qty"] / merged["Sales_Qty"]) * days_p, 1),
          0.0,
      )
      sort_col = "Turnover_Days"
      y_c = "Turnover_Days"
      t_str = f"Turnover Days ({days_p} Days) by {' + '.join(s_disps)}"
    else:
      merged["Stock_Need"] = np.maximum(
          merged["Sales_Qty"] - merged["Stock_Qty"], 0.0
      )
      merged["Stock_Diff"] = merged["Stock_Qty"] - merged["Sales_Qty"]
      sort_col = "Stock_Need" if "need" in q else "Stock_Qty"
      y_c = "Stock_Qty"
      t_str = f"Stock vs Sales by {' + '.join(s_disps)}"

    if len(s_disps) > 1:
      merged["Label"] = merged[s_disps].astype(str).agg(" | ".join, axis=1)
      x_c = "Label"
    else:
      x_c = s_disps[0]

    merged = merged.drop(columns=s_keys, errors="ignore").sort_values(
        by=sort_col, ascending=False
    )
    return merged, x_c, y_c, t_str

  # শুধুমাত্র Stock / Inventory রিপোর্ট
  if want_stock and has_inv:
    i_grp, i_keys, i_disps = prep_grouped(i_sub, grp_cols, "Stock_Qty")
    if len(i_disps) > 1:
      i_grp["Label"] = i_grp[i_disps].astype(str).agg(" | ".join, axis=1)
      x_c = "Label"
    else:
      x_c = i_disps[0]
    i_grp = i_grp.drop(columns=i_keys, errors="ignore").sort_values(
        by="Stock_Qty", ascending=False
    )
    return i_grp, x_c, "Stock_Qty", f"Stock by {' + '.join(i_disps)}"

  # Sales / Target vs Achievement / যেকোনো কমান্ড
  if has_sales:
    s_grp, s_keys, s_disps = prep_grouped(
        s_sub, grp_cols, "Sales_Qty", "Sales_Amount", "Active_Shops"
    )
    if df_t is not None and not df_t.empty:
      tk = s_keys[0]
      s_grp = pd.merge(
          s_grp, df_t, left_on=tk, right_on="_tgt_key", how="left"
      )
      s_grp["Target_Qty"] = to_num(s_grp["Target_Qty"])
      s_grp["Achievement_%"] = np.where(
          s_grp["Target_Qty"] > 0,
          np.round((s_grp["Sales_Qty"] / s_grp["Target_Qty"]) * 100, 2),
          0.0,
      )
      s_grp = s_grp.drop(columns=["_tgt_key"], errors="ignore")

    if len(s_disps) > 1:
      s_grp["Label"] = s_grp[s_disps].astype(str).agg(" | ".join, axis=1)
      x_c = "Label"
    else:
      x_c = s_disps[0]

    s_grp = s_grp.drop(columns=s_keys, errors="ignore").sort_values(
        by="Sales_Qty", ascending=False
    )
    return s_grp, x_c, "Sales_Qty", f"Sales Report by {' + '.join(s_disps)}"

  if has_inv:
    i_grp, i_keys, i_disps = prep_grouped(i_sub, grp_cols, "Stock_Qty")
    x_c = i_disps[0]
    i_grp = i_grp.drop(columns=i_keys, errors="ignore").sort_values(
        by="Stock_Qty", ascending=False
    )
    return i_grp, x_c, "Stock_Qty", f"Stock Report by {' + '.join(i_disps)}"

  return None, None, None, "Report"


st.markdown("### 💬 কমান্ড বক্স (যেকোনো রিপোর্ট বা কলামের নাম লিখুন)")
cmd = st.text_input(
    "কমান্ড লিখুন",
    placeholder=(
        "যা চাইবেন লিখুন — যেমন: Shop Model Wise Turn Over days, MD Stock"
        " Need, SE Wise Target vs Achievement, ASM wise sales, Model wise stock"
        " ইত্যাদি"
    ),
)

c1, c2 = st.columns(2)
report, x_col, y_col, chart_title = None, None, None, "Report"

q = (cmd or "").strip().lower()
days_period = get_days(df_sales)

if df_sales is not None or df_inv is not None:
  try:
    report, x_col, y_col, chart_title = build_dynamic_report(
        q, df_sales, df_inv, df_target, days_period
    )
    limit_n = top_n_from_cmd(q)
    if report is not None and limit_n and limit_n > 0:
      report = report.head(limit_n)
  except Exception as e:
    st.info(f"রিপোর্ট তৈরি করার সময় নোট: {e}")

if report is not None and not report.empty:
  with c1:
    st.subheader("📋 রিপোর্ট টেবিল")
    st.dataframe(report, use_container_width=True)

    excel_buf = io.BytesIO()
    with pd.ExcelWriter(excel_buf, engine="openpyxl") as writer:
      report.to_excel(writer, index=False, sheet_name="Report")
    st.download_button(
        label="⬇️ Download Excel (.xlsx)",
        data=excel_buf.getvalue(),
        file_name="vivo_generated_report.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ),
    )

  with c2:
    st.subheader("🖼️ চার্ট ও ছবি ডাউনলোড")
    if x_col and y_col and x_col in report.columns and y_col in report.columns:
      fig = px.bar(
          report.head(25), x=x_col, y=y_col, text=y_col, title=chart_title
      )
      st.plotly_chart(fig, use_container_width=True)

      try:
        img_bytes = fig.to_image(format="png", scale=2)
        st.download_button(
            label="⬇️ Download Chart Image (.png)",
            data=img_bytes,
            file_name="vivo_report_chart.png",
            mime="image/png",
        )
      except Exception:
        st.caption(
            "PNG ডাউনলোডের জন্য রিপোজিটরিতে `requirements.txt` ফাইলে `kaleido`"
            " রাখুন।"
        )

if pbi_link:
  st.divider()
  st.subheader("⚡ Power BI Dashboard")
  st.components.v1.iframe(pbi_link, height=600, scrolling=True)
