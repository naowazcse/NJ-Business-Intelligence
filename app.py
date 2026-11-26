import io
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

st.title("📊 vivo NJ Region — Sales, Inventory & Report Generator")

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
  unique_cols = []
  for c in df.columns:
    if c in seen:
      seen[c] += 1
      unique_cols.append(f"{c}_{seen[c]}")
    else:
      seen[c] = 0
      unique_cols.append(c)
  df.columns = unique_cols
  return df


def norm_key(series):
  return (
      series.fillna("")
      .astype(str)
      .str.replace(r"\.0$", "", regex=True)
      .str.strip()
  )


def to_num(series):
  return pd.to_numeric(series, errors="coerce").fillna(0)


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
  df = clean_cols(read_excel_safe(raw, sheet_name="NJ Shop Information", header=1))
  store_col = find_col(df, ["store code", "商店代码", "store_code"])
  if not store_col:
    return None

  keep_cols = [store_col]
  for kw in [
      "shop name",
      "se ",
      "se_name",
      "asm",
      "rsm",
      "business area",
      "upazola",
      "thana",
      "md name",
      "md code",
  ]:
    col = find_col(df, [kw])
    if col and col not in keep_cols:
      keep_cols.append(col)

  out = df[keep_cols].copy()
  out["_store_key"] = norm_key(out[store_col])
  out = out[out["_store_key"] != ""].drop_duplicates(
      subset=["_store_key"], keep="first"
  )
  return out


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
            "se name",
            "se_name",
            "se ",
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

# Shop তথ্য Sales ও Inventory-তে নিরাপদে যুক্ত করা
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


def detect_period_days(df_s):
  if df_s is not None:
    dt_col = find_col(df_s, ["sales date", "date"])
    if dt_col:
      dts = pd.to_datetime(df_s[dt_col], errors="coerce").dropna()
      if not dts.empty:
        span = (dts.max() - dts.min()).days + 1
        if span > 0:
          return int(span)
  return 30


def detect_group_col(q, df_s, df_i):
  # কমান্ডে যে ডাইমেনশন চাওয়া হয়েছে সেটি আগে খোঁজা হবে
  rules = [
      (["shop", "store", "দোকান", "শপ"], ["shop name", "_store_key"]),
      (["md", "এমডি", "distributor", "ডিস্ট্রিবিউটর"], ["md name", "md code"]),
      (["model", "মডেল", "product", "পণ্য"], ["model name", "model"]),
      (["asm", "এএসএম"], ["asm name", "asm"]),
      (["rsm", "আরএসএম"], ["rsm name", "rsm"]),
      (["area", "এরিয়া", "thana", "upazila"], ["business area", "upazola"]),
      (
          ["se", "এসই", "target", "টার্গেট", "achievement"],
          ["se name", "se_name", "se ", "md name", "model name"],
      ),
  ]
  for kws, col_kws in rules:
    if any(k in q for k in kws):
      for d in [df_s, df_i]:
        col = find_col(d, col_kws)
        if col:
          return col

  # কমান্ডে নির্দিষ্ট শব্দ না থাকলে উপলব্ধ কলাম থেকেই একটি বেছে নেওয়া
  for d in [df_s, df_i]:
    for col_kws in [
        ["se name", "se ", "md name", "model name", "shop name", "_store_key"]
    ]:
      col = find_col(d, col_kws)
      if col:
        return col
  return None


def group_sum(df, g_col, val_col, out_name):
  grp_vals = df[g_col].fillna("Unmapped").astype(str).str.strip()
  res = df.groupby(grp_vals, as_index=False)[val_col].sum()
  res.columns = [g_col, out_name]
  return res


st.markdown("### 💬 কমান্ড বক্স (বাংলা বা ইংরেজিতে লিখুন)")
cmd = st.text_input(
    "কমান্ড লিখুন",
    placeholder=(
        "যেমন: Shop Wise Turn Over Days / MD Stock Need / SE Wise Target vs"
        " Achievement / Model wise sales"
    ),
)

c1, c2 = st.columns(2)
report = None
x_col, y_col, chart_title = None, None, "Report"

q = (cmd or "").strip().lower()
days_period = detect_period_days(df_sales)

if df_sales is not None or df_inv is not None:
  try:
    is_turnover = any(
        k in q for k in ["turnover", "turn over", "turnaround", "টার্নওভার"]
    )
    is_stock_q = any(
        k in q for k in ["stock", "inventory", "need", "স্টক", "ইনভেন্টরি"]
    )
    is_shop_q = any(
        k in q for k in ["shop", "store", "দোকান", "শপ"]
    )
    is_md_q = any(
        k in q for k in ["md", "এমডি", "distributor", "ডিস্ট্রিবিউটর"]
    )

    # ১. Turnover Days রিপোর্ট (Shop বা যেকোনো লেভেলে)
    if is_turnover and df_sales is not None and df_inv is not None:
      s_key = "_store_key" if "_store_key" in df_sales.columns else None
      i_key = "_store_key" if "_store_key" in df_inv.columns else None

      if s_key and i_key:
        s_grp = (
            df_sales.groupby(s_key, as_index=False)["_qty"]
            .sum()
            .rename(columns={"_qty": "Sales_Qty"})
        )
        i_grp = (
            df_inv.groupby(i_key, as_index=False)["_qty"]
            .sum()
            .rename(columns={"_qty": "Stock_Qty"})
        )
        m = pd.merge(i_grp, s_grp, on="_store_key", how="outer")
        m["Stock_Qty"] = to_num(m["Stock_Qty"])
        m["Sales_Qty"] = to_num(m["Sales_Qty"])

        meta_cols = ["_store_key"]
        for src in [df_inv, df_sales]:
          if src is not None and "_store_key" in src.columns:
            for kw in ["shop name", "se ", "se_name", "md name", "asm"]:
              mc = find_col(src, [kw])
              if mc and mc not in meta_cols:
                meta_cols.append(mc)
        if len(meta_cols) > 1:
          src_m = df_inv if df_inv is not None else df_sales
          m = pd.merge(
              m,
              src_m[meta_cols].drop_duplicates(subset=["_store_key"]),
              on="_store_key",
              how="left",
          )

        m["Daily_Sales"] = (m["Sales_Qty"] / days_period).round(2)
        m["Turnover_Days"] = [
            round((stk / sls) * days_period, 1) if sls > 0 else 0.0
            for stk, sls in zip(m["Stock_Qty"], m["Sales_Qty"])
        ]

        # MD বা SE অনুযায়ী চাইলে গ্রুপিং করে দেওয়া
        if is_md_q and find_col(m, ["md name"]):
          md_c = find_col(m, ["md name"])
          report = (
              m.groupby(
                  m[md_c].fillna("Unmapped").astype(str).str.strip(),
                  as_index=False,
              )[["Stock_Qty", "Sales_Qty", "Daily_Sales"]]
              .sum()
          )
          report.columns = [md_c, "Stock_Qty", "Sales_Qty", "Daily_Sales"]
          report["Turnover_Days"] = [
              round((stk / sls) * days_period, 1) if sls > 0 else 0.0
              for stk, sls in zip(report["Stock_Qty"], report["Sales_Qty"])
          ]
          report = report.sort_values(by="Turnover_Days", ascending=False)
          x_col, y_col, chart_title = (
              md_c,
              "Turnover_Days",
              f"MD Wise Turnover Days ({days_period} Days)",
          )
        else:
          lbl = find_col(m, ["shop name"]) or "_store_key"
          report = m.sort_values(by="Turnover_Days", ascending=False)
          x_col, y_col, chart_title = (
              lbl,
              "Turnover_Days",
              f"Shop Wise Turnover Days ({days_period} Days)",
          )

    # ২. MD Stock Need / MD Stock রিপোর্ট
    elif is_md_q and is_stock_q:
      md_i = find_col(df_inv, ["md name", "md code"])
      md_s = find_col(df_sales, ["md name", "md code"])

      if df_inv is not None and md_i:
        inv_m = group_sum(df_inv, md_i, "_qty", "Stock_Qty")
        inv_m["_k"] = norm_key(inv_m[md_i])

        if df_sales is not None and md_s:
          sls_m = group_sum(df_sales, md_s, "_qty", "Sales_Qty")
          sls_m["_k"] = norm_key(sls_m[md_s])
          m = pd.merge(
              inv_m, sls_m[["_k", "Sales_Qty"]], on="_k", how="left"
          )
          m["Sales_Qty"] = to_num(m["Sales_Qty"])
          m["Stock_Need"] = (m["Sales_Qty"] - m["Stock_Qty"]).clip(lower=0)
          m["Stock_vs_Sales_Diff"] = m["Stock_Qty"] - m["Sales_Qty"]
          report = m.drop(columns=["_k"]).sort_values(
              by=["Stock_Need", "Stock_Qty"], ascending=[False, False]
          )
          x_col, y_col, chart_title = (
              md_i,
              "Stock_Qty",
              "MD Wise Stock vs Sales",
          )
        else:
          report = inv_m.drop(columns=["_k"]).sort_values(
              by="Stock_Qty", ascending=False
          )
          x_col, y_col, chart_title = md_i, "Stock_Qty", "MD Wise Stock"

    # ৩. যেকোনো Stock / Inventory কমান্ড
    elif is_stock_q and df_inv is not None:
      g_col = detect_group_col(q, None, df_inv)
      if g_col:
        report = group_sum(df_inv, g_col, "_qty", "Stock_Qty").sort_values(
            by="Stock_Qty", ascending=False
        )
        if df_sales is not None:
          s_g = detect_group_col(q, df_sales, None)
          if s_g:
            s_m = group_sum(df_sales, s_g, "_qty", "Sales_Qty")
            s_m["_k"] = norm_key(s_m[s_g])
            report["_k"] = norm_key(report[g_col])
            report = pd.merge(
                report, s_m[["_k", "Sales_Qty"]], on="_k", how="left"
            )
            report["Sales_Qty"] = to_num(report["Sales_Qty"])
            report["Stock_Need"] = (
                report["Sales_Qty"] - report["Stock_Qty"]
            ).clip(lower=0)
            report = report.drop(columns=["_k"], errors="ignore")
        x_col, y_col, chart_title = g_col, "Stock_Qty", f"Stock by {g_col}"

    # ৪. Sales / Target / SE / Model / Shop বা অন্য যেকোনো কমান্ড
    elif df_sales is not None:
      g_col = detect_group_col(q, df_sales, df_inv)
      if g_col:
        grp_vals = df_sales[g_col].fillna("Unmapped").astype(str).str.strip()
        agg_spec = {"Sales_Qty": ("_qty", "sum")}
        if find_col(df_sales, ["sales amount", "amount"]):
          agg_spec["Sales_Amount"] = ("_amt", "sum")
        if "_store_key" in df_sales.columns:
          agg_spec["Active_Shops"] = ("_store_key", "nunique")

        report = (
            df_sales.groupby(grp_vals, as_index=False)
            .agg(**agg_spec)
            .sort_values(by="Sales_Qty", ascending=False)
        )
        report = report.rename(columns={report.columns[0]: g_col})

        if df_target is not None:
          report["_tk"] = norm_key(report[g_col])
          report = pd.merge(report, df_target, on="_tk", how="left")
          report["Target_Qty"] = to_num(report["Target_Qty"])
          report["Achievement_%"] = [
              round((s / t) * 100, 2) if t > 0 else 0.0
              for s, t in zip(report["Sales_Qty"], report["Target_Qty"])
          ]
          report = report.drop(columns=["_tk"], errors="ignore")

        x_col, y_col, chart_title = (
            g_col,
            "Sales_Qty",
            f"Sales Performance by {g_col}",
        )

    # ৫. শুধু Inventory থাকলে ফলব্যাক Stock রিপোর্ট
    elif df_inv is not None:
      g_col = detect_group_col(q, None, df_inv)
      if g_col:
        report = group_sum(df_inv, g_col, "_qty", "Stock_Qty").sort_values(
            by="Stock_Qty", ascending=False
        )
        x_col, y_col, chart_title = g_col, "Stock_Qty", f"Stock by {g_col}"

  except Exception as e:
    st.info(f"রিপোর্ট প্রসেস করার সময় নোট: {e}")

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
