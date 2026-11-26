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


def norm_key(series):
  return (
      series.fillna("")
      .astype(str)
      .str.replace(r"\.0$", "", regex=True)
      .str.strip()
  )


def load_shop_df(file):
  if file is None:
    return None
  try:
    df = pd.read_excel(file, sheet_name="NJ Shop Information", header=1)
  except Exception:
    df = pd.read_excel(file, header=1)

  df.columns = [str(c).strip() for c in df.columns]
  store_col = next(
      (
          c
          for c in df.columns
          if "store code" in c.lower() or "商店代码" in c
      ),
      None,
  )
  if not store_col:
    return None

  desired = [
      store_col,
      "Shop Name As Per VK Updated",
      "SE 名称",
      "ASM Name",
      "RSM Name",
      "Business Area",
      "Upazola/Tana乡",
  ]
  keep_cols = [c for c in desired if c in df.columns]
  df = df[keep_cols].copy()
  df["_store_key"] = norm_key(df[store_col])
  df = df[df["_store_key"] != ""].drop_duplicates(subset=["_store_key"])
  if "SE 名称" in df.columns:
    df["SE 名称"] = df["SE 名称"].fillna("Unmapped").astype(str).str.strip()
    df.loc[df["SE 名称"] == "", "SE 名称"] = "Unmapped"
  return df


def load_target_df(file):
  if file is None:
    return None
  xls = pd.ExcelFile(file)
  best_df = None
  for sh in xls.sheet_names:
    tmp = pd.read_excel(file, sheet_name=sh)
    tmp.columns = [str(c).strip() for c in tmp.columns]
    if all(str(c).startswith("Unnamed") for c in tmp.columns) and len(tmp) > 0:
      tmp.columns = [str(x).strip() for x in tmp.iloc[0]]
      tmp = tmp.iloc[1:].reset_index(drop=True)
    se_col = next(
        (
            c
            for c in tmp.columns
            if any(
                k in str(c).lower()
                for k in ["se name", "se_name", "se", "এসই", "sales executive"]
            )
        ),
        None,
    )
    tgt_col = next(
        (
            c
            for c in tmp.columns
            if any(
                k in str(c).lower()
                for k in ["target", "টার্গেট", "qty target", "tgt"]
            )
        ),
        None,
    )
    if se_col and tgt_col:
      t_clean = pd.DataFrame({
          "_se_key": norm_key(tmp[se_col]),
          "Target_Qty": pd.to_numeric(tmp[tgt_col], errors="coerce").fillna(0),
      })
      t_clean = t_clean[t_clean["_se_key"] != ""]
      best_df = t_clean.groupby("_se_key", as_index=False)["Target_Qty"].sum()
      break
  return best_df


df_sales = pd.read_excel(sales_up) if sales_up else None
df_inv = pd.read_excel(inv_up) if inv_up else None
df_shop = load_shop_df(shop_up) if shop_up else None
df_target = load_target_df(target_up) if target_up else None

if df_sales is not None:
  df_sales.columns = [str(c).strip() for c in df_sales.columns]
  s_store_col = next(
      (c for c in df_sales.columns if "store code" in c.lower()), None
  )
  if "Quantity" in df_sales.columns:
    df_sales["Quantity"] = pd.to_numeric(
        df_sales["Quantity"], errors="coerce"
    ).fillna(0)
  if "Sales amount" in df_sales.columns:
    df_sales["Sales amount"] = pd.to_numeric(
        df_sales["Sales amount"], errors="coerce"
    ).fillna(0)
  if s_store_col and df_shop is not None:
    df_sales["_store_key"] = norm_key(df_sales[s_store_col])
    df_sales = pd.merge(
        df_sales,
        df_shop.drop(
            columns=[
                c
                for c in df_shop.columns
                if "code" in c.lower() or "商店代码" in c
            ],
            errors="ignore",
        ),
        on="_store_key",
        how="left",
    )
  if "SE 名称" not in df_sales.columns:
    df_sales["SE 名称"] = "Unmapped"
  else:
    df_sales["SE 名称"] = (
        df_sales["SE 名称"].fillna("Unmapped").astype(str).str.strip()
    )

if df_inv is not None:
  df_inv.columns = [str(c).strip() for c in df_inv.columns]
  i_store_col = next(
      (c for c in df_inv.columns if "store code" in c.lower()), None
  )
  if "Quantity" in df_inv.columns:
    df_inv["Quantity"] = pd.to_numeric(
        df_inv["Quantity"], errors="coerce"
    ).fillna(0)
  if i_store_col and df_shop is not None:
    df_inv["_store_key"] = norm_key(df_inv[i_store_col])
    df_inv = pd.merge(
        df_inv,
        df_shop.drop(
            columns=[
                c
                for c in df_shop.columns
                if "code" in c.lower() or "商店代码" in c
            ],
            errors="ignore",
        ),
        on="_store_key",
        how="left",
    )
  if "SE 名称" not in df_inv.columns:
    df_inv["SE 名称"] = "Unmapped"
  else:
    df_inv["SE 名称"] = (
        df_inv["SE 名称"].fillna("Unmapped").astype(str).str.strip()
    )

st.markdown("### 💬 কমান্ড বক্স (বাংলা বা ইংরেজিতে লিখুন)")
cmd = st.text_input(
    "কমান্ড লিখুন",
    placeholder=(
        "যেমন: MD Stock Need / SE Wise Target vs Achievement / Model wise"
        " sales"
    ),
)

c1, c2 = st.columns(2)
report = None
x_col, y_col, title = None, None, "Report"

if cmd or df_sales is not None or df_inv is not None:
  q = (cmd or "").lower()

  # ১. MD Stock Need / MD-wise Stock vs Sales
  if ("md" in q) and (
      "stock" in q or "inventory" in q or "need" in q or "স্টক" in q
  ):
    if df_inv is not None and "MD name" in df_inv.columns:
      md_stock = (
          df_inv.groupby("MD name", as_index=False)["Quantity"]
          .sum()
          .rename(columns={"Quantity": "Stock_Qty"})
      )
      md_stock["_md_key"] = norm_key(md_stock["MD name"])

      if df_sales is not None and "MD name" in df_sales.columns:
        md_sales = (
            df_sales.groupby("MD name", as_index=False)["Quantity"]
            .sum()
            .rename(columns={"Quantity": "Sales_Qty"})
        )
        md_sales["_md_key"] = norm_key(md_sales["MD name"])
        report = pd.merge(
            md_stock,
            md_sales[["_md_key", "Sales_Qty"]],
            on="_md_key",
            how="left",
        )
        report["Sales_Qty"] = (
            pd.to_numeric(report["Sales_Qty"], errors="coerce").fillna(0)
        )
        # Sales এর তুলনায় Stock ঘাটতি থাকলে Stock_Need পজিটিভ আসবে, নাহলে 0
        report["Stock_Need"] = (
            report["Sales_Qty"] - report["Stock_Qty"]
        ).clip(lower=0)
        report["Stock_Gap_Net"] = report["Stock_Qty"] - report["Sales_Qty"]
        report = report.drop(columns=["_md_key"]).sort_values(
            by=["Stock_Need", "Stock_Qty"], ascending=[False, False]
        )
        x_col, y_col, title = "MD name", "Stock_Qty", "MD Wise Stock vs Sales"
      else:
        report = md_stock.drop(columns=["_md_key"]).sort_values(
            by="Stock_Qty", ascending=False
        )
        x_col, y_col, title = "MD name", "Stock_Qty", "MD Wise Inventory Stock"

  # ২. সাধারণ স্টক / ইনভেন্টরি রিপোর্ট (Model বা SE)
  elif (
      ("stock" in q or "inventory" in q or "ইনভেন্টরি" in q or "স্টক" in q)
      and df_inv is not None
  ):
    grp_col = (
        "SE 名称"
        if (("se" in q or "এসই" in q) and "SE 名称" in df_inv.columns)
        else ("Model name" if "Model name" in df_inv.columns else "SE 名称")
    )
    report = (
        df_inv.groupby(grp_col, as_index=False)["Quantity"]
        .sum()
        .sort_values("Quantity", ascending=False)
    )
    x_col, y_col, title = grp_col, "Quantity", f"Inventory Stock by {grp_col}"

  # ৩. মডেল ওয়াইজ সেলস
  elif (
      ("model" in q or "মডেল" in q)
      and df_sales is not None
      and "Model name" in df_sales.columns
  ):
    agg_dict = {"Total_Qty": ("Quantity", "sum")}
    if "Sales amount" in df_sales.columns:
      agg_dict["Total_Amount"] = ("Sales amount", "sum")
    report = (
        df_sales.groupby("Model name", as_index=False)
        .agg(**agg_dict)
        .sort_values("Total_Qty", ascending=False)
    )
    x_col, y_col, title = "Model name", "Total_Qty", "Model Wise Sales Quantity"

  # ৪. SE ওয়াইজ সেলস ও টার্গেট ভার্সেস অ্যাচিভমেন্ট
  elif df_sales is not None:
    agg_dict = {"Sales_Qty": ("Quantity", "sum")}
    if "Sales amount" in df_sales.columns:
      agg_dict["Sales_Amount"] = ("Sales amount", "sum")
    if "_store_key" in df_sales.columns:
      agg_dict["Active_Shops"] = ("_store_key", "nunique")

    report = (
        df_sales.groupby("SE 名称", as_index=False)
        .agg(**agg_dict)
        .sort_values("Sales_Qty", ascending=False)
    )
    report["_se_key"] = norm_key(report["SE 名称"])

    if df_target is not None:
      report = pd.merge(report, df_target, on="_se_key", how="left")
      report["Target_Qty"] = pd.to_numeric(
          report["Target_Qty"], errors="coerce"
      ).fillna(0)
      report["Achievement_%"] = report.apply(
          lambda r: round((r["Sales_Qty"] / r["Target_Qty"]) * 100, 2)
          if r["Target_Qty"] > 0
          else 0.0,
          axis=1,
      )

    report = report.drop(columns=["_se_key"], errors="ignore")
    x_col, y_col, title = "SE 名称", "Sales_Qty", "SE Wise Sales Performance"

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
    fig = px.bar(report, x=x_col, y=y_col, text=y_col, title=title)
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
