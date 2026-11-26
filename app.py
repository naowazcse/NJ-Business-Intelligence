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


def load_shop_df(file):
  try:
    df = pd.read_excel(file, sheet_name="NJ Shop Information", header=1)
  except Exception:
    df = pd.read_excel(file, header=1)
  keep_cols = [
      c
      for c in [
          "Store Code 商店代码",
          "Shop Name As Per VK Updated",
          "SE 名称",
          "ASM Name",
          "RSM Name",
          "Business Area",
          "Upazola/Tana乡",
      ]
      if c in df.columns
  ]
  df = df[keep_cols].dropna(subset=["Store Code 商店代码"]).drop_duplicates()
  df["Store Code 商店代码"] = df["Store Code 商店代码"].astype(str).str.strip()
  return df


df_sales = pd.read_excel(sales_up) if sales_up else None
df_inv = pd.read_excel(inv_up) if inv_up else None
df_shop = load_shop_df(shop_up) if shop_up else None
df_target = pd.read_excel(target_up) if target_up else None

if df_sales is not None:
  df_sales["Store code"] = df_sales["Store code"].astype(str).str.strip()
  if "Sales date" in df_sales.columns:
    df_sales["Sales date_dt"] = pd.to_datetime(
        df_sales["Sales date"], errors="coerce"
    )
  if df_shop is not None:
    df_sales = pd.merge(
        df_sales,
        df_shop,
        left_on="Store code",
        right_on="Store Code 商店代码",
        how="left",
    )
    df_sales["SE 名称"] = df_sales["SE 名称"].fillna("Unmapped")

if df_inv is not None:
  df_inv["Store code"] = df_inv["Store code"].astype(str).str.strip()
  if df_shop is not None:
    df_inv = pd.merge(
        df_inv,
        df_shop,
        left_on="Store code",
        right_on="Store Code 商店代码",
        how="left",
    )
    df_inv["SE 名称"] = df_inv["SE 名称"].fillna("Unmapped")

st.markdown("### 💬 কমান্ড বক্স (বাংলা বা ইংরেজিতে লিখুন)")
cmd = st.text_input(
    "কমান্ড লিখুন",
    placeholder=(
        "যেমন: SE wise target vs achievement দেখাও / Model wise sales / Stock"
        " summary"
    ),
)

c1, c2 = st.columns(2)
if cmd or df_sales is not None or df_inv is not None:
  q = (cmd or "").lower()

  # ১. স্টক / ইনভেন্টরি রিপোর্ট
  if (
      ("stock" in q or "inventory" in q or "ইনভেন্টরি" in q or "স্টক" in q)
      and df_inv is not None
  ):
    grp_col = "Model name" if "Model name" in df_inv.columns else None
    if ("se" in q or "এসই" in q) and "SE 名称" in df_inv.columns:
      grp_col = "SE 名称"
    report = (
        df_inv.groupby(grp_col, dropna=False)["Quantity"]
        .sum()
        .reset_index()
        .sort_values("Quantity", ascending=False)
    )
    x_col, y_col, title = grp_col, "Quantity", f"Inventory Stock by {grp_col}"

  # ২. মডেল ওয়াইজ সেলস
  elif (
      ("model" in q or "মডেল" in q)
      and df_sales is not None
      and "Model name" in df_sales.columns
  ):
    report = (
        df_sales.groupby("Model name", dropna=False)
        .agg(
            Total_Qty=("Quantity", "sum"),
            Total_Amount=(
                "Sales amount" if "Sales amount" in df_sales.columns else "Quantity",
                "sum",
            ),
        )
        .reset_index()
        .sort_values("Total_Qty", ascending=False)
    )
    x_col, y_col, title = "Model name", "Total_Qty", "Model Wise Sales Quantity"

  # ৩. ডিফল্ট: SE ওয়াইজ সেলস ও টার্গেট ভার্সেস অ্যাচিভমেন্ট
  elif df_sales is not None and "SE 名称" in df_sales.columns:
    report = (
        df_sales.groupby("SE 名称", dropna=False)
        .agg(
            Sales_Qty=("Quantity", "sum"),
            Sales_Amount=(
                "Sales amount" if "Sales amount" in df_sales.columns else "Quantity",
                "sum",
            ),
            Active_Shops=("Store code", "nunique"),
        )
        .reset_index()
        .sort_values("Sales_Qty", ascending=False)
    )

    if df_target is not None:
      t_se_col = next(
          (
              c
              for c in df_target.columns
              if "se" in c.lower() or "name" in c.lower()
          ),
          None,
      )
      t_val_col = next(
          (c for c in df_target.columns if "target" in c.lower()), None
      )
      if t_se_col and t_val_col:
        report = pd.merge(
            report,
            df_target[[t_se_col, t_val_col]],
            left_on="SE 名称",
            right_on=t_se_col,
            how="left",
        )
        report["Achievement_%"] = (
            (report["Sales_Qty"] / report[t_val_col]) * 100
        ).round(2)

    x_col, y_col, title = "SE 名称", "Sales_Qty", "SE Wise Sales Performance"
  else:
    report = None

  if report is not None:
    with c1:
      st.subheader("📋 রিপোর্ট টেবিল")
      st.dataframe(report, use_container_width=True)

      # Excel ডাউনলোড
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

      # PNG ছবি হিসেবে ডাউনলোড
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
            "PNG এক্সপোর্টের জন্য `requirements.txt`-এ `kaleido` যুক্ত করুন।"
        )

if pbi_link:
  st.divider()
  st.subheader("⚡ Power BI Dashboard")
  st.components.v1.iframe(pbi_link, height=600, scrolling=True)
