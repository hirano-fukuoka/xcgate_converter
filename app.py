import streamlit as st
import pandas as pd
from openpyxl import load_workbook, Workbook
from io import BytesIO
from datetime import datetime, timedelta
from tag_utils import detect_tag_from_cell, month_to_daily_df

st.set_page_config(page_title="XC-GATE帳票変換", layout="wide")
st.title("📋 Excel点検表 → XC-GATE帳票変換アプリ")

# --- サイドバー：取扱説明 ---
st.sidebar.title("ℹ️ 取扱説明")
with st.sidebar.expander("▶️ アプリの使い方", expanded=True):
    st.markdown("""
### 📝 アプリ概要
このアプリは、**Excel点検表をXC-GATE帳票に自動変換**するツールです。

---

### ✅ 入力条件
- 点検表は `.xlsx` 形式
- 1列目：点検項目名
- 2列目以降：日別または月単位の点検データ

---

### 🎨 タグの変換ルール（背景色）
下記で最大10件まで自由に設定できます。

---

### 📤 出力
- 出力形式：`.xlsx`（XC-GATE帳票形式）
- 各日付1行、項目ごとにタグが付きます
""")

# --- サイドバー：色とタグ対応設定（最大10件） ---
st.sidebar.markdown("### 🎨 色とタグの対応設定（最大10件）")

tag_options = ["*日付", "*数値", "*入力", "*実績", "*選択", "*送信", "*日時"]
default_colors = ["FFFF00", "00B0F0", "00FF00", "BFBFBF", "FF0000", "C0C0C0", "800080", "FFA500", "008000", "000000"]
default_tags = ["*日付", "*数値", "*入力", "*実績", "*選択", "*送信", "*日時", "*入力", "*入力", "*入力"]

user_mapping = {}

for i in range(10):
    color_hex = st.sidebar.text_input(f"{i+1}. 背景色コード（6桁 HEX）", value=default_colors[i], key=f"color_{i}")
    color_hex = color_hex.upper().strip().replace("#", "")[:6]

    # 色のプレビュー付き表示
    st.sidebar.markdown(
        f"""
        <div style='display:flex; align-items:center; margin-bottom:4px'>
            <div style='width:20px; height:20px; background-color:#{color_hex}; border:1px solid #ccc; margin-right:8px'></div>
            <span style='font-weight:bold'>背景色 #{color_hex}</span>
        </div>
        """,
        unsafe_allow_html=True
    )

    tag = st.sidebar.selectbox(
        f"→ 上記の色に対応するタグ",
        tag_options,
        index=tag_options.index(default_tags[i]),
        key=f"tag_{i}"
    )
    user_mapping[color_hex] = tag

# --- メイン処理 ---
uploaded_file = st.file_uploader("点検表Excelをアップロード", type=["xlsx"])
if uploaded_file:
    wb = load_workbook(uploaded_file, data_only=True)
    ws = wb.active

    st.subheader("📄 元Excelプレビュー")
    df_raw = pd.DataFrame([[cell.value for cell in row] for row in ws.iter_rows()])
    st.dataframe(df_raw)

    st.subheader("🗓️ 日次＋タグ付き帳票")
    df_daily = month_to_daily_df(ws, user_mapping)
    st.dataframe(df_daily)

    # Excel帳票出力
    out_wb = Workbook()
    out_ws = out_wb.active
    out_ws.title = "XC-GATE帳票"

    for i, row in df_daily.iterrows():
        for j, val in enumerate(row):
            out_ws.cell(row=i+1, column=j+1, value=val)

    # ダウンロードボタン
    output = BytesIO()
    out_wb.save(output)
    st.download_button("📥 XC-GATE帳票をダウンロード", data=output.getvalue(), file_name="xcgate_output.xlsx")
