import streamlit as st
import pandas as pd
from openpyxl import load_workbook, Workbook
from io import BytesIO
from datetime import datetime, timedelta
from tag_utils import detect_tag_from_cell, month_to_daily_df

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
以下で自由に変更可能

---

### 📤 出力
- 出力形式：`.xlsx`（XC-GATE帳票形式）
- 各日付1行、項目ごとにタグが付きます
""")

# --- サイドバー：色とタグ対応のカスタマイズ ---
st.sidebar.markdown("### 🎨 色とタグの対応設定")

default_mapping = {
    "FFFF00": "*日付",    # 黄色
    "00B0F0": "*数値",    # 青
    "00FF00": "*入力",    # 緑
    "BFBFBF": "*実績",    # グレー
}

tag_options = ["*日付", "*数値", "*入力", "*実績", "*選択", "*送信", "*日時"]
user_mapping = {}
for color_hex, default_tag in default_mapping.items():
    tag = st.sidebar.selectbox(
        f"背景色 {color_hex} に対応するタグ",
        tag_options,
        index=tag_options.index(default_tag)
    )
    user_mapping[color_hex.upper()] = tag

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
