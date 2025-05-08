import streamlit as st
import pandas as pd
from openpyxl import load_workbook, Workbook
from io import BytesIO
from tag_utils import detect_tag_from_cell, month_to_daily_df

st.title("📋 Excel点検表 → XC-GATE帳票変換")

uploaded_file = st.file_uploader("点検表Excelをアップロード", type=["xlsx"])
if uploaded_file:
    wb = load_workbook(uploaded_file, data_only=True)
    ws = wb.active

    st.subheader("元データプレビュー")
    df_raw = pd.DataFrame([[cell.value for cell in row] for row in ws.iter_rows()])
    st.dataframe(df_raw)

    st.subheader("日次＋タグ付き帳票")
    df_daily = month_to_daily_df(ws)
    st.dataframe(df_daily)

    # 出力ブック作成
    out_wb = Workbook()
    out_ws = out_wb.active
    out_ws.title = "XC-GATE帳票"

    for i, row in df_daily.iterrows():
        for j, val in enumerate(row):
            out_ws.cell(row=i+1, column=j+1, value=val)

    output = BytesIO()
    out_wb.save(output)
    st.download_button("📥 XC-GATE帳票をダウンロード", data=output.getvalue(), file_name="xcgate_output.xlsx")
