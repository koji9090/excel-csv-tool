# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import openpyxl
import io
import zipfile

# === アプリのタイトル ===
st.title("📂 Excel列分解＆CSV作成ツール")
st.write("基準列（店舗名）の順番を完全に維持してCSV化します。")

# === サイドバー：設定エリア ===
st.sidebar.header("⚙️ 設定")
anchor_col = st.sidebar.text_input("1. 固定して使う基準列 (例: A)", value="A")
skip_rows = st.sidebar.number_input("2. 最初に削除する行数", min_value=0, value=2)
ignore_col_start = st.sidebar.text_input("3. 除外したい開始列 (例: B)", value="")
ignore_col_end = st.sidebar.text_input("4. 除外したい終了列 (例: G)", value="")

# === メインエリア ===
uploaded_file = st.file_uploader("Excelファイルをアップロード", type=['xlsx'])

if uploaded_file:
    try:
        file_bytes = uploaded_file.getvalue()
        wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=False)
        ws = wb.active
        
        anchor_idx = openpyxl.utils.column_index_from_string(anchor_col) - 1

        # --- 【内部処理】基準となる店舗マスターリストを作成 ---
        df_full = pd.read_excel(
            io.BytesIO(file_bytes), 
            header=None, 
            skiprows=skip_rows, 
            engine='openpyxl'
        )
        # 店舗名をクレンジングし、重複を排除して「正しい順番」を固定
        df_full.iloc[:, anchor_idx] = df_full.iloc[:, anchor_idx].astype(str).str.strip()
        # 店舗名列（anchor_idx）だけで重複削除し、順番を保持
        master_stores = df_full.iloc[:, [anchor_idx]].drop_duplicates(subset=df_full.columns[anchor_idx], keep='first')
        master_list = master_stores.iloc[:, 0].tolist()

        # --- 数式列の自動検出 ---
        ignore_indices = []
        if ignore_col_start and ignore_col_end:
            try:
                s = openpyxl.utils.column_index_from_string(ignore_col_start) - 1
                e = openpyxl.utils.column_index_from_string(ignore_col_end)
                ignore_indices = list(range(s, e))
            except: pass

        formula_candidates = []
        for c in range(1, ws.max_column + 1):
            if (c-1)
