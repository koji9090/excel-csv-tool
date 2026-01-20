# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import openpyxl
import io
import zipfile

# === アプリのタイトル ===
st.title("📂 Excel列分解＆CSV作成ツール")
st.write("Excelの並び順を1行目から最後まで完全に維持してCSV化します。")

# === サイドバー：設定エリア ===
st.sidebar.header("⚙️ 設定")
anchor_col_letter = st.sidebar.text_input("1. 基準列 (例: A)", value="A")
skip_rows = st.sidebar.number_input("2. データ開始前の除外行数", min_value=0, value=2)
ignore_col_start = st.sidebar.text_input("3. 除外開始列", value="")
ignore_col_end = st.sidebar.text_input("4. 除外終了列", value="")

# === メインエリア ===
uploaded_file = st.file_uploader("Excelファイルをアップロード", type=['xlsx'])

if uploaded_file:
    try:
        file_bytes = uploaded_file.getvalue()
        
        # 1. ファイル名取得用に openpyxl で読み込み (2行目の値を取得するため)
        wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
        ws = wb.active
        
        # 2. データ処理用に Pandas で全データを読み込み (header=None で全ての行を対象にする)
        # engine='openpyxl' を明示し、計算後の値を読み込む
        df_raw = pd.read_excel(io.BytesIO(file_bytes), header=None, engine='openpyxl')

        # 基準列のインデックス (A -> 0)
        anchor_idx = openpyxl.utils.column_index_from_string(anchor_col_letter) - 1

        # 3. 数式列の自動検出 (openpyxlの数式フラグを使用)
        # 再度、数式を確認するために数式保持モードで読み込み
        wb_formula = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=False)
        ws_f = wb_formula.active
        
        formula_candidates = []
        ignore_indices = []
        if ignore_col_start and ignore_col_end:
            try:
                s = openpyxl.utils.column_index_from_string(ignore_col_start) - 1
                e = openpyxl.utils.column_index_from_string(ignore_col_end)
                ignore_indices = list(range(s, e))
            except: pass

        # データ開始行（skip_rowsの次）から数行チェックして数式列を探す
        check_start = skip_rows + 1
        check_end = min(check_start + 10, ws_f.max_row)
        for c in range(1, ws_f.max_column + 1):
            if (c-1) == anchor_idx or (c-1) in ignore_indices: continue
            is_f = False
            for r in range(check_start, check_end + 1):
                cell = ws_f.cell(row=r, column=c)
                if cell.data_type == 'f' or (cell.value and str(cell.value).startswith('=')):
                    is_f = True; break
            if is_f:
                formula_candidates.append({"idx": c-1, "name": openpyxl.utils.get_column_letter(c)})

        if formula_candidates:
            st.success(f"✅ {len(formula_candidates)} 個の数式列を検出しました。")
            
            st.subheader("🛠️ 出力する列を選択")
            selected_indices = []
            cols_ui = st.columns(4)
            for i, cand in enumerate(formula_candidates):
                with cols_ui[i % 4]:
                    if st.checkbox(f"{cand['name']} 列", value=True, key=cand['idx']):
                        selected_indices.append(cand['idx'])

            # --- CSV作成実行 ---
            if st.button("🚀 CSVを作成"):
                if not selected_indices:
                    st.error("列を選択してください。")
                else:
                    zip_buffer = io.BytesIO()
                    with zipfile.ZipFile(zip_buffer, 'w') as myzip:
                        for target_idx in selected_indices:
                            # ファイル名生成: 2行目(row=2)のセルの値を取得
                            col_letter = openpyxl.utils.get_column_letter(target_idx + 1)
                            row2_val = ws.cell(row=2, column=target_idx + 1).value
                            suffix = f"_{row2_val}" if row2_val is not None else ""
