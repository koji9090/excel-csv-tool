# -*- coding: utf-8 -*-
import streamlit as st
import openpyxl
import io
import zipfile
import csv

# === アプリのタイトル ===
st.title("📂 Excel列分解＆CSV作成ツール")
st.write("Excelの並び順を100%維持してCSV化します。")

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
        
        # 1. データ読み込み (計算後の値を取得)
        wb_data = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
        ws = wb_data.active
        
        # 2. 数式チェック用に読み込み
        wb_formula = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=False)
        ws_f = wb_formula.active
        
        anchor_idx = openpyxl.utils.column_index_from_string(anchor_col_letter)

        # --- 数式列の自動検出 ---
        formula_candidates = []
        check_start = skip_rows + 1
        check_end = min(check_start + 10, ws_f.max_row)
        for c in range(1, ws_f.max_column + 1):
            if c == anchor_idx: continue
            is_f = False
            for r in range(check_start, check_end + 1):
                cell = ws_f.cell(row=r, column=c)
                if cell.data_type == 'f' or (cell.value and str(cell.value).startswith('=')):
                    is_f = True; break
            if is_f:
                formula_candidates.append({"idx": c, "name": openpyxl.utils.get_column_letter(c)})

        if formula_candidates:
            st.success(f"✅ {len(formula_candidates)} 個の数式列を検出しました。")
            selected_cols = []
            cols_ui = st.columns(4)
            for i, cand in enumerate(formula_candidates):
                with cols_ui[i % 4]:
                    if st.checkbox(f"{cand['name']} 列", value=True, key=cand['idx']):
                        selected_cols.append(cand['idx'])

            if st.button("🚀 CSVを作成"):
                with st.spinner('Excelの順番を維持して処理中...'):
                    zip_buffer = io.BytesIO()
                    with zipfile.ZipFile(zip_buffer, 'w') as myzip:
                        
                        for target_idx in selected_cols:
                            # 2行目の値を取得してファイル名にする
                            col_letter = openpyxl.utils.get_column_letter(target_idx)
                            row2_val = ws.cell(row=2, column=target_idx).value
                            suffix = f"_{row2_val}" if row2_val is not None else ""
                            filename = f"output_column_{col_letter}{suffix}.csv"
                            
                            # --- データ抽出（Excelの順番を絶対維持） ---
                            final_rows = []
                            seen_full_row = set() # 「名前＋数値」の重複チェック用
                            
                            # 指定された開始行から、Excelの最後の行までループ
                            for r in range(skip_rows + 1, ws.max_row + 1):
                                name = ws.cell(row=r, column=anchor_idx).value
                                val = ws.cell(row=r, column=target_idx).value
                                
                                # 店舗名が完全に空の行は、Excel末尾の不要な行としてスキップ
                                if name is None or str(name).strip() == "":
                                    continue
                                
                                # 重複チェック（名前と数値がどちらも同じ場合のみ飛ばす）
                                # タプルにしてsetで管理（順番は変えない）
                                row_identifier = (str(name).strip(), str(val).strip())
                                
                                if row_identifier not in seen_full_row:
                                    seen_full_row.add(row_identifier)
                                    final_rows.append([name, val])

                            # CSV作成
                            output = io.StringIO()
                            writer = csv.writer(output, quoting=csv.QUOTE_MINIMAL)
                            for row in final_rows:
                                writer.writerow(row)
                            
                            myzip.writestr(filename, output.getvalue().encode('utf-8-sig'))
                            output.close()
                    
                    st.success("✅ 完了しました。Excelの順番通りです。")
                    st.download_button(label="📥 ダウンロード", data=zip_buffer.getvalue(), file_name="処理結果.zip")
    except Exception as e:
        st.error(f"エラー: {e}")
