# -*- coding: utf-8 -*-
import streamlit as st
import openpyxl
import io
import zipfile
import csv

# === アプリのタイトル ===
st.title("📂 Excel列分解＆CSV作成ツール")
st.write("Excelの行順序を絶対的に維持し、重複を排除してCSV化します。")

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
        # 1. データを読み込む (計算済みの値を取得)
        file_bytes = uploaded_file.getvalue()
        wb_data = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
        ws = wb_data.active
        
        # 2. 数式を確認するために別途読み込む
        wb_formula = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=False)
        ws_f = wb_formula.active
        
        # 基準列のインデックス (Aなら1)
        anchor_idx = openpyxl.utils.column_index_from_string(anchor_col_letter)

        # --- 除外設定 ---
        ignore_indices = []
        if ignore_col_start and ignore_col_end:
            try:
                s = openpyxl.utils.column_index_from_string(ignore_col_start)
                e = openpyxl.utils.column_index_from_string(ignore_col_end)
                ignore_indices = list(range(s, e + 1))
            except: pass

        # --- 数式列の検出 ---
        formula_candidates = []
        check_start = skip_rows + 1
        check_end = min(check_start + 10, ws_f.max_row)
        for c in range(1, ws_f.max_column + 1):
            if c == anchor_idx or c in ignore_indices: continue
            
            is_f = False
            for r in range(check_start, check_end + 1):
                cell = ws_f.cell(row=r, column=c)
                if cell.data_type == 'f' or (cell.value and str(cell.value).startswith('=')):
                    is_f = True; break
            if is_f:
                formula_candidates.append({"idx": c, "name": openpyxl.utils.get_column_letter(c)})

        if formula_candidates:
            st.success(f"✅ {len(formula_candidates)} 個の数式列を検出しました。")
            st.subheader("🛠️ 出力する列を選択")
            selected_cols = []
            cols_ui = st.columns(4)
            for i, cand in enumerate(formula_candidates):
                with cols_ui[i % 4]:
                    if st.checkbox(f"{cand['name']} 列", value=True, key=cand['idx']):
                        selected_cols.append(cand['idx'])

            # --- CSV作成実行 ---
            if st.button("🚀 CSVを作成"):
                if not selected_cols:
                    st.error("列を選択してください。")
                else:
                    with st.spinner('Excelの行を順番に解析中...'):
                        zip_buffer = io.BytesIO()
                        with zipfile.ZipFile(zip_buffer, 'w') as myzip:
                            
                            for target_idx in selected_cols:
                                # ファイル名の設定: 2行目の値を取得
                                col_letter = openpyxl.utils.get_column_letter(target_idx)
                                row2_val = ws.cell(row=2, column=target_idx).value
                                suffix = f"_{row2_val}" if row2_val else ""
                                filename = f"output_column_{col_letter}{suffix}.csv"
                                
                                # --- 行の抽出（ここが順番維持の核心） ---
                                rows_to_csv = []
                                seen_names_normalized = set() # 重複判定用
                                
                                # skip_rowsの次の行から、Excelの最終行まで順番に走査
                                for r in range(skip_rows + 1, ws.max_row + 1):
                                    store_name_raw = ws.cell(row=r, column=anchor_idx).value
                                    value_data = ws.cell(row=r, column=target_idx).value
                                    
                                    if store_name_raw is None:
                                        continue # 名前が空ならスキップ
                                    
                                    # 判定用に名前を「超正規化」する
                                    # 全角半角スペース、改行をすべて除去して比較
                                    name_str = str(store_name_raw)
                                    normalized_name = "".join(name_str.split()) 
                                    
                                    if normalized_name == "":
                                        continue

                                    # 初めて出た名前なら採用（Excelの上にある行が優先される）
                                    if normalized_name not in seen_names_normalized:
                                        seen_names_normalized.add(normalized_name)
                                        # 保存するのは「正規化前」の元の文字とデータ
                                        rows_to_csv.append([store_name_raw, value_data])

                                # CSV文字列の生成
                                output = io.StringIO()
                                writer = csv.writer(output, quoting=csv.QUOTE_MINIMAL)
                                for row in rows_to_csv:
                                    writer.writerow(row)
                                
                                # ZIPに追加 (BOM付きUTF-8)
                                myzip.writestr(filename, output.getvalue().encode('utf-8-sig'))
                                output.close()
                        
                        st.success("✅ 完了しました。Excelの行順序を100%維持して作成しました。")
                        st.download_button(label="📥 ダウンロード", data=zip_buffer.getvalue(), file_name="処理結果.zip")
        else:
            st.warning("数式列が見つかりませんでした。")

    except Exception as e:
        st.error(f"エラーが発生しました: {e}")
