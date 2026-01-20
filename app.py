# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import openpyxl
import io
import zipfile

# === アプリのタイトル ===
st.title("📂 Excel列分解＆CSV作成ツール")
st.write("基準列と数式列をペアにしてCSV化します。（店舗名の整合性チェック付き）")

# === サイドバー：設定エリア ===
st.sidebar.header("⚙️ 設定")

# 1. 基準列の設定
st.sidebar.subheader("1. 基準列（店舗名）")
anchor_col = st.sidebar.text_input("固定して使う列 (例: A)", value="A")

# 2. 行の削除設定
st.sidebar.subheader("2. 行の削除")
skip_rows = st.sidebar.number_input("最初に削除する行数", min_value=0, value=2)

# 3. 列の除外設定
st.sidebar.subheader("3. 列の除外設定")
ignore_col_start = st.sidebar.text_input("除外したい開始列 (例: B)", value="")
ignore_col_end = st.sidebar.text_input("除外したい終了列 (例: G)", value="")

# === メインエリア ===
uploaded_file = st.file_uploader("Excelファイルをアップロード", type=['xlsx'])

if uploaded_file:
    try:
        file_bytes = uploaded_file.getvalue()
        wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=False)
        ws = wb.active
        
        start_row = skip_rows + 1
        
        try:
            anchor_idx = openpyxl.utils.column_index_from_string(anchor_col) - 1
        except:
            st.error("基準列の指定が間違っています。")
            st.stop()

        # --- 基準となる店舗リストの作成（重複を除いた正しい順番を保持） ---
        df_full = pd.read_excel(
            io.BytesIO(file_bytes), 
            header=None, 
            skiprows=skip_rows, 
            engine='openpyxl'
        )
        # 元データのA列から、重複を除いた「本来あるべき店舗の並び順」を取得
        original_series = df_full.iloc[:, anchor_idx].astype(str).str.strip()
        base_store_list = original_series.drop_duplicates(keep='first').tolist()
        
        st.info(f"📊 抽出対象の総店舗数: {len(base_store_list)} 件")

        # --- 数式列の自動検出 ---
        ignore_indices = []
        if ignore_col_start and ignore_col_end:
            try:
                start_ignore = openpyxl.utils.column_index_from_string(ignore_col_start)
                end_ignore = openpyxl.utils.column_index_from_string(ignore_col_end)
                ignore_indices = list(range(start_ignore - 1, end_ignore))
            except:
                pass

        formula_candidates = []
        max_check = min(start_row + 10, ws.max_row)
        for col_idx_1based in range(1, ws.max_column + 1): 
            col_idx_0based = col_idx_1based - 1
            if col_idx_0based == anchor_idx or col_idx_0based in ignore_indices:
                continue
            is_formula = False
            for r in range(start_row, max_check + 1):
                cell = ws.cell(row=r, column=col_idx_1based)
                if cell.data_type == 'f' or (str(cell.value).startswith('=')):
                    is_formula = True
                    break
            if is_formula:
                col_letter = openpyxl.utils.get_column_letter(col_idx_1based)
                formula_candidates.append({"idx": col_idx_0based, "name": col_letter})

        if formula_candidates:
            st.success(f"✅ {len(formula_candidates)} 個の数式列が見つかりました。")
            
            st.subheader("🛠️ 出力する列を選択")
            cols = st.columns(4)
            selected_indices = []
            for i, candidate in enumerate(formula_candidates):
                with cols[i % 4]:
                    if st.checkbox(f"{candidate['name']} 列", value=True, key=candidate['idx']):
                        selected_indices.append(candidate['idx'])

            # --- CSV作成実行 ---
            if st.button("🚀 選択した列のCSVを作成"):
                if not selected_indices:
                    st.error("列が選択されていません。")
                else:
                    with st.spinner('作成中...'):
                        zip_buffer = io.BytesIO()
                        check_passed = True
                        error_cols = []

                        with zipfile.ZipFile(zip_buffer, 'w') as myzip:
                            for target_idx in selected_indices:
                                # 1. ファイル名作成 (H列_2行目の値.csv)
                                col_letter = openpyxl.utils.get_column_letter(target_idx + 1)
                                cell_value_row2 = ws.cell(row=2, column=target_idx + 1).value
                                suffix = f"_{cell_value_row2}" if cell_value_row2 is not None else ""
                                filename = f"output_column_{col_letter}{suffix}.csv"
                                
                                # 2. データ抽出と重複削除（最初を残す）
                                output_df = df_full.iloc[:, [anchor_idx, target_idx]].copy()
                                output_df.iloc[:, 0] = output_df.iloc[:, 0].astype(str).str.strip()
                                output_df = output_df.drop_duplicates(keep='first')

                                # 3. 内部チェック：店舗リストと順番が一致するか
                                current_store_list = output_df.iloc[:, 0].tolist()
                                if current_store_list != base_store_list:
                                    check_passed = False
                                    error_cols.append(col_letter)

                                # 4. 書き出し
                                csv_data = output_df.to_csv(header=False, index=False, encoding='utf-8-sig')
                                myzip.writestr(filename, csv_data)
                        
                        if check_passed:
                            st.success(f"✅ チェック完了: すべての店舗（{len(base_store_list)}件）が正しい順番で出力されました。")
                        else:
                            st.warning(f"⚠️ 一部の列（{', '.join(error_cols)}）で、元の店舗リストと件数や順番が一致しませんでした。出力結果を確認してください。")

                        st.download_button(
                            label="📥 ZIPファイルをダウンロード",
                            data=zip_buffer.getvalue(),
                            file_name="処理結果.zip",
                            mime="application/zip"
                        )
    except Exception as e:
        st.error(f"エラーが発生しました: {e}")
