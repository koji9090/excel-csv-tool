# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import openpyxl
import io
import zipfile

# === アプリのタイトル ===
st.title("📂 Excel列分解＆CSV作成ツール")
st.write("基準列（店舗名）の順番を維持したままCSV化します。")

# === サイドバー：設定エリア ===
st.sidebar.header("⚙️ 設定")

# 1. 基準列の設定
anchor_col = st.sidebar.text_input("1. 固定して使う基準列 (例: A)", value="A")

# 2. 行の削除設定
skip_rows = st.sidebar.number_input("2. 最初に削除する行数", min_value=0, value=2)

# 3. 列の除外設定
ignore_col_start = st.sidebar.text_input("3. 除外したい開始列 (例: B)", value="")
ignore_col_end = st.sidebar.text_input("4. 除外したい終了列 (例: G)", value="")

# === メインエリア ===
uploaded_file = st.file_uploader("Excelファイルをアップロード", type=['xlsx'])

if uploaded_file:
    try:
        file_bytes = uploaded_file.getvalue()
        wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=False)
        ws = wb.active
        
        # 基準列のインデックス
        anchor_idx = openpyxl.utils.column_index_from_string(anchor_col) - 1

        # --- 【重要】基準となる正しい店舗リスト（順番）を生成 ---
        df_full = pd.read_excel(
            io.BytesIO(file_bytes), 
            header=None, 
            skiprows=skip_rows, 
            engine='openpyxl'
        )
        # A列の生データを取得してクレンジング
        raw_stores = df_full.iloc[:, anchor_idx].astype(str).str.strip()
        # 重複を除いた「正しい順番」をマスターとする
        master_store_order = raw_stores.drop_duplicates(keep='first').tolist()

        # --- 数式列の検出 ---
        ignore_indices = []
        if ignore_col_start and ignore_col_end:
            try:
                s = openpyxl.utils.column_index_from_string(ignore_col_start) - 1
                e = openpyxl.utils.column_index_from_string(ignore_col_end)
                ignore_indices = list(range(s, e))
            except: pass

        formula_candidates = []
        for c in range(1, ws.max_column + 1):
            if (c-1) == anchor_idx or (c-1) in ignore_indices: continue
            # データの開始行付近をチェック
            is_f = False
            for r in range(skip_rows + 1, min(skip_rows + 10, ws.max_row) + 1):
                cell = ws.cell(row=r, column=c)
                if cell.data_type == 'f' or str(cell.value).startswith('='):
                    is_f = True; break
            if is_f:
                formula_candidates.append({"idx": c-1, "name": openpyxl.utils.get_column_letter(c)})

        if formula_candidates:
            st.success(f"✅ {len(formula_candidates)} 個の数式列を検出しました。")
            st.subheader("🛠️ 出力する列を選択")
            selected_indices = []
            cols = st.columns(4)
            for i, cand in enumerate(formula_candidates):
                with cols[i % 4]:
                    if st.checkbox(f"{cand['name']} 列", value=True, key=cand['idx']):
                        selected_indices.append(cand['idx'])

            # --- CSV作成実行 ---
            if st.button("🚀 CSVを作成してチェックを実行"):
                if not selected_indices:
                    st.error("列が選択されていません。")
                else:
                    zip_buffer = io.BytesIO()
                    error_count = 0  # 順番や件数が狂った列を数える

                    with zipfile.ZipFile(zip_buffer, 'w') as myzip:
                        for target_idx in selected_indices:
                            col_letter = openpyxl.utils.get_column_letter(target_idx + 1)
                            
                            # ファイル名 (2行目の値を取得)
                            row2_val = ws.cell(row=2, column=target_idx + 1).value
                            filename = f"output_column_{col_letter}_{row2_val}.csv" if row2_val else f"output_column_{col_letter}.csv"
                            
                            # データ抽出と重複削除
                            output_df = df_full.iloc[:, [anchor_idx, target_idx]].copy()
                            output_df.iloc[:, 0] = output_df.iloc[:, 0].astype(str).str.strip()
                            output_df = output_df.drop_duplicates(keep='first')

                            # 【順番と件数のチェック】
                            current_list = output_df.iloc[:, 0].tolist()
                            if current_list != master_store_order:
                                error_count += 1
                            
                            csv_data = output_df.to_csv(header=False, index=False, encoding='utf-8-sig')
                            myzip.writestr(filename, csv_data)
                    
                    # 結果表示
                    if error_count == 0:
                        st.success(f"✨ 全 {len(selected_indices)} ファイル、順番・件数ともに完璧に一致しました！")
                    else:
                        st.error(f"⚠️ {error_count} 個のファイルで店舗名の順番または件数がExcelと異なります。")
                    
                    st.download_button(
                        label="📥 ZIPファイルをダウンロード",
                        data=zip_buffer.getvalue(),
                        file_name="処理結果.zip",
                        mime="application/zip"
                    )
    except Exception as e:
        st.error(f"エラー: {e}")
