# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import openpyxl
import io
import zipfile

# === アプリのタイトル ===
st.title("📂 Excel列分解＆CSV作成ツール")
st.write("Excelの行順序を完全に維持してCSV化します。")

# === サイドバー：設定エリア ===
st.sidebar.header("⚙️ 設定")
anchor_col = st.sidebar.text_input("1. 基準列 (例: A)", value="A")
skip_rows = st.sidebar.number_input("2. 削除する行数", min_value=0, value=2)
ignore_col_start = st.sidebar.text_input("3. 除外開始列", value="")
ignore_col_end = st.sidebar.text_input("4. 除外終了列", value="")

# === メインエリア ===
uploaded_file = st.file_uploader("Excelファイルをアップロード", type=['xlsx'])

if uploaded_file:
    try:
        file_bytes = uploaded_file.getvalue()
        wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=False)
        ws = wb.active
        
        # 基準列のインデックス（Aなら0）
        try:
            anchor_idx = openpyxl.utils.column_index_from_string(anchor_col) - 1
        except:
            st.error("基準列の指定が正しくありません。")
            st.stop()

        # --- 【重要】Excelを読み込み、元の行番号を保持する ---
        df_full = pd.read_excel(
            io.BytesIO(file_bytes), 
            header=None, 
            skiprows=skip_rows, 
            engine='openpyxl'
        )
        # 行番号をリセットして、元の並び順（0, 1, 2...）をインデックスに固定
        df_full = df_full.reset_index(drop=True)

        # 店舗名のクレンジング（重複判定のミスを防ぐため）
        df_full.iloc[:, anchor_idx] = df_full.iloc[:, anchor_idx].astype(str).str.strip()

        # --- 数式列の検出 ---
        ignore_indices = []
        if ignore_col_start and ignore_col_end:
            try:
                s = openpyxl.utils.column_index_from_string(ignore_col_start) - 1
                e = openpyxl.utils.column_index_from_string(ignore_col_end)
                ignore_indices = list(range(s, e))
            except: pass

        formula_candidates = []
        max_check = min(skip_rows + 10, ws.max_row)
        for c in range(1, ws.max_column + 1):
            if (c-1) == anchor_idx or (c-1) in ignore_indices: continue
            is_f = False
            for r in range(skip_rows + 1, max_check + 1):
                cell = ws.cell(row=r, column=c)
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
                    with st.spinner('作成中...'):
                        zip_buffer = io.BytesIO()
                        with zipfile.ZipFile(zip_buffer, 'w') as myzip:
                            for target_idx in selected_indices:
                                # ファイル名 (列記号 + 2行目の値)
                                col_letter = openpyxl.utils.get_column_letter(target_idx + 1)
                                row2_val = ws.cell(row=2, column=target_idx + 1).value
                                suffix = f"_{row2_val}" if row2_val is not None else ""
                                filename = f"output_column_{col_letter}{suffix}.csv"
                                
                                # データの抽出
                                output_df = df_full.iloc[:, [anchor_idx, target_idx]].copy()
                                
                                # 【重複削除】店舗名（anchor列）で重複を削除。keep='first'で最初の行を残す。
                                # Pandasのインデックス（元の行番号）はそのまま維持される
                                output_df = output_df.drop_duplicates(subset=output_df.columns[0], keep='first')

                                # 【並び替え】念のため、元のインデックス（行番号）順に並び替えてExcelの順番を保証する
                                output_df = output_df.sort_index()

                                # CSV書き出し
                                csv_data = output_df.to_csv(header=False, index=False, encoding='utf-8-sig')
                                myzip.writestr(filename, csv_data)
                        
                        st.success("✅ 完了しました！Excelの順序通りに出力されています。")
                        st.download_button(label="📥 ダウンロード", data=zip_buffer.getvalue(), file_name="処理結果.zip")
        else:
            st.warning("数式列が見つかりませんでした。")
    except Exception as e:
        st.error(f"エラー: {e}")
