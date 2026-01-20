# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import openpyxl
import io
import zipfile

# === アプリのタイトル ===
st.title("📂 Excel列分解＆CSV作成ツール")
st.write("基準となる列（店舗名など）と、数式列をペアにしてCSV化します。")

# === サイドバー：設定エリア ===
st.sidebar.header("⚙️ 設定")

# 1. 基準列の設定
st.sidebar.subheader("1. 基準列（店舗名）")
anchor_col = st.sidebar.text_input(
    "固定して使う列 (例: A)",
    value="A"
)

# 2. 行の削除設定
st.sidebar.subheader("2. 行の削除")
skip_rows = st.sidebar.number_input(
    "最初に削除する行数",
    min_value=0,
    value=2
)

# 3. 重複削除の設定（追加機能）
st.sidebar.subheader("3. データの整理")
remove_dup = st.sidebar.checkbox("重複した行を自動で削除する", value=True)

# 4. 列の除外設定
st.sidebar.subheader("4. 列の除外設定")
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
        max_check = min(start_row + 10, ws.max_row)
        
        try:
            anchor_idx = openpyxl.utils.column_index_from_string(anchor_col) - 1
        except:
            st.error("基準列の指定が間違っています。")
            st.stop()

        ignore_indices = []
        if ignore_col_start and ignore_col_end:
            try:
                start_ignore = openpyxl.utils.column_index_from_string(ignore_col_start)
                end_ignore = openpyxl.utils.column_index_from_string(ignore_col_end)
                ignore_indices = list(range(start_ignore - 1, end_ignore))
            except:
                st.error("除外列の指定が間違っています。")

        # --- 数式列の検出 ---
        formula_candidates = []
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

        if not formula_candidates:
            st.warning("⚠️ 数式列が見つかりませんでした。")
        else:
            st.success(f"✅ {len(formula_candidates)} 個の数式列が見つかりました！")
            
            st.subheader("🛠️ 出力する列を選択")
            cols = st.columns(4)
            selected_indices = []
            for i, candidate in enumerate(formula_candidates):
                with cols[i % 4]:
                    if st.checkbox(f"{candidate['name']} 列", value=True, key=candidate['idx']):
                        selected_indices.append(candidate['idx'])

            # --- CSV作成 ---
            st.markdown("---")
            if st.button("🚀 選択した列のCSVを作成"):
                if not selected_indices:
                    st.error("列が選択されていません。")
                else:
                    with st.spinner('CSVを作成中...'):
                        # 抽出用に pandas で読み込み
                        df = pd.read_excel(
                            io.BytesIO(file_bytes), 
                            header=None, 
                            skiprows=skip_rows, 
                            engine='openpyxl'
                        )
                        
                        max_idx = len(df.columns) - 1
                        zip_buffer = io.BytesIO()

                        with zipfile.ZipFile(zip_buffer, 'w') as myzip:
                            for target_idx in selected_indices:
                                col_letter = openpyxl.utils.get_column_letter(target_idx + 1)
                                
                                # ファイル名の設定（2行目の値を取得）
                                cell_value_row2 = ws.cell(row=2, column=target_idx + 1).value
                                suffix = f"_{cell_value_row2}" if cell_value_row2 is not None else ""
                                filename = f"output_column_{col_letter}{suffix}.csv"
                                
                                if target_idx <= max_idx:
                                    # 抽出
                                    output_df = df.iloc[:, [anchor_idx, target_idx]]
                                    
                                    # 【修正ポイント】重複行を削除
                                    if remove_dup:
                                        output_df = output_df.drop_duplicates()
                                    
                                    # 空白行（店舗名が空の行など）も除外したい場合はここに追加
                                    # output_df = output_df.dropna(subset=[output_df.columns[0]])

                                    csv_data = output_df.to_csv(header=False, index=False, encoding='utf-8-sig')
                                    myzip.writestr(filename, csv_data)
                        
                        st.success("完了しました！")
                        st.download_button(
                            label="📥 ZIPファイルをダウンロード",
                            data=zip_buffer.getvalue(),
                            file_name="処理結果.zip",
                            mime="application/zip"
                        )

    except Exception as e:
        st.error(f"エラーが発生しました: {e}")
