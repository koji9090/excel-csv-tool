# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import openpyxl
import io
import zipfile

# === アプリのタイトルと説明 ===
st.title("📂 Excel列分解＆CSV作成ツール")
st.write("Excelファイルをアップロードすると、数式が入っている列を自動検出し、A列（店舗名）とセットにしてCSV化します。")

# === サイドバー：設定エリア ===
st.sidebar.header("⚙️ 設定")

# 1. 行数の指定機能
skip_rows = st.sidebar.number_input(
    "最初に削除する行数（ヘッダー上の不要行）",
    min_value=0,
    value=2,
    help="データが始まる前の不要な行数を指定します。"
)

# === メインエリア：ファイルアップロード ===
uploaded_file = st.file_uploader("ExcelまたはCSVファイルをアップロード", type=['xlsx', 'csv'])

if uploaded_file:
    try:
        # Excelとして読み込むための前処理
        file_bytes = uploaded_file.getvalue()
        
        # Openpyxlで開いて数式列を探す
        wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=False)
        ws = wb.active
        
        # データの開始行
        start_row = skip_rows + 1
        max_check = min(start_row + 10, ws.max_row)
        
        # 数式列の候補を探す
        formula_candidates = []
        for col_idx in range(2, ws.max_column + 1): # B列(2)以降
            is_formula = False
            for r in range(start_row, max_check + 1):
                cell = ws.cell(row=r, column=col_idx)
                if cell.data_type == 'f' or (str(cell.value).startswith('=')):
                    is_formula = True
                    break
            
            if is_formula:
                col_letter = openpyxl.utils.get_column_letter(col_idx)
                formula_candidates.append({"idx": col_idx - 1, "name": col_letter})

        if not formula_candidates:
            st.warning("⚠️ 数式が入っている列が見つかりませんでした。")
        else:
            st.success(f"✅ {len(formula_candidates)} 個の数式列を検出しました！")

            # 2. 列の選択機能
            st.subheader("出力する列を選択")
            options = [c["name"] for c in formula_candidates]
            selected_names = st.multiselect(
                "必要な列だけ残してください:",
                options=options,
                default=options
            )
            selected_indices = [c["idx"] for c in formula_candidates if c["name"] in selected_names]

            if st.button("🚀 CSVを作成してダウンロード"):
                if not selected_indices:
                    st.error("列が選択されていません。")
                else:
                    # データ読み込み
                    df = pd.read_excel(
                        io.BytesIO(file_bytes), 
                        header=None, 
                        skiprows=skip_rows, 
                        engine='openpyxl'
                    )
                    
                    # ZIP作成
                    zip_buffer = io.BytesIO()
                    with zipfile.ZipFile(zip_buffer, 'w') as myzip:
                        for col_idx in selected_indices:
                            col_name = openpyxl.utils.get_column_letter(col_idx + 1)
                            if col_idx < len(df.columns):
                                output_df = df.iloc[:, [0, col_idx]]
                                filename = f"output_column_{col_name}.csv"
                                csv_data = output_df.to_csv(header=False, index=False, encoding='utf-8-sig')
                                myzip.writestr(filename, csv_data)
                    
                    st.download_button(
                        label="📥 ZIPファイルをダウンロード",
                        data=zip_buffer.getvalue(),
                        file_name="処理結果.zip",
                        mime="application/zip"
                    )

    except Exception as e:
        st.error(f"エラーが発生しました: {e}")
