# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import openpyxl
import io
import zipfile

# === アプリのタイトルと説明 ===
st.title("📂 Excel列分解＆CSV作成ツール")
st.write("Excelファイルをアップロードすると、数式が入っている列を自動検出し、A列（店舗名）とセットにしてCSV化します。")

# === サイドバー：詳細設定エリア ===
st.sidebar.header("⚙️ 詳細設定")

# 1. 行の削除設定
st.sidebar.subheader("1. 行の削除")
skip_rows = st.sidebar.number_input(
    "最初に削除する行数",
    min_value=0,
    value=2,
    help="データが始まる前の不要なヘッダー行数を指定します。"
)

# 2. 列の除外設定
st.sidebar.subheader("2. 列の除外設定")
st.sidebar.write("数式列の検出対象から外したい列があれば指定してください。")
ignore_col_start = st.sidebar.text_input("除外したい開始列 (例: B)", value="")
ignore_col_end = st.sidebar.text_input("除外したい終了列 (例: G)", value="")

# === メインエリア：ファイルアップロード ===
uploaded_file = st.file_uploader("Excelファイルをアップロード", type=['xlsx'])

if uploaded_file:
    try:
        # Excel読み込みの前処理
        file_bytes = uploaded_file.getvalue()
        
        # Openpyxlで開いて分析
        wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=False)
        ws = wb.active
        
        # データの開始行
        start_row = skip_rows + 1
        max_check = min(start_row + 10, ws.max_row)
        
        # 除外列の範囲を計算
        ignore_indices = []
        if ignore_col_start and ignore_col_end:
            try:
                start_idx = openpyxl.utils.column_index_from_string(ignore_col_start)
                end_idx = openpyxl.utils.column_index_from_string(ignore_col_end)
                ignore_indices = list(range(start_idx, end_idx + 1))
                st.info(f"ℹ️ {ignore_col_start}列 から {ignore_col_end}列 は無視します。")
            except:
                st.error("列の指定が間違っています（半角大文字で入力してください）。")

        # 数式列の候補を探す
        formula_candidates = []
        for col_idx in range(2, ws.max_column + 1): # B列(2)以降
            if col_idx in ignore_indices:
                continue

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
            st.success(f"✅ {len(formula_candidates)} 個の数式列が見つかりました！")
            
            # --- 出力する列を選ぶエリア ---
            st.subheader("🛠️ 出力する列を選択")
            st.write("チェックを外した列はCSVに出力されません。")
            
            cols = st.columns(4)
            selected_indices = []
            
            for i, candidate in enumerate(formula_candidates):
                col_name = candidate["name"]
                col_idx = candidate["idx"]
                with cols[i % 4]:
                    if st.checkbox(f"{col_name} 列", value=True, key=col_idx):
                        selected_indices.append(col_idx)

            # --- CSV作成ボタン ---
            st.markdown("---")
            if st.button("🚀 選択した列のCSVを作成"):
                if not selected_indices:
                    st.error("出力する列が一つも選ばれていません。")
                else:
                    with st.spinner('CSVを作成中...'):
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
                                # エラーが出ていた箇所を修正（改行せず1行で記述）
                                col_name = openpyxl.utils.get_column_letter(col_idx + 1)
                                
                                if col_idx < len(df.columns):
                                    output_df = df.iloc[:, [0, col_idx]]
                                    filename = f"output_column_{col_name}.csv"
                                    csv_data = output_df.to_csv(header=False, index=False, encoding='utf-8-sig')
                                    myzip.writestr(filename, csv_data)
                        
                        st.success("完了しました！下のボタンからダウンロードしてください。")
                        st.download_button(
                            label="📥 ZIPファイルをダウンロード",
                            data=zip_buffer.getvalue(),
                            file_name="処理結果.zip",
                            mime="application/zip"
                        )

    except Exception as e:
        st.error(f"エラーが発生しました: {e}")
