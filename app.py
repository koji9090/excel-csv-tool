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

# 1. 基準列の設定（新機能）
st.sidebar.subheader("1. 基準列（店舗名）")
anchor_col = st.sidebar.text_input(
    "固定して使う列 (例: A)",
    value="A",
    help="すべてのCSVの左側に入る列です。通常は店舗名などの列を指定します。"
)

# 2. 行の削除設定
st.sidebar.subheader("2. 行の削除")
skip_rows = st.sidebar.number_input(
    "最初に削除する行数",
    min_value=0,
    value=2,
    help="データが始まる前の不要なヘッダー行数を指定します。"
)

# 3. 列の除外設定
st.sidebar.subheader("3. 列の除外設定")
st.sidebar.write("数式列の検出対象から外したい列があれば指定してください。")
ignore_col_start = st.sidebar.text_input("除外したい開始列 (例: B)", value="")
ignore_col_end = st.sidebar.text_input("除外したい終了列 (例: G)", value="")

# === メインエリア ===
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
        
        # --- 設定値の計算 ---
        # 基準列（アンカー）のインデックス化
        try:
            anchor_idx = openpyxl.utils.column_index_from_string(anchor_col) - 1
        except:
            st.error("基準列の指定が間違っています（半角大文字で入力してください）。")
            st.stop()

        # 除外列の範囲を計算
        ignore_indices = []
        if ignore_col_start and ignore_col_end:
            try:
                start_ignore = openpyxl.utils.column_index_from_string(ignore_col_start)
                end_ignore = openpyxl.utils.column_index_from_string(ignore_col_end)
                ignore_indices = list(range(start_ignore - 1, end_ignore)) # 0始まりに合わせる
                st.info(f"ℹ️ {ignore_col_start}列 から {ignore_col_end}列 は検索対象から外します。")
            except:
                st.error("除外列の指定が間違っています。")

        # --- 数式列の検出ループ ---
        formula_candidates = []
        
        # 全列を走査（1列目から最終列まで）
        for col_idx_1based in range(1, ws.max_column + 1): 
            col_idx_0based = col_idx_1based - 1
            
            # 1. 基準列（店舗名）自体は数式チェックの対象外
            if col_idx_0based == anchor_idx:
                continue

            # 2. 除外リストに含まれていたらスキップ
            if col_idx_0based in ignore_indices:
                continue

            # 3. 数式チェック
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
            st.warning("⚠️ 数式が入っている列が見つかりませんでした。設定を確認してください。")
        else:
            st.success(f"✅ {len(formula_candidates)} 個の数式列が見つかりました！")
            
            # --- 出力選択 ---
            st.subheader("🛠️ 出力する列を選択")
            st.write(f"基準列【 {anchor_col}列 】とペアにして出力します。")
            
            cols = st.columns(4)
            selected_indices = []
            
            for i, candidate in enumerate(formula_candidates):
                col_name = candidate["name"]
                col_idx = candidate["idx"]
                with cols[i % 4]:
                    if st.checkbox(f"{col_name} 列", value=True, key=col_idx):
                        selected_indices.append(col_idx)

            # --- CSV作成 ---
            st.markdown("---")
            if st.button("🚀 選択した列のCSVを作成"):
                if not selected_indices:
                    st.error("出力する列が選ばれていません。")
                else:
                    with st.spinner('CSVを作成中...'):
                        # データ読み込み
                        df = pd.read_excel(
                            io.BytesIO(file_bytes), 
                            header=None, 
                            skiprows=skip_rows, 
                            engine='openpyxl'
                        )
                        
                        # 列の範囲チェック（エラー防止）
                        max_idx = len(df.columns) - 1
                        if anchor_idx > max_idx:
                            st.error(f"エラー：基準列（{anchor_col}）がデータ範囲外です。")
                            st.stop()

                        zip_buffer = io.BytesIO()
                        with zipfile.ZipFile(zip_buffer, 'w') as myzip:
                            for target_idx in selected_indices:
                                col_name = openpyxl.utils.get_column_letter(target_idx + 1)
                                
                                if target_idx <= max_idx:
                                    # 基準列 と ターゲット列 を抽出
                                    output_df = df.iloc[:, [anchor_idx, target_idx]]
                                    
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
