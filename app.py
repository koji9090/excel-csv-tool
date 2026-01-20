# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import openpyxl
import io
import zipfile

# === アプリのタイトル ===
st.title("📂 Excel列分解ツール (最終行優先版)")
st.write("店舗名の重複は「最後の行」を優先して残し、エクセルの順番通りに出力します。")

# === サイドバー：設定エリア ===
st.sidebar.header("⚙️ 設定")
anchor_col_letter = st.sidebar.text_input("1. 基準列 (例: A)", value="A")
skip_rows = st.sidebar.number_input("2. 除外する行数", min_value=0, value=2)

# === メインエリア ===
uploaded_file = st.file_uploader("Excelファイルをアップロード", type=['xlsx'])

if uploaded_file:
    try:
        file_bytes = uploaded_file.getvalue()
        
        # 1. データ読み込み（計算後の値）
        df_raw = pd.read_excel(io.BytesIO(file_bytes), header=None, engine='openpyxl')
        
        # 2. ファイル名取得用 (2行目の値)
        wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
        ws = wb.active
        
        # 3. 数式検出用
        wb_f = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=False)
        ws_f = wb_f.active

        anchor_idx = openpyxl.utils.column_index_from_string(anchor_col_letter) - 1

        # --- 数式列の検出 ---
        formula_candidates = []
        check_start = skip_rows + 1
        check_end = min(check_start + 10, ws_f.max_row)
        for c in range(1, ws_f.max_column + 1):
            if (c-1) == anchor_idx: continue
            is_f = False
            for r in range(check_start, check_end + 1):
                cell = ws_f.cell(row=r, column=c)
                if cell.data_type == 'f' or (cell.value and str(cell.value).startswith('=')):
                    is_f = True; break
            if is_f:
                formula_candidates.append({"idx": c-1, "name": openpyxl.utils.get_column_letter(c)})

        if formula_candidates:
            st.success(f"✅ {len(formula_candidates)} 個の数式列を検出しました。")
            selected_indices = []
            cols_ui = st.columns(4)
            for i, cand in enumerate(formula_candidates):
                with cols_ui[i % 4]:
                    if st.checkbox(f"{cand['name']} 列", value=True, key=cand['idx']):
                        selected_indices.append(cand['idx'])

            # --- CSV作成実行 ---
            if st.button("🚀 CSVを作成"):
                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, 'w') as myzip:
                    
                    for target_idx in selected_indices:
                        col_letter = openpyxl.utils.get_column_letter(target_idx + 1)
                        row2_val = ws.cell(row=2, column=target_idx + 1).value
                        suffix = f"_{row2_val}" if row2_val is not None else ""
                        filename = f"output_column_{col_letter}{suffix}.csv"
                        
                        # データ抽出
                        df_data = df_raw.iloc[skip_rows:].copy()
                        df_target = df_data.iloc[:, [anchor_idx, target_idx]]
                        
                        # 店舗名のクリーニング
                        df_target.iloc[:, 0] = df_target.iloc[:, 0].astype(str).str.strip()
                        # 空行の除外
                        df_target = df_target[df_target.iloc[:, 0] != "nan"]
                        df_target = df_target[df_target.iloc[:, 0] != ""]

                        # 【ここが最重要：修正ポイント】
                        # keep='last' を指定することで、重複がある場合は「下の行」を残す。
                        # これで、最後にあるべき「BYD AUTO 東京品川」が正しく残ります。
                        df_target = df_target.drop_duplicates(subset=df_target.columns[0], keep='last')

                        # 出力（並び順はエクセルの出現順を維持）
                        csv_data = df_target.to_csv(header=False, index=False, encoding='utf-8-sig')
                        myzip.writestr(filename, csv_data)
                
                st.success("✅ 完了しました！重複は「下の行」を優先して1つにまとめました。")
                st.download_button("📥 ダウンロード", data=zip_buffer.getvalue(), file_name="処理結果.zip")
        else:
            st.warning("数式列が見つかりません。")

    except Exception as e:
        st.error(f"エラーが発生しました: {e}")
