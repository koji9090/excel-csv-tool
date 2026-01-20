# -*- coding: utf-8 -*-
import streamlit as st
import openpyxl
import io
import zipfile
import csv

st.set_page_config(page_title="Excel CSV Tool", layout="wide")

# === アプリのタイトル ===
st.title("📂 Excel列分解ツール (完全順序維持版)")
st.write("エクセルの上から下までの並び順を「1行も入れ替えず」にCSV化します。")

# === サイドバー：設定エリア ===
st.sidebar.header("⚙️ 設定")
anchor_col_letter = st.sidebar.text_input("1. 基準列 (例: A)", value="A", help="店舗名などがある列")
skip_rows = st.sidebar.number_input("2. データの開始行まで飛ばす行数", min_value=0, value=2, help="1行目がタイトル、2行目がヘッダーなら「2」")

# === メインエリア ===
uploaded_file = st.file_uploader("Excelファイルをアップロードしてください", type=['xlsx'])

if uploaded_file:
    try:
        file_bytes = uploaded_file.getvalue()
        
        # 1. データ読み込み (計算後の値を取得)
        wb_data = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
        ws = wb_data.active
        
        # 2. 数式チェック用
        wb_formula = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=False)
        ws_f = wb_formula.active
        
        anchor_idx = openpyxl.utils.column_index_from_string(anchor_col_letter)

        # ---------------------------------------------------------
        # 【重要】エクセルの並び順をそのままプレビュー
        # ---------------------------------------------------------
        master_rows = []
        for r in range(skip_rows + 1, ws.max_row + 1):
            name = ws.cell(row=r, column=anchor_idx).value
            if name is not None:
                master_rows.append({"row_num": r, "name": str(name)})

        if not master_rows:
            st.error("指定された列にデータが見つかりません。設定を確認してください。")
            st.stop()

        st.success(f"📊 エクセルから {len(master_rows)} 行のデータを検出しました。")
        with st.expander("店舗名の並び順（上から順）を確認する"):
            st.table([{"行": d["row_num"], "店舗名": d["name"]} for d in master_rows])

        # --- 数式列の自動検出 ---
        formula_candidates = []
        # データ開始行の数行をサンプルチェック
        for c in range(1, ws_f.max_column + 1):
            if c == anchor_idx: continue
            is_f = False
            for r in range(skip_rows + 1, min(skip_rows + 10, ws_f.max_row) + 1):
                cell = ws_f.cell(row=r, column=c)
                if cell.data_type == 'f' or (cell.value and str(cell.value).startswith('=')):
                    is_f = True; break
            if is_f:
                formula_candidates.append({"idx": c, "name": openpyxl.utils.get_column_letter(c)})

        if formula_candidates:
            st.subheader("🛠️ 出力する列を選択")
            selected_cols = []
            cols_ui = st.columns(4)
            for i, cand in enumerate(formula_candidates):
                with cols_ui[i % 4]:
                    if st.checkbox(f"{cand['name']} 列", value=True, key=cand['idx']):
                        selected_cols.append(cand['idx'])

            # --- CSV作成実行 ---
            if st.button("🚀 CSVを作成 (エクセルの順番を死守)"):
                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, 'w') as myzip:
                    
                    for target_idx in selected_cols:
                        # ファイル名設定
                        col_letter = openpyxl.utils.get_column_letter(target_idx)
                        row2_val = ws.cell(row=2, column=target_idx).value
                        suffix = f"_{row2_val}" if row2_val is not None else ""
                        filename = f"output_column_{col_letter}{suffix}.csv"
                        
                        # --- データ抽出（1行ずつ順番に追加するだけ） ---
                        output_data = io.StringIO()
                        writer = csv.writer(output_data, quoting=csv.QUOTE_MINIMAL)
                        
                        # master_rows（エクセルを上から順に読み込んだリスト）をそのまま回す
                        for item in master_rows:
                            r_num = item["row_num"]
                            name = ws.cell(row=r_num, column=anchor_idx).value
                            val = ws.cell(row=r_num, column=target_idx).value
                            writer.writerow([name, val])
                        
                        # ZIPに追加
                        myzip.writestr(filename, output_data.getvalue().encode('utf-8-sig'))
                        output_data.close()
                
                st.success("✅ 作成完了！エクセルと全く同じ順番で書き出しました。")
                st.download_button("📥 ZIPファイルをダウンロード", data=zip_buffer.getvalue(), file_name="処理結果.zip")
        else:
            st.warning("数式列が見つかりません。")

    except Exception as e:
        st.error(f"エラーが発生しました: {e}")
