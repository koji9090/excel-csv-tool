#@title 簡易Webアプリ風フォーム
#@markdown 以下の設定をしてから、左の再生ボタンを押してください。

import pandas as pd
import openpyxl
from google.colab import files
import io
import zipfile

# フォーム設定
skip_rows = 2 #@param {type:"number"}
#@markdown ※ファイルを選択すると自動で処理が始まります

uploaded = files.upload()

if uploaded:
    filename = list(uploaded.keys())[0]
    print(f"処理中: {filename} ...")
    
    # Excelとして処理
    process_filename = "temp.xlsx"
    with open(process_filename, 'wb') as f:
        f.write(uploaded[filename])

    try:
        # 数式列の検出
        wb = openpyxl.load_workbook(process_filename, data_only=False)
        ws = wb.active
        
        formula_cols = []
        start_r = skip_rows + 1
        
        for c in range(2, ws.max_column + 1):
            is_f = False
            for r in range(start_r, min(start_r+10, ws.max_row)):
                cell = ws.cell(row=r, column=c)
                if cell.data_type == 'f' or str(cell.value).startswith('='):
                    is_f = True; break
            if is_f: formula_cols.append(c-1)
        
        if not formula_cols:
            print("数式列が見つかりませんでした。")
        else:
            print(f"検出された列: {[openpyxl.utils.get_column_letter(c+1) for c in formula_cols]}")
            
            df = pd.read_excel(process_filename, skiprows=skip_rows, header=None, engine='openpyxl')
            output_files = []
            
            for c_idx in formula_cols:
                if c_idx < len(df.columns):
                    c_name = openpyxl.utils.get_column_letter(c_idx+1)
                    out_df = df.iloc[:, [0, c_idx]]
                    fname = f"output_{c_name}.csv"
                    out_df.to_csv(fname, header=False, index=False, encoding='utf-8-sig')
                    output_files.append(fname)
            
            if output_files:
                zname = "result.zip"
                with zipfile.ZipFile(zname, 'w') as z:
                    for f in output_files: z.write(f)
                files.download(zname)
                print("ダウンロード完了！")
                
    except Exception as e:
        print(f"エラー: {e}")