import os
import io
import fitz  # PyMuPDF
from datetime import datetime
from collections import defaultdict
import unicodedata
from tqdm import tqdm
import pandas as pd

# 入力フォルダの指定
input_folder_path = './input_combine'  # input_combineフォルダのパス
# 現在の日付でoutputフォルダを作成
output_folder_name = f"{datetime.now().strftime('%Y%m%d')}_output"
output_folder_path = f"./{output_folder_name}"
os.makedirs(output_folder_path, exist_ok=True)

print(f"入力フォルダ: {input_folder_path}")
print(f"出力フォルダ: {output_folder_path}")

# 50音のカタカナ行を定義
rows = {
    'ア行': 'アイウエオ',
    'カ行': 'カキクケコ',
    'サ行': 'サシスセソ',
    'タ行': 'タチツテト',
    'ナ行': 'ナニヌネノ',
    'ハ行': 'ハヒフヘホ',
    'マ行': 'マミムメモ',
    'ヤ行': 'ヤユヨ',
    'ラ行': 'ラリルレロ',
    'ワ行': 'ワヲン'
}

file_statuses = []

subfolders = os.listdir(input_folder_path)
print(f"サブフォルダの数: {len(subfolders)}")

print("ファイルを分類中...")
for subfolder_name in tqdm(subfolders):
    subfolder_path = os.path.join(input_folder_path, subfolder_name)
    if os.path.isdir(subfolder_path):
        print(f"現在処理中のサブフォルダ: {subfolder_name}")

        sub_output_folder_path = os.path.join(output_folder_path, subfolder_name)
        os.makedirs(sub_output_folder_path, exist_ok=True)

        file_groups = defaultdict(list)
        subfolder_file_statuses = []

        for file_name in os.listdir(subfolder_path):
            if file_name.endswith('.pdf'):
                print(f"処理中のファイル: {file_name}")
                file_path = os.path.join(subfolder_path, file_name)
                status_record = {
                    'ファイルパス': file_path,
                    '状態': '未結合',
                    '分類': 'なし'
                }

                first_char = unicodedata.normalize('NFKC', file_name[0])

                for row, chars in rows.items():
                    if first_char in chars:
                        file_groups[row].append(file_path)
                        status_record['状態'] = '結合予定'
                        status_record['分類'] = row
                        print(f"{file_name} は {row} に分類されました")
                        break
                else:
                    print(f"{file_name} は 50音順に対応しません")

                subfolder_file_statuses.append(status_record)

        print("PDFを結合中...")
        for idx, (row, files) in enumerate(file_groups.items(), 1):
            if files:
                print(f"{row} に含まれるファイル数: {len(files)}")
                sorted_files = sorted(files, key=lambda f: unicodedata.normalize('NFKC', os.path.basename(f)))

                merger = fitz.open()
                for pdf_file in sorted_files:
                    print(f"結合中のPDFファイル: {pdf_file}")
                    try:
                        with fitz.open(pdf_file) as src_doc:
                            merger.insert_pdf(src_doc)
                        for status in subfolder_file_statuses:
                            if status['ファイルパス'] == pdf_file:
                                status['状態'] = '結合済'
                                break
                    except Exception as e:
                        print(f"{pdf_file} の処理中にエラーが発生しました: {e}")
                        for status in subfolder_file_statuses:
                            if status['ファイルパス'] == pdf_file:
                                status['状態'] = f'エラー: {e}'
                                break
                        continue

                output_pdf_path = os.path.join(sub_output_folder_path, f"{subfolder_name}_{row}.pdf")
                # PyMuPDFで基本最適化
                merger.save(output_pdf_path, garbage=4, deflate=True)
                merger.close()
                print(f"{output_pdf_path} に保存しました (PyMuPDF最適化済)")


        file_statuses.extend(subfolder_file_statuses)

print("ファイルの結合と整理が完了しました。")

log_file_name = f"{output_folder_name}ログ.xlsx"
log_file_path = os.path.join(output_folder_path, log_file_name)
df = pd.DataFrame(file_statuses)
df.to_excel(log_file_path, index=False)
print(f"ログファイルが {log_file_path} にExcel形式で出力されました。")

