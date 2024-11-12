import os
from PyPDF2 import PdfMerger
from datetime import datetime
from collections import defaultdict
import unicodedata
from tqdm import tqdm

# 入力フォルダの指定
input_folder_path = './input_combine'  # input_combine フォルダのパスに変更
# 現在の日付でoutputフォルダを作成
output_folder_path = f"./{datetime.now().strftime('%Y%m%d')}_output"
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

# 行ごとにファイルを分類するための辞書を作成
file_groups = defaultdict(list)

# ファイルの状態を記録する辞書を作成
file_statuses = {}

# inputフォルダ内のすべてのサブフォルダを取得
subfolders = os.listdir(input_folder_path)
print(f"サブフォルダの数: {len(subfolders)}")

# サブフォルダごとにファイルを確認して分類する部分の進捗表示
print("ファイルを分類中...")
for subfolder_name in tqdm(subfolders):
    subfolder_path = os.path.join(input_folder_path, subfolder_name)

    if os.path.isdir(subfolder_path):
        print(f"現在処理中のサブフォルダ: {subfolder_name}")

        # サブフォルダ内のファイルを確認
        for file_name in os.listdir(subfolder_path):
            if file_name.endswith('.pdf'):  # PDFファイルに限定
                print(f"処理中のファイル: {file_name}")
                file_path = os.path.join(subfolder_path, file_name)
                file_statuses[file_path] = '未結合'  # 初期状態は未結合

                first_char = file_name[0]  # ファイル名の最初の文字を取得
                first_char = unicodedata.normalize('NFKC', first_char)

                # カタカナ行に基づいてファイルを分類
                for row, chars in rows.items():
                    if first_char in chars:
                        file_groups[row].append(file_path)
                        file_statuses[file_path] = f'結合予定: {row}'
                        print(f"{file_name} は {row} に分類されました")
                        break
                else:
                    print(f"{file_name} は 50音順に対応しません")

# 各行ごとにファイルをカタカナの50音順にソートしてPDFを結合
print("PDFを結合中...")
for idx, (row, files) in enumerate(tqdm(file_groups.items()), 1):
    if files:
        print(f"{row} に含まれるファイル数: {len(files)}")
        sorted_files = sorted(files, key=lambda f: unicodedata.normalize('NFKC', os.path.basename(f)))

        merger = PdfMerger()
        for pdf_file in sorted_files:
            print(f"結合中のPDFファイル: {pdf_file}")
            try:
                merger.append(pdf_file)
                # 結合済みに更新
                file_statuses[pdf_file] = '結合済'
            except Exception as e:
                print(f"{pdf_file} の処理中にエラーが発生しました: {e}")
                file_statuses[pdf_file] = f'error: {e}'
                continue

        # 結合後のPDFファイルを保存するパス
        output_pdf_path = os.path.join(output_folder_path, f"{subfolder_name}_{row}.pdf")
        with open(output_pdf_path, 'wb') as f_out:
            print(f"{output_pdf_path} に保存中...")
            merger.write(f_out)

        # 結合処理を終了
        merger.close()
        print(f"{output_pdf_path} の結合が完了しました")

print("ファイルの結合と整理が完了しました。")

# ログをファイルに出力
log_file_path = os.path.join(output_folder_path, 'log.txt')
with open(log_file_path, 'w', encoding='utf-8') as log_file:
    for file_path, status in file_statuses.items():
        log_file.write(f"{file_path}: {status}\n")

print(f"ログファイルが {log_file_path} に出力されました。")