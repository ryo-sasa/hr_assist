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

# inputフォルダ内のすべてのサブフォルダを取得
subfolders = os.listdir(input_folder_path)
print(f"サブフォルダの数: {len(subfolders)}")

print("ファイルを分類中...")

def split_pdf_if_large(pdf_bytes: bytes, base_output_path: str, limit_size=9*(1024*1024)):
    """
    PDFを9MBごとに分割する関数(以前使用したロジックを再掲)
    """
    if len(pdf_bytes) <= limit_size:
        with open(base_output_path, 'wb') as f:
            f.write(pdf_bytes)
        print(f"{base_output_path} の結合が完了しました (分割不要)")
        return

    print(f"PDFが大きすぎるため分割を開始します: {base_output_path}")
    reader = fitz.open(stream=pdf_bytes, filetype="pdf")
    total_pages = len(reader)
    base_name, ext = os.path.splitext(base_output_path)

    part_number = 1
    page_start = 0

    while page_start < total_pages:
        part_doc = fitz.open()
        part_pages = 0
        current_size = 0

        for i in range(page_start, total_pages):
            part_doc.insert_pdf(reader, from_page=i, to_page=i)
            temp_stream = io.BytesIO()
            part_doc.save(temp_stream, garbage=4, deflate=True)
            new_size = len(temp_stream.getvalue())

            if new_size > limit_size:
                # このページ追加でオーバーするので削除して確定出力
                part_doc.delete_page(-1)
                break
            else:
                current_size = new_size
                part_pages += 1

        if part_pages == 0:
            # 1ページも追加できない場合(単ページ9MB超え)
            part_doc = fitz.open()
            part_doc.insert_pdf(reader, from_page=page_start, to_page=page_start)
            temp_stream = io.BytesIO()
            part_doc.save(temp_stream, garbage=4, deflate=True)
            current_size = len(temp_stream.getvalue())
            part_pages = 1

        output_part_path = f"{base_name}-{part_number}{ext}"
        with open(output_part_path, 'wb') as f_out:
            part_doc.save(f_out, garbage=4, deflate=True)
        print(f"{output_part_path} に分割保存しました ({part_pages}ページ, 約{current_size/1024/1024:.2f}MB)")

        part_number += 1
        page_start += part_pages

    reader.close()
    print(f"{base_output_path} の分割が完了しました")


for subfolder_name in tqdm(subfolders):
    subfolder_path = os.path.join(input_folder_path, subfolder_name)
    if os.path.isdir(subfolder_path):
        print(f"現在処理中のサブフォルダ: {subfolder_name}")

        # サブフォルダ専用の出力フォルダを作成
        sub_output_folder_path = os.path.join(output_folder_path, subfolder_name)
        os.makedirs(sub_output_folder_path, exist_ok=True)

        file_groups = defaultdict(list)
        subfolder_file_statuses = []

        # サブフォルダ内のファイルを確認
        for file_name in os.listdir(subfolder_path):
            if file_name.endswith('.pdf'):
                print(f"処理中のファイル: {file_name}")
                file_path = os.path.join(subfolder_path, file_name)
                status_record = {
                    'ファイルパス': file_path,
                    '状態': '未結合',
                    '分類': 'なし'
                }

                first_char = file_name[0]
                first_char = unicodedata.normalize('NFKC', first_char)

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

                # 結合後PDFをメモリに書き出し
                pdf_stream = io.BytesIO()
                merger.save(pdf_stream, garbage=4, deflate=True)
                merger.close()

                pdf_bytes = pdf_stream.getvalue()
                output_pdf_path = os.path.join(sub_output_folder_path, f"{subfolder_name}_{row}.pdf")

                # フォルダ名に「履歴書」が含まれるか判定
                if "履歴書" in subfolder_name:
                    # そのまま保存
                    with open(output_pdf_path, 'wb') as f_out:
                        f_out.write(pdf_bytes)
                    print(f"{output_pdf_path} に保存しました (分割なし)")
                else:
                    # 9MB以下に分割
                    split_pdf_if_large(pdf_bytes, output_pdf_path, limit_size=9*(1024*1024))

        file_statuses.extend(subfolder_file_statuses)

print("ファイルの結合と整理が完了しました。")

log_file_name = f"{output_folder_name}ログ.xlsx"
log_file_path = os.path.join(output_folder_path, log_file_name)
df = pd.DataFrame(file_statuses)
df.to_excel(log_file_path, index=False)
print(f"ログファイルが {log_file_path} にExcel形式で出力されました。")

