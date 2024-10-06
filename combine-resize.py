import os
import fitz  # PyMuPDF
from PyPDF2 import PdfMerger
from datetime import datetime
from collections import defaultdict
import unicodedata
from tqdm import tqdm

# A4サイズに統一するための関数
def convert_to_a4(input_pdf, output_pdf):
    doc = fitz.open(input_pdf)
    new_doc = fitz.open()  # 新しいドキュメントを作成
    a4_width, a4_height = fitz.paper_size("A4")  # A4の幅と高さを取得

    for page_num in range(len(doc)):
        page = doc.load_page(page_num)  # 元のページを取得
        rect = page.rect  # ページの元のサイズを取得
        scale_x = a4_width / rect.width  # 横方向のスケール
        scale_y = a4_height / rect.height  # 縦方向のスケール
        scale = min(scale_x, scale_y)  # 縦横比を保つために小さい方を使う

        # ページをA4サイズにスケーリングして適用
        mat = fitz.Matrix(scale, scale)  # スケーリングの行列
        pix = page.get_pixmap(matrix=mat)  # スケーリングされたページをピクセルデータとして取得
        new_page = new_doc.new_page(width=a4_width, height=a4_height)  # A4サイズの新しいページ
        new_page.insert_image(new_page.rect, pixmap=pix)  # 新しいページに元のページを画像として描画

    new_doc.save(output_pdf)  # 結果を保存
    new_doc.close()
    doc.close()

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

                first_char = file_name[0]  # ファイル名の最初の文字を取得
                first_char = unicodedata.normalize('NFKC', first_char)
                
                # カタカナ行に基づいてファイルを分類
                for row, chars in rows.items():
                    if first_char in chars:
                        file_groups[row].append(os.path.join(subfolder_path, file_name))
                        print(f"{file_name} は {row} に分類されました")
                        break
                else:
                    print(f"{file_name} は 50音順に対応しません")

# 各行ごとにファイルをカタカナの50音順にソートしてPDFを結合
print("PDFを結合中...")
for idx, (row, files) in enumerate(tqdm(file_groups.items()), 1):
    if files:
        print(f"{row} に含まれるファイル数: {len(files)}")
        sorted_files = sorted(files, key=lambda f: unicodedata.normalize('NFKC', f))
        
        merger = PdfMerger()
        for pdf_file in sorted_files:
            print(f"結合中のPDFファイル: {pdf_file}")
            try:
                merger.append(pdf_file)
            except Exception as e:
                print(f"{pdf_file} の処理中にエラーが発生しました: {e}")
                continue

        # 結合後のPDFファイルを保存するパス
        combined_pdf_path = os.path.join(output_folder_path, f"{subfolder_name}_{row}.pdf")
        with open(combined_pdf_path, 'wb') as f_out:
            print(f"{combined_pdf_path} に保存中...")
            merger.write(f_out)
        
        # A4サイズに変換
        a4_pdf_path = combined_pdf_path.replace(".pdf", "_A4.pdf")
        convert_to_a4(combined_pdf_path, a4_pdf_path)
        
        # 結合処理を終了
        merger.close()
        print(f"{a4_pdf_path} のA4サイズ変換が完了しました")

print("ファイルの結合とA4サイズへの変換が完了しました。")