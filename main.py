import os
import io
import fitz
from datetime import datetime
from collections import defaultdict
import unicodedata
from tqdm import tqdm
import pandas as pd
from PIL import Image
import numpy as np
import cv2

# ========== 設定セクション ==========
input_folder_path = './input_combine'
limit_size = 9 * (1024 * 1024)  # 9MB上限

dnn_prototxt = 'deploy.prototxt'
dnn_model = 'res10_300x300_ssd_iter_140000.caffemodel'
dnn_confidence_threshold = 0.3  # 顔とみなす最低信頼度

dpi_for_face_detection = 300  # 高DPIで画像化

output_folder_name = f"{datetime.now().strftime('%Y%m%d')}_output"
output_folder_path = f"./{output_folder_name}"
os.makedirs(output_folder_path, exist_ok=True)

print(f"入力フォルダ: {input_folder_path}")
print(f"出力フォルダ: {output_folder_path}")

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

file_groups = defaultdict(list)
file_statuses = []

subfolders = os.listdir(input_folder_path)
print(f"サブフォルダの数: {len(subfolders)}")

print("ファイルを分類中...")
for subfolder_name in tqdm(subfolders):
    subfolder_path = os.path.join(input_folder_path, subfolder_name)
    if os.path.isdir(subfolder_path):
        print(f"現在処理中のサブフォルダ: {subfolder_name}")
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

                file_statuses.append(status_record)

# DNNによる顔検出モデル読み込み
net = cv2.dnn.readNetFromCaffe(dnn_prototxt, dnn_model)

def is_photo_page(page, dpi=300, net=None, confidence_threshold=0.5):
    """
    DNNを用いて顔検出を行い、confidence_threshold以上の信頼度で顔が検出されれば写真ページとみなす
    """
    pix = page.get_pixmap(dpi=dpi)
    img = Image.open(io.BytesIO(pix.tobytes("png"))).convert("RGB")
    img_cv = np.array(img)
    (h, w) = img_cv.shape[:2]

    # DNN用にblob作成
    blob = cv2.dnn.blobFromImage(cv2.resize(img_cv, (300, 300)), 1.0, (300, 300), (104.0, 177.0, 123.0))
    net.setInput(blob)
    detections = net.forward()

    face_count = 0
    for i in range(detections.shape[2]):
        confidence = detections[0, 0, i, 2]
        if confidence > confidence_threshold:
            face_count += 1

    print(f"顔検出数: {face_count}, 閾値: {confidence_threshold}")
    return face_count > 0


def split_pdf_if_large(pdf_bytes: bytes, base_output_path: str, limit_size=9*(1024*1024), photo_info=None):
    if len(pdf_bytes) <= limit_size:
        with open(base_output_path, 'wb') as f:
            f.write(pdf_bytes)
        print(f"{base_output_path} の結合が完了しました (分割不要)")
        return

    print(f"PDFが大きすぎるため分割を開始します: {base_output_path}")
    reader = fitz.open(stream=pdf_bytes, filetype="pdf")
    total_pages = len(reader)
    print(f"元のPDFページ数: {total_pages}ページ")

    if photo_info is None or len(photo_info) != total_pages:
        photo_info = [False]*total_pages

    base_name, ext = os.path.splitext(base_output_path)
    part_number = 1
    page_index = 0

    while page_index < total_pages:
        # 写真ページを探す
        while page_index < total_pages and not photo_info[page_index]:
            print(f"ページ{page_index+1}: 写真ページでないためスキップ")
            page_index += 1

        if page_index >= total_pages:
            print("写真ページが見つからずファイル末尾。これ以上分割パートなし。")
            break

        part_doc = fitz.open()
        pages_in_part = 0
        part_size = 0

        while page_index < total_pages:
            part_doc.insert_pdf(reader, from_page=page_index, to_page=page_index)
            temp_stream = io.BytesIO()
            part_doc.save(temp_stream, garbage=4, deflate=True)
            new_size = len(temp_stream.getvalue())
            if new_size > limit_size:
                if pages_in_part == 0:
                    print(f"ページ{page_index+1}が単独で9MB超。1ページのみで出力します。")
                    pages_in_part = 1
                    page_index += 1
                else:
                    part_doc.delete_page(-1)
                break
            else:
                pages_in_part += 1
                part_size = new_size
                page_index += 1

        output_part_path = f"{base_name}-{part_number}{ext}"
        print(f"出力パート: {output_part_path}, {pages_in_part}ページ, 約{part_size/1024/1024:.2f}MB")
        with open(output_part_path, 'wb') as f_out:
            part_doc.save(f_out, garbage=4, deflate=True)
        part_number += 1

    reader.close()
    print(f"{base_output_path} の分割が完了しました")


print("PDFを結合中...")
for idx, (row, files) in enumerate(tqdm(file_groups.items()), 1):
    if files:
        print(f"{row} に含まれるファイル数: {len(files)}")
        sorted_files = sorted(files, key=lambda f: unicodedata.normalize('NFKC', os.path.basename(f)))
        print("結合対象ファイル一覧:", sorted_files)

        merged_doc = fitz.open()
        for pdf_file in sorted_files:
            print(f"結合中のPDFファイル: {pdf_file}")
            try:
                with fitz.open(pdf_file) as src_doc:
                    merged_doc.insert_pdf(src_doc)
                for status in file_statuses:
                    if status['ファイルパス'] == pdf_file:
                        status['状態'] = '結合済'
                        break
            except Exception as e:
                print(f"{pdf_file} の処理中にエラー: {e}")
                for status in file_statuses:
                    if status['ファイルパス'] == pdf_file:
                        status['状態'] = f'エラー: {e}'
                        break

        print(f"結合後PDFページ数: {len(merged_doc)}ページ")

        # 顔検出による写真ページ判定
        photo_info = []
        for i, page in enumerate(merged_doc):
            face_page = is_photo_page(page, dpi=dpi_for_face_detection, net=net, confidence_threshold=dnn_confidence_threshold)
            photo_info.append(face_page)
            print(f"ページ{i+1}: {'写真ページ(顔検出成功)' if face_page else '非写真ページ'}")

        output_pdf_path = os.path.join(output_folder_path, f"{subfolder_name}_{row}.pdf")
        pdf_stream = io.BytesIO()
        merged_doc.save(pdf_stream, garbage=4, deflate=True)
        merged_doc.close()

        pdf_bytes = pdf_stream.getvalue()
        print(f"出力PDF: {output_pdf_path}, サイズ約{len(pdf_bytes)/1024/1024:.2f}MB")
        split_pdf_if_large(pdf_bytes, output_pdf_path, limit_size=limit_size, photo_info=photo_info)

print("ファイルの結合と整理が完了しました。")

log_file_name = f"{output_folder_name}ログ.xlsx"
log_file_path = os.path.join(output_folder_path, log_file_name)
df = pd.DataFrame(file_statuses)
df.to_excel(log_file_path, index=False)
print(f"ログファイルが {log_file_path} に出力されました。")
