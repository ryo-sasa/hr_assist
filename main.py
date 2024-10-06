import pytesseract
from pdf2image import convert_from_path
import pandas as pd
import os
import re
import cv2
import numpy as np
import fitz  # PyMuPDF
import unicodedata

# OCRの設定
pytesseract.pytesseract.tesseract_cmd = '/opt/homebrew/bin/tesseract'

# A4サイズのPDFの実寸とDPI設定
A4_WIDTH_CM = 21.0
A4_HEIGHT_CM = 29.7
DPI = 300  # 解像度 (dots per inch)
CM_TO_INCH = 2.54

# cm単位からピクセル単位に変換
def cm_to_px(cm, dpi=DPI):
    return int((cm / CM_TO_INCH) * dpi)

# 指定領域から画像を抽出
def extract_image_region(pdf_path, page_num, top_left_cm, bottom_right_cm):
    doc = fitz.open(pdf_path)
    page = doc.load_page(page_num)  # ページを取得
    
    # A4サイズで指定の領域をピクセルに変換
    top_left_px = (cm_to_px(top_left_cm[0]), cm_to_px(top_left_cm[1]))
    bottom_right_px = (cm_to_px(bottom_right_cm[0]), cm_to_px(bottom_right_cm[1]))
    
    print(f"Extracting region from {top_left_px} to {bottom_right_px} in pixels")  # デバッグログ
    
    # 領域を指定して画像を抽出
    pix = page.get_pixmap(clip=fitz.Rect(top_left_px[0], top_left_px[1], bottom_right_px[0], bottom_right_px[1]))
    
    # 画像を numpy 配列に変換
    if pix.samples:  # pix.samples が空でないか確認
        img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)
        if pix.n == 4:  # RGBAの場合
            img = cv2.cvtColor(img, cv2.COLOR_RGBA2RGB)
    else:
        print("Error: Failed to extract image region. The extracted region is empty.")
        img = None
    
    doc.close()
    return img

# 画像を前処理（グレースケール化、ノイズ除去、二値化など）
def preprocess_image(image):
    if image is None:
        raise ValueError("The input image is empty, cannot preprocess")
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    denoised = cv2.GaussianBlur(gray, (5, 5), 0)
    _, binary = cv2.threshold(denoised, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    equalized = cv2.equalizeHist(binary)
    kernel = np.array([[0, -1, 0], [-1, 5,-1], [0, -1, 0]])
    sharpened = cv2.filter2D(equalized, -1, kernel)
    return sharpened

# 銀行コード、支店コード、口座番号を抽出する正規表現を使った解析
def extract_bank_info(text):
    bank_code_pattern = r'[^\d]([0-9]{4})[^\d]'  # 4桁の銀行コード
    branch_code_pattern = r'[^\d]([0-9]{3})[^\d]'  # 3桁の支店コード
    account_number_pattern = r'[^\d]([0-9]{6})[^\d]'  # 6桁の口座番号
    
    bank_code = re.search(bank_code_pattern, text)
    branch_code = re.search(branch_code_pattern, text)
    account_number = re.search(account_number_pattern, text)
    
    return {
        '銀行コード': bank_code.group(1) if bank_code else None,
        '支店コード': branch_code.group(1) if branch_code else None,
        '口座番号': account_number.group(1) if account_number else None
    }

# 画像からテキストを抽出
def ocr_image(image):
    preprocessed_image = preprocess_image(image)
    return pytesseract.image_to_string(preprocessed_image, lang='jpn')

# PDFから指定領域をOCRするメイン処理
def process_pdf(pdf_path):
    # 4つの領域の指定 (縦横: 左上、右上、左下、右下) in cm
    regions = [
        ((13.5, 1), (18, 19)),   # 領域1
        ((13.5, 19), (18, 19)),  # 領域2
        ((13.5, 1), (18, 19)),   # 領域3
        ((13.5, 19), (18, 19))   # 領域4
    ]
    
    extracted_text = ""
    for idx, region in enumerate(regions):
        top_left, bottom_right = region
        img = extract_image_region(pdf_path, 0, top_left, bottom_right)  # 0ページ目を処理
        
        if img is None:
            print(f"Skipping region {idx + 1} due to extraction failure.")
            continue
        
        extracted_text += ocr_image(img)
    
    # OCR結果を確認
    print("OCR結果:\n", extracted_text)
    
    # 銀行コード、支店コード、口座番号を抽出
    bank_info = extract_bank_info(extracted_text)
    
    return bank_info

# ファイル名の最初の文字に基づいて50音順にソートするためのヘルパー関数
def sort_key(file_name):
    first_char = file_name[0]
    first_char = unicodedata.normalize('NFKC', first_char)  # 文字を正規化
    return first_char

# 指定フォルダ内のすべてのPDFファイルを処理する関数
def process_all_pdfs_in_folder(input_folder, output_csv):
    all_data = []
    
    for file_name in os.listdir(input_folder):
        if file_name.endswith('.pdf'):
            pdf_path = os.path.join(input_folder, file_name)
            print(f"処理中: {pdf_path}")
            
            bank_info = process_pdf(pdf_path)  # PDFを処理
            
            all_data.append({
                'ファイル名': file_name,
                '銀行コード': bank_info['銀行コード'],
                '支店コード': bank_info['支店コード'],
                '口座番号': bank_info['口座番号']
            })
    
    # ファイル名の50音順でソート
    all_data = sorted(all_data, key=lambda x: sort_key(x['ファイル名']))
    
    df = pd.DataFrame(all_data)
    df.to_csv(output_csv, index=False, encoding='utf-8-sig')
    print(f"結果が {output_csv} に保存されました")

# 実行例
input_folder_path = './input_ocr'  # PDFファイルが含まれるフォルダ
output_csv_path = '銀行情報.csv'  # 保存先のCSVファイル名
process_all_pdfs_in_folder(input_folder_path, output_csv_path)