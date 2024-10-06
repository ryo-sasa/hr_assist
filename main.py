import pytesseract
from pdf2image import convert_from_path
import pandas as pd
import os
import re

# OCRの設定
pytesseract.pytesseract.tesseract_cmd = '/usr/local/bin/tesseract'

# PDFファイルを画像に変換
def pdf_to_images(pdf_path):
    return convert_from_path(pdf_path, poppler_path='/usr/local/bin')  # Popplerのパスを指定

# 画像からテキストを抽出
def ocr_image(image):
    return pytesseract.image_to_string(image, lang='jpn')

# 銀行コード、支店コード、口座番号を抽出する正規表現を使った解析
def extract_bank_info(text):
    # 銀行コード、支店コード、口座番号を抽出するための簡易的な正規表現
    bank_code_pattern = r'銀行コード[：: ]*([0-9]{3,4})'
    branch_code_pattern = r'支店コード[：: ]*([0-9]{3,4})'
    account_number_pattern = r'口座番号[：: ]*([0-9]{6,})'
    
    bank_code = re.search(bank_code_pattern, text)
    branch_code = re.search(branch_code_pattern, text)
    account_number = re.search(account_number_pattern, text)
    
    return {
        '銀行コード': bank_code.group(1) if bank_code else None,
        '支店コード': branch_code.group(1) if branch_code else None,
        '口座番号': account_number.group(1) if account_number else None
    }

# メイン処理
def process_pdf(pdf_path):
    images = pdf_to_images(pdf_path)
    extracted_text = ""
    
    # 各ページをOCRで処理
    for image in images:
        extracted_text += ocr_image(image)
    
    # OCR結果を確認
    print("OCR結果:\n", extracted_text)
    
    # 銀行コード、支店コード、口座番号を抽出
    bank_info = extract_bank_info(extracted_text)
    
    return bank_info

# PDFファイル名とその内容をDataFrameにまとめる
def create_csv_from_pdf(pdf_path, output_csv):
    file_name = os.path.basename(pdf_path)
    bank_info = process_pdf(pdf_path)
    
    # ファイル名とともに銀行情報をまとめる
    data = {
        'ファイル名': [file_name],
        '銀行コード': [bank_info['銀行コード']],
        '支店コード': [bank_info['支店コード']],
        '口座番号': [bank_info['口座番号']]
    }
    
    df = pd.DataFrame(data)
    
    # CSVに保存
    df.to_csv(output_csv, index=False, encoding='utf-8-sig')
    print(f"結果が {output_csv} に保存されました")

# 実行例
pdf_file_path = '/Users/sasagawaryousuke/Development/private_dev/nagase/HR_OCR/test.pdf'  # PDFのパスを指定
output_csv_path = '銀行情報.csv'  # 保存先のCSVファイル名
create_csv_from_pdf(pdf_file_path, output_csv_path)