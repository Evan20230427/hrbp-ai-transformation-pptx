import os
import fitz  # PyMuPDF
import json
import pytesseract
from PIL import Image
import pdfplumber
import glob

# 設定 Tesseract 路徑 (Windows 必要)
# pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

def extract_high_quality_content():
    # 動態搜尋 input 資料夾下的第一個 pdf 檔案
    pdf_files = glob.glob(os.path.join("input", "*.pdf"))
    if not pdf_files:
        print("Error: No PDF files found in the 'input' folder.")
        return
    
    pdf_path = pdf_files[0]
    output_path = "extracted_content.json"
    
    print(f"Starting High-Quality Extraction (pdfplumber + OCR) for: {pdf_path}")
    
    results = []
    
    # 使用 pdfplumber 進行結構化擷取
    with pdfplumber.open(pdf_path) as pdf:
        for i, page in enumerate(pdf.pages):
            # 優先嘗試直接提取文字 (數位版 PDF 效果最好)
            text = page.extract_text(layout=True)
            
            # 如果直接提取失敗或字數太少，則啟動 OCR 補正
            if not text or len(text.strip()) < 50:
                print(f"Page {i+1}: Text sparse, falling back to OCR...")
                # 取得該頁的圖像
                doc = fitz.open(pdf_path)
                pdf_page = doc[i]
                pix = pdf_page.get_pixmap(matrix=fitz.Matrix(2, 2))
                img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                text = pytesseract.image_to_string(img, lang='chi_tra+eng')
                doc.close()
            
            # 清理 OCR 或提取出的文字 (過濾掉零散字元與修正切斷線)
            cleaned_text = " ".join(text.split())
            
            results.append({
                "page": i + 1,
                "text": cleaned_text,
                "has_tables": len(page.find_tables()) > 0
            })
            print(f"Processed page {i+1}")

    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(results, f, ensure_ascii=False, indent=2)
    
    print(f"Successfully extracted content to {output_path}")

if __name__ == "__main__":
    extract_high_quality_content()
