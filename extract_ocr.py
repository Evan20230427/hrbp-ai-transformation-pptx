import os
import json
import fitz  # PyMuPDF
import pytesseract
from PIL import Image
import io

# 如果 Tesseract 路徑不在 PATH 中，請於此處指定
# pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

def extract_ocr_content():
    pdf_path = os.path.join("input", "AI 時代重新設計人力資源業務夥伴（HRBP）角色的最佳實務.pdf")
    output_path = "extracted_content.json"
    
    if not os.path.exists(pdf_path):
        print(f"Error: {pdf_path} not found.")
        return

    print(f"Starting OCR extraction with Tesseract for: {pdf_path}")
    
    doc = fitz.open(pdf_path)
    structured_data = []
    
    for page_index in range(len(doc)):
        page = doc.load_page(page_index)
        # 將頁面渲染為圖像 (提高解析度以利 OCR)
        pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
        img_data = pix.tobytes("png")
        img = Image.open(io.BytesIO(img_data))
        
        # 執行 OCR (支援繁體中文)
        text = pytesseract.image_to_string(img, lang='chi_tra')
        
        structured_data.append({
            "page": page_index + 1,
            "text": text.strip()
        })
        print(f"Processed page {page_index + 1}")

    doc.close()

    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(structured_data, f, ensure_ascii=False, indent=2)
    
    print(f"Successfully extracted content to {output_path}")

if __name__ == "__main__":
    extract_ocr_content()
