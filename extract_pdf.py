import pdfplumber
import json
import os

pdf_path = os.path.join("input", "AI 時代重新設計人力資源業務夥伴（HRBP）角色的最佳實務.pdf")
output_path = "extracted_content.json"

result = []
try:
    with pdfplumber.open(pdf_path) as pdf:
        for i, page in enumerate(pdf.pages):
            text = page.extract_text()
            result.append({
                "page": i + 1,
                "text": text or ""
            })
    
    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    print(f"Extraction successful! Saved to {output_path}")
except Exception as e:
    print(f"Error: {e}")
