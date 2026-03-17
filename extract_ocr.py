import os
import fitz  # PyMuPDF
import json
import pytesseract
from PIL import Image
import pdfplumber
import glob
import re
from collections import Counter

# 設定 Tesseract 路徑 (Windows 必要)
# pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# ======== doc-coauthoring 整合：關鍵字權重表 ========
# 定義領域關鍵字群組，用於自動分類頁面主題
KEYWORD_GROUPS = {
    "AI 技能培養": ["AI", "技能", "培養", "學習", "訓練", "計劃"],
    "行銷轉型": ["行銷", "marketing", "品牌", "客戶", "傳達"],
    "銷售賦能": ["銷售", "sales", "MCAPS", "客戶互動", "業務"],
    "工程創新": ["工程", "engineering", "技術", "工程師", "建置"],
    "互動探索": ["Garage", "hackathon", "挑戰", "Copilot", "探索", "互動"],
    "負責任 AI": ["負責任", "道德", "信任", "風險", "政策"],
    "學習途徑": ["途徑", "課程", "飛行計劃", "模組", "基礎"],
    "組織文化": ["文化", "協作", "知識分享", "創新", "社群"]
}

def classify_page(text):
    """根據關鍵字群組對頁面進行主題分類 (doc-coauthoring: 結構化分析)"""
    scores = {}
    for group, keywords in KEYWORD_GROUPS.items():
        score = sum(1 for kw in keywords if kw.lower() in text.lower())
        if score > 0:
            scores[group] = score
    if scores:
        return max(scores, key=scores.get)
    return "概述"

def extract_page_keywords(text, top_n=5):
    """從頁面文字中擷取高頻關鍵字 (doc-coauthoring: 內容精煉)"""
    # 過濾掉短詞與常見虛詞
    stopwords = {"的", "了", "在", "是", "和", "與", "以", "及", "或", "來", "到",
                 "這", "我", "們", "您", "他", "她", "它", "不", "有", "為", "能",
                 "可", "會", "將", "更", "已", "被", "都", "也", "要", "讓", "從",
                 "其", "並", "對", "如", "之", "透過", "方面", "方式"}
    words = re.findall(r'[\u4e00-\u9fff]{2,4}|[A-Za-z]{3,}', text)
    filtered = [w for w in words if w not in stopwords and len(w) >= 2]
    counter = Counter(filtered)
    return [word for word, _ in counter.most_common(top_n)]

def generate_page_summary(text, max_len=80):
    """為頁面生成精煉摘要 (doc-coauthoring: 每句話的重量感)"""
    sentences = re.split(r'[。！？\n]', text)
    meaningful = [s.strip() for s in sentences if len(s.strip()) > 15]
    if meaningful:
        return meaningful[0][:max_len]
    return text[:max_len]

def extract_high_quality_content():
    """高品質內容擷取主函式 (整合 pdf + doc-coauthoring Skills)"""
    pdf_files = glob.glob(os.path.join("input", "*.pdf"))
    if not pdf_files:
        print("Error: No PDF files found in the 'input' folder.")
        return
    
    pdf_path = pdf_files[0]
    output_path = "extracted_content.json"
    
    print(f"Starting High-Quality Extraction for: {pdf_path}")
    
    results = []
    all_keywords = Counter()
    
    with pdfplumber.open(pdf_path) as pdf:
        for i, page in enumerate(pdf.pages):
            # 優先使用 pdfplumber 直接提取
            text = page.extract_text(layout=True)
            
            # 文字不足時啟動 OCR 補正
            if not text or len(text.strip()) < 50:
                print(f"  Page {i+1}: Sparse text, falling back to OCR...")
                doc = fitz.open(pdf_path)
                pdf_page = doc[i]
                pix = pdf_page.get_pixmap(matrix=fitz.Matrix(2, 2))
                img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                text = pytesseract.image_to_string(img, lang='chi_tra+eng')
                doc.close()
            
            cleaned_text = " ".join(text.split())
            
            # doc-coauthoring 結構化分析
            topic = classify_page(cleaned_text)
            keywords = extract_page_keywords(cleaned_text)
            summary = generate_page_summary(cleaned_text)
            
            # 累積全文關鍵字
            for kw in keywords:
                all_keywords[kw] += 1
            
            results.append({
                "page": i + 1,
                "text": cleaned_text,
                "topic": topic,
                "keywords": keywords,
                "summary": summary,
                "has_tables": len(page.find_tables()) > 0
            })
            print(f"  Page {i+1}: [{topic}] {', '.join(keywords[:3])}")

    # 生成全文關鍵字排行
    top_global_keywords = [w for w, _ in all_keywords.most_common(15)]
    
    # 建立主題->頁碼的邏輯對應表
    topic_page_map = {}
    for r in results:
        t = r["topic"]
        if t not in topic_page_map:
            topic_page_map[t] = []
        topic_page_map[t].append(r["page"])
    
    # 輸出結構化 JSON (含分析結果供 JS 引擎使用)
    output = {
        "pages": results,
        "analysis": {
            "top_keywords": top_global_keywords,
            "topic_page_map": topic_page_map,
            "total_pages": len(results)
        }
    }
    
    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, indent=2)
    
    print(f"\nSuccessfully extracted content to {output_path}")
    print(f"  Global Keywords: {', '.join(top_global_keywords[:8])}")
    print(f"  Topic Map: {json.dumps(topic_page_map, ensure_ascii=False)}")

if __name__ == "__main__":
    extract_high_quality_content()
