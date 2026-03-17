import os
import fitz  # PyMuPDF
import json
import pytesseract
from PIL import Image
import pdfplumber
import glob
import re
from collections import Counter

def classify_page(text):
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
    scores = {}
    for group, keywords in KEYWORD_GROUPS.items():
        score = sum(1 for kw in keywords if kw.lower() in text.lower())
        if score > 0: scores[group] = score
    return max(scores, key=scores.get) if scores else "概述"

def extract_page_keywords(text, top_n=5):
    stopwords = {"的", "了", "在", "是", "和", "與", "以", "及", "或", "來", "到", "這", "我", "們", "您", "他", "她", "它", "不", "有", "為", "能", "可", "會", "將", "更", "已", "被", "都", "也", "要", "讓", "從", "其", "並", "對", "如", "之", "透過", "方面", "方式", "以及"}
    words = re.findall(r'[\u4e00-\u9fff]{2,4}|[A-Za-z]{3,}', text)
    filtered = [w for w in words if w not in stopwords and len(w) >= 2]
    return [word for word, _ in Counter(filtered).most_common(top_n)]

def generate_quality_summary(text, max_len=25):
    sentences = re.split(r'[。！？；\n]', text)
    meaningful = [s.strip() for s in sentences if len(s.strip()) > 5]
    if not meaningful: return text[:max_len].strip()
    summary = meaningful[0]
    if len(summary) > max_len:
        parts = re.split(r'[，、]', summary)
        short_summary = ""
        for p in parts:
            if len(short_summary) + len(p) < max_len: short_summary += p + " "
            else: break
        summary = short_summary.strip() or summary[:max_len]
    return summary[:max_len].strip()

def extract_high_quality_content():
    pdf_files = glob.glob(os.path.join("input", "*.pdf"))
    if not pdf_files: return
    pdf_path = pdf_files[0]
    results = []
    all_keywords = Counter()
    with pdfplumber.open(pdf_path) as pdf:
        for i, page in enumerate(pdf.pages):
            text = page.extract_text(layout=True)
            if not text or len(text.strip()) < 50:
                doc = fitz.open(pdf_path)
                pix = doc[i].get_pixmap(matrix=fitz.Matrix(2, 2))
                img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                text = pytesseract.image_to_string(img, lang='chi_tra+eng')
                doc.close()
            cleaned = " ".join(text.split())
            topic = classify_page(cleaned)
            keywords = extract_page_keywords(cleaned)
            summary = generate_quality_summary(cleaned, 25)
            for kw in keywords: all_keywords[kw] += 1
            results.append({"page": i+1, "text": cleaned, "topic": topic, "keywords": keywords, "summary": summary})
            # print(f"  Page {i+1}: [{topic}] {summary}") # 移除 print 以避免 Windows 編碼錯誤
    top_global_keywords = [w for w, _ in all_keywords.most_common(15)]
    topic_map = {}
    for r in results:
        t = r["topic"]
        if t not in topic_map: topic_map[t] = []
        topic_map[t].append(r["page"])
    output = {"pages": results, "analysis": {"top_keywords": top_global_keywords, "topic_page_map": topic_map, "total_pages": len(results)}}
    with open("extracted_content.json", "w", encoding="utf-8") as f: json.dump(output, f, ensure_ascii=False, indent=2)

if __name__ == "__main__": extract_high_quality_content()
