import os
import fitz  # PyMuPDF
import json
import pytesseract
from PIL import Image
import pdfplumber
import glob
import re
from collections import Counter

# ======== v14 Prometheus: 核心主題群組 ========
KEYWORD_GROUPS = {
    "AI 技能培訓": ["AI", "技能", "培養", "學習", "訓練", "計劃"],
    "行銷賦能": ["行銷", "marketing", "品牌", "客戶", "傳達"],
    "銷售與 MCAPS": ["銷售", "sales", "MCAPS", "客戶互動", "業務"],
    "工程與技術創新": ["工程", "engineering", "技術", "創新", "Coding"],
    "探索與實驗 (Garage)": ["Garage", "hackathon", "挑戰", "Copilot", "實驗"],
    "負責任與治理": ["負責任", "道德", "信任", "風險", "政策"],
    "學習途徑與社群": ["途徑", "課程", "飛行計劃", "社群", "分享"],
    "文化與未來趨勢": ["文化", "趨勢", "數位轉型", "人才", "驅動"]
}

def classify_page(text):
    scores = {}
    for group, keywords in KEYWORD_GROUPS.items():
        score = sum(1 for kw in keywords if kw.lower() in text.lower())
        if score > 0: scores[group] = score
    return max(scores, key=scores.get) if scores else "概述"

def extract_page_keywords(text, top_n=5):
    stopwords = {"的", "了", "在", "是", "和", "與", "以", "及", "或", "來", "到", "這", "我", "們", "您", "他", "她", "它", "不", "有", "為", "能", "可", "會", "將", "更", "已", "被", "都", "也", "要", "讓", "從", "其", "並", "對", "如", "之", "透過", "方面", "方式", "以及"}
    # 擷取 2-4 字中文或 3 字以上英文
    words = re.findall(r'[\u4e00-\u9fff]{2,4}|[A-Za-z]{3,}', text)
    filtered = [w for w in words if w not in stopwords and len(w) >= 2]
    return [word for word, _ in Counter(filtered).most_common(top_n)]

def generate_prometheus_summary(text, max_len=15):
    """
    v14: 極致 15 字摘要，用於簡報下方字塊與 Prompt
    """
    # 移除頁碼與冗餘標頭
    clean_text = re.sub(r'\d+', '', text).replace("加速員工 AI 技能的 10 種 最佳做法", "").strip()
    sentences = re.split(r'[。！？；\n]', clean_text)
    meaningful = [s.strip() for s in sentences if len(s.strip()) > 3]
    
    source = meaningful[0] if meaningful else clean_text[:30]
    # 精煉至 15 字
    summary = source[:max_len]
    if "，" in summary: summary = summary.split("，")[0]
    return summary.strip()

def extract_v14_content():
    pdf_files = glob.glob(os.path.join("input", "*.pdf"))
    if not pdf_files: return
    pdf_path = pdf_files[0]
    pdf_name = os.path.basename(pdf_path)
    results = []
    all_keywords = Counter()
    
    print(f"--- [v14] Starting Content Refinement for: {pdf_name} ---")
    
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
            keywords = extract_page_keywords(cleaned, top_n=3)
            # v14: 15 字重點字塊
            summary_15 = generate_prometheus_summary(cleaned, 15)
            
            for kw in keywords: all_keywords[kw] += 1
            
            results.append({
                "page": i + 1,
                "text": cleaned,
                "topic": topic,
                "keywords": keywords,
                "summary": summary_15,
                "image_prompt": f"{summary_15} in technology workspace"
            })
            # print(f"  P{i+1} [15字]: {summary_15}") # 避免 Windows 編碼報錯

    top_global_keywords = [w for w, _ in all_keywords.most_common(12)]
    topic_map = {}
    for r in results:
        t = r["topic"]
        if t not in topic_map: topic_map[t] = []
        topic_map[t].append(r["page"])
    
    output = {
        "pdf_name": pdf_name,
        "pages": results,
        "analysis": {
            "top_keywords": top_global_keywords,
            "topic_page_map": topic_map,
            "total_pages": len(results)
        }
    }
    with open("extracted_content.json", "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, indent=2)
    print(f"[SUCCESS] v14 Data Ready. Extracted {len(results)} pages.")

if __name__ == "__main__":
    extract_v14_content()
