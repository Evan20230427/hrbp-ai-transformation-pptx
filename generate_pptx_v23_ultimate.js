const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

/**
 * [Dynamic Engine v11] 損壞修復與長度補強版 (Rescue & Length Edition)
 * 修正：損壞座標處理、長度最小值 25 頁限制、全篇 PDF 內容處理
 */

// ======== 視覺風格定義 (維持單一鎖定) ========
const STYLES = {
    ghibli: {
        name: "吉卜力漫畫風格",
        images: [
            "C:/Users/TW-Evan.Chen/.gemini/antigravity/brain/ac6f3712-34eb-4b1e-b0f3-0fb2ebf77cf6/style_ghibli_learning_1773732831828.png",
            "C:/Users/TW-Evan.Chen/.gemini/antigravity/brain/ac6f3712-34eb-4b1e-b0f3-0fb2ebf77cf6/slide_marketing_ai_1773730529257.png",
            "C:/Users/TW-Evan.Chen/.gemini/antigravity/brain/ac6f3712-34eb-4b1e-b0f3-0fb2ebf77cf6/slide_garage_hackathon_1773730597833.png"
        ],
        theme: { primary: "1E1E1E", secondary: "0066FF", accent: "00FFFF", text: "FFFFFF", cardBg: "2D2D2D" }
    },
    photography: {
        name: "攝影風格",
        images: [
            "C:/Users/TW-Evan.Chen/.gemini/antigravity/brain/ac6f3712-34eb-4b1e-b0f3-0fb2ebf77cf6/style_photography_tech_innovation_1773732856137.png",
            "C:/Users/TW-Evan.Chen/.gemini/antigravity/brain/ac6f3712-34eb-4b1e-b0f3-0fb2ebf77cf6/style_architectural_realism_ai_1773732476963.png"
        ],
        theme: { primary: "1A2332", secondary: "2D8B8B", accent: "A8DADC", text: "F1FAEE", cardBg: "243447" }
    },
    lineart: {
        name: "線條簡潔風格",
        images: [
            "C:/Users/TW-Evan.Chen/.gemini/antigravity/brain/ac6f3712-34eb-4b1e-b0f3-0fb2ebf77cf6/style_minimalist_swiss_logic_1773732459484.png",
            "C:/Users/TW-Evan.Chen/.gemini/antigravity/brain/ac6f3712-34eb-4b1e-b0f3-0fb2ebf77cf6/style_fine_line_statue_v23_1773732497567.png"
        ],
        theme: { primary: "F1F5F9", secondary: "B45309", accent: "475569", text: "1E293B", cardBg: "FFFFFF" }
    }
};

const JSON_PATH = "extracted_content.json";
if (!fs.existsSync(JSON_PATH)) process.exit(1);
const jsonData = JSON.parse(fs.readFileSync(JSON_PATH, "utf-8"));
const rawData = jsonData.pages;
const analysis = jsonData.analysis;

// v11: 鎖定單一風格
const styleKeys = Object.keys(STYLES);
const lockedStyleKey = styleKeys[Math.floor(Math.random() * styleKeys.length)];
const S = STYLES[lockedStyleKey];
const T = S.theme;

let pres = new pptxgen();
pres.layout = 'LAYOUT_16x9';
pres.defineSlideMaster({
    title: 'MASTER_V11',
    background: { color: T.primary },
    objects: [{ rect: { x: 0, y: 0, w: 0.04, h: "100%", fill: { color: T.secondary } } }]
});

let usedImages = new Set();
function getUniqueImage(idx) {
    const pool = S.images.filter(img => !usedImages.has(img));
    if (pool.length === 0) return S.images[idx % S.images.length]; 
    const selected = pool[Math.floor(Math.random() * pool.length)];
    usedImages.add(selected);
    return selected;
}

// 1. 封面
let cover = pres.addSlide({ masterName: 'MASTER_V11' });
const cImg = getUniqueImage(0);
if (cImg) cover.addImage({ path: cImg, x: 5, y: 0.5, w: 4.5, h: 4.5, sizing: { type: 'cover' } });
cover.addText(rawData[0].text.substring(0, 35), { x: 0.6, y: 1.5, w: 4, fontSize: 32, bold: true, color: T.text });

// 2. 內容頁 (解除 slice 限制，處理完整 1-24)
rawData.slice(1).forEach((pData, idx) => {
    let slide = pres.addSlide({ masterName: 'MASTER_V11' });
    slide.addText(pData.summary || "核心主題總結", { x: 0.6, y: 0.4, w: 8, fontSize: 22, bold: true, color: T.text });
    const chunks = [ pData.summary, `組織層次：${pData.topic}`, `關鍵技術：${pData.keywords[0] || "AI"}`, `發展目標：${pData.keywords[1] || "創新"}` ];
    chunks.slice(0, 4).forEach((c, ci) => {
        slide.addText(c.substring(0, 25), { x: 0.6, y: 1.2 + ci * 0.8, w: 4.5, h: 0.6, fontSize: 13, color: T.text, bullet: true, fill: { color: T.cardBg, transparency: 85 } });
    });
    const iPath = getUniqueImage(idx + 1);
    if (iPath) slide.addImage({ path: iPath, x: 5.5, y: 1.0, w: 4, h: 4, sizing: { type: 'contain' } });
});

// v11: 最小值 25 頁補全邏輯
let currentSlides = pres.slides.length;
const minSlides = 25;
if (currentSlides < minSlides - 1) { // 留一頁給思維導圖
    const needed = (minSlides - 1) - currentSlides;
    const topics = Object.keys(analysis.topic_page_map);
    for (let i = 0; i < needed; i++) {
        let gapSlide = pres.addSlide({ masterName: 'MASTER_V11' });
        const topic = topics[i % topics.length];
        gapSlide.addText(`${topic} - 深度洞察補強`, { x: 0.6, y: 0.4, w: 8, fontSize: 22, bold: true, color: T.accent });
        gapSlide.addText(`針對「${topic}」章節之核心關鍵字：${analysis.top_keywords.slice(i, i+3).join(", ")} 進行深度摘要與價值總結。`, { x: 0.6, y: 1.5, w: 8, fontSize: 14, color: T.text });
        gapSlide.addText("此頁為自動生成之分段補強，確保簡報完整度。項目：領導、實踐、創新。", { x: 0.6, y: 3, w: 8, fontSize: 12, color: T.muted, italic: true });
    }
}

// 3. 末頁：思維導圖 (修正座標損壞)
let mindSlide = pres.addSlide({ masterName: 'MASTER_V11' });
const centerX = 5.0, centerY = 2.8;
mindSlide.addShape(pres.shapes.OVAL, { x: centerX - 0.7, y: centerY - 0.4, w: 1.4, h: 0.8, fill: { color: T.secondary } });
mindSlide.addText("思維導圖", { x: centerX - 0.7, y: centerY - 0.4, w: 1.4, h: 0.8, fontSize: 13, bold: true, color: "#FFFFFF", align: "center", valign: "middle" });

const topics = Object.keys(analysis.topic_page_map).slice(0, 8);
topics.forEach((topic, ti) => {
    const rad = (ti * (360 / topics.length)) * (Math.PI / 180);
    const dist = 2.4;
    const dx = Math.cos(rad) * dist, dy = Math.sin(rad) * dist;
    // 修正：修正線條 w, h 絕對值過小的問題
    const lineW = Math.abs(dx) < 0.1 ? 0.1 : dx;
    const lineH = Math.abs(dy) < 0.1 ? 0.1 : dy;
    
    mindSlide.addShape(pres.shapes.LINE, { x: centerX, y: centerY, w: lineW, h: lineH, line: { color: T.accent, width: 2, dashType: 'dash' } });
    const tx = centerX + dx - 0.6, ty = centerY + dy - 0.25;
    mindSlide.addShape(pres.shapes.RECTANGLE, { x: tx, y: ty, w: 1.2, h: 0.5, fill: { color: T.cardBg }, shadow: { type: 'outer', blur: 3, offset: 2 } });
    mindSlide.addText(topic, { x: tx, y: ty, w: 1.2, h: 0.5, fontSize: 10, color: T.text, align: "center", valign: "middle" });
});

pres.writeFile({ fileName: path.join(__dirname, "output", "v11_Rescue_Final.pptx") }).then(fn => console.log(`[SUCCESS] v11 QC: ${fn}`));
