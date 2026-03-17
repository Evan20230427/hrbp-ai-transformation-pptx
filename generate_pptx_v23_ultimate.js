const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

// ======== 視覺風格定義 ========
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

// v10: 單一風格鎖定
const styleKeys = Object.keys(STYLES);
const lockedStyleKey = styleKeys[Math.floor(Math.random() * styleKeys.length)];
const S = STYLES[lockedStyleKey];
const T = S.theme;
console.log(`[QC v10] Style Locked: ${S.name}`);

let pres = new pptxgen();
pres.layout = 'LAYOUT_16x9';
pres.defineSlideMaster({
    title: 'MASTER_V10',
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

// 內容頁
rawData.slice(1, 10).forEach((pData, idx) => {
    let slide = pres.addSlide({ masterName: 'MASTER_V10' });
    slide.addText(pData.summary, { x: 0.6, y: 0.4, w: 8, fontSize: 24, bold: true, color: T.text });
    
    // v10 限制：4 字塊，不超過 25 字，不重疊
    const chunks = [ pData.summary, `主題脈絡：${pData.topic}`, `關鍵特徵：${pData.keywords[0] || "AI"}`, `重點標籤：${pData.keywords[1] || "精煉"}` ];
    chunks.forEach((chunk, ci) => {
        slide.addText(chunk.substring(0, 25), { x: 0.6, y: 1.2 + ci * 0.8, w: 4.5, h: 0.6, fontSize: 13, color: T.text, bullet: true, fill: { color: T.cardBg, transparency: 80 } });
    });
    
    const imgPath = getUniqueImage(idx);
    if (imgPath) slide.addImage({ path: imgPath, x: 5.5, y: 1.0, w: 4, h: 4, sizing: { type: 'contain' } });
});

// 末頁：思維導圖
if (analysis) {
    let mindSlide = pres.addSlide({ masterName: 'MASTER_V10' });
    const centerX = 5.0, centerY = 2.8;
    mindSlide.addShape(pres.shapes.OVAL, { x: centerX - 0.8, y: centerY - 0.5, w: 1.6, h: 1.0, fill: { color: T.secondary } });
    mindSlide.addText("思維導圖", { x: centerX - 0.8, y: centerY - 0.5, w: 1.6, h: 1.0, fontSize: 14, bold: true, color: "#FFFFFF", align: "center", valign: "middle" });
    const topics = Object.keys(analysis.topic_page_map);
    topics.forEach((topic, ti) => {
        const rad = (ti * (360 / topics.length)) * (Math.PI / 180);
        const tx = centerX + Math.cos(rad) * 2.5 - 0.7, ty = centerY + Math.sin(rad) * 2.5 - 0.3;
        mindSlide.addShape(pres.shapes.LINE, { x: centerX, y: centerY, w: Math.cos(rad) * 1.8, h: Math.sin(rad) * 1.8, line: { color: T.accent, width: 2 } });
        mindSlide.addShape(pres.shapes.RECTANGLE, { x: tx, y: ty, w: 1.4, h: 0.6, fill: { color: T.cardBg } });
        mindSlide.addText(topic, { x: tx, y: ty, w: 1.4, h: 0.6, fontSize: 10, color: T.text, align: "center", valign: "middle" });
    });
}

pres.writeFile({ fileName: path.join(__dirname, "output", "v10_Final_QC.pptx") }).then(fn => console.log(`[SUCCESS] v10 QC: ${fn}`));
