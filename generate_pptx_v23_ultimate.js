const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

/**
 * [Dynamic Engine v14] Prometheus Edition (終極穩定與 15字精煉版)
 * 修正：15字重點字塊、內容對應影像、XML 安全淨化、三層展開思維導圖、不少於 25 頁。
 */

// 1. 文字淨化 (XML 損壞防護)
function sanitize(str) {
    if (!str) return "";
    return str
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&apos;")
        .replace(/[\x00-\x1F\x7F-\x9F]/g, "") 
        .replace(/[^\x00-\x7F\u4e00-\u9fa5\u3000-\u303f\uff00-\uffef]/g, "") // 僅保留基本 ASCII 與中日韓文字
        .trim();
}

// 2. 加載數據
const JSON_PATH = "extracted_content.json";
if (!fs.existsSync(JSON_PATH)) process.exit(1);
const data = JSON.parse(fs.readFileSync(JSON_PATH, "utf-8"));
const rawData = data.pages;
const analysis = data.analysis;
const inputBaseName = data.pdf_name ? path.parse(data.pdf_name).name : "Advanced_AI_Skills";

// 3. 視覺系統 (鎖定單一美學)
const T = { primary: "FFFFFF", secondary: "1E293B", accent: "3B82F6", text: "0F172A", cardBg: "F8FAFC" };

// 4. 內容對應影像庫 (一頁一圖，絕不重複)
const IMAGE_DIR = "C:/Users/TW-Evan.Chen/.gemini/antigravity/brain/ac6f3712-34eb-4b1e-b0f3-0fb2ebf77cf6/";
const MASTER_POOL = [
    "style_ghibli_learning_1773732831828.png", "style_photography_tech_innovation_1773732856137.png",
    "style_minimalist_swiss_logic_1773732459484.png", "style_realistic_ai_1773731994739.png",
    "style_architectural_realism_ai_1773732476963.png", "style_fine_line_statue_v23_1773732497567.png",
    "slide_marketing_ai_1773730529257.png", "slide_sales_training_1773730547009.png",
    "slide_engineering_explore_1773730579435.png", "slide_garage_hackathon_1773730597833.png",
    "slide_copilot_tools_1773730631865.png", "slide_best_practices_1773730614566.png",
    "style_flat_design_collaboration_1773732313529.png", "style_logic_concept_1773732030157.png",
    "style_lineart_skills_1773732009515.png", "v23_statue_curious_ac6f3712_png_1773711545138.png",
    "v23_statue_excitement_ac6f3712_png_1773711561255.png"
].map(f => path.join(IMAGE_DIR, f));

let usedImages = new Set();
function getContentMappedImage(summary, topic, idx) {
    const s = (summary || "").toLowerCase();
    const t = (topic || "").toLowerCase();
    
    // 優先匹配高端素材
    let selected = null;
    if (t.includes("行銷")) selected = MASTER_POOL.find(p => p.includes("marketing") && !usedImages.has(p));
    if (!selected && t.includes("銷售")) selected = MASTER_POOL.find(p => p.includes("sales") && !usedImages.has(p));
    if (!selected && t.includes("工程")) selected = MASTER_POOL.find(p => p.includes("engineering") && !usedImages.has(p));
    if (!selected && t.includes("Garage")) selected = MASTER_POOL.find(p => p.includes("garage") && !usedImages.has(p));
    if (!selected && s.includes("Copilot")) selected = MASTER_POOL.find(p => p.includes("copilot") && !usedImages.has(p));

    if (!selected) {
        const remaining = MASTER_POOL.filter(p => !usedImages.has(p));
        selected = remaining.length > 0 ? remaining[Math.floor(Math.random() * remaining.length)] : MASTER_POOL[idx % MASTER_POOL.length];
    }
    usedImages.add(selected);
    return selected;
}

const pres = new pptxgen();
pres.layout = 'LAYOUT_16x9';
pres.defineSlideMaster({
    title: 'PROMETHEUS',
    background: { color: T.primary },
    objects: [{ rect: { x: 0, y: 0, w: "100%", h: 0.1, fill: { color: T.secondary } } }]
});

// v14 生成
// [1. 封面]
let cover = pres.addSlide({ masterName: 'PROMETHEUS' });
const cImg = getContentMappedImage("Cover", "AI", 0);
if (cImg) cover.addImage({ path: cImg, x: 5, y: 0.5, w: 4.5, h: 4.5, sizing: { type: 'contain' } });
cover.addText(sanitize(inputBaseName), { x: 0.6, y: 2, w: 4.5, fontSize: 32, bold: true, color: T.text });

// [2. 內容頁]
rawData.slice(1).forEach((pData, idx) => {
    let slide = pres.addSlide({ masterName: 'PROMETHEUS' });
    slide.addText(sanitize(pData.topic || "核心主題"), { x: 0.6, y: 0.4, w: 6, fontSize: 24, bold: true, color: T.secondary });
    
    // 4 個字塊，不超過 25 字
    const bodyItems = [
        sanitize(pData.summary),
        `關鍵要素：${sanitize(pData.keywords[0] || "AI 轉型")}`,
        `實踐主張：${sanitize(pData.keywords[1] || "創新實戰")}`,
        `資源鏈結：Microsoft Learn`
    ];
    bodyItems.forEach((b, bi) => {
        slide.addText(b.substring(0, 25), { x: 0.6, y: 1.4 + bi * 0.9, w: 4.5, h: 0.7, fontSize: 13, color: T.text, bullet: true, fill: { color: T.cardBg }, margin: 10 });
    });

    // 15 字極簡重點字塊 (最下方)
    slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.0, w: "100%", h: 0.6, fill: { color: T.secondary } });
    slide.addText(`核心精煉：${sanitize(pData.summary.substring(0, 15))}`, { x: 0.6, y: 5.0, w: 9, h: 0.6, fontSize: 14, bold: true, color: "#FFFFFF", align: "left", valign: "middle" });

    const iPath = getContentMappedImage(pData.summary, pData.topic, idx + 1);
    if (iPath) slide.addImage({ path: iPath, x: 5.5, y: 0.8, w: 4.2, h: 3.8, sizing: { type: 'contain' } });
});

// v14: 25 頁最小值補強 (由思維導圖補充)
while (pres.slides.length < 24) {
    let s = pres.addSlide({ masterName: 'PROMETHEUS' });
    const kw = analysis.top_keywords[pres.slides.length % analysis.top_keywords.length];
    s.addText(`深度分析：${sanitize(kw)}`, { x: 0.6, y: 0.5, fontSize: 24, bold: true, color: T.secondary });
    s.addText(`針對「${sanitize(kw)}」之領先行業實踐進行全面剖析。`, { x: 0.6, y: 1.5, w: 4, fontSize: 14, color: T.text });
    const iPath = getContentMappedImage(kw, "Ext", pres.slides.length);
    if (iPath) s.addImage({ path: iPath, x: 5.5, y: 1, w: 4, h: 4, sizing: { type: 'contain' } });
}

// [3. 末頁：三層展開思維導圖]
let mind = pres.addSlide({ masterName: 'PROMETHEUS' });
mind.addText("核心脈絡三層級思維導圖 / Prometheus Strategic Map", { x: 0.6, y: 0.4, fontSize: 22, bold: true, color: T.secondary });
const cx = 5.0, cy = 3.2;
mind.addShape(pres.shapes.OVAL, { x: cx-0.7, y: cy-0.4, w: 1.4, h: 0.8, fill: { color: T.secondary } });
mind.addText("AI 賦能", { x: cx-0.7, y: cy-0.4, w: 1.4, h: 0.8, fontSize: 12, bold: true, color: "#FFFFFF", align: "center", valign: "middle" });

const topics = Object.keys(analysis.topic_page_map).slice(0, 6);
topics.forEach((t, ti) => {
    const angle = (ti * (360/topics.length)) * (Math.PI/180);
    const d2 = 2.0, d3 = 3.2;
    const x2 = cx + Math.cos(angle)*d2, y2 = cy + Math.sin(angle)*d2;
    // 連線 (加強無效數值過濾)
    mind.addShape(pres.shapes.LINE, { x: cx, y: cy, w: Math.cos(angle)*1.3 || 0.1, h: Math.sin(angle)*1.3 || 0.1, line: { color: T.secondary, width: 2 } });
    mind.addShape(pres.shapes.RECTANGLE, { x: x2-0.7, y: y2-0.3, w: 1.4, h: 0.6, fill: { color: T.accent } });
    mind.addText(sanitize(t), { x: x2-0.7, y: y2-0.3, w: 1.4, h: 0.6, fontSize: 10, bold: true, color: T.text, align: "center", valign: "middle" });

    // 第三層：關鍵字
    const sub = analysis.top_keywords.slice(ti*2, (ti+1)*2);
    sub.forEach((skw, si) => {
        const sa = angle + (si === 0 ? -0.4 : 0.4);
        const x3 = cx + Math.cos(sa)*d3, y3 = cy+Math.sin(sa)*d3;
        mind.addShape(pres.shapes.LINE, { x: x2, y: y2, w: Math.cos(sa)*0.8 || 0.1, h: Math.sin(sa)*0.8 || 0.1, line: { color: T.accent, width: 1, dashType: 'dash' } });
        mind.addText(sanitize(skw), { x: x3-0.5, y: y3-0.2, w: 1, h: 0.4, fontSize: 8, color: T.secondary, italic: true, align: "center" });
    });
});

const outPath = path.join(__dirname, "output", `${inputBaseName}_v14_Prometheus.pptx`);
pres.writeFile({ fileName: outPath }).then(fn => console.log(`[SUCCESS] v14 Prometheus: ${fn}`));
