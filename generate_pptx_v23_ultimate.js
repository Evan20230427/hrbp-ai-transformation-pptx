const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

/**
 * [Dynamic Engine v8] 高端極簡美學版 (Museum Edition)
 * 修正：徹底移除繽紛向量風格，回歸「瑞士極簡」、「建築寫實」、「精細線稿」之高品質美學。
 */

// ======== Theme Factory ========
const THEMES = {
    tech: {
        name: "Tech Innovation",
        primary: "1E1E1E", secondary: "0066FF", accent: "00FFFF",
        text: "FFFFFF", muted: "94A3B8", cardBg: "2D2D2D",
        fontTitle: "Microsoft JhengHei", fontBody: "Calibri"
    },
    ocean: {
        name: "Ocean Depths",
        primary: "1A2332", secondary: "2D8B8B", accent: "A8DADC",
        text: "F1FAEE", muted: "8E9AAF", cardBg: "243447",
        fontTitle: "Microsoft JhengHei", fontBody: "Calibri"
    },
    classic: {
        name: "Roman Classic",
        primary: "F1F5F9", secondary: "B45309", accent: "475569",
        text: "1E293B", muted: "64748B", cardBg: "FFFFFF",
        fontTitle: "Microsoft JhengHei", fontBody: "Calibri"
    }
};

// ======== Museum Quality 插圖庫 (極簡、專業、高端) ========
const STYLE_IMAGES = {
    swiss:   "C:\\Users\\TW-Evan.Chen\\.gemini\\antigravity\\brain\\ac6f3712-34eb-4b1e-b0f3-0fb2ebf77cf6\\style_minimalist_swiss_logic_1773732459484.png",
    arch:    "C:\\Users\\TW-Evan.Chen\\.gemini\\antigravity\\brain\\ac6f3712-34eb-4b1e-b0f3-0fb2ebf77cf6\\style_architectural_realism_ai_1773732476963.png",
    statue:  "C:\\Users\\TW-Evan.Chen\\.gemini\\antigravity\\brain\\ac6f3712-34eb-4b1e-b0f3-0fb2ebf77cf6\\style_fine_line_statue_v23_1773732497567.png",
    content: {
        marketing:   "C:\\Users\\TW-Evan.Chen\\.gemini\\antigravity\\brain\\ac6f3712-34eb-4b1e-b0f3-0fb2ebf77cf6\\slide_marketing_ai_1773730529257.png",
        sales:       "C:\\Users\\TW-Evan.Chen\\.gemini\\antigravity\\brain\\ac6f3712-34eb-4b1e-b0f3-0fb2ebf77cf6\\slide_sales_training_1773730547009.png",
        engineering: "C:\\Users\\TW-Evan.Chen\\.gemini\\antigravity\\brain\\ac6f3712-34eb-4b1e-b0f3-0fb2ebf77cf6\\slide_engineering_explore_1773730579435.png",
        hackathon:   "C:\\Users\\TW-Evan.Chen\\.gemini\\antigravity\\brain\\ac6f3712-34eb-4b1e-b0f3-0fb2ebf77cf6\\slide_garage_hackathon_1773730597833.png"
    }
};

/**
 * 高端選圖邏輯：移除繽紛向量，僅保留極簡專業風格
 */
function selectImage(text, idx) {
    const t = text.toLowerCase();
    
    // 1. 內容具象匹配 (僅在具備高度關聯時)
    if (t.includes("行銷") || t.includes("marketing")) return STYLE_IMAGES.content.marketing;
    if (t.includes("銷售") || t.includes("sales")) return STYLE_IMAGES.content.sales;
    if (t.includes("工程") || t.includes("engineering")) return STYLE_IMAGES.content.engineering;
    if (t.includes("garage") || t.includes("hackathon")) return STYLE_IMAGES.content.hackathon;
    
    // 2. 風格循環 (高端極簡)
    const museumPool = [STYLE_IMAGES.swiss, STYLE_IMAGES.arch, STYLE_IMAGES.statue];
    return museumPool[idx % museumPool.length];
}

function makeShadow() {
    return { type: "outer", color: "000000", blur: 4, offset: 2, angle: 135, opacity: 0.1 };
}

/**
 * 自然斷句：確保文字不跨塊，維持語意完整
 */
function splitToSentences(text, maxLen = 65) {
    const rawSegments = text.split(/(?<=[。！？；\n])/g).filter(s => s.trim().length > 0);
    const lines = [];
    let buffer = "";

    for (const seg of rawSegments) {
        if (buffer.length + seg.length <= maxLen) {
            buffer += seg;
        } else {
            if (buffer.trim()) lines.push(buffer.trim());
            buffer = seg;
        }
    }
    if (buffer.trim()) lines.push(buffer.trim());
    return lines;
}

// ======== 載入數據 ========
const JSON_PATH = path.join(__dirname, "extracted_content.json");
if (!fs.existsSync(JSON_PATH)) process.exit(1);
const jsonData = JSON.parse(fs.readFileSync(JSON_PATH, "utf-8"));
const rawData = jsonData.pages || jsonData;
const analysis = jsonData.analysis || null;

// ======== 初始化 ========
const fullText = rawData.map(p => p.text).join(" ");
let T = THEMES.tech; // 微軟專案預設高端深色
if (fullText.includes("Classic")) T = THEMES.classic;

let pres = new pptxgen();
pres.layout = 'LAYOUT_16x9';
pres.author = 'Antigravity Museum Engine v8';

pres.defineSlideMaster({
    title: 'MASTER_V8',
    background: { color: T.primary },
    objects: [{ rect: { x: 0, y: 0, w: 0.05, h: "100%", fill: { color: T.secondary } } }]
});

const mainTitle = rawData[0].text.substring(0, 35);
const LAYOUTS = ["TWO_COLUMN", "HALF_BLEED_RIGHT", "STAT_CALLOUT", "ICON_ROWS"];

// ======== 封面：建築寫實風 ========
let cover = pres.addSlide({ masterName: 'MASTER_V8' });
if (fs.existsSync(STYLE_IMAGES.arch)) {
    cover.addImage({ path: STYLE_IMAGES.arch, x: 4.5, y: 0, w: 5.5, h: 5.625, sizing: { type: 'cover' } });
}
cover.addText(mainTitle, { x: 0.6, y: 1.5, w: 4.0, h: 2.0, fontSize: 36, bold: true, color: T.text, fontFace: T.fontTitle, align: "left" });
cover.addText("Leadership & Innovation Academy", { x: 0.6, y: 3.6, w: 4.0, fontSize: 14, color: T.accent, fontFace: T.fontBody });

// ======== 內容頁 (自然排版) ========
rawData.slice(1).forEach((pData, idx) => {
    let slide = pres.addSlide({ masterName: 'MASTER_V8' });
    const pageTitle = pData.summary ? pData.summary.substring(0, 35) : pData.text.substring(0, 35);
    const bodyLines = splitToSentences(pData.text, 60);
    const layoutMode = LAYOUTS[idx % LAYOUTS.length];
    const imgPath = selectImage(pData.text, idx);
    const hasImg = imgPath && fs.existsSync(imgPath);

    slide.addText(pageTitle, { x: 0.6, y: 0.4, w: 8, h: 0.5, fontSize: 22, bold: true, color: T.text, fontFace: T.fontTitle, align: "left" });
    
    if (layoutMode === "TWO_COLUMN" || layoutMode === "HALF_BLEED_RIGHT") {
        const isLeft = (layoutMode === "TWO_COLUMN");
        if (hasImg) slide.addImage({ path: imgPath, x: isLeft ? 0.6 : 5.8, y: 1.2, w: 3.6, h: 3.6, sizing: { type: 'contain' } });
        slide.addText(bodyLines.slice(0, 7).map(l => ({ text: l, options: { bullet: true, breakLine: true } })), { 
            x: isLeft ? 4.8 : 0.6, y: 1.2, w: 4.6, h: 3.6, fontSize: 13, color: T.text, fontFace: T.fontBody, valign: "top", paraSpaceAfter: 8 
        });
    } else {
        slide.addText(bodyLines.slice(0, 10).map(l => ({ text: l, options: { bullet: true, breakLine: true } })), { 
            x: 0.6, y: 1.2, w: 6.5, h: 3.6, fontSize: 13, color: T.text, fontFace: T.fontBody, valign: "top", paraSpaceAfter: 6 
        });
        if (hasImg) slide.addImage({ path: imgPath, x: 7.2, y: 3.5, w: 2.3, h: 1.8, sizing: { type: 'contain' } });
    }
    slide.addText(`P.${idx+2}`, { x: 8.8, y: 5.2, w: 0.7, fontSize: 10, color: T.muted, align: "right" });
});

// ======== 末頁：脈絡主張結構 (極簡版) ========
if (analysis) {
    let idxSlide = pres.addSlide({ masterName: 'MASTER_V8' });
    idxSlide.addText("核心脈絡與策略主張 / Insights Strategy", { x: 0.6, y: 0.3, w: 8, fontSize: 22, bold: true, color: T.secondary });
    
    const topicMap = analysis.topic_page_map || {};
    const topicEntries = Object.entries(topicMap);
    const startY = 1.0;
    const rowH = 0.8;

    idxSlide.addShape(pres.shapes.LINE, { x: 0.8, y: startY, w: 0, h: topicEntries.length * rowH, line: { color: T.secondary, width: 2 } });

    topicEntries.forEach(([topic, pages], ti) => {
        const y = startY + ti * rowH;
        idxSlide.addShape(pres.shapes.OVAL, { x: 0.65, y: y, w: 0.3, h: 0.3, fill: { color: T.secondary } });
        idxSlide.addText(topic, { x: 1.1, y: y-0.05, w: 3, fontSize: 14, bold: true, color: T.text });
        
        const topicKws = analysis.top_keywords.slice(ti * 3, (ti + 1) * 3);
        idxSlide.addText(`主題主張：${topicKws.join(" | ")}`, { x: 1.1, y: y + 0.3, w: 5, fontSize: 10, color: T.muted, italic: true });
        idxSlide.addText(`Pages: ${pages.join(", ")}`, { x: 7, y: y, w: 2.5, fontSize: 11, bold: true, color: T.accent, align: "right" });
        idxSlide.addShape(pres.shapes.LINE, { x: 0.8, y: y + 0.65, w: 8.7, h: 0, line: { color: T.muted, width: 0.5, dashType: 'dash' } });
    });
}

const safeName = mainTitle.replace(/[\\/:"*?<>|]/g, "_").trim();
const outPath = path.join(__dirname, "output", `${safeName}_v8_Museum.pptx`);
pres.writeFile({ fileName: outPath }).then(fn => console.log(`[SUCCESS] v8 Museum at: ${fn}`));
