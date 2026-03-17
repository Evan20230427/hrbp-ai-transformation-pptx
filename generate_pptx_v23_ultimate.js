const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

/**
 * [Dynamic Engine v3] 整合 Theme Factory + Canvas Design + Frontend Design
 * 核心能力：根據內容關鍵字動態匹配主題、佈局與內容感知插圖
 */

// ======== Theme Factory 主題庫 ========
const THEMES = {
    tech: {
        name: "Tech Innovation",
        primary: "1E1E1E", secondary: "0066FF", accent: "00FFFF", text: "FFFFFF",
        fontTitle: "Microsoft JhengHei", fontBody: "Arial"
    },
    ocean: {
        name: "Ocean Depths",
        primary: "1A2332", secondary: "2D8B8B", accent: "A8DADC", text: "F1FAEE",
        fontTitle: "Microsoft JhengHei", fontBody: "Arial"
    },
    classic: {
        name: "Roman Classic v23",
        primary: "F1F5F9", secondary: "B45309", accent: "475569", text: "1E293B",
        fontTitle: "Microsoft JhengHei", fontBody: "Arial"
    }
};

// ======== Canvas Design 插圖庫  ========
// 根據主題區段對應的內容感知插圖 (由 generate_image 產出)
const CANVAS_IMAGES = {
    cover: "C:\\Users\\TW-Evan.Chen\\.gemini\\antigravity\\brain\\ac6f3712-34eb-4b1e-b0f3-0fb2ebf77cf6\\slide_ai_skills_cover_1773730513816.png",
    marketing: "C:\\Users\\TW-Evan.Chen\\.gemini\\antigravity\\brain\\ac6f3712-34eb-4b1e-b0f3-0fb2ebf77cf6\\slide_marketing_ai_1773730529257.png",
    sales: "C:\\Users\\TW-Evan.Chen\\.gemini\\antigravity\\brain\\ac6f3712-34eb-4b1e-b0f3-0fb2ebf77cf6\\slide_sales_training_1773730547009.png",
    engineering: "C:\\Users\\TW-Evan.Chen\\.gemini\\antigravity\\brain\\ac6f3712-34eb-4b1e-b0f3-0fb2ebf77cf6\\slide_engineering_explore_1773730579435.png",
    hackathon: "C:\\Users\\TW-Evan.Chen\\.gemini\\antigravity\\brain\\ac6f3712-34eb-4b1e-b0f3-0fb2ebf77cf6\\slide_garage_hackathon_1773730597833.png",
    bestPractices: "C:\\Users\\TW-Evan.Chen\\.gemini\\antigravity\\brain\\ac6f3712-34eb-4b1e-b0f3-0fb2ebf77cf6\\slide_best_practices_1773730614566.png",
    copilot: "C:\\Users\\TW-Evan.Chen\\.gemini\\antigravity\\brain\\ac6f3712-34eb-4b1e-b0f3-0fb2ebf77cf6\\slide_copilot_tools_1773730631865.png"
};

/**
 * 根據頁面文字內容智慧匹配最佳插圖
 * @param {string} text - 頁面文字內容
 * @param {number} idx - 頁面索引 (用於 fallback 輪替)
 * @returns {string|null} 匹配的圖片路徑
 */
function matchImageByContent(text, idx) {
    const t = text.toLowerCase();
    if (t.includes("行銷") || t.includes("marketing") || t.includes("習慣")) return CANVAS_IMAGES.marketing;
    if (t.includes("銷售") || t.includes("sales") || t.includes("mcaps")) return CANVAS_IMAGES.sales;
    if (t.includes("工程") || t.includes("engineering") || t.includes("全球學習")) return CANVAS_IMAGES.engineering;
    if (t.includes("garage") || t.includes("hackathon") || t.includes("駭客松") || t.includes("skillup")) return CANVAS_IMAGES.hackathon;
    if (t.includes("copilot") || t.includes("杯挑戰")) return CANVAS_IMAGES.copilot;
    if (t.includes("最佳做法") || t.includes("best practice") || t.includes("計劃")) return CANVAS_IMAGES.bestPractices;
    // fallback：輪替所有插圖
    const allImgs = Object.values(CANVAS_IMAGES);
    return allImgs[idx % allImgs.length];
}

// ======== 載入 OCR 數據 ========
const JSON_PATH = path.join(__dirname, "extracted_content.json");
if (!fs.existsSync(JSON_PATH)) {
    console.error("[ERROR] extracted_content.json not found.");
    process.exit(1);
}
const rawData = JSON.parse(fs.readFileSync(JSON_PATH, "utf-8"));

// ======== 智慧主題偵測 ========
const fullText = rawData.map(p => p.text).join(" ");
let theme = THEMES.classic;
if (fullText.includes("AI") || fullText.includes("Microsoft") || fullText.includes("技術")) {
    theme = THEMES.tech;
} else if (fullText.includes("管理") || fullText.includes("商務") || fullText.includes("專業")) {
    theme = THEMES.ocean;
}
console.log(`[THEME] Applied: ${theme.name}`);

// ======== 初始化簡報 ========
let pres = new pptxgen();
pres.layout = 'LAYOUT_16x9';
pres.author = 'Antigravity Dynamic Engine v3';
pres.title = rawData[0].text.substring(0, 40);

// Master Slide
pres.defineSlideMaster({
    title: 'DYNAMIC_V3',
    background: { color: theme.primary },
    objects: [
        { rect: { x: 0, y: 0, w: 0.08, h: "100%", fill: { color: theme.accent } } },
        { rect: { x: 0, y: 5.1, w: "100%", h: 0.4, fill: { color: theme.secondary, transparency: 80 } } }
    ]
});

// ======== 封面頁 ========
const mainTitle = rawData[0].text.substring(0, 35);
let cover = pres.addSlide({ masterName: 'DYNAMIC_V3' });

// 封面全屏背景圖 (canvas-design 產出)
if (fs.existsSync(CANVAS_IMAGES.cover)) {
    cover.addImage({ path: CANVAS_IMAGES.cover, x: 0, y: 0, w: '100%', h: '100%', sizing: { type: 'cover' } });
}
// 半透明文字底框 (glassmorphism)
cover.addShape(pres.shapes.RECTANGLE, { x: 0.3, y: 2.8, w: 9.4, h: 2.2, fill: { color: theme.primary, transparency: 25 }, rectRadius: 0.1 });
cover.addText(mainTitle, { x: 0.6, y: 3.0, w: 8.8, fontSize: 36, bold: true, color: theme.text, fontFace: theme.fontTitle, align: "center" });
cover.addShape(pres.shapes.LINE, { x: 3, y: 3.9, w: 4, h: 0, line: { color: theme.accent, width: 2 } });
cover.addText(`${theme.name} | Canvas Design Edition`, { x: 0.6, y: 4.1, w: 8.8, fontSize: 16, color: theme.accent, align: "center" });

// ======== 內容頁生成 ========
rawData.slice(1).forEach((pData, idx) => {
    let slide = pres.addSlide({ masterName: 'DYNAMIC_V3' });
    let pageNum = idx + 2;

    // 智慧標題擷取
    const words = pData.text.split(/\s+/).filter(w => w.length > 2);
    const pageTitle = words.slice(0, 5).join(" ").substring(0, 30) || `分析 ${pageNum}`;
    const bodyText = words.slice(5, 60).join(" ").substring(0, 500);

    // 頂部導航條 (frontend-design: Navy 深色頂欄)
    slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: '100%', h: 0.85, fill: { color: theme.secondary } });
    slide.addText(pageTitle, { x: 0.5, y: 0.2, w: 7, h: 0.45, fontSize: 22, bold: true, color: theme.text, fontFace: theme.fontTitle });
    slide.addText(`Page ${pageNum}`, { x: 8.5, y: 0.3, w: 1, h: 0.3, fontSize: 12, color: theme.accent, align: "right" });

    // 左右交替佈局 (frontend-design: 8px grid + asymmetric)
    const imgLeft = (pageNum % 2 === 0);
    const imgX = imgLeft ? 0.3 : 5.2;
    const txtX = imgLeft ? 5.2 : 0.3;

    // 內容感知插圖 (canvas-design)
    const matchedImg = matchImageByContent(pData.text, idx);
    if (matchedImg && fs.existsSync(matchedImg)) {
        slide.addImage({ path: matchedImg, x: imgX, y: 1.1, w: 4.5, h: 3.8, sizing: { type: 'contain' } });
    }

    // 文字區塊
    slide.addText(bodyText, {
        x: txtX, y: 1.1, w: 4.5, h: 3.8,
        fontSize: 14, color: theme.text, fontFace: theme.fontBody,
        lineSpacing: 20, valign: "top", paraSpaceAfter: 6
    });

    // 底部裝飾線 (frontend-design: accent separator)
    slide.addShape(pres.shapes.LINE, { x: 0.3, y: 5.05, w: 9.4, h: 0, line: { color: theme.accent, width: 1, dashType: 'dash' } });
});

// ======== 輸出 ========
const safeFilename = mainTitle.replace(/[\\/:"*?<>|]/g, "_").trim();
const outputDir = path.join(__dirname, "output");
if (!fs.existsSync(outputDir)) fs.mkdirSync(outputDir, { recursive: true });
const outPath = path.join(outputDir, `${safeFilename}_CanvasDesign_${theme.name.replace(/ /g, "_")}.pptx`);

pres.writeFile({ fileName: outPath }).then(fn => {
    console.log(`[SUCCESS] Canvas Design PPTX at: ${fn}`);
}).catch(err => {
    console.error(`[ERROR] Render failed: ${err.message}`);
    process.exit(1);
});
