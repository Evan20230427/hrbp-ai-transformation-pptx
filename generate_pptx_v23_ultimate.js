const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

/**
 * [Dynamic Engine v4] PPTX Skill 排版引擎
 * 整合：Theme Factory + Canvas Design + Frontend Design + PPTX Skill
 * 核心升級：5 種佈局模式循環、卡片陰影、統計面板、精密對齊
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

// ======== Canvas Design 內容感知插圖 ========
const CANVAS_IMAGES = {
    cover:       "C:\\Users\\TW-Evan.Chen\\.gemini\\antigravity\\brain\\ac6f3712-34eb-4b1e-b0f3-0fb2ebf77cf6\\slide_ai_skills_cover_1773730513816.png",
    marketing:   "C:\\Users\\TW-Evan.Chen\\.gemini\\antigravity\\brain\\ac6f3712-34eb-4b1e-b0f3-0fb2ebf77cf6\\slide_marketing_ai_1773730529257.png",
    sales:       "C:\\Users\\TW-Evan.Chen\\.gemini\\antigravity\\brain\\ac6f3712-34eb-4b1e-b0f3-0fb2ebf77cf6\\slide_sales_training_1773730547009.png",
    engineering: "C:\\Users\\TW-Evan.Chen\\.gemini\\antigravity\\brain\\ac6f3712-34eb-4b1e-b0f3-0fb2ebf77cf6\\slide_engineering_explore_1773730579435.png",
    hackathon:   "C:\\Users\\TW-Evan.Chen\\.gemini\\antigravity\\brain\\ac6f3712-34eb-4b1e-b0f3-0fb2ebf77cf6\\slide_garage_hackathon_1773730597833.png",
    bestPractices:"C:\\Users\\TW-Evan.Chen\\.gemini\\antigravity\\brain\\ac6f3712-34eb-4b1e-b0f3-0fb2ebf77cf6\\slide_best_practices_1773730614566.png",
    copilot:     "C:\\Users\\TW-Evan.Chen\\.gemini\\antigravity\\brain\\ac6f3712-34eb-4b1e-b0f3-0fb2ebf77cf6\\slide_copilot_tools_1773730631865.png"
};

/**
 * 根據頁面文字內容智慧匹配插圖
 * @param {string} text - 頁面文字
 * @param {number} idx - 頁面索引
 * @returns {string|null} 圖片路徑
 */
function matchImage(text, idx) {
    const t = text.toLowerCase();
    if (t.includes("行銷") || t.includes("marketing")) return CANVAS_IMAGES.marketing;
    if (t.includes("銷售") || t.includes("sales") || t.includes("mcaps")) return CANVAS_IMAGES.sales;
    if (t.includes("工程") || t.includes("engineering") || t.includes("全球學習")) return CANVAS_IMAGES.engineering;
    if (t.includes("garage") || t.includes("hackathon") || t.includes("駭客松")) return CANVAS_IMAGES.hackathon;
    if (t.includes("copilot")) return CANVAS_IMAGES.copilot;
    if (t.includes("最佳做法") || t.includes("計劃")) return CANVAS_IMAGES.bestPractices;
    const all = Object.values(CANVAS_IMAGES);
    return all[idx % all.length];
}

/**
 * 建立不重複的 shadow 物件 (PPTX Skill: 禁止物件複用)
 */
function makeShadow() {
    return { type: "outer", color: "000000", blur: 4, offset: 2, angle: 135, opacity: 0.12 };
}

// ======== 載入 OCR 數據 ========
const JSON_PATH = path.join(__dirname, "extracted_content.json");
if (!fs.existsSync(JSON_PATH)) { console.error("[ERROR] extracted_content.json not found."); process.exit(1); }
const rawData = JSON.parse(fs.readFileSync(JSON_PATH, "utf-8"));

// ======== 智慧主題偵測 ========
const fullText = rawData.map(p => p.text).join(" ");
let T = THEMES.classic;
if (fullText.includes("AI") || fullText.includes("Microsoft") || fullText.includes("技術")) T = THEMES.tech;
else if (fullText.includes("管理") || fullText.includes("商務")) T = THEMES.ocean;
console.log(`[THEME] ${T.name}`);

// ======== 初始化簡報 ========
let pres = new pptxgen();
pres.layout = 'LAYOUT_16x9';
pres.author = 'Antigravity Engine v4';
pres.title = rawData[0].text.substring(0, 40);

// Master: 深色基底 + 左側 accent 邊條
pres.defineSlideMaster({
    title: 'MASTER_V4',
    background: { color: T.primary },
    objects: [
        { rect: { x: 0, y: 0, w: 0.06, h: "100%", fill: { color: T.secondary } } }
    ]
});

const mainTitle = rawData[0].text.substring(0, 35);

// ======== [SLIDE 1] 封面 - 半出血圖片佈局 ========
let cover = pres.addSlide({ masterName: 'MASTER_V4' });
const coverImg = CANVAS_IMAGES.cover;
if (fs.existsSync(coverImg)) {
    cover.addImage({ path: coverImg, x: 4.5, y: 0, w: 5.5, h: 5.625, sizing: { type: 'cover', w: 5.5, h: 5.625 } });
}
// 左側文字區
cover.addText(mainTitle, {
    x: 0.6, y: 1.5, w: 4.0, h: 2.0, fontSize: 36, bold: true,
    color: T.text, fontFace: T.fontTitle, align: "left", margin: 0
});
cover.addText(`${T.name} | Canvas Design Edition`, {
    x: 0.6, y: 3.6, w: 4.0, fontSize: 14, color: T.muted, fontFace: T.fontBody, align: "left"
});

// ======== 5 種佈局模式 (PPTX Skill 規範: 禁止重複佈局) ========
const LAYOUTS = [
    "TWO_COLUMN",       // 雙欄：左圖右文
    "CARD_GRID",        // 卡片格柵：2x2 卡片
    "HALF_BLEED_RIGHT", // 半出血：右側滿版圖
    "STAT_CALLOUT",     // 統計面板：大型數字 + 文字
    "ICON_ROWS"         // 圖標行列：左側 accent 條 + 縮排文字
];

/**
 * 從文字中嘗試提取數字統計 (用於 STAT_CALLOUT 佈局)
 */
function extractStats(text) {
    const nums = text.match(/\d+%?/g);
    if (nums && nums.length >= 1) return nums.slice(0, 3);
    return ["AI", "10", "Best"];
}

// ======== 內容頁生成 ========
rawData.slice(1).forEach((pData, idx) => {
    let slide = pres.addSlide({ masterName: 'MASTER_V4' });
    let pageNum = idx + 2;

    // 文字預處理
    const words = pData.text.split(/\s+/).filter(w => w.length > 1);
    const pageTitle = words.slice(0, 6).join(" ").substring(0, 35);
    const bodyLines = [];
    let buf = "";
    words.slice(6).forEach(w => {
        buf += w + " ";
        if (buf.length > 60) { bodyLines.push(buf.trim()); buf = ""; }
    });
    if (buf.trim()) bodyLines.push(buf.trim());

    // 格式化頁碼 (PPTX Skill: caption 10-12pt muted)
    const pageLabel = `Page ${pageNum}`;

    // 選擇佈局模式 (循環)
    const layoutMode = LAYOUTS[idx % LAYOUTS.length];
    const imgPath = matchImage(pData.text, idx);
    const hasImg = imgPath && fs.existsSync(imgPath);

    switch (layoutMode) {
        case "TWO_COLUMN": {
            // 左圖右文雙欄 (PPTX Skill: two-column layout)
            slide.addText(pageTitle, {
                x: 0.5, y: 0.4, w: 9, h: 0.5, fontSize: 24, bold: true,
                color: T.text, fontFace: T.fontTitle, align: "left", margin: 0
            });
            if (hasImg) {
                slide.addImage({ path: imgPath, x: 0.5, y: 1.2, w: 4.2, h: 3.5, sizing: { type: 'contain', w: 4.2, h: 3.5 } });
            }
            slide.addText(bodyLines.slice(0, 5).map(l => ({
                text: l, options: { bullet: true, breakLine: true, fontSize: 14, color: T.text, fontFace: T.fontBody }
            })), { x: 5.0, y: 1.2, w: 4.5, h: 3.5, valign: "top", paraSpaceAfter: 8 });
            break;
        }

        case "CARD_GRID": {
            // 2x2 卡片格柵 (PPTX Skill: shadow cards + grid)
            slide.addText(pageTitle, {
                x: 0.5, y: 0.3, w: 9, h: 0.5, fontSize: 24, bold: true,
                color: T.text, fontFace: T.fontTitle, align: "left", margin: 0
            });
            const cardPositions = [
                { x: 0.5, y: 1.1, w: 4.3, h: 1.9 },
                { x: 5.2, y: 1.1, w: 4.3, h: 1.9 },
                { x: 0.5, y: 3.2, w: 4.3, h: 1.9 },
                { x: 5.2, y: 3.2, w: 4.3, h: 1.9 }
            ];
            cardPositions.forEach((pos, ci) => {
                const cardText = bodyLines[ci] || "";
                slide.addShape(pres.shapes.RECTANGLE, {
                    ...pos, fill: { color: T.cardBg }, shadow: makeShadow()
                });
                // 卡片左側 accent 邊條
                slide.addShape(pres.shapes.RECTANGLE, {
                    x: pos.x, y: pos.y, w: 0.06, h: pos.h, fill: { color: T.secondary }
                });
                slide.addText(cardText, {
                    x: pos.x + 0.2, y: pos.y + 0.2, w: pos.w - 0.4, h: pos.h - 0.4,
                    fontSize: 13, color: T.text, fontFace: T.fontBody, valign: "top", align: "left"
                });
            });
            break;
        }

        case "HALF_BLEED_RIGHT": {
            // 右側滿版圖 + 左側文字 (PPTX Skill: half-bleed image)
            if (hasImg) {
                slide.addImage({ path: imgPath, x: 5, y: 0, w: 5, h: 5.625, sizing: { type: 'cover', w: 5, h: 5.625 } });
            }
            slide.addText(pageTitle, {
                x: 0.5, y: 0.5, w: 4.2, h: 0.6, fontSize: 24, bold: true,
                color: T.text, fontFace: T.fontTitle, align: "left", margin: 0
            });
            slide.addText(bodyLines.slice(0, 6).join("\n"), {
                x: 0.5, y: 1.3, w: 4.2, h: 3.5, fontSize: 14,
                color: T.text, fontFace: T.fontBody, valign: "top", align: "left", paraSpaceAfter: 6
            });
            break;
        }

        case "STAT_CALLOUT": {
            // 大型統計數字面板 (PPTX Skill: 60-72pt stat callouts)
            const stats = extractStats(pData.text);
            slide.addText(pageTitle, {
                x: 0.5, y: 0.3, w: 9, h: 0.5, fontSize: 24, bold: true,
                color: T.text, fontFace: T.fontTitle, align: "left", margin: 0
            });
            // 三格統計面板
            stats.forEach((stat, si) => {
                const sx = 0.5 + si * 3.2;
                slide.addShape(pres.shapes.RECTANGLE, {
                    x: sx, y: 1.1, w: 2.8, h: 2.0, fill: { color: T.cardBg }, shadow: makeShadow()
                });
                slide.addText(stat, {
                    x: sx, y: 1.2, w: 2.8, h: 1.2, fontSize: 48, bold: true,
                    color: T.secondary, fontFace: T.fontTitle, align: "center", margin: 0
                });
                const label = bodyLines[si] ? bodyLines[si].substring(0, 30) : "";
                slide.addText(label, {
                    x: sx + 0.2, y: 2.4, w: 2.4, h: 0.5, fontSize: 11,
                    color: T.muted, fontFace: T.fontBody, align: "center"
                });
            });
            // 下方補充文字
            slide.addText(bodyLines.slice(3, 6).join(" "), {
                x: 0.5, y: 3.4, w: 9, h: 1.6, fontSize: 13,
                color: T.text, fontFace: T.fontBody, valign: "top", align: "left"
            });
            break;
        }

        case "ICON_ROWS": {
            // 圖標行列：左側色條 + 分段文字 (PPTX Skill: icon + text rows)
            slide.addText(pageTitle, {
                x: 0.5, y: 0.3, w: 9, h: 0.5, fontSize: 24, bold: true,
                color: T.text, fontFace: T.fontTitle, align: "left", margin: 0
            });
            bodyLines.slice(0, 4).forEach((line, li) => {
                const ry = 1.1 + li * 1.05;
                // 圓形色標
                slide.addShape(pres.shapes.OVAL, {
                    x: 0.6, y: ry + 0.15, w: 0.35, h: 0.35, fill: { color: T.secondary }
                });
                // 序號
                slide.addText(`${li + 1}`, {
                    x: 0.6, y: ry + 0.15, w: 0.35, h: 0.35, fontSize: 14, bold: true,
                    color: T.text, fontFace: T.fontBody, align: "center", valign: "middle", margin: 0
                });
                // 文字
                slide.addText(line, {
                    x: 1.2, y: ry, w: 8.3, h: 0.7, fontSize: 14,
                    color: T.text, fontFace: T.fontBody, valign: "middle", align: "left"
                });
            });
            // 右下角小圖
            if (hasImg) {
                slide.addImage({ path: imgPath, x: 7, y: 3.8, w: 2.5, h: 1.5, sizing: { type: 'contain', w: 2.5, h: 1.5 } });
            }
            break;
        }
    }

    // 頁碼 (統一置底右側, PPTX Skill: caption 10-12pt muted)
    slide.addText(pageLabel, {
        x: 8.5, y: 5.2, w: 1, h: 0.3, fontSize: 10, color: T.muted, align: "right", fontFace: T.fontBody
    });
});

// ======== 輸出 ========
const safeFilename = mainTitle.replace(/[\\/:"*?<>|]/g, "_").trim();
const outputDir = path.join(__dirname, "output");
if (!fs.existsSync(outputDir)) fs.mkdirSync(outputDir, { recursive: true });
const outPath = path.join(outputDir, `${safeFilename}_v4_${T.name.replace(/ /g, "_")}.pptx`);

pres.writeFile({ fileName: outPath }).then(fn => {
    console.log(`[SUCCESS] v4 PPTX at: ${fn}`);
}).catch(err => {
    console.error(`[ERROR] Render failed: ${err.message}`);
    process.exit(1);
});
