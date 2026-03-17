const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

/**
 * [Dynamic Engine v5] 整合全部 Skills
 * PDF + Doc-Coauthoring + Theme Factory + Canvas Design + Frontend Design + PPTX Skill
 * 新增：結構化摘要、關鍵字分析、末頁索引頁
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

// ======== Canvas Design 插圖 ========
const CANVAS_IMAGES = {
    cover:        "C:\\Users\\TW-Evan.Chen\\.gemini\\antigravity\\brain\\ac6f3712-34eb-4b1e-b0f3-0fb2ebf77cf6\\slide_ai_skills_cover_1773730513816.png",
    marketing:    "C:\\Users\\TW-Evan.Chen\\.gemini\\antigravity\\brain\\ac6f3712-34eb-4b1e-b0f3-0fb2ebf77cf6\\slide_marketing_ai_1773730529257.png",
    sales:        "C:\\Users\\TW-Evan.Chen\\.gemini\\antigravity\\brain\\ac6f3712-34eb-4b1e-b0f3-0fb2ebf77cf6\\slide_sales_training_1773730547009.png",
    engineering:  "C:\\Users\\TW-Evan.Chen\\.gemini\\antigravity\\brain\\ac6f3712-34eb-4b1e-b0f3-0fb2ebf77cf6\\slide_engineering_explore_1773730579435.png",
    hackathon:    "C:\\Users\\TW-Evan.Chen\\.gemini\\antigravity\\brain\\ac6f3712-34eb-4b1e-b0f3-0fb2ebf77cf6\\slide_garage_hackathon_1773730597833.png",
    bestPractices:"C:\\Users\\TW-Evan.Chen\\.gemini\\antigravity\\brain\\ac6f3712-34eb-4b1e-b0f3-0fb2ebf77cf6\\slide_best_practices_1773730614566.png",
    copilot:      "C:\\Users\\TW-Evan.Chen\\.gemini\\antigravity\\brain\\ac6f3712-34eb-4b1e-b0f3-0fb2ebf77cf6\\slide_copilot_tools_1773730631865.png"
};

function matchImage(text, idx) {
    const t = text.toLowerCase();
    if (t.includes("行銷") || t.includes("marketing")) return CANVAS_IMAGES.marketing;
    if (t.includes("銷售") || t.includes("sales") || t.includes("mcaps")) return CANVAS_IMAGES.sales;
    if (t.includes("工程") || t.includes("engineering")) return CANVAS_IMAGES.engineering;
    if (t.includes("garage") || t.includes("hackathon") || t.includes("駭客松")) return CANVAS_IMAGES.hackathon;
    if (t.includes("copilot")) return CANVAS_IMAGES.copilot;
    if (t.includes("最佳做法") || t.includes("計劃")) return CANVAS_IMAGES.bestPractices;
    const all = Object.values(CANVAS_IMAGES);
    return all[idx % all.length];
}

function makeShadow() {
    return { type: "outer", color: "000000", blur: 4, offset: 2, angle: 135, opacity: 0.12 };
}

// ======== 載入結構化 JSON (v5 新格式) ========
const JSON_PATH = path.join(__dirname, "extracted_content.json");
if (!fs.existsSync(JSON_PATH)) { console.error("[ERROR] extracted_content.json not found."); process.exit(1); }
const jsonData = JSON.parse(fs.readFileSync(JSON_PATH, "utf-8"));

// 相容新舊 JSON 格式
const rawData = jsonData.pages || jsonData;
const analysis = jsonData.analysis || null;

// ======== 智慧主題偵測 ========
const fullText = rawData.map(p => p.text).join(" ");
let T = THEMES.classic;
if (fullText.includes("AI") || fullText.includes("Microsoft") || fullText.includes("技術")) T = THEMES.tech;
else if (fullText.includes("管理") || fullText.includes("商務")) T = THEMES.ocean;
console.log(`[THEME] ${T.name}`);

// ======== 初始化簡報 ========
let pres = new pptxgen();
pres.layout = 'LAYOUT_16x9';
pres.author = 'Antigravity Engine v5';
pres.title = rawData[0].text.substring(0, 40);

pres.defineSlideMaster({
    title: 'MASTER_V5',
    background: { color: T.primary },
    objects: [
        { rect: { x: 0, y: 0, w: 0.06, h: "100%", fill: { color: T.secondary } } }
    ]
});

const mainTitle = rawData[0].text.substring(0, 35);
const LAYOUTS = ["TWO_COLUMN", "CARD_GRID", "HALF_BLEED_RIGHT", "STAT_CALLOUT", "ICON_ROWS"];

function extractStats(text) {
    const nums = text.match(/\d+%?/g);
    if (nums && nums.length >= 1) return nums.slice(0, 3);
    return ["AI", "10", "Best"];
}

// ======== [SLIDE 1] 封面 ========
let cover = pres.addSlide({ masterName: 'MASTER_V5' });
const coverImg = CANVAS_IMAGES.cover;
if (fs.existsSync(coverImg)) {
    cover.addImage({ path: coverImg, x: 4.5, y: 0, w: 5.5, h: 5.625, sizing: { type: 'cover', w: 5.5, h: 5.625 } });
}
cover.addText(mainTitle, {
    x: 0.6, y: 1.5, w: 4.0, h: 2.0, fontSize: 36, bold: true,
    color: T.text, fontFace: T.fontTitle, align: "left", margin: 0
});
cover.addText(`${T.name} | Engine v5`, {
    x: 0.6, y: 3.6, w: 4.0, fontSize: 14, color: T.muted, fontFace: T.fontBody, align: "left"
});

// ======== 內容頁 (5 種佈局循環) ========
rawData.slice(1).forEach((pData, idx) => {
    let slide = pres.addSlide({ masterName: 'MASTER_V5' });
    let pageNum = idx + 2;

    // doc-coauthoring: 使用結構化摘要作為標題
    const pageTitle = pData.summary
        ? pData.summary.substring(0, 35)
        : pData.text.split(/\s+/).filter(w => w.length > 1).slice(0, 6).join(" ").substring(0, 35);

    const words = pData.text.split(/\s+/).filter(w => w.length > 1);
    const bodyLines = [];
    let buf = "";
    words.slice(6).forEach(w => {
        buf += w + " ";
        if (buf.length > 60) { bodyLines.push(buf.trim()); buf = ""; }
    });
    if (buf.trim()) bodyLines.push(buf.trim());

    const layoutMode = LAYOUTS[idx % LAYOUTS.length];
    const imgPath = matchImage(pData.text, idx);
    const hasImg = imgPath && fs.existsSync(imgPath);

    // 主題標籤 (doc-coauthoring: 結構化分類)
    const topicTag = pData.topic || "";

    switch (layoutMode) {
        case "TWO_COLUMN": {
            slide.addText(pageTitle, {
                x: 0.5, y: 0.4, w: 9, h: 0.5, fontSize: 24, bold: true,
                color: T.text, fontFace: T.fontTitle, align: "left", margin: 0
            });
            if (topicTag) {
                slide.addShape(pres.shapes.RECTANGLE, { x: 8.0, y: 0.4, w: 1.5, h: 0.35, fill: { color: T.secondary } });
                slide.addText(topicTag, { x: 8.0, y: 0.4, w: 1.5, h: 0.35, fontSize: 9, color: T.text, align: "center", fontFace: T.fontBody });
            }
            if (hasImg) slide.addImage({ path: imgPath, x: 0.5, y: 1.2, w: 4.2, h: 3.5, sizing: { type: 'contain', w: 4.2, h: 3.5 } });
            slide.addText(bodyLines.slice(0, 5).map(l => ({
                text: l, options: { bullet: true, breakLine: true, fontSize: 14, color: T.text, fontFace: T.fontBody }
            })), { x: 5.0, y: 1.2, w: 4.5, h: 3.5, valign: "top", paraSpaceAfter: 8 });
            break;
        }
        case "CARD_GRID": {
            slide.addText(pageTitle, {
                x: 0.5, y: 0.3, w: 9, h: 0.5, fontSize: 24, bold: true,
                color: T.text, fontFace: T.fontTitle, align: "left", margin: 0
            });
            const cardPos = [
                { x: 0.5, y: 1.1, w: 4.3, h: 1.9 }, { x: 5.2, y: 1.1, w: 4.3, h: 1.9 },
                { x: 0.5, y: 3.2, w: 4.3, h: 1.9 }, { x: 5.2, y: 3.2, w: 4.3, h: 1.9 }
            ];
            cardPos.forEach((pos, ci) => {
                slide.addShape(pres.shapes.RECTANGLE, { ...pos, fill: { color: T.cardBg }, shadow: makeShadow() });
                slide.addShape(pres.shapes.RECTANGLE, { x: pos.x, y: pos.y, w: 0.06, h: pos.h, fill: { color: T.secondary } });
                slide.addText(bodyLines[ci] || "", {
                    x: pos.x + 0.2, y: pos.y + 0.2, w: pos.w - 0.4, h: pos.h - 0.4,
                    fontSize: 13, color: T.text, fontFace: T.fontBody, valign: "top", align: "left"
                });
            });
            break;
        }
        case "HALF_BLEED_RIGHT": {
            if (hasImg) slide.addImage({ path: imgPath, x: 5, y: 0, w: 5, h: 5.625, sizing: { type: 'cover', w: 5, h: 5.625 } });
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
            const stats = extractStats(pData.text);
            slide.addText(pageTitle, {
                x: 0.5, y: 0.3, w: 9, h: 0.5, fontSize: 24, bold: true,
                color: T.text, fontFace: T.fontTitle, align: "left", margin: 0
            });
            stats.forEach((stat, si) => {
                const sx = 0.5 + si * 3.2;
                slide.addShape(pres.shapes.RECTANGLE, { x: sx, y: 1.1, w: 2.8, h: 2.0, fill: { color: T.cardBg }, shadow: makeShadow() });
                slide.addText(stat, { x: sx, y: 1.2, w: 2.8, h: 1.2, fontSize: 48, bold: true, color: T.secondary, fontFace: T.fontTitle, align: "center", margin: 0 });
                slide.addText(bodyLines[si] ? bodyLines[si].substring(0, 30) : "", { x: sx + 0.2, y: 2.4, w: 2.4, h: 0.5, fontSize: 11, color: T.muted, fontFace: T.fontBody, align: "center" });
            });
            slide.addText(bodyLines.slice(3, 6).join(" "), { x: 0.5, y: 3.4, w: 9, h: 1.6, fontSize: 13, color: T.text, fontFace: T.fontBody, valign: "top", align: "left" });
            break;
        }
        case "ICON_ROWS": {
            slide.addText(pageTitle, {
                x: 0.5, y: 0.3, w: 9, h: 0.5, fontSize: 24, bold: true,
                color: T.text, fontFace: T.fontTitle, align: "left", margin: 0
            });
            bodyLines.slice(0, 4).forEach((line, li) => {
                const ry = 1.1 + li * 1.05;
                slide.addShape(pres.shapes.OVAL, { x: 0.6, y: ry + 0.15, w: 0.35, h: 0.35, fill: { color: T.secondary } });
                slide.addText(`${li + 1}`, { x: 0.6, y: ry + 0.15, w: 0.35, h: 0.35, fontSize: 14, bold: true, color: T.text, fontFace: T.fontBody, align: "center", valign: "middle", margin: 0 });
                slide.addText(line, { x: 1.2, y: ry, w: 8.3, h: 0.7, fontSize: 14, color: T.text, fontFace: T.fontBody, valign: "middle", align: "left" });
            });
            if (hasImg) slide.addImage({ path: imgPath, x: 7, y: 3.8, w: 2.5, h: 1.5, sizing: { type: 'contain', w: 2.5, h: 1.5 } });
            break;
        }
    }
    slide.addText(`Page ${pageNum}`, { x: 8.5, y: 5.2, w: 1, h: 0.3, fontSize: 10, color: T.muted, align: "right", fontFace: T.fontBody });
});

// ======== [末頁] 索引頁：關鍵字 + 邏輯關係 + 頁碼對應 ========
if (analysis) {
    let idxSlide = pres.addSlide({ masterName: 'MASTER_V5' });
    const totalPages = rawData.length + 1; // +1 for this index slide

    // 標題
    idxSlide.addText("簡報索引 / Presentation Index", {
        x: 0.5, y: 0.3, w: 9, h: 0.5, fontSize: 28, bold: true,
        color: T.secondary, fontFace: T.fontTitle, align: "left", margin: 0
    });

    // ── 區塊 A：全文關鍵字 ──
    idxSlide.addText("全文關鍵字 Top Keywords", {
        x: 0.5, y: 0.95, w: 4, h: 0.3, fontSize: 14, bold: true,
        color: T.accent, fontFace: T.fontTitle, align: "left", margin: 0
    });
    const kwTags = (analysis.top_keywords || []).slice(0, 12);
    kwTags.forEach((kw, ki) => {
        const col = ki % 4;
        const row = Math.floor(ki / 4);
        const tx = 0.5 + col * 2.3;
        const ty = 1.35 + row * 0.45;
        idxSlide.addShape(pres.shapes.RECTANGLE, {
            x: tx, y: ty, w: 2.1, h: 0.35, fill: { color: T.cardBg }, shadow: makeShadow()
        });
        idxSlide.addText(kw, {
            x: tx, y: ty, w: 2.1, h: 0.35, fontSize: 12, bold: true,
            color: T.text, fontFace: T.fontBody, align: "center", valign: "middle"
        });
    });

    // ── 區塊 B：主題 → 頁碼對應表 (PPTX Skill: addTable) ──
    idxSlide.addText("主題邏輯關係 Topic-Page Map", {
        x: 0.5, y: 2.85, w: 9, h: 0.3, fontSize: 14, bold: true,
        color: T.accent, fontFace: T.fontTitle, align: "left", margin: 0
    });

    const topicMap = analysis.topic_page_map || {};
    const tableHeader = [
        [
            { text: "主題 (Topic)", options: { bold: true, color: "FFFFFF", fill: { color: T.secondary }, fontSize: 12, fontFace: T.fontTitle } },
            { text: "對應頁碼 (Pages)", options: { bold: true, color: "FFFFFF", fill: { color: T.secondary }, fontSize: 12, fontFace: T.fontTitle } },
            { text: "頁數", options: { bold: true, color: "FFFFFF", fill: { color: T.secondary }, fontSize: 12, fontFace: T.fontTitle } }
        ]
    ];
    const tableRows = Object.entries(topicMap).map(([topic, pages]) => [
        { text: topic, options: { fontSize: 11, fontFace: T.fontBody, color: T.text } },
        { text: pages.join(", "), options: { fontSize: 11, fontFace: T.fontBody, color: T.muted } },
        { text: `${pages.length}`, options: { fontSize: 11, fontFace: T.fontBody, color: T.accent, bold: true, align: "center" } }
    ]);

    idxSlide.addTable([...tableHeader, ...tableRows], {
        x: 0.5, y: 3.2, w: 9, colW: [3.5, 4, 1.5],
        border: { pt: 0.5, color: T.muted },
        rowH: 0.35,
        fill: { color: T.cardBg }
    });

    // 頁碼
    idxSlide.addText(`Page ${totalPages}`, { x: 8.5, y: 5.2, w: 1, h: 0.3, fontSize: 10, color: T.muted, align: "right", fontFace: T.fontBody });

    console.log(`[INDEX] Generated index slide with ${kwTags.length} keywords & ${Object.keys(topicMap).length} topics`);
} else {
    console.log("[INDEX] No analysis data found, skipping index slide.");
}

// ======== 輸出 ========
const safeFilename = mainTitle.replace(/[\\/:"*?<>|]/g, "_").trim();
const outputDir = path.join(__dirname, "output");
if (!fs.existsSync(outputDir)) fs.mkdirSync(outputDir, { recursive: true });
const outPath = path.join(outputDir, `${safeFilename}_v5_${T.name.replace(/ /g, "_")}.pptx`);

pres.writeFile({ fileName: outPath }).then(fn => {
    console.log(`[SUCCESS] v5 PPTX at: ${fn}`);
}).catch(err => {
    console.error(`[ERROR] Render failed: ${err.message}`);
    process.exit(1);
});
