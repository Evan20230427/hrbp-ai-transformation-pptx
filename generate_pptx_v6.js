const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

let pres = new pptxgen();
pres.layout = 'LAYOUT_16x9';
pres.author = 'Antigravity AI';
pres.title = 'HRBP AI 轉型最佳實務 - 終極精修版';

// Professional High-End Theme
const THEME = {
    primary: "0F172A",
    secondary: "2563EB",
    accent: "10B981",
    bg: "FFFFFF",
    text: "1E293B",
    subtle: "64748B",
    white: "FFFFFF",
    line: "CBD5E1"
};

const FONT_TCH = "Microsoft JhengHei";
const FONT_BODY = "Arial";

/**
 * Helper: Process text to scale English in brackets and remove duplicates
 * @param {string[]} lines - Array of bullet strings
 * @param {number} baseSize - Base font size for Chinese
 * @returns {Array} Array of PPTX text object arrays
 */
function formatLines(lines, baseSize = 18) {
    // 1. Remove duplicates while preserving order
    const uniqueLines = [...new Set(lines)];
    const smallSize = Math.max(12, baseSize - 4);

    return uniqueLines.map(line => {
        const parts = [];
        // Regex to split by (English Content)
        const regex = /(\([A-Za-z\s\d,/.%-]+\))/g;
        const tokens = line.split(regex);

        tokens.forEach(token => {
            if (token.match(regex)) {
                parts.push({ text: token, options: { fontSize: smallSize, color: THEME.secondary, fontFace: FONT_BODY, italic: true } });
            } else if (token.trim()) {
                parts.push({ text: token, options: { fontSize: baseSize, fontFace: FONT_TCH, color: THEME.text } });
            }
        });
        return parts;
    });
}

// Master Slide
pres.defineSlideMaster({
    title: 'FIXED_MASTER',
    background: { color: THEME.bg },
    objects: [
        { rect: { x: 0, y: 0, w: "100%", h: 0.1, fill: { color: THEME.secondary } } },
        { text: { text: "© 2026 Gartner Insight | HRBP AI Transformation Strategy", options: { x: 0.5, y: 5.3, w: 9, h: 0.25, fontSize: 10, color: THEME.subtle, align: "right", fontFace: FONT_BODY } } }
    ]
});

function addSlideHeader(slide, title) {
    slide.addText(title, { x: 0.5, y: 0.35, w: 9, h: 0.6, fontSize: 26, bold: true, color: THEME.primary, fontFace: FONT_TCH });
    slide.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 0.95, w: 9, h: 0.02, fill: { color: THEME.secondary } });
}

// --- Slides Generation ---
// Slide 1: Title
let slide1 = pres.addSlide();
slide1.background = { color: THEME.primary };
slide1.addText("重塑 HRBP 角色：\n引領企業 AI 轉型的策略實務", { x: 0, y: 1.8, w: "100%", h: 1.2, fontSize: 44, bold: true, color: THEME.white, fontFace: FONT_TCH, align: "center" });
slide1.addText("2026 全球實務指南 (Gartner Insights)", { x: 0, y: 3.1, w: "100%", h: 0.5, fontSize: 18, color: THEME.secondary, fontFace: FONT_TCH, align: "center" });

// Slide 2: Challenges
let slide2 = pres.addSlide({ masterName: "FIXED_MASTER" });
addSlideHeader(slide2, "現狀挑戰 (The Paradox)");
const s2Lines = [
    "目前僅有 51% 領導者滿意其 HRBP 的策略貢獻 (Strategic Discussions)。",
    "行政事務產能分配不均 (Admin Capacity Allocation)。",
    "AI 正快速吸收事務性工作 (Automation Absorption)。",
    "面臨角色被重複執行與未充分利用的風險 (Risk of Underleverage)。"
];
formatLines(s2Lines).forEach((parts, idx) => {
    slide2.addText(parts, { x: 0.8, y: 1.5 + (idx * 0.7), w: 8.5, h: 0.6, bullet: true, lineSpacing: 28 });
});

// Slide 3: Roles
let slide3 = pres.addSlide({ masterName: "FIXED_MASTER" });
addSlideHeader(slide3, "未來角色：策略人才領袖 (STL)");
const s3Lines = [
    "引導人力重新設計 (Workforce Redesign)。",
    "監測 AI 倫理與偏見方案 (AI Bias & Ethics)。",
    "優化人機協作效率 (Human-Machine Collaboration)。",
    "在重大變革中嵌入人力條款 (Embed STL Remit)。"
];
formatLines(s3Lines).forEach((parts, idx) => {
    slide3.addText(parts, { x: 0.8, y: 1.5 + (idx * 0.7), w: 8.5, h: 0.6, bullet: true, lineSpacing: 28 });
});

// ... (Other slides truncated for script brevity, focusing on the refined logic)
let slide4 = pres.addSlide({ masterName: "FIXED_MASTER" });
addSlideHeader(slide4, "成功指標 (Success Metrics)");
const s4Lines = [
    "週期時間顯著縮短 (Cycle Time Reduction)。",
    "繼任人才準備率提升 (% Successor Readiness)。",
    "遺憾離職率優化 (Regrettable Attrition)。",
    "偏差案例監測與降低 (AI Bias Mitigation)。"
];
formatLines(s4Lines).forEach((parts, idx) => {
    slide4.addText(parts, { x: 0.8, y: 1.5 + (idx * 0.7), w: 8.5, h: 0.6, bullet: true, lineSpacing: 28 });
});

// --- Slide 5: Final Mind Map (3-Level Native) ---
let slideMap = pres.addSlide({ masterName: "FIXED_MASTER" });
addSlideHeader(slideMap, "全課精華：三層深度心智圖");

const CX = 5.0, CY = 2.8;
const data = {
    text: "HRBP AI 轉型",
    children: [
        { text: "挑戰現狀", children: ["51%策略不足", "行政作業佔據", "AI自動化威脅"] },
        { text: "STL 定義", children: ["人力重設計", "倫理監測", "人機協作"] },
        { text: "實作路徑", children: ["P1 剔除行政", "P2 AI 強化", "P3 擴展開拓"] },
        { text: "成功指標", children: ["週期時間縮短", "繼任準備率", "離職率優化"] }
    ]
};

// Draw Root
slideMap.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: CX - 0.8, y: CY - 0.35, w: 1.6, h: 0.7, fill: { color: THEME.primary }, rectRadius: 0.1 });
slideMap.addText(data.text, { x: CX - 0.8, y: CY - 0.35, w: 1.6, h: 0.7, color: THEME.white, bold: true, align: "center", fontSize: 13, fontFace: FONT_TCH });

// Calculate and Draw Layers
data.children.forEach((l1, i) => {
    const angle = (i * 90) - 45; // Spread around center
    const dist1 = 2.2;
    const l1x = CX + dist1 * Math.cos(angle * Math.PI / 180) - 0.7;
    const l1y = CY + dist1 * Math.sin(angle * Math.PI / 180) - 0.25;

    // Line Root -> L1
    slideMap.addShape(pres.shapes.LINE, { x: CX, y: CY, w: (l1x + 0.7 - CX), h: (l1y + 0.25 - CY), line: { color: THEME.secondary, width: 2 } });

    // L1 Node
    slideMap.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: l1x, y: l1y, w: 1.4, h: 0.5, fill: { color: THEME.secondary }, rectRadius: 0.1 });
    slideMap.addText(l1.text, { x: l1x, y: l1y, w: 1.4, h: 0.5, color: THEME.white, bold: true, align: "center", fontSize: 11, fontFace: FONT_TCH });

    // L2 Children
    l1.children.forEach((l2, j) => {
        const dist2 = 1.0;
        const l2angle = angle + (j - 1) * 25; // Narrow fan out
        const l2x = l1x + 0.7 + dist2 * Math.cos(l2angle * Math.PI / 180) - 0.6;
        const l2y = l1y + 0.25 + dist2 * Math.sin(l2angle * Math.PI / 180) - 0.15;

        // Line L1 -> L2
        slideMap.addShape(pres.shapes.LINE, { x: l1x + 0.7, y: l1y + 0.25, w: (l2x + 0.6 - (l1x + 0.7)), h: (l2y + 0.15 - (l1y + 0.25)), line: { color: THEME.line, width: 1 } });

        // L2 Node
        slideMap.addShape(pres.shapes.RECTANGLE, { x: l2x, y: l2y, w: 1.2, h: 0.35, fill: { color: "F8FAFC" }, line: { color: THEME.line, width: 1 } });
        slideMap.addText(l2, { x: l2x, y: l2y, w: 1.2, h: 0.35, color: THEME.text, fontSize: 9, align: "center", fontFace: FONT_TCH });
    });
});

// Save
const outDir = path.join("C:\\Users\\TW-Evan.Chen\\.gemini\\antigravity\\scratch\\Skills_Workspace", "Output");
if (!fs.existsSync(outDir)) fs.mkdirSync(outDir);
const outPath = path.join(outDir, "HRBP_AI_Transformation_Refined_v6.pptx");

pres.writeFile({ fileName: outPath }).then(fn => {
    console.log(`Successfully generated Refined PPTX at ${fn}`);
}).catch(err => console.error(err));
