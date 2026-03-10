const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

let pres = new pptxgen();
pres.layout = 'LAYOUT_16x9';
pres.author = 'Antigravity AI';
pres.title = 'AI 時代重新設計 HRBP 角色的最佳實務 - 奢華旗艦版';

// CURATED LUXURY THEME
const THEME = {
    primary: "0F172A",    // Deep Navy
    secondary: "3B82F6",  // Vibrant Blue
    accent: "10B981",     // Emerald
    highlight: "F59E0B",  // Amber
    bg: "FFFFFF",
    bg_alt: "F8FAFC",
    text: "1E293B",
    subtle: "64748B",
    white: "FFFFFF",
    line: "E2E8F0"
};

const FONT_TCH = "Microsoft JhengHei";
const FONT_BODY = "Arial";

// Image Paths
const IMG_DIR = "C:\\Users\\TW-Evan.Chen\\.gemini\\antigravity\\brain\\cb95dffe-33bd-4e40-a98b-feaff376ea1a";
const SCRATCH_DIR = "C:\\Users\\TW-Evan.Chen\\.gemini\\antigravity\\scratch\\pdf-xlsx-to-pptx";

const IMAGES = {
    gif: path.join(SCRATCH_DIR, "hrbp_animated_visual.gif"),
    dashboard: path.join(IMG_DIR, "hrbp_luxury_dashboard_v9_1773159099389.png"),
    collab: path.join(IMG_DIR, "hrbp_ai_collaboration_v9_1773159117982.png"),
    ladder: path.join(IMG_DIR, "hrbp_visionary_ladder_v9_1773159135984.png"),
    map_bg: path.join(IMG_DIR, "hrbp_transformation_mindmap_infographic_1773157697477.png")
};

/**
 * Text Engine for Overlap Prevention & Style
 */
function createRichText(lines, baseSize = 16) {
    const finalContent = [];
    const smallSize = Math.max(12, baseSize - 3);
    [...new Set(lines)].forEach((line, idx) => {
        const regex = /(\([^)]+\))/g;
        const tokens = line.split(regex);
        tokens.forEach((token, tIdx) => {
            const isEng = token.match(regex);
            finalContent.push({
                text: token,
                options: {
                    fontSize: isEng ? smallSize : baseSize,
                    color: isEng ? THEME.secondary : THEME.text,
                    fontFace: isEng ? FONT_BODY : FONT_TCH,
                    italic: isEng ? true : false,
                    bullet: (tIdx === 0) ? true : false,
                    breakLine: (tIdx === tokens.length - 1)
                }
            });
        });
    });
    return finalContent;
}

// MASTER SLIDE: LUXURY GRID
pres.defineSlideMaster({
    title: 'LUXURY_GRID',
    background: { color: THEME.bg },
    objects: [
        { rect: { x: 0, y: 0, w: 0.2, h: "100%", fill: { color: THEME.secondary } } }, // Left Accent Bar
        { text: { text: "Gartner Strategy | Luxury Edition v9", options: { x: 0.5, y: 5.3, w: 9, h: 0.25, fontSize: 10, color: THEME.subtle, align: "right", fontFace: FONT_BODY } } }
    ]
});

function applyHeader(slide, title) {
    slide.addText(title, { x: 0.5, y: 0.4, w: 9, h: 0.6, fontSize: 28, bold: true, color: THEME.primary, fontFace: FONT_TCH });
    slide.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.0, w: 5, h: 0.04, fill: { color: THEME.secondary } });
}

// --- SLIDES ---

// 1. Title Slide (Dynamic Layout)
let s1 = pres.addSlide();
s1.background = { color: THEME.primary };
if (fs.existsSync(IMAGES.ladder)) s1.addImage({ path: IMAGES.ladder, x: 5, y: 0, w: 5, h: "100%", sizing: { type: "cover" } });
s1.addText("重塑 HRBP 角色：\n引領 AI 轉型的策略實務", { x: 0.5, y: 1.8, w: 4.5, h: 1.5, fontSize: 38, bold: true, color: THEME.white, fontFace: FONT_TCH });
s1.addText("2026 年度奢華旗艦版 | Gartner 核心洞察", { x: 0.5, y: 3.3, w: 4.5, h: 0.5, fontSize: 16, color: THEME.secondary, fontFace: FONT_TCH });

// 2. The Paradox (Grid Split)
let s2 = pres.addSlide({ masterName: 'LUXURY_GRID' });
applyHeader(s2, "核心挑戰：策略貢獻與事務束縛");
s2.addText(createRichText([
    "數據揭露：僅 51% (Only 51%) 的領導層同意 HRBP 參與了重大策略討論。",
    "轉型瓶頸：仍被鎖定在 AI 與自動化正在吸收的任務中 (Transactional Heavy)。",
    "自動化威脅：涵蓋職務描述與數據摘要等日常事務 (Standard Tasks)。"
]), { x: 0.6, y: 1.5, w: 5.5, h: 3.5, lineSpacing: 24 });
if (fs.existsSync(IMAGES.dashboard)) s2.addImage({ path: IMAGES.dashboard, x: 6.5, y: 1.2, w: 3, h: 3.8, sizing: { type: "contain" } });

// 3. The STL Role (Dynamic Visual)
let s3 = pres.addSlide({ masterName: 'LUXURY_GRID' });
applyHeader(s3, "未來角色：策略人才領袖 (STL)");
if (fs.existsSync(IMAGES.gif)) s3.addImage({ path: IMAGES.gif, x: 0.6, y: 1.4, w: 4, h: 2.5 });
s3.addText(createRichText([
    "定位：引導 AI 驅動轉型的人員面向 (Consultative Advantage)。",
    "進化：從「解釋人員策略」到「主導轉型對話」。"
], 20), { x: 5.0, y: 1.4, w: 4.5, h: 3, lineSpacing: 30 });

// 4. Three Pillars Deep Dive
const PILLARS = [
    { t: "職責 1：人力重新設計", c: ["主導職能重塑決策 (Redesign)", "決定人才培訓與部署 (Reskill/Redeploy)"], img: IMAGES.collab },
    { t: "職責 2：AI 倫理與偏見", c: ["監測人才決策中的偏見 (Addressing Bias)", "維護企業文化倫理邊界 (Ethics Boundary)"], img: null },
    { t: "職責 3：人機協作效率", c: ["優化人類直覺與 AI 計算之平衡", "重新設計工作流以賦能員工 (Empowerment)"], img: null }
];

PILLARS.forEach(p => {
    let slide = pres.addSlide({ masterName: 'LUXURY_GRID' });
    applyHeader(slide, p.t);
    slide.addText(createRichText(p.c), { x: 0.7, y: 1.8, w: 8.5, h: 2.5, lineSpacing: 28 });
    if (p.img && fs.existsSync(p.img)) {
        slide.addImage({ path: p.img, x: 5.5, y: 2.5, w: 4, h: 2.5, sizing: { type: "cover" } });
    }
});

// ... (Phase slides using Grid Layout to prevent overlap)
const PHASES = [
    { t: "P1 剔除行動：回收策略產能", c: ["定義策略重點區域 (Strategic Focus)", "建立 12-24 個月自動化藍圖 (Roadmap)"] },
    { t: "P2 強化行動：AI 賦能高價值", c: ["更新職能模型，使 AI 準備度透明化", "利用預測洞察作為領導對話基準 (Data-driven)"] },
    { t: "P3 擴展行動：主導新型策略", c: ["啟動 STL Pods 小組領航", "在轉型決策中嵌入強力條款 (Governance)"] }
];
PHASES.forEach(ph => {
    let s = pres.addSlide({ masterName: 'LUXURY_GRID' });
    applyHeader(s, ph.t);
    s.addText(createRichText(ph.c), { x: 1.0, y: 2.0, w: 8, h: 2, lineSpacing: 30 });
});

// Metrics Slide
let sM = pres.addSlide({ masterName: 'LUXURY_GRID' });
applyHeader(sM, "關鍵衡量指標 (Success Metrics)");
sM.addText(createRichText([
    "決策週期縮短 (Cycle Time Reduction)",
    "繼任人才準備率提升 (% Readiness)",
    "遺憾離職率優化 (Regrettable Attrition)",
    "AI 偏見降解指標 (Bias Mitigation)"
], 22), { x: 0.7, y: 1.8, w: 8.5, h: 3, lineSpacing: 32 });

// 15. The Mind Map (Native v8 logic but on v9 luxury bg)
let slideMap = pres.addSlide({ masterName: 'LUXURY_GRID' });
applyHeader(slideMap, "全課精華：三層層級心智圖");
const RX = 0.8, RY = 2.4;
const L1 = [
    { t: "現狀解析", c: ["策略參與不足", "行政作業負擔"] },
    { t: "STL 定義", c: ["人力重設計", "倫理與協作"] },
    { t: "轉型三階", c: ["P1 剔除/P2 強化", "P3 擴展開拓"] },
    { t: "指標量測", c: ["效率週期", "準備/流失"] }
];
slideMap.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: RX, y: RY, w: 1.4, h: 0.6, fill: { color: THEME.primary }, rectRadius: 0.1 });
slideMap.addText("HRBP AI 轉型", { x: RX, y: RY, w: 1.4, h: 0.6, color: THEME.white, bold: true, align: "center", fontFace: FONT_TCH, fontSize: 11 });
L1.forEach((node, i) => {
    let nx = RX + 2.2, ny = 1.0 + (i * 1.25);
    slideMap.addShape(pres.shapes.LINE, { x: RX + 1.4, y: RY + 0.3, w: 0.4, h: 0, line: { color: THEME.secondary, width: 2 } });
    slideMap.addShape(pres.shapes.LINE, { x: RX + 1.8, y: Math.min(RY + 0.3, ny + 0.25), w: 0, h: Math.abs(RY + 0.3 - (ny + 0.25)), line: { color: THEME.secondary, width: 2 } });
    slideMap.addShape(pres.shapes.LINE, { x: RX + 1.8, y: ny + 0.25, w: 0.4, h: 0, line: { color: THEME.secondary, width: 2 } });
    slideMap.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: nx, y: ny, w: 1.4, h: 0.5, fill: { color: THEME.secondary }, rectRadius: 0.1 });
    slideMap.addText(node.t, { x: nx, y: ny, w: 1.4, h: 0.5, color: THEME.white, bold: true, align: "center", fontFace: FONT_TCH, fontSize: 10 });
    node.c.forEach((child, j) => {
        let cx = nx + 2.0, cy = ny - 0.2 + (j * 0.45);
        slideMap.addShape(pres.shapes.LINE, { x: nx + 1.4, y: ny + 0.25, w: 0.2, h: 0, line: { color: THEME.line, width: 1 } });
        slideMap.addShape(pres.shapes.LINE, { x: nx + 1.6, y: Math.min(ny + 0.25, cy + 0.15), w: 0, h: Math.abs(ny + 0.25 - (cy + 0.15)), line: { color: THEME.line, width: 1 } });
        slideMap.addShape(pres.shapes.LINE, { x: nx + 1.6, y: cy + 0.15, w: 0.4, h: 0, line: { color: THEME.line, width: 1 } });
        slideMap.addText(child, { x: cx, y: cy, w: 1.8, h: 0.35, color: THEME.text, fontSize: 9, fontFace: FONT_TCH, valign: "middle" });
    });
});

// Final Slide
let slideLast = pres.addSlide();
slideLast.background = { color: THEME.primary };
slideLast.addText("啟動您的數據領航之旅", { x: 0, y: 2.3, w: "100%", h: 0.6, bold: true, fontSize: 36, color: THEME.white, align: "center", fontFace: FONT_TCH });

const outPath = path.join(SCRATCH_DIR, "output", "HRBP_AI_Transformation_Luxury_v9.pptx");
pres.writeFile({ fileName: outPath }).then(fn => {
    console.log(`Success: Generated Luxury v9 at ${fn}`);
}).catch(err => console.error(err));
