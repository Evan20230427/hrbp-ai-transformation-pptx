const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

let pres = new pptxgen();
pres.layout = 'LAYOUT_16x9';
pres.author = 'Antigravity AI';
pres.title = 'AI 時代重新設計 HRBP 角色的最佳實務 (視覺增強版)';

// Solid Professional Theme
const THEME = {
    primary: "0F172A",    // Deep Navy
    secondary: "2563EB",  // Vibrant Blue
    accent: "059669",     // Emerald
    bg: "FFFFFF",
    text: "1E293B",
    subtle: "64748B",
    white: "FFFFFF",
    line: "CBD5E1"
};

const FONT_BODY = "Arial";
const FONT_TCH = "Microsoft JhengHei";

// Image Paths (Absolute Paths from brainstorming session)
const IMG_DIR = "C:\\Users\\TW-Evan.Chen\\.gemini\\antigravity\\brain\\cb95dffe-33bd-4e40-a98b-feaff376ea1a";
const IMAGES = {
    title: path.join(IMG_DIR, "hrbp_ai_title_visual_1773157399585.png"),
    stl: path.join(IMG_DIR, "hrbp_stl_consultant_1773157416001.png"),
    roadmap: path.join(IMG_DIR, "hrbp_transformation_roadmap_1773157436793.png"),
    success: path.join(IMG_DIR, "hrbp_success_metrics_data_1773157454088.png")
};

// Define Slide Master
pres.defineSlideMaster({
    title: 'STL_MASTER',
    background: { color: THEME.bg },
    objects: [
        { rect: { x: 0, y: 0, w: "100%", h: 0.1, fill: { color: THEME.secondary } } },
        { text: { text: "Gartner | Redesigning HRBP to Fuel AI Transformation", options: { x: 0.5, y: 5.3, w: 9, h: 0.25, fontSize: 10, color: THEME.subtle, align: "right", fontFace: FONT_BODY } } }
    ]
});

function addTitle(slide, text) {
    slide.addText(text, { x: 0.5, y: 0.3, w: 9, h: 0.6, fontSize: 26, bold: true, color: THEME.primary, fontFace: FONT_TCH });
    slide.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 0.9, w: 9, h: 0.02, fill: { color: THEME.secondary } });
}

// --- Slide 1: Title ---
let slide1 = pres.addSlide();
slide1.background = { color: THEME.primary };
if (fs.existsSync(IMAGES.title)) {
    slide1.addImage({ path: IMAGES.title, x: 0, y: 0, w: "100%", h: "100%", sizing: { type: "cover" }, transparency: 30 });
}
slide1.addText("重塑 HRBP 角色：\n引領企業 AI 轉型的策略實務", { x: 0, y: 1.8, w: "100%", h: 1.2, fontSize: 44, bold: true, color: THEME.white, fontFace: FONT_TCH, align: "center", shadow: { type: "outer", color: "000000", opacity: 0.5 } });
slide1.addText("From Administrator to AI Transformation Consultant", { x: 0, y: 3.1, w: "100%", h: 0.5, fontSize: 18, color: THEME.secondary, fontFace: FONT_BODY, align: "center" });

// --- Slide 2: Challenge ---
let slide2 = pres.addSlide({ masterName: "STL_MASTER" });
addTitle(slide2, "現狀挑戰：HRBP 的策略鴻溝");
slide2.addText([
    { text: "關鍵數據：只有 51% 的領導者滿意其 HRBP 的策略轉化能力。\n", options: { bold: true, color: THEME.secondary, breakLine: true } },
    { text: "困境：行政瑣事佔據過多產能。\n", options: { bullet: true, breakLine: true } },
    { text: "AI 模型崛起：正在快速自動化傳統的 HRBP 任務。", options: { bullet: true } }
], { x: 0.7, y: 1.5, w: 8.5, h: 3, fontSize: 20, fontFace: FONT_TCH, lineSpacing: 34 });

// --- Slide 3: STL Role Visual ---
let slide3 = pres.addSlide({ masterName: "STL_MASTER" });
addTitle(slide3, "新定位：策略人才領袖 (STL)");
if (fs.existsSync(IMAGES.stl)) {
    slide3.addImage({ path: IMAGES.stl, x: 5.5, y: 1.5, w: 4, h: 3 });
}
slide3.addText("HRBP 不再僅是內部諮詢者，而是轉型顧問，主導 AI 引發的組織變革。", { x: 0.7, y: 1.5, w: 4.5, h: 3, fontSize: 22, fontFace: FONT_TCH, color: THEME.primary, valign: "middle" });

// --- Slides 4-6: Responsibilities (Text with subtle icons simulated) ---
// (Keeping v3 richness)
const stl_details = [
    { t: "1. 引導人力重新設計", d: "判斷職務重塑與技能再培訓的最佳時機。" },
    { t: "2. 解決 AI 倫理與偏見", d: "確保 AI 建構的人才數據透明、客觀且公平。" },
    { t: "3. 塑造人機協作", d: "平衡技術產出與員工心理健康，創造雙贏文化。" }
];
stl_details.forEach(item => {
    let s = pres.addSlide({ masterName: "STL_MASTER" });
    addTitle(s, item.t);
    s.addShape(pres.shapes.RECTANGLE, { x: 0.7, y: 2.0, w: 0.15, h: 1, fill: { color: THEME.accent } });
    s.addText(item.d, { x: 1.0, y: 2.0, w: 8, h: 1, fontSize: 24, fontFace: FONT_TCH, color: THEME.text, valign: "middle" });
});

// --- Slide 7: Roadmap Visual ---
let slide7 = pres.addSlide({ masterName: "STL_MASTER" });
addTitle(slide7, "調整任務：轉型三階段路徑");
if (fs.existsSync(IMAGES.roadmap)) {
    slide7.addImage({ path: IMAGES.roadmap, x: 0.5, y: 2.5, w: 4.5, h: 2.5 });
}
const phases = ["1. 剔除舊務", "2. AI 強化", "3. 擴展策略"];
phases.forEach((p, i) => {
    slide7.addText("● " + p, { x: 5.5, y: 1.8 + (i * 0.8), w: 4, h: 0.5, fontSize: 20, fontFace: FONT_TCH, bold: true, color: THEME.primary });
});

// --- Phase Slides (Rich detail from v3) ---
const phase_data = [
    { t: "P1：剔除與 AI 重疊的行政工作", c: "定義「停止行動」清單，騰出 20-30% 產能。" },
    { t: "P2：透過 AI 強化核心責任", c: "使用 AI 洞察作為 BU 對話基礎，縮短決策週期。" },
    { t: "P3：擴展至由 AI 推動的新領域", c: "啟動 STL 小組，在變革決策中嵌入人員洞察。" }
];
phase_data.forEach(p => {
    let s = pres.addSlide({ masterName: "STL_MASTER" });
    addTitle(s, p.t);
    s.addShape(pres.shapes.RECTANGLE, { x: 0.7, y: 1.8, w: 8.5, h: 2.5, fill: { color: "F8FAFC" } });
    s.addText(p.c, { x: 1.0, y: 2.0, w: 8, h: 2, fontSize: 22, fontFace: FONT_TCH, color: THEME.text, align: "center", valign: "middle" });
});

// --- Slide 11: Success Metrics Visual ---
let slide11 = pres.addSlide({ masterName: "STL_MASTER" });
addTitle(slide11, "成功衡量：數據驅動成效");
if (fs.existsSync(IMAGES.success)) {
    slide11.addImage({ path: IMAGES.success, x: 5.5, y: 1.5, w: 4, h: 3 });
}
slide11.addText([
    { text: "📊 週期時間縮短\n", options: { bullet: true, breakLine: true } },
    { text: "📊 關鍵人才準備率提升\n", options: { bullet: true, breakLine: true } },
    { text: "📊 高風險離職率降低", options: { bullet: true } }
], { x: 0.7, y: 1.5, w: 4.5, h: 3, fontSize: 20, fontFace: FONT_TCH, color: THEME.accent, valign: "middle" });

// --- Slide 12: Mind Map (Optimized Connector Node) ---
let slideMap = pres.addSlide({ masterName: "STL_MASTER" });
addTitle(slideMap, "全課精華心智圖 (v4)");
const CX = 5.0, CY = 2.8;
slideMap.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: CX - 1, y: CY - 0.4, w: 2, h: 0.8, fill: { color: THEME.primary }, rectRadius: 0.1 });
slideMap.addText("HRBP\nAI 轉型", { x: CX - 1, y: CY - 0.4, w: 2, h: 0.8, color: THEME.white, bold: true, align: "center", fontFace: FONT_TCH, fontSize: 16 });

const nodes = [
    ["1.行政革新", 1.5, 1.2, THEME.secondary],
    ["2.STL 顧問", 7.0, 1.2, THEME.secondary],
    ["3.實作三階", 1.5, 4.4, THEME.accent],
    ["4.成效指標", 7.0, 4.4, THEME.accent]
];
nodes.forEach(node => {
    let nx = node[1], ny = node[2];
    slideMap.addShape(pres.shapes.RECTANGLE, { x: Math.min(CX, nx + 0.9), y: CY, w: Math.abs(CX - (nx + 0.9)), h: 0.02, fill: { color: THEME.line } });
    slideMap.addShape(pres.shapes.RECTANGLE, { x: nx + 0.9, y: Math.min(CY, ny + 0.3), w: 0.02, h: Math.abs(CY - (ny + 0.3)), fill: { color: THEME.line } });
    slideMap.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: nx, y: ny, w: 1.8, h: 0.6, fill: { color: node[3] }, rectRadius: 0.1 });
    slideMap.addText(node[0], { x: nx, y: ny, w: 1.8, h: 0.6, color: THEME.white, bold: true, align: "center", fontFace: FONT_TCH, fontSize: 12 });
});

// Final Slide
let slideLast = pres.addSlide();
slideLast.background = { color: THEME.primary };
slideLast.addText("啟動您的 AI 轉型旅程", { x: 0, y: 2.3, w: "100%", h: 0.6, fontSize: 36, bold: true, color: THEME.white, align: "center", fontFace: FONT_TCH });

// Output
const outDir = path.join("C:\\Users\\TW-Evan.Chen\\.gemini\\antigravity\\scratch\\Skills_Workspace", "Output");
if (!fs.existsSync(outDir)) fs.mkdirSync(outDir);
const outPath = path.join(outDir, "HRBP_AI_Transformation_Illustrated_v4.pptx");

pres.writeFile({ fileName: outPath }).then(fn => {
    console.log(`Success: Generated ${fn}`);
}).catch(err => console.error(err));
