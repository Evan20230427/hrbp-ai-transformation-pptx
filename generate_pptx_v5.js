const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

let pres = new pptxgen();
pres.layout = 'LAYOUT_16x9';
pres.author = 'Antigravity AI';
pres.title = 'AI 時代 HRBP 角色重塑最佳實務 - 旗艦版';

// Professional High-End Theme
const THEME = {
    primary: "0F172A",
    secondary: "2563EB",
    accent: "10B981",
    bg: "FFFFFF",
    text: "1E293B",
    subtle: "64748B",
    white: "FFFFFF",
    line: "E2E8F0"
};

const FONT_TCH = "Microsoft JhengHei";
const FONT_BODY = "Arial";

// Image Directory and Paths
const IMG_DIR = "C:\\Users\\TW-Evan.Chen\\.gemini\\antigravity\\brain\\cb95dffe-33bd-4e40-a98b-feaff376ea1a";
const IMAGES = {
    title: path.join(IMG_DIR, "hrbp_ai_title_visual_1773157399585.png"),
    stl: path.join(IMG_DIR, "hrbp_stl_consultant_1773157416001.png"),
    roadmap: path.join(IMG_DIR, "hrbp_transformation_roadmap_1773157436793.png"),
    success: path.join(IMG_DIR, "hrbp_success_metrics_data_1773157454088.png"),
    mindmap: path.join(IMG_DIR, "hrbp_transformation_mindmap_infographic_1773157697477.png")
};

// Define Standard Master Slide
pres.defineSlideMaster({
    title: 'FIXED_MASTER',
    background: { color: THEME.bg },
    objects: [
        { rect: { x: 0, y: 0, w: "100%", h: 0.1, fill: { color: THEME.secondary } } },
        { text: { text: "Gartner Research | HRBP AI Transformation Strategy", options: { x: 0.5, y: 5.3, w: 9, h: 0.25, fontSize: 10, color: THEME.subtle, align: "right", fontFace: FONT_BODY } } }
    ]
});

function addSlideTitle(slide, title) {
    slide.addText(title, { x: 0.5, y: 0.35, w: 9, h: 0.6, fontSize: 26, bold: true, color: THEME.primary, fontFace: FONT_TCH });
    slide.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 0.95, w: 9, h: 0.02, fill: { color: THEME.secondary } });
}

// --- Slide 1: Title ---
let slide1 = pres.addSlide();
slide1.background = { color: THEME.primary };
if (fs.existsSync(IMAGES.title)) {
    slide1.addImage({ path: IMAGES.title, x: 0, y: 0, w: "100%", h: "100%", sizing: { type: "cover" }, transparency: 30 });
}
slide1.addText("重塑 HRBP 角色：\n成為 AI 轉型顧問之最佳實務", { x: 0, y: 1.8, w: "100%", h: 1.2, fontSize: 44, bold: true, color: THEME.white, fontFace: FONT_TCH, align: "center" });
slide1.addText("2026 年度策略指南 | Gartner 全球實務匯總", { x: 0, y: 3.1, w: "100%", h: 0.5, fontSize: 18, color: THEME.secondary, fontFace: FONT_TCH, align: "center" });

// --- Slide 2: The Paradox ---
let slide2 = pres.addSlide({ masterName: "FIXED_MASTER" });
addSlideTitle(slide2, "HRBP 的生存危機：策略貢獻與事實不符");
slide2.addText([
    { text: "● 現況：目前僅有 51% 的商業領袖同意他們的 HRBP 參與了重要的策略討論。\n", options: { bold: true, color: THEME.secondary, breakLine: true } },
    { text: "● 痛點：HRBP 仍被困在 AI 已經開始吸收的事務性工作中。\n", options: { bullet: true, breakLine: true } },
    { text: "● 工具化威脅：擬定職務說明、摘要參與數據、回答政策 FAQ，這些任務已逐漸由 AI 主導。\n", options: { bullet: true, breakLine: true } },
    { text: "● 警訊：若不重新定義，即使是資深 HRBP 也面臨「未充分利用」的邊緣化風險。", options: { bullet: true } }
], { x: 0.8, y: 1.5, w: 8.5, h: 3, fontSize: 20, fontFace: FONT_TCH, lineSpacing: 34 });

// --- Slide 3-5: The STL Definition ---
let slide3 = pres.addSlide({ masterName: "FIXED_MASTER" });
addSlideTitle(slide3, "核心轉型：策略人才領袖 (STL)");
if (fs.existsSync(IMAGES.stl)) {
    slide3.addImage({ path: IMAGES.stl, x: 5.5, y: 1.5, w: 4, h: 3 });
}
slide3.addText("HRBP 必須從「解釋人員策略」進化為「主導轉型領袖」。", { x: 0.7, y: 1.3, fontSize: 16, color: THEME.subtle, fontFace: FONT_TCH });
slide3.addText("1. 引導人力設計\n2. 監測 AI 倫理偏見\n3. 優化人機協作效益", { x: 0.7, y: 2.0, w: 4.5, h: 2, fontSize: 24, fontFace: FONT_TCH, bold: true, color: THEME.primary, lineSpacing: 40 });

// --- Slide 6: Roadmap Visual ---
let slide6 = pres.addSlide({ masterName: "FIXED_MASTER" });
addSlideTitle(slide6, "轉型導圖：三階段調整任務");
if (fs.existsSync(IMAGES.roadmap)) {
    slide6.addImage({ path: IMAGES.roadmap, x: 0.5, y: 1.5, w: 4, h: 3 });
}
slide6.addText([
    { text: "階段 1: 剔除 - 移除行政舊務\n", options: { bold: true, breakLine: true } },
    { text: "階段 2: 強化 - AI 賦能核心責任\n", options: { bold: true, breakLine: true } },
    { text: "階段 3: 擴展 - 開拓新策略邊界", options: { bold: true } }
], { x: 5.0, y: 1.5, w: 4.5, h: 3, fontSize: 22, fontFace: FONT_TCH, color: THEME.primary, lineSpacing: 38, valign: "middle" });

// --- Slide 7: P1 Deep Dive ---
let slide7 = pres.addSlide({ masterName: "FIXED_MASTER" });
addSlideTitle(slide7, "第一階段：剔除 AI 可代勞的內容");
slide7.addTable([
    ["關鍵行動", "具體實施細節"],
    ["明確策略重點", "事先定義 HRBP 每日應有的策略時間配比。"],
    ["建立自動化路線圖", "根據任務準備度，設定 12-24 個月的自動化移交期。"],
    ["實施「停止行動」清單", "明確定義 AI vs HRBP 任務邊界，嚴禁重複投產。"]
], { x: 0.5, y: 1.5, w: 9, rowH: 0.8, fill: "F1F5F9", fontSize: 16, fontFace: FONT_TCH, border: { color: THEME.line } });

// --- Slide 8: P2 Deep Dive ---
let slide8 = pres.addSlide({ masterName: "FIXED_MASTER" });
addSlideTitle(slide8, "第二階段：透過 AI 強化核心責任");
slide8.addText([
    { text: "● 更新 HRBP 職能：確保「AI 運用能力」成為關鍵考核指標。\n", options: { bullet: true, breakLine: true } },
    { text: "● 數據賦能對話：以 AI 生成的洞察作為領導層溝通的「基礎基線」。\n", options: { bullet: true, breakLine: true } },
    { text: "● 協作流設計：定義 HRBP-CoE-AI 工作流，確保專業指導不缺席。\n", options: { bullet: true, breakLine: true } },
    { text: "● 目標：顯著縮短人才決策與繼任規劃的週期時間。", options: { bullet: true } }
], { x: 0.8, y: 1.5, w: 8.5, h: 3, fontSize: 20, fontFace: FONT_TCH, lineSpacing: 34 });

// --- Slide 9: P3 Deep Dive ---
let slide9 = pres.addSlide({ masterName: "FIXED_MASTER" });
addSlideTitle(slide9, "第三階段：擴展至全新策略領域");
slide9.addText([
    { text: "● 啟動 STL 小組：選拔具備 AI 洞察力的 HRBP 擔任試點先鋒。\n", options: { bullet: true, breakLine: true } },
    { text: "● 嵌入 STL 條款：在組織變更、收購、大規模 AI 部署中嵌入強大的人才條款。\n", options: { bullet: true, breakLine: true } },
    { text: "● 轉型指標：監控因 AI 重新設計而被「重新部署」而非「裁減」的角色比例。\n", options: { bullet: true, breakLine: true } },
    { text: "● 價值產出：降低 AI 標記的高風險角色流失率。", options: { bullet: true } }
], { x: 0.8, y: 1.5, w: 8.5, h: 3, fontSize: 20, fontFace: FONT_TCH, lineSpacing: 34 });

// --- Slide 10: Metrics Details ---
let slide10 = pres.addSlide({ masterName: "FIXED_MASTER" });
addSlideTitle(slide10, "衡量成功的量化指標");
if (fs.existsSync(IMAGES.success)) {
    slide10.addImage({ path: IMAGES.success, x: 5.5, y: 1.5, w: 4, h: 3 });
}
slide10.addText("1. 週期時間縮短 (Cycle Time Reduction)\n2. 繼任者準備率提升 (% Successor Ready)\n3. 遺憾離職率降低 (Regrettable Attrition)\n4. AI 偏見案例下降 (Bias Case Mitigation)", { x: 0.8, y: 1.5, w: 4.5, h: 3, fontSize: 18, fontFace: FONT_TCH, bold: true, color: THEME.accent, lineSpacing: 42 });

// --- Slide 11: Final Mind Map (Infographic Hybrid) ---
let slideMap = pres.addSlide({ masterName: "FIXED_MASTER" });
addSlideTitle(slideMap, "全課精華：AI 時代 HRBP 轉型心智圖");

if (fs.existsSync(IMAGES.mindmap)) {
    // Background Infographic
    slideMap.addImage({ path: IMAGES.mindmap, x: 1, y: 1.2, w: 8, h: 4 });

    // Precise Logical Overlays for Richness
    const labels = [
        { t: "現狀：51% 策略參與度\n面臨行政自動化衝擊", x: 0.5, y: 1.2, w: 2.5, c: THEME.primary },
        { t: "顧問定位：STL 角色\n主導人機協作與倫理", x: 7.0, y: 1.2, w: 2.5, c: THEME.primary },
        { t: "三階路徑：從剔除、\n強化到開拓新領域", x: 0.5, y: 4.5, w: 2.5, c: THEME.accent },
        { t: "指標：週期、準備率、\n離職率與偏見防治", x: 7.0, y: 4.5, w: 2.5, c: THEME.accent }
    ];

    labels.forEach(lb => {
        slideMap.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: lb.x, y: lb.y, w: lb.w, h: 0.8, fill: { color: THEME.bg }, line: { color: lb.c, width: 2 }, rectRadius: 0.1 });
        slideMap.addText(lb.t, { x: lb.x, y: lb.y, w: lb.w, h: 0.8, fontSize: 13, bold: true, color: lb.c, align: "center", fontFace: FONT_TCH });
    });
}

// Final Slide
let slideLast = pres.addSlide();
slideLast.background = { color: THEME.primary };
slideLast.addText("感謝您的聆聽\n啟動您的 HR 轉型之旅", { x: 0, y: 2.3, w: "100%", h: 1, fontSize: 36, bold: true, color: THEME.white, align: "center", fontFace: FONT_TCH });

// Save File
const outDir = path.join("C:\\Users\\TW-Evan.Chen\\.gemini\\antigravity\\scratch\\Skills_Workspace", "Output");
if (!fs.existsSync(outDir)) fs.mkdirSync(outDir);
const finalPath = path.join(outDir, "HRBP_AI_Transformation_Flagship_v5.pptx");

pres.writeFile({ fileName: finalPath }).then(fn => {
    console.log(`Success: Generated Flagship PPTX at ${fn}`);
}).catch(err => console.error(err));
