const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

let pres = new pptxgen();
pres.layout = 'LAYOUT_16x9';
pres.author = 'Antigravity AI';
pres.title = 'AI 時代重新設計 HRBP 角色的最佳實務';

// Solid Professional Theme
const THEME = {
    primary: "0F172A",    // Deep Navy
    secondary: "2563EB",  // Vibrant Blue
    accent: "059669",     // Emerald
    bg: "FFFFFF",         // Standard Pure White for maximum compatibility
    text: "1E293B",       // Slate 800
    subtle: "64748B",     // Slate 500
    white: "FFFFFF",
    line: "CBD5E1"
};

const FONT_BODY = "Arial";
const FONT_TCH = "Microsoft JhengHei";

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
    slide.addText(text, { x: 0.5, y: 0.4, w: 9, h: 0.6, fontSize: 28, bold: true, color: THEME.primary, fontFace: FONT_TCH });
    slide.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.0, w: 9, h: 0.02, fill: { color: THEME.secondary } });
}

// Slide 1: Title
let slide1 = pres.addSlide();
slide1.background = { color: THEME.primary };
slide1.addText("重塑 HRBP 角色：\n引領企業 AI 轉型的策略實務", { x: 0, y: 1.8, w: "100%", h: 1.2, fontSize: 44, bold: true, color: THEME.white, fontFace: FONT_TCH, align: "center" });
slide1.addText("From Overstretched Administrator to AI Transformation Consultant", { x: 0, y: 3.1, w: "100%", h: 0.5, fontSize: 18, color: THEME.secondary, fontFace: FONT_BODY, align: "center" });

// Slide 2: Paradox
let slide2 = pres.addSlide({ masterName: "STL_MASTER" });
addTitle(slide2, "現狀挑戰：HRBP 的策略鴻溝");
slide2.addText([
    { text: "關鍵數據：只有 51% 的領導者認為其 HRBP 參與了重要的策略討論。\n", options: { bold: true, color: THEME.secondary, breakLine: true } },
    { text: "困境：HRBP 仍被鎖定在「事務性工作」中。\n", options: { bullet: true, breakLine: true } },
    { text: "AI 風險：AI 正在吸收 HRBP 的傳統任務（如職務說明、數據摘要、問答）。\n", options: { bullet: true, breakLine: true } },
    { text: "核心問題：如果 HRBP 的角色不變，他們將面臨日益嚴重的「未充分利用」風險。", options: { bullet: true } }
], { x: 0.7, y: 1.5, w: 8.5, h: 3, fontSize: 20, fontFace: FONT_TCH, lineSpacing: 34 });

// Slide 3: The New Identity
let slide3 = pres.addSlide({ masterName: "STL_MASTER" });
addTitle(slide3, "新身份：AI 轉型顧問 (STL)");
slide3.addText("未來的 HRBP 應該被定義為「策略人才領袖 (Strategic Talent Leaders)」。", { x: 0.7, y: 1.3, fontSize: 16, color: THEME.subtle, fontFace: FONT_TCH });
slide3.addShape(pres.shapes.RECTANGLE, { x: 0.7, y: 1.8, w: 8.5, h: 2.5, fill: { color: "F1F5F9" }, rectRadius: 0.1 });
slide3.addText("隨著 AI 重塑價值鏈，HRBP 的職權擴展為直接引導人工智慧驅動轉型的人員面向。", { x: 1.0, y: 2.1, w: 8, h: 1.5, fontSize: 22, color: THEME.primary, fontFace: FONT_TCH, align: "center", valign: "middle" });

// Slides 4-6: STL Responsibilities
const stl_details = [
    { t: "1. 引導人力重新設計 (Workforce Redesign)", d: "隨著 AI 改變職務，決定何時進行技能再培訓、重新部署或逐步淘汰職位。" },
    { t: "2. 解決 AI 倫理與偏見 (AI Bias & Ethics)", d: "解決 AI 驅動的人才決策中的偏見與倫理挑戰，確保決策透明度。" },
    { t: "3. 塑造人機協作 (Human-Machine Collab)", d: "確保生產力提升不會以犧牲員工參與度為代價，優化人類與機器的動態協作。" }
];

stl_details.forEach(item => {
    let s = pres.addSlide({ masterName: "STL_MASTER" });
    addTitle(s, "STL 核心職責：" + item.t);
    s.addShape(pres.shapes.RECTANGLE, { x: 0.7, y: 2.0, w: 0.15, h: 1, fill: { color: THEME.accent } });
    s.addText(item.d, { x: 1.0, y: 2.0, w: 8, h: 1, fontSize: 24, fontFace: FONT_TCH, color: THEME.text, valign: "middle" });
});

// Slide 7: Roadmap Overview
let slide7 = pres.addSlide({ masterName: "STL_MASTER" });
addTitle(slide7, "調整任務：轉型三階段路徑");
const phases = ["階段 1: 剔除舊有工作", "階段 2: 透過 AI 強化核心", "階段 3: 擴展新策略領域"];
phases.forEach((p, i) => {
    let x = 0.5 + (i * 3.1);
    slide7.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: x, y: 2.0, w: 2.8, h: 1.5, fill: { color: THEME.primary }, rectRadius: 0.1 });
    slide7.addText(p, { x: x, y: 2.0, w: 2.8, h: 1.5, color: THEME.white, bold: true, align: "center", fontSize: 20, fontFace: FONT_TCH });
});

// Phase Details
const phase_data = [
    {
        title: "第一階段：剔除 (Strip Out Legacy)",
        actions: ["透過釐清優先順序，定義策略重點。", "根據自動化準備度建立路線圖。", "實施「停止行動」清單以界定 AI 與人力的邊界。"],
        metrics: ["從特定任務（如入職清單、報告）中回收的工時。", "重新分配至策略優先事項（如繼任計畫）的時間百分比。"]
    },
    {
        title: "第二階段：強化 (Augment Core)",
        actions: ["更新職能模型，確保「AI 準備度」可見。", "利用 AI 生成的洞察作為領導對話的基準。", "定義 HRBP-CoE-AI 工作流程以避免重疊。"],
        metrics: ["人力/繼任決策的週期時間大幅縮短。", "18 個月內準備好擔任關鍵職位的繼任者百分比增加。"]
    },
    {
        title: "第三階段：擴展 (Expand New)",
        actions: ["啟動與試行「策略人才領袖 (STL)」小組。", "在轉型決策中嵌入 STL 條款，確保人力洞察具法律約束力。"],
        metrics: ["因 AI 重新設計而被重新部署（而非裁員）的角色比例。", "AI 識別出的高風險職位流失率降低。"]
    }
];

phase_data.forEach(p => {
    // Actions Slide
    let s_act = pres.addSlide({ masterName: "STL_MASTER" });
    addTitle(s_act, p.title + " - 主要行動");
    s_act.addText(p.actions.map(a => "● " + a).join("\n"), { x: 0.8, y: 1.5, w: 8.5, h: 3, fontSize: 20, fontFace: FONT_TCH, lineSpacing: 34 });

    // Metrics Slide
    let s_met = pres.addSlide({ masterName: "STL_MASTER" });
    addTitle(s_met, p.title + " - 成功衡量標準");
    s_met.addText(p.metrics.map(m => "📊 " + m).join("\n"), { x: 0.8, y: 1.5, w: 8.5, h: 2, fontSize: 20, fontFace: FONT_TCH, color: THEME.accent, lineSpacing: 34 });
});

// Slide 14: Gartner's 4 Pillars
let slide14 = pres.addSlide({ masterName: "STL_MASTER" });
addTitle(slide14, "Gartner 建議：HR 專門手冊");
const pillars = [
    "1. 增強 HR 專業知識：保持對 AI 趨勢的最新洞察。",
    "2. 識別開發機會：基於勝任力模型與診斷工具。",
    "3. 加速項目執行：使用最佳實務指南與手冊。",
    "4. 數據驅動決策：利用 DataHub 數據進行同業標竿對比。"
];
slide14.addText(pillars.join("\n"), { x: 0.8, y: 1.5, h: 3, fontSize: 20, fontFace: FONT_TCH, lineSpacing: 32 });

// Slide 15: Mind Map (Hyper-Robust)
let slideMap = pres.addSlide({ masterName: "STL_MASTER" });
addTitle(slideMap, "全課精華：HRBP 轉型心智圖");

const CX = 5.0, CY = 2.8; // Center point
// Draw Center Node
slideMap.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: CX - 1, y: CY - 0.4, w: 2.0, h: 0.8, fill: { color: THEME.primary }, rectRadius: 0.1 });
slideMap.addText("HRBP\nAI 轉型", { x: CX - 1, y: CY - 0.4, w: 2.0, h: 0.8, color: THEME.white, bold: true, align: "center", fontFace: FONT_TCH, fontSize: 16 });

// Branches: [text, x, y, color]
const nodes = [
    ["1.現狀與挑戰", 1.5, 1.2, THEME.secondary],
    ["2.STL新角色", 7.0, 1.2, THEME.secondary],
    ["3.三階段實作", 1.5, 4.4, THEME.accent],
    ["4.成功指標", 7.0, 4.4, THEME.accent]
];

nodes.forEach(node => {
    let nx = node[1], ny = node[2];
    // Connection lines using Rects for stability
    let lineX = Math.min(CX, nx + 0.75);
    let lineW = Math.abs(CX - (nx + 0.75));
    // Horizontal part
    slideMap.addShape(pres.shapes.RECTANGLE, { x: lineX, y: CY, w: lineW, h: 0.02, fill: { color: THEME.line } });
    // Vertical part
    slideMap.addShape(pres.shapes.RECTANGLE, { x: nx + 0.75, y: Math.min(CY, ny + 0.25), w: 0.02, h: Math.abs(CY - (ny + 0.25)), fill: { color: THEME.line } });

    // Node
    slideMap.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: nx, y: ny, w: 1.8, h: 0.6, fill: { color: node[3] }, rectRadius: 0.1 });
    slideMap.addText(node[0], { x: nx, y: ny, w: 1.8, h: 0.6, color: THEME.white, bold: true, align: "center", fontFace: FONT_TCH, fontSize: 12 });
});

// Final Slide
let slideLast = pres.addSlide();
slideLast.background = { color: THEME.primary };
slideLast.addText("Ready for the Future?", { x: 0, y: 2.3, w: "100%", h: 0.6, fontSize: 36, bold: true, color: THEME.white, align: "center", fontFace: FONT_BODY });
slideLast.addText("HRBP 轉型實務 - 成果交付", { x: 0, y: 3.0, w: "100%", h: 0.4, fontSize: 18, color: THEME.secondary, align: "center", fontFace: FONT_TCH });

// Output
const outDir = path.join("C:\\Users\\TW-Evan.Chen\\.gemini\\antigravity\\scratch\\Skills_Workspace", "Output");
if (!fs.existsSync(outDir)) fs.mkdirSync(outDir);
const outPath = path.join(outDir, "HRBP_AI_Transformation_Full_v3.pptx");

pres.writeFile({ fileName: outPath }).then(fn => {
    console.log(`Successfully created: ${fn}`);
}).catch(err => console.error(err));
