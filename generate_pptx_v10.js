const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

let pres = new pptxgen();
pres.layout = 'LAYOUT_16x9';
pres.author = 'Antigravity AI';
pres.title = 'AI 時代重新設計 HRBP 角色的最佳實務 - 精英版 v10';

const THEME = {
    primary: "0F172A",
    secondary: "3B82F6",
    accent: "10B981",
    bg: "FFFFFF",
    text: "1E293B",
    subtle: "64748B",
    white: "FFFFFF"
};

const FONT_TCH = "Microsoft JhengHei";
const FONT_BODY = "Arial";

// Resource Paths
const IMG_DIR = "C:\\Users\\TW-Evan.Chen\\.gemini\\antigravity\\brain\\cb95dffe-33bd-4e40-a98b-feaff376ea1a";
const SCRATCH_DIR = "C:\\Users\\TW-Evan.Chen\\.gemini\\antigravity\\scratch\\pdf-xlsx-to-pptx";

const IMAGES = {
    gif: path.join(SCRATCH_DIR, "hrbp_office_dynamic_v10.gif"),
    human_ai: path.join(IMG_DIR, "hrbp_professional_human_ai_v10_1773159552183.png"),
    journey: path.join(IMG_DIR, "hrbp_transformation_journey_v10_1773159569827.png"),
    meeting: path.join(IMG_DIR, "hrbp_data_strategy_meeting_v10_1773159591017.png"),
    balance: path.join(IMG_DIR, "hrbp_ethical_ai_balance_v10_1773159607684.png"),
    dash: path.join(IMG_DIR, "hrbp_luxury_dashboard_v9_1773159099389.png")
};

/**
 * Enhanced Text Logic: Uniform, Clean, Scaled English.
 */
function createEliteText(lines, baseSize = 17) {
    const finalContent = [];
    const smallSize = Math.max(12, baseSize - 3);
    [...new Set(lines)].forEach(line => {
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

// MASTER: DENSE PROFESSIONAL
pres.defineSlideMaster({
    title: 'ELITE_MASTER',
    background: { color: THEME.bg },
    objects: [
        { rect: { x: 0, y: 0, w: 0.15, h: "100%", fill: { color: THEME.primary } } },
        { text: { text: "Gartner Research | HRBP AI Transformation - Elite Edition v10", options: { x: 0.5, y: 5.3, w: 9, h: 0.25, fontSize: 10, color: THEME.subtle, align: "right", fontFace: FONT_BODY } } }
    ]
});

function addTitle(slide, title) {
    slide.addText(title, { x: 0.5, y: 0.4, w: 9, h: 0.6, fontSize: 26, bold: true, color: THEME.primary, fontFace: FONT_TCH });
    slide.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.0, w: 4, h: 0.03, fill: { color: THEME.secondary } });
}

// --- SLIDES: MINIMIZING WHITE SPACE ---

// 1. Title
let s1 = pres.addSlide();
s1.background = { color: THEME.primary };
if (fs.existsSync(IMAGES.journey)) s1.addImage({ path: IMAGES.journey, x: 0, y: 0, w: "100%", h: "100%", sizing: { type: "cover" }, transparency: 40 });
s1.addText("重塑 HRBP 角色：\n引領 AI 轉型的策略實務", { x: 0, y: 1.8, w: "100%", h: 1.2, fontSize: 42, bold: true, color: THEME.white, fontFace: FONT_TCH, align: "center" });
s1.addText("2026 年度精英旗艦版 | 完整深度實作大綱", { x: 0, y: 3.1, w: "100%", h: 0.5, fontSize: 16, color: THEME.secondary, fontFace: FONT_TCH, align: "center" });

// 2. Paradox (Fill with Dash)
let s2 = pres.addSlide({ masterName: 'ELITE_MASTER' });
addTitle(s2, "核心挑戰：策略貢獻與事務束縛");
s2.addText(createEliteText([
    "數據揭露：僅 51% (Only 51%) 的領袖同意其 HRBP 參與了重大策略討論。",
    "轉型瓶頸：仍被鎖定在 AI 與自動化正在吸收的任務中 (Transactional Burden)。",
    "核心風險：領袖對 HRBP 策略價值的認知鴻溝 (Value Perception Gap)。"
]), { x: 0.6, y: 1.5, w: 5.8, h: 3.5, lineSpacing: 25 });
if (fs.existsSync(IMAGES.dash)) s2.addImage({ path: IMAGES.dash, x: 6.8, y: 1.2, w: 2.8, h: 3.8, sizing: { type: "cover" } });

// 3. STL Definition (Realistic GIF)
let s3 = pres.addSlide({ masterName: 'ELITE_MASTER' });
addTitle(s3, "未來定位：策略人才領袖 (STL)");
if (fs.existsSync(IMAGES.gif)) s3.addImage({ path: IMAGES.gif, x: 0.6, y: 1.4, w: 5, h: 2.8 });
s3.addText(createEliteText([
    "定位：引導人工智慧轉型中的人員設計。",
    "權責：從解释者進化為「主導轉型對話者」。"
], 20), { x: 5.8, y: 1.4, w: 3.8, h: 2.8, lineSpacing: 30, valign: "middle" });

// 4. Responsibility 1 (Human-AI Image)
let s4 = pres.addSlide({ masterName: 'ELITE_MASTER' });
addTitle(s4, "職責 1：人力重新設計 (Workforce Redesign)");
s4.addText(createEliteText([
    "主導隨 AI 改變的職能重塑決策。",
    "優化人才再培訓與部署 (Reskilling/Redeployment)。",
    "確保組織架構與技術能量對齊。"
]), { x: 0.6, y: 1.5, w: 5.0, h: 3, lineSpacing: 28 });
if (fs.existsSync(IMAGES.human_ai)) s4.addImage({ path: IMAGES.human_ai, x: 5.8, y: 1.2, w: 3.8, h: 3.8, sizing: { type: "cover" } });

// 5. Responsibility 2 (Balance Image)
let s5 = pres.addSlide({ masterName: 'ELITE_MASTER' });
addTitle(s5, "職責 2：應對 AI 倫理與偏見 (Ethics)");
s5.addText(createEliteText([
    "監測人才決策中的算法偏見 (Addressing Bias)。",
    "確保數據推薦透明度與公平性。",
    "維護企業文化在技術浪潮中的倫理基石。"
]), { x: 0.6, y: 1.5, w: 5.0, h: 3, lineSpacing: 28 });
if (fs.existsSync(IMAGES.balance)) s5.addImage({ path: IMAGES.balance, x: 5.8, y: 1.2, w: 3.8, h: 3.8, sizing: { type: "cover" } });

// 6. Responsibility 3 (Meeting Image)
let s6 = pres.addSlide({ masterName: 'ELITE_MASTER' });
addTitle(s6, "職責 3：優化人機協作效率");
s6.addText(createEliteText([
    "提升生產力的同時維護員工參與度。",
    "平衡人類直覺與 AI 計算之互補關係。",
    "重新流設計工作流 (Workflow Redesign)。"
]), { x: 0.6, y: 1.5, w: 5.0, h: 3, lineSpacing: 28 });
if (fs.existsSync(IMAGES.meeting)) s6.addImage({ path: IMAGES.meeting, x: 5.8, y: 1.2, w: 3.8, h: 3.8, sizing: { type: "cover" } });

// ... (Phase slides & Metrics leveraging elite layout)
const PHASES = [
    { t: "P1 剔除行動：回收策略產能", c: ["定義策略優先權 (Define Priorities)", "建立 1-2 年自動化藍圖 (Roadmap)"] },
    { t: "P2 強化行動：AI 賦能高價值", c: ["更新職能模型，使 AI 準備度透明化", "利用預測洞察強化領導溝通層次"] },
    { t: "P3 擴展行動：開拓新型策略", c: ["啟動 STL Pods 小組引領變革", "在收購與重組決策中嵌入 STL 條款"] }
];
PHASES.forEach(ph => {
    let s = pres.addSlide({ masterName: 'ELITE_MASTER' });
    addTitle(s, ph.t);
    s.addText(createEliteText(ph.c, 21), { x: 1.0, y: 2.2, w: 8, h: 1.5, lineSpacing: 35 });
});

// Final Map (Elite Visualization)
let slideMap = pres.addSlide({ masterName: 'ELITE_MASTER' });
addTitle(slideMap, "全課精華：三層層級心智圖 (Elite)");
const RX = 0.6, RY = 2.4;
const L1 = [
    { t: "現狀挑戰", c: ["策略不足(51%)", "行政負擔"] },
    { t: "STL 定義", c: ["人力設計", "倫理/協作"] },
    { t: "轉型三階", c: ["P1 剔除/P2 強化", "P3 開拓域"] },
    { t: "成效量測", c: ["週期/產能", "準備/流失"] }
];
slideMap.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: RX, y: RY, w: 1.4, h: 0.6, fill: { color: THEME.primary }, rectRadius: 0.1 });
slideMap.addText("HRBP AI 轉型", { x: RX, y: RY, w: 1.4, h: 0.6, color: THEME.white, bold: true, align: "center", fontFace: FONT_TCH, fontSize: 11 });
L1.forEach((node, i) => {
    let nx = RX + 2.1, ny = 1.0 + (i * 1.25);
    slideMap.addShape(pres.shapes.LINE, { x: RX + 1.4, y: RY + 0.3, w: 0.4, h: 0, line: { color: THEME.secondary, width: 2 } });
    slideMap.addShape(pres.shapes.LINE, { x: RX + 1.8, y: Math.min(RY + 0.3, ny + 0.25), w: 0, h: Math.abs(RY + 0.3 - (ny + 0.25)), line: { color: THEME.secondary, width: 2 } });
    slideMap.addShape(pres.shapes.LINE, { x: RX + 1.8, y: ny + 0.25, w: 0.3, h: 0, line: { color: THEME.secondary, width: 2 } });
    slideMap.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: nx, y: ny, w: 1.4, h: 0.5, fill: { color: THEME.secondary }, rectRadius: 0.1 });
    slideMap.addText(node.t, { x: nx, y: ny, w: 1.4, h: 0.5, color: THEME.white, bold: true, align: "center", fontFace: FONT_TCH, fontSize: 10 });
    node.c.forEach((child, j) => {
        let cx = nx + 1.8, cy = ny - 0.2 + (j * 0.45);
        slideMap.addShape(pres.shapes.LINE, { x: nx + 1.4, y: ny + 0.25, w: 0.2, h: 0, line: { color: THEME.subtle, width: 1 } });
        slideMap.addShape(pres.shapes.LINE, { x: nx + 1.6, y: Math.min(ny + 0.25, cy + 0.15), w: 0, h: Math.abs(ny + 0.25 - (cy + 0.15)), line: { color: THEME.subtle, width: 1 } });
        slideMap.addShape(pres.shapes.LINE, { x: nx + 1.6, y: cy + 0.15, w: 0.2, h: 0, line: { color: THEME.subtle, width: 1 } });
        slideMap.addText(child, { x: cx, y: cy, w: 2.2, h: 0.3, color: THEME.text, fontSize: 9, fontFace: FONT_TCH, valign: "middle" });
    });
});

// Final
let slideLast = pres.addSlide();
slideLast.background = { color: THEME.primary };
slideLast.addText("成功引領組織，邁向 AI 轉型巔峰", { x: 0, y: 2.3, w: "100%", h: 0.6, bold: true, fontSize: 36, color: THEME.white, align: "center", fontFace: FONT_TCH });

const outPath = path.join(SCRATCH_DIR, "output", "HRBP_AI_Transformation_Elite_v10.pptx");
pres.writeFile({ fileName: outPath }).then(fn => {
    console.log(`Successfully generated Elite v10 at ${fn}`);
}).catch(err => console.error(err));
