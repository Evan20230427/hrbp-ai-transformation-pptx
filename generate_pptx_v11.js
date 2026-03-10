const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

let pres = new pptxgen();
pres.layout = 'LAYOUT_16x9';
pres.author = 'Antigravity AI';
pres.title = 'HRBP AI 轉型最佳實務 - 終極大師版 v11';

/**
 * THEME: TECH INNOVATION (Refined for Digital Brutalism)
 * Principles: Bold Navy, Vibrant Blue, High Alignment, Heavy Contrast.
 */
const THEME = {
    primary: "0F172A",    // Deep Navy (Foundation)
    secondary: "3B82F6",  // Vibrant Blue (Digital Energy)
    accent: "10B981",     // Emerald (Success)
    bg: "FFFFFF",
    bg_dark: "0F172A",
    text: "1E293B",
    white: "FFFFFF",
    line: "E2E8F0"
};

const FONT_TITLE = "Microsoft JhengHei";
const FONT_BODY = "Arial";

// Resource Management
const IMG_DIR = "C:\\Users\\TW-Evan.Chen\\.gemini\\antigravity\\brain\\cb95dffe-33bd-4e40-a98b-feaff376ea1a";
const SCRATCH_DIR = "C:\\Users\\TW-Evan.Chen\\.gemini\\antigravity\\scratch\\pdf-xlsx-to-pptx";

const ASSETS = {
    gif: path.join(SCRATCH_DIR, "hrbp_office_dynamic_v10.gif"),
    human_ai: path.join(IMG_DIR, "hrbp_professional_human_ai_v10_1773159552183.png"),
    journey: path.join(IMG_DIR, "hrbp_transformation_journey_v10_1773159569827.png"),
    meeting: path.join(IMG_DIR, "hrbp_data_strategy_meeting_v10_1773159591017.png"),
    balance: path.join(IMG_DIR, "hrbp_ethical_ai_balance_v10_1773159607684.png"),
    dash: path.join(IMG_DIR, "hrbp_luxury_dashboard_v9_1773159099389.png")
};

/**
 * COMPONENT-BASED TEXT ENGINE (Frontend-Design Concept)
 * Uniform alignment, precise scaling, no overlap.
 */
function renderComponentText(slide, lines, opts = {}) {
    // Deduplication & Scaling Logic
    const uniqueLines = [...new Set(lines)];
    const baseSize = opts.fontSize || 18;
    const smallSize = Math.max(11, baseSize - 5);
    const content = [];

    uniqueLines.forEach(line => {
        const regex = /(\([^)]+\))/g;
        const tokens = line.split(regex);
        tokens.forEach((token, tIdx) => {
            const isEng = token.match(regex);
            content.push({
                text: token,
                options: {
                    fontSize: isEng ? smallSize : baseSize,
                    color: isEng ? THEME.secondary : THEME.text,
                    fontFace: isEng ? FONT_BODY : FONT_TITLE,
                    italic: isEng ? true : false,
                    bullet: (tIdx === 0),
                    breakLine: (tIdx === tokens.length - 1)
                }
            });
        });
    });

    slide.addText(content, {
        x: opts.x || 0.6,
        y: opts.y || 1.4,
        w: opts.w || 5.5,
        h: opts.h || 3.8,
        lineSpacing: opts.spacing || 24,
        valign: opts.valign || "top"
    });
}

// MASTER SLIDE: BRUTALIST GRID
pres.defineSlideMaster({
    title: 'BRUTALIST_V11',
    background: { color: THEME.bg },
    objects: [
        { rect: { x: 0, y: 0, w: 0.1, h: "100%", fill: { color: THEME.primary } } },
        { rect: { x: 0, y: 0, w: "100%", h: 0.05, fill: { color: THEME.secondary } } },
        { text: { text: "Gartner | STL AI Transformation Deliverable v11", options: { x: 0.5, y: 5.3, w: 9, h: 0.25, fontSize: 9, color: THEME.secondary, align: "right", fontFace: FONT_BODY } } }
    ]
});

function addHeaderWithLine(slide, title) {
    slide.addText(title, { x: 0.5, y: 0.4, w: 9, h: 0.5, fontSize: 28, bold: true, color: THEME.primary, fontFace: FONT_TITLE });
    slide.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.0, w: 3, h: 0.05, fill: { color: THEME.secondary } });
}

// --- SLIDES GENERATION ---

// 1. Title (Full Coverage Luxury)
let s1 = pres.addSlide();
s1.background = { color: THEME.bg_dark };
if (fs.existsSync(ASSETS.journey)) s1.addImage({ path: ASSETS.journey, x: 0, y: 0, w: "100%", h: "100%", sizing: { type: "cover" }, transparency: 50 });
s1.addText("重塑 HRBP 角色：\n引領企業 AI 轉型的策略實務", { x: 0, y: 1.6, w: "100%", h: 1.5, fontSize: 44, bold: true, color: THEME.white, fontFace: FONT_TITLE, align: "center" });
s1.addText("MASTER DESIGNER EDITION v11 | 2026 年度旗艦", { x: 0, y: 3.2, w: "100%", h: 0.5, fontSize: 16, color: THEME.secondary, fontFace: FONT_TITLE, align: "center", bold: true });

// 2. Paradox (Split Layout)
let s2 = pres.addSlide({ masterName: 'BRUTALIST_V11' });
addHeaderWithLine(s2, "現狀與挑戰：策略參與的鴻溝");
renderComponentText(s2, [
    "關鍵數據：僅 51% (Only 51%) 的領導層同意其 HRBP 參與了策略決策。",
    "行政束縛：大量工時被 AI 正在取代的「事務性舊務」佔據 (Transactional Work)。",
    "自動化威脅：涵蓋職務說明、數據摘要等高頻重複勞動 (Standardized Automation)。",
    "生存風險：角色定位被「工具化」而非「策略化」的職涯危機。"
], { x: 0.6, y: 1.5, w: 5.8, h: 3.5 });
if (fs.existsSync(ASSETS.dash)) s2.addImage({ path: ASSETS.dash, x: 6.6, y: 1.0, w: 3.2, h: 4.6, sizing: { type: "cover" } });

// 3. STL Definition (GIF Integration)
let s3 = pres.addSlide({ masterName: 'BRUTALIST_V11' });
addHeaderWithLine(s3, "未來定位：策略人才領袖 (STL)");
if (fs.existsSync(ASSETS.gif)) s3.addImage({ path: ASSETS.gif, x: 0.5, y: 1.3, w: 4.8, h: 3.0 });
renderComponentText(s3, [
    "目標：主導 AI 驅動轉型的人員面向。",
    "職責：引導人力設計、倫理監測與協作 (Expertise)。",
    "進化：從人力解釋者轉向「戰略對話推動者」。"
], { x: 5.5, y: 1.5, w: 4.0, h: 3, spacing: 30 });

// 4-6: Responsibilities (Full Visual Grid)
const ROLES = [
    { t: "職責 1：人力重新設計", c: ["主導職能重塑決策 (Workforce Redesign)", "優化人才再培訓與部署策略 (Reskill)"], img: ASSETS.human_ai },
    { t: "職責 2：應對倫理與偏見", c: ["監測 AI 決策中的算法偏見 (Bias Detection)", "建立組織內部的 AI 使用透明度規範"], img: ASSETS.balance },
    { t: "職責 3：優化人機協作效率", c: ["設計高效之「人類直覺 + AI」工作流", "在提升生產力時維護員工體驗 (Engagement)"], img: ASSETS.meeting }
];

ROLES.forEach(r => {
    let s = pres.addSlide({ masterName: 'BRUTALIST_V11' });
    addHeaderWithLine(s, r.t);
    renderComponentText(s, r.c, { x: 0.6, y: 1.6, w: 5.0, h: 2.5, spacing: 30 });
    if (fs.existsSync(r.img)) s.addImage({ path: r.img, x: 5.8, y: 1.0, w: 3.8, h: 4.5, sizing: { type: "cover" } });
});

// ... (Phase slides & Metrics leveraging the same elite v10 logic but updated branding)
const PH = [
    { t: "P1 剔除：移除行政負擔", c: ["定義策略優先權 (Strategy First)", "建立 1-2 年自動化藍圖 (Automation Roadmap)"] },
    { t: "P2 強化：AI 賦能高價值", c: ["更新模型，使 AI 準備度透明化 (Readiness)", "利用預測洞察大幅提升溝通深度"] },
    { t: "P3 開拓：主導新型策略領域", c: ["啟動 STL Pods 核心小組試行", "在重大決策中嵌入 STL 條款 (Governance)"] }
];

PH.forEach(p => {
    let s = pres.addSlide({ masterName: 'BRUTALIST_V11' });
    addHeaderWithLine(s, p.t);
    renderComponentText(s, p.c, { x: 0.6, y: 2.0, w: 8.5, h: 2.5, spacing: 35, fontSize: 20 });
});

// 15. The Mind Map (Brutalist Architecture)
let slideMap = pres.addSlide({ masterName: 'BRUTALIST_V11' });
addHeaderWithLine(slideMap, "全課精華：三層層級心智圖 (Master v11)");
const OX = 0.6, OY = 2.4;
const MAP = [
    { t: "挑戰觀測", c: ["策略不足(51%)", "行政作業負擔"] },
    { t: "STL 顧問力", c: ["人力重新設計", "倫理與人機協作"] },
    { t: "轉型三階", c: ["P1 剔除/P2 強化", "P3 全新領域擴展"] },
    { t: "成效衡量", c: ["週期產能回收", "繼任準備率提升"] }
];

slideMap.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: OX, y: OY, w: 1.5, h: 0.6, fill: { color: THEME.primary } });
slideMap.addText("HRBP AI 轉型", { x: OX, y: OY, w: 1.5, h: 0.6, color: THEME.white, bold: true, align: "center", fontFace: FONT_TITLE, fontSize: 11 });

MAP.forEach((n, i) => {
    let nx = OX + 2.0, ny = 1.0 + (i * 1.25);
    slideMap.addShape(pres.shapes.LINE, { x: OX + 1.5, y: OY + 0.3, w: 0.3, h: 0, line: { color: THEME.secondary, width: 2 } });
    slideMap.addShape(pres.shapes.LINE, { x: OX + 1.8, y: Math.min(OY + 0.3, ny + 0.25), w: 0, h: Math.abs(OY + 0.3 - (ny + 0.25)), line: { color: THEME.secondary, width: 2 } });
    slideMap.addShape(pres.shapes.LINE, { x: OX + 1.8, y: ny + 0.25, w: 0.2, h: 0, line: { color: THEME.secondary, width: 2 } });

    slideMap.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: nx, y: ny, w: 1.4, h: 0.5, fill: { color: THEME.secondary } });
    slideMap.addText(n.t, { x: nx, y: ny, w: 1.4, h: 0.5, color: THEME.white, bold: true, align: "center", fontFace: FONT_TITLE, fontSize: 10 });

    n.c.forEach((ch, j) => {
        let cx = nx + 1.6, cy = ny - 0.2 + (j * 0.45);
        slideMap.addShape(pres.shapes.LINE, { x: nx + 1.4, y: ny + 0.25, w: 0.1, h: 0, line: { color: THEME.line, width: 1 } });
        slideMap.addShape(pres.shapes.LINE, { x: nx + 1.5, y: Math.min(ny + 0.25, cy + 0.15), w: 0, h: Math.abs(ny + 0.25 - (cy + 0.15)), line: { color: THEME.line, width: 1 } });
        slideMap.addShape(pres.shapes.LINE, { x: nx + 1.5, y: cy + 0.15, w: 0.1, h: 0, line: { color: THEME.line, width: 1 } });
        slideMap.addText(ch, { x: cx, y: cy, w: 2.5, h: 0.3, color: THEME.text, fontSize: 9, fontFace: FONT_TITLE, valign: "middle" });
    });
});

// Final Slide
let sL = pres.addSlide();
sL.background = { color: THEME.primary };
sL.addText("啟動您的精英與 AI 共生之路", { x: 0, y: 2.3, w: "100%", h: 0.6, bold: true, fontSize: 36, color: THEME.white, align: "center", fontFace: FONT_TITLE });

const outP = path.join(SCRATCH_DIR, "output", "HRBP_AI_Transformation_Master_v11.pptx");
pres.writeFile({ fileName: outP }).then(fn => {
    console.log(`Success: Generated Master v11 at ${fn}`);
}).catch(err => console.error(err));
