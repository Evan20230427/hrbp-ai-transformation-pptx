const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

let pres = new pptxgen();
pres.layout = 'LAYOUT_16x9';
pres.author = 'Antigravity AI';
pres.title = 'HRBP AI 轉型最佳實務 - 終極穩定修復版 v12';

/**
 * THEME: COLOR HARMONY & ZEN CLARITY
 * Principles: Avoid complementary colors (no blue/orange, no red/green overlap).
 * Palette: Deep Navy, Azure Blue, Silver, White.
 */
const THEME = {
    primary: "0F172A",    // Deep Navy
    secondary: "3B82F6",  // Azure Blue
    accent: "64748B",     // Cool Slate (Neutral contrast)
    bg: "FFFFFF",
    bg_alt: "F8FAFC",
    text: "1E293B",
    white: "FFFFFF",
    line: "CBD5E1"
};

const FONT_TITLE = "Microsoft JhengHei";
const FONT_BODY = "Arial";

// Resource Management
const IMG_DIR = "C:\\Users\\TW-Evan.Chen\\.gemini\\antigravity\\brain\\cb95dffe-33bd-4e40-a98b-feaff376ea1a";
const SCRATCH_DIR = "C:\\Users\\TW-Evan.Chen\\.gemini\\antigravity\\scratch\\pdf-xlsx-to-pptx";

const ASSETS = {
    cover: path.join(IMG_DIR, "hrbp_simple_minimalist_cover_v12_1773160264391.png"),
    human: path.join(IMG_DIR, "hrbp_professional_human_v12_simple_1773160281938.png"),
    dash: path.join(IMG_DIR, "hrbp_luxury_dashboard_v9_1773159099389.png"),
    journey: path.join(IMG_DIR, "hrbp_transformation_journey_v10_1773159569827.png"),
    meeting: path.join(IMG_DIR, "hrbp_data_strategy_meeting_v10_1773159591017.png")
};

/**
 * Safe Text Engine: Avoids overlap and ensures color harmony.
 */
function renderSafeText(slide, lines, opts = {}) {
    const uniqueLines = [...new Set(lines)];
    const baseSize = opts.fontSize || 18;
    const smallSize = Math.max(12, baseSize - 4);
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
        w: opts.w || 6.5,
        h: opts.h || 3.5,
        lineSpacing: opts.spacing || 26,
        valign: "top"
    });
}

// MASTER SLIDE: ZEN CLARITY
pres.defineSlideMaster({
    title: 'ZEN_MASTER_V12',
    background: { color: THEME.bg },
    objects: [
        { rect: { x: 0, y: 0, w: 0.1, h: "100%", fill: { color: THEME.secondary } } },
        { text: { text: "© 2026 Gartner Insight | HRBP AI Transformation Zenith Guide", options: { x: 0.5, y: 5.3, w: 9, h: 0.25, fontSize: 10, color: THEME.accent, align: "right", fontFace: FONT_BODY } } }
    ]
});

function addHeader(slide, title) {
    slide.addText(title, { x: 0.5, y: 0.4, w: 9, h: 0.6, fontSize: 28, bold: true, color: THEME.primary, fontFace: FONT_TITLE });
    slide.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.0, w: 2.5, h: 0.04, fill: { color: THEME.secondary } });
}

// --- SLIDES ---

// 1. Cover (Minimalist)
let s1 = pres.addSlide();
s1.background = { color: THEME.white };
if (fs.existsSync(ASSETS.cover)) s1.addImage({ path: ASSETS.cover, x: 5, y: 0.5, w: 4.5, h: 4.5 });
s1.addText("重塑 HRBP 角色：\n引領企業 AI 轉型的策略實務", { x: 0.5, y: 1.8, w: 5.5, h: 1.5, fontSize: 36, bold: true, color: THEME.primary, fontFace: FONT_TITLE });
s1.addText("2026 年度終極穩定版 | 專業深度內容 (PRISTINE EDITION)", { x: 0.5, y: 3.3, w: 5.5, h: 0.5, fontSize: 15, color: THEME.secondary, fontFace: FONT_TITLE });

// 2. Paradox
let s2 = pres.addSlide({ masterName: 'ZEN_MASTER_V12' });
addHeader(s2, "核心挑戰：策略參與的空洞化");
renderSafeText(s2, [
    "數據實例：僅 51% (Only 51%) 的領導者認同 HRBP 具備足夠策略影響力。",
    "轉型束縛：行政瑣事佔據過多產能，且 AI 正加速吸收此类任務 (Automation Absorption)。",
    "生存威脅：必須在 AI 極致工具化的浪潮中，重新定義「高階價值感」。"
]);
if (fs.existsSync(ASSETS.dash)) s2.addImage({ path: ASSETS.dash, x: 7.2, y: 1.2, w: 2.4, h: 3.8, sizing: { type: "contain" } });

// 3. STL Role
let s3 = pres.addSlide({ masterName: 'ZEN_MASTER_V12' });
addHeader(s3, "未來視角：策略人才領袖 (STL)");
renderSafeText(s3, [
    "重新定位：從解釋人員策略進化為「主導轉型對話」 (Transformation Leader)。",
    "核心價值：引導人力設計、監測 AI 倫理與人機協作效率提升能力。",
    "行動導向：直接嵌入業務變革的核心策略決策圈。"
], { fontSize: 20, spacing: 32 });
if (fs.existsSync(ASSETS.human)) s3.addImage({ path: ASSETS.human, x: 7.2, y: 1.2, w: 2.4, h: 3.8, sizing: { type: "cover" } });

// 4-6: Responsibilities
const ROLES = [
    { t: "職責 1：人力重新設計", c: ["主導 AI 時代下的職能重塑與定義 (Redesign)", "決策人才培訓 (Reskill) 與資源分配優先級"], img: ASSETS.journey },
    { t: "職責 2：應對倫理與偏見", c: ["監測人才決策中的算法偏見 (Address Bias)", "確保 AI 洞察之透明度與企業倫理對齊"], img: null },
    { t: "職責 3：優化人機協作效率", c: ["設計具備生產力且維持員工參與度的工作流", "在技術導入過程中平衡人類直覺與 AI 計算"], img: ASSETS.meeting }
];
ROLES.forEach(r => {
    let s = pres.addSlide({ masterName: 'ZEN_MASTER_V12' });
    addHeader(s, r.t);
    renderSafeText(s, r.c);
    if (r.img && fs.existsSync(r.img)) s.addImage({ path: r.img, x: 7.0, y: 1.5, w: 2.5, h: 3 });
});

// ... Phase & Metric slides ... (Kept dense but clean)
const PH = [
    { t: "P1 剔除行動：回收產能", c: ["精確定義策略重點區", "建立 12-24 月自動化路線圖 (Automation Roadmap)"] },
    { t: "P2 強化行動：AI 賦能", c: ["更新模型使 AI 準備度透明化", "利用預測大數據強化領導溝通層次 (Data-driven)"] },
    { t: "P3 擴展開拓：新策略域", c: ["試行 STL Pods 小組領航計畫", "在重大收購與變革中嵌入 STL 治理條款"] }
];
PH.forEach(p => {
    let s = pres.addSlide({ masterName: 'ZEN_MASTER_V12' });
    addHeader(s, p.t);
    renderSafeText(s, p.c, { fontSize: 21, spacing: 38 });
});

// --- Slide 15: FINAL HIERARCHICAL MIND MAP (V12 PRISTINE) ---
// User requirement: 3rd level font size >= 16.
let sMap = pres.addSlide({ masterName: 'ZEN_MASTER_V12' });
addHeader(sMap, "全課精華：層級深度心智圖 (V12 Elite)");

const ROOT_X = 0.5, ROOT_Y = 2.4;
const DATA = [
    { t: "現狀挑戰", c: ["策略參與不足(51%)", "行政作業佔據產能", "AI自動化之威脅"] },
    { t: "STL 定義", c: ["人力重新設計優化", "倫理監測與治理力", "人機協作產出效益"] },
    { t: "轉型三階", c: ["P1 剔除舊有事務", "P2 AI 強化核心", "P3 開拓新策略域"] },
    { t: "成功指標", c: ["週期產能回收率", "繼任準備率提升", "遺憾離職率優化"] }
];

// Root Node
sMap.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: ROOT_X, y: ROOT_Y, w: 1.3, h: 0.6, fill: { color: THEME.primary } });
sMap.addText("HRBP\nAI 轉型", { x: ROOT_X, y: ROOT_Y, w: 1.3, h: 0.6, color: THEME.white, bold: true, align: "center", fontFace: FONT_TITLE, fontSize: 13 });

// Vertical Logic to prevent overlap with large fonts
DATA.forEach((node, i) => {
    const L1_X = ROOT_X + 1.8;
    const L1_Y = 0.8 + (i * 1.35); // Ample spacing

    // Connect Root -> L1
    sMap.addShape(pres.shapes.LINE, { x: ROOT_X + 1.3, y: ROOT_Y + 0.3, w: 0.25, h: 0, line: { color: THEME.secondary, width: 2 } });
    sMap.addShape(pres.shapes.LINE, { x: ROOT_X + 1.55, y: Math.min(ROOT_Y + 0.3, L1_Y + 0.25), w: 0, h: Math.abs(ROOT_Y + 0.3 - (L1_Y + 0.25)), line: { color: THEME.secondary, width: 2 } });
    sMap.addShape(pres.shapes.LINE, { x: ROOT_X + 1.55, y: L1_Y + 0.25, w: 0.25, h: 0, line: { color: THEME.secondary, width: 2 } });

    // L1 Node
    sMap.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: L1_X, y: L1_Y, w: 1.6, h: 0.5, fill: { color: THEME.secondary } });
    sMap.addText(node.t, { x: L1_X, y: L1_Y, w: 1.6, h: 0.5, color: THEME.white, bold: true, align: "center", fontFace: FONT_TITLE, fontSize: 13 });

    // L2 Nodes (3rd Level) - Font Size 16+
    node.c.forEach((child, j) => {
        const L2_X = L1_X + 2.0;
        const L2_Y = L1_Y - 0.25 + (j * 0.45); // Spread children vertically

        // Connect L1 -> L2
        sMap.addShape(pres.shapes.LINE, { x: L1_X + 1.6, y: L1_Y + 0.25, w: 0.2, h: 0, line: { color: THEME.accent, width: 1 } });
        sMap.addShape(pres.shapes.LINE, { x: L1_X + 1.8, y: Math.min(L1_Y + 0.25, L2_Y + 0.15), w: 0, h: Math.abs(L1_Y + 0.25 - (L2_Y + 0.15)), line: { color: THEME.accent, width: 1 } });
        sMap.addShape(pres.shapes.LINE, { x: L1_X + 1.8, y: L2_Y + 0.15, w: 0.2, h: 0, line: { color: THEME.accent, width: 1 } });

        // User requirement: 3rd level Font Size >= 16
        sMap.addText(child, {
            x: L2_X, y: L2_Y, w: 4.0, h: 0.35,
            fontSize: 16, // MANDATORY 16+
            color: THEME.text,
            fontFace: FONT_TITLE,
            align: "left",
            valign: "middle"
        });
    });
});

// Final Slide
let sL = pres.addSlide();
sL.background = { color: THEME.primary };
sL.addText("啟動您的數據領航之旅", { x: 0, y: 2.3, w: "100%", h: 0.6, bold: true, fontSize: 36, color: THEME.white, align: "center", fontFace: FONT_TITLE });

const outP = path.join(SCRATCH_DIR, "output", "HRBP_AI_Transformation_Zenith_v12.pptx");
pres.writeFile({ fileName: outP }).then(fn => {
    console.log(`Success: Generated v12 Pristine at ${fn}`);
}).catch(err => console.error(err));
