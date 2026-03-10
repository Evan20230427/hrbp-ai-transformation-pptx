const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

let pres = new pptxgen();
pres.layout = 'LAYOUT_16x9';
pres.author = 'Antigravity AI';
pres.title = 'HRBP AI 轉型最佳實務 - 最高完整性與階梯層級版 v18';

/**
 * THEME: ZENITH MAXIMUM INTEGRITY (V18 Final Master)
 */
const THEME = {
    primary: "0F172A",
    secondary: "3B82F6",
    text: "1E293B",
    white: "FFFFFF",
    line: "CBD5E1",
    accent: "64748B"
};

const FONT_TITLE = "Microsoft JhengHei";
const FONT_BODY = "Arial";

// Resource Management
const IMG_DIR = "C:\\Users\\TW-Evan.Chen\\.gemini\\antigravity\\brain\\cb95dffe-33bd-4e40-a98b-feaff376ea1a";
const SCRATCH_DIR = "C:\\Users\\TW-Evan.Chen\\.gemini\\antigravity\\scratch\\pdf-xlsx-to-pptx";

const ASSETS = {
    cover: path.join(IMG_DIR, "hrbp_simple_minimalist_cover_v12_1773160264391.png"),
    paradox: path.join(IMG_DIR, "hrbp_luxury_dashboard_v9_1773159099389.png"),
    human: path.join(IMG_DIR, "hrbp_professional_human_v12_simple_1773160281938.png"),
    role1: path.join(IMG_DIR, "hrbp_workforce_redesign_v13_simple_1773160725232.png"),
    role2: path.join(IMG_DIR, "hrbp_ai_ethics_v13_simple_1773160703050.png"),
    role3: path.join(IMG_DIR, "hrbp_professional_human_ai_v10_1773159552183.png"),
    p1: path.join(IMG_DIR, "hrbp_roadmap_v13_simple_1773160725232.png"),
    p2: path.join(IMG_DIR, "hrbp_data_strategy_meeting_v10_1773159591017.png"),
    p3: path.join(IMG_DIR, "hrbp_transformation_journey_v10_1773159569827.png")
};

/**
 * BILINGUAL TERM MAP (V16 Rules: Single language in parens)
 */
const BILINGUAL_NOTES = {
    "HRBP": "人力資源業務夥伴",
    "STL": "策略人才領袖",
    "AI 轉型": "AI Transformation",
    "人力重新設計": "Workforce Redesign",
    "AI 倫理": "AI Ethics",
    "人機協作": "Human-Machine Collaboration",
    "技能再造": "Reskilling",
    "自動化路徑": "Automation Roadmap",
    "成功指標": "Success Metrics",
    "產能回收": "Cycle Time Recovery"
};

/**
 * Enhanced Content Engine: Maximum Integrity + Boundary Guard.
 */
function renderContent(slide, lines, opts = {}) {
    let safeX = opts.x || 0.6;
    let safeY = opts.y || 1.4;
    let safeW = Math.min(opts.w || 6.2, 9.6 - safeX);
    let safeH = Math.min(opts.h || 3.4, 5.2 - safeY);

    const uniqueLines = [...new Set(lines)];
    const baseSize = opts.fontSize || 17;
    const content = [];

    uniqueLines.forEach(line => {
        let enhancedLine = line;
        Object.keys(BILINGUAL_NOTES).forEach(key => {
            if (enhancedLine.includes(key) && !enhancedLine.includes(`${key} (`)) {
                enhancedLine = enhancedLine.replace(key, `${key} (${BILINGUAL_NOTES[key]})`);
            }
        });

        const regex = /(\([^)]+\))/g;
        const tokens = enhancedLine.split(regex);
        tokens.forEach((token, tIdx) => {
            const isEngParenthetical = token.match(regex);
            content.push({
                text: token,
                options: {
                    fontSize: isEngParenthetical ? Math.max(11, baseSize - 5) : baseSize,
                    color: isEngParenthetical ? THEME.secondary : THEME.text,
                    fontFace: (isEngParenthetical || token.match(/[a-zA-Z]/)) ? FONT_BODY : FONT_TITLE,
                    italic: isEngParenthetical ? true : false,
                    bullet: (tIdx === 0),
                    breakLine: (tIdx === tokens.length - 1)
                }
            });
        });
    });

    slide.addText(content, {
        x: safeX, y: safeY, w: safeW, h: safeH,
        lineSpacing: opts.spacing || 24,
        valign: "top"
    });
}

pres.defineSlideMaster({
    title: 'MAX_INTEGRITY',
    background: { color: THEME.white },
    objects: [
        { rect: { x: 0, y: 0, w: 0.1, h: "100%", fill: { color: THEME.secondary } } }
    ]
});

function applyHeader(slide, title, pageNum) {
    slide.addText(title, { x: 0.5, y: 0.4, w: 9, h: 0.6, fontSize: 28, bold: true, color: THEME.primary, fontFace: FONT_TITLE });
    slide.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.0, w: 2.8, h: 0.04, fill: { color: THEME.secondary } });
    if (pageNum > 0) {
        slide.addText(pageNum.toString(), { x: 9.3, y: 5.2, w: 0.5, h: 0.3, fontSize: 11, color: THEME.secondary, align: "right", fontFace: FONT_BODY });
    }
}

let pg = 0;

// 1. Cover
let sTitle = pres.addSlide();
sTitle.background = { color: THEME.white };
if (fs.existsSync(ASSETS.cover)) sTitle.addImage({ path: ASSETS.cover, x: 5.8, y: 0.8, w: 3.8, h: 3.8, sizing: { type: 'contain' } });
sTitle.addText("重塑 HRBP 角色：\n引領企業 AI 轉型的策略實務", { x: 0.5, y: 1.8, w: 5.0, h: 1.5, fontSize: 34, bold: true, color: THEME.primary, fontFace: FONT_TITLE });
sTitle.addText("2026 年度最高完整性版 | 內容絕不精簡 & 階梯層級 (v18)", { x: 0.5, y: 3.3, w: 5.2, h: 0.5, fontSize: 15, color: THEME.secondary, fontFace: FONT_TITLE });

// 2. Challenges
pg++;
let s2 = pres.addSlide({ masterName: 'MAX_INTEGRITY' });
applyHeader(s2, "核心挑戰：策略參與的空洞化", pg);
renderContent(s2, [
    "數據揭露：高達 51% 的經理人認為 HRBP 參與了關鍵的策略討論，顯示仍有巨大的專業缺口。",
    "行政束縛：HRBP 目前仍深陷於低價值的行政事務性工作，如手動編寫職缺描述或彙整基礎數據摘要。",
    "轉型迫切：必須在 AI 轉型 浪潮中確立 STL 的核心治理地位，將人力資產轉化為企業競爭力。"
]);
if (fs.existsSync(ASSETS.paradox)) s2.addImage({ path: ASSETS.paradox, x: 7.0, y: 1.3, w: 2.6, h: 3.8, sizing: { type: "contain" } });

// 3. New Role STL
pg++;
let s3 = pres.addSlide({ masterName: 'MAX_INTEGRITY' });
applyHeader(s3, "未來定位：策略人才領袖 (STL)", pg);
if (fs.existsSync(ASSETS.human)) s3.addImage({ path: ASSETS.human, x: 0.5, y: 1.3, w: 3.5, h: 3.8, sizing: { type: "cover" } });
renderContent(s3, [
    "角色升級：從人員策略的「翻譯者」轉型為具備商業洞察力的「AI 轉型 顧問」。",
    "核心價值：主導高階策略對話，解決 AI 決策中可能產生的算法偏見與倫理治理問題。",
    "賦能轉型：主導 人力重新設計 與 技能再造 方案，確保組織人才結構與未來技術佈局高度契合。"
], { x: 4.5, w: 5.0, fontSize: 18 });

// 4. Responsibility 1: Workforce Redesign
pg++;
let s4 = pres.addSlide({ masterName: 'MAX_INTEGRITY' });
applyHeader(s4, "職責 1：主導 人力重新設計", pg);
renderContent(s4, [
    "職能建模：前瞻性規劃隨 AI 技術演進而改變的高階人才角色與職能能力模型。",
    "技能再造：主導企業內部的 技能再造 決策流程，確保人才轉型步調領先於技術更迭。",
    "價值定義：清晰定義各職務中可被自動化替代的部分，並將人力資源聚焦於創造高增值的策略領域。 "
], { y: 3.8, h: 1.4, w: 9, x: 0.6 });
if (fs.existsSync(ASSETS.role1)) s4.addImage({ path: ASSETS.role1, x: 0.5, y: 1.2, w: 9, h: 2.3, sizing: { type: "cover" } });

// 5. Responsibility 2: Ethics
pg++;
let s5 = pres.addSlide({ masterName: 'MAX_INTEGRITY' });
applyHeader(s5, "職責 2：應對 AI 倫理 與偏見", pg);
if (fs.existsSync(ASSETS.role2)) s5.addImage({ path: ASSETS.role2, x: 5.8, y: 1.3, w: 3.7, h: 3.7, sizing: { type: "contain" } });
renderContent(s5, [
    "算法監測：持續追蹤與監測 AI 技術輔助決策中的算法公平性、透明度與潛在的數據偏見。",
    "治理標準：建立負責任的 數位化治理 標準，確保 HR 數據在自動化流程中的合法性與透明化。 ",
    "技術透明：確保人才招聘與績效篩選機制中，不存在隱性的技術偏見，維護企業雇用品牌公正性。"
], { x: 0.6, w: 4.8, y: 1.8 });

// 6. Responsibility 3: Collaboration
pg++;
let s6 = pres.addSlide({ masterName: 'MAX_INTEGRITY' });
applyHeader(s6, "職責 3：優化 人機協作 效率", pg);
renderContent(s6, [
    "工作流設計：設計能平衡 AI 自動化帶來的生產力增益與員工實際參與度的數位化工作流體系。",
    "文化一致性：在轉型過程中維護組織文化的一致性，動態降低技術轉型引發的基層心理壓力和不安感。",
    "效能評估：建立 科學的 人機協作 效能評估反饋體系，持續優化人與技術的最佳配置。 "
], { x: 4.5, w: 5.0, y: 1.5 });
if (fs.existsSync(ASSETS.role3)) s6.addImage({ path: ASSETS.role3, x: 0.5, y: 1.5, w: 3.7, h: 3.7, sizing: { type: "contain" } });

// 7-9: Phases (Full Content preservation)
const PHASES = [
    { t: "P1 剔除行動：產能回收 計畫", c: ["核心重點：明確策略區域，徹底剔除低價值、可重複的行政舊務。", "進程標準：建立視覺化的 自動化路徑 進程標竿與具體轉型時程表。"], img: ASSETS.p1, flip: false },
    { t: "P2 強化行動：賦能 核心責任", c: ["預測模型：利用高階 AI 預測模型優化組織繼任計劃與人才準備度。", "二次投資：將行政效率提升後的 產能回收 轉化為對關鍵人才資本的二次開發決策。"], img: ASSETS.p2, flip: true },
    { t: "P3 開拓階段：新型 策略領域", c: ["嵌入決策：在關鍵業務變革中深度嵌入 STL 策略條款與人才建議。", "治理主導：確立 HR 部門在企業級 AI 轉型 重大決策中的核心主導權位。"], img: ASSETS.p3, flip: false }
];

PHASES.forEach(item => {
    pg++;
    let s = pres.addSlide({ masterName: 'MAX_INTEGRITY' });
    applyHeader(s, item.t, pg);
    if (!item.flip) {
        renderContent(s, item.c, { x: 0.6, w: 5.8, y: 1.8 });
        if (fs.existsSync(item.img)) s.addImage({ path: item.img, x: 6.8, y: 1.2, w: 2.8, h: 4, sizing: { type: "contain" } });
    } else {
        if (fs.existsSync(item.img)) s.addImage({ path: item.img, x: 0.5, y: 1.2, w: 2.8, h: 4, sizing: { type: "contain" } });
        renderContent(s, item.c, { x: 3.6, w: 6.0, y: 1.8 });
    }
});

// Final Slide: MIND MAP (Safe Boundaries, v18 Rich Content Matching, Level Font Rules)
pg++;
let sMM = pres.addSlide();
sMM.background = { color: THEME.white };
applyHeader(sMM, "全課精華：內容完整呼應心智圖 (v18)", pg);

const MX = 0.5, MY = 2.4;
const M_DATA = [
    { t: "現狀挑戰與策略缺口", c: ["策略參與(51%)不足", "行政事務性束縛", "AI 轉型之急迫性"] },
    { t: "STL 定位與核心責任", c: ["主導人力重新設計", "應對 AI 倫理治理", "優化人機協作效能"] },
    { t: "轉型三階實務路徑", c: ["P1 剔除行政回收產能", "P2 強化預測賦能責任", "P3 開拓開拓策略主导"] },
    { t: "驗收指標與成功衡量", c: ["產能回收與效率提升", "繼任人才庫之準備度", "員工投入與技術參與"] }
];

// Root Node
sMM.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: MX, y: MY, w: 1.4, h: 0.6, fill: { color: THEME.primary }, rectRadius: 0.1 });
sMM.addText("HRBP\nAI 轉型", { x: MX, y: MY, w: 1.4, h: 0.6, color: THEME.white, bold: true, align: "center", fontFace: FONT_TITLE, fontSize: 13 });

M_DATA.forEach((n, i) => {
    let nx = MX + 2.0, ny = 0.82 + (i * 1.15);
    // Root -> L1
    sMM.addShape(pres.shapes.LINE, { x: MX + 1.4, y: MY + 0.3, w: 0.2, h: 0, line: { color: THEME.secondary, width: 1.5 } });
    sMM.addShape(pres.shapes.LINE, { x: MX + 1.6, y: Math.min(MY + 0.3, ny + 0.25), w: 0, h: Math.abs(MY + 0.3 - (ny + 0.25)), line: { color: THEME.secondary, width: 1.5 } });
    sMM.addShape(pres.shapes.LINE, { x: MX + 1.6, y: ny + 0.25, w: 0.4, h: 0, line: { color: THEME.secondary, width: 1.5 } });

    // L1 Node (Level 2) - MUST BE >= Level 3. Setting L2=16pt, L3=14pt
    sMM.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: nx, y: ny, w: 1.8, h: 0.5, fill: { color: THEME.secondary }, rectRadius: 0.1 });
    sMM.addText(n.t, { x: nx, y: ny, w: 1.8, h: 0.5, color: THEME.white, bold: true, align: "center", fontFace: FONT_TITLE, fontSize: 16 }); // L2 @ 16pt

    // L2 Nodes (Level 3 Child) - Font 14
    n.c.forEach((ch, j) => {
        let cx = nx + 2.1, cy = ny - 0.22 + (j * 0.42);
        // L1 -> L2
        sMM.addShape(pres.shapes.LINE, { x: nx + 1.8, y: ny + 0.25, w: 0.1, h: 0, line: { color: THEME.line, width: 1 } });
        sMM.addShape(pres.shapes.LINE, { x: nx + 1.9, y: Math.min(ny + 0.25, cy + 0.15), w: 0, h: Math.abs(ny + 0.25 - (cy + 0.15)), line: { color: THEME.line, width: 1 } });
        sMM.addShape(pres.shapes.LINE, { x: nx + 1.9, y: cy + 0.15, w: 0.2, h: 0, line: { color: THEME.line, width: 1 } });

        sMM.addText(ch, { x: cx, y: cy, w: 3.5, h: 0.35, color: THEME.text, fontSize: 14, fontFace: FONT_TITLE, valign: "middle" }); // L3 @ 14pt
    });
});

const outP = path.join(SCRATCH_DIR, "output", "HRBP_AI_Transformation_Zenith_v18.pptx");
pres.writeFile({ fileName: outP }).then(fn => {
    console.log(`Success: Zenith v18 (Maximum Integrity) Generated at ${fn}`);
}).catch(err => console.error(err));
