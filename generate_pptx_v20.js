const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

let pres = new pptxgen();
pres.layout = 'LAYOUT_16x9';
pres.author = 'Antigravity AI';
pres.title = 'HRBP AI 轉型最佳實務 - 大理石之生 v20';

/**
 * THEME: MARMOREAL LIFE (V20 Ultimate Roman)
 */
const THEME = {
    primary: "1E293B",
    secondary: "3B82F6",
    text: "334155",
    white: "FFFFFF",
    line: "E2E8F0",
    accent: "64748B",
    highlight: "FFFF00" // Yellow Background for 1:1 Mapping
};

const FONT_TITLE = "Microsoft JhengHei";
const FONT_BODY = "Arial";

// Resource Management (Roman Series v20)
const IMG_DIR = "C:\\Users\\TW-Evan.Chen\\.gemini\\antigravity\\brain\\cb95dffe-33bd-4e40-a98b-feaff376ea1a";
const SCRATCH_DIR = "C:\\Users\\TW-Evan.Chen\\.gemini\\antigravity\\scratch\\pdf-xlsx-to-pptx";

const ASSETS = {
    cover: path.join(IMG_DIR, "hrbp_roman_cover_v20_1773173000010.png"),
    paradox: path.join(IMG_DIR, "hrbp_roman_paradox_v20_1773173000011.png"),
    stl: path.join(IMG_DIR, "hrbp_roman_stl_v20_1773173000012.png"),
    workforce: path.join(IMG_DIR, "hrbp_roman_workforce_v20_1773173000013.png"),
    ethics: path.join(IMG_DIR, "hrbp_roman_ethics_v20_1773173000014.png"),
    collab: path.join(IMG_DIR, "hrbp_roman_collab_v20_1773173000015.png"),
    p1: path.join(IMG_DIR, "hrbp_roman_p1_v20_1773173000016.png"),
    p2: path.join(IMG_DIR, "hrbp_roman_p2_v20_1773173000017.png"),
    p3: path.join(IMG_DIR, "hrbp_roman_p3_v20_1773173000018.png")
};

const BILINGUAL_NOTES = {
    "HRBP": "人力資源業務夥伴",
    "STL": "策略人才領袖",
    "AI 轉型": "AI Transformation",
    "人力重新設計": "Workforce Redesign",
    "AI 倫理治理": "AI Ethics Governance",
    "人機協作效率": "Human-Machine Collaboration",
    "技能再造": "Reskilling",
    "自動化路徑專案": "Automation Roadmap",
    "行政事務性束縛": "Administrative Constraints",
    "產能回收計畫": "Cycle Time Recovery"
};

const HIGHLIGHT_LIST = [
    "策略參與", "挑戰與契機", "行政事務性束縛", "AI 轉型",
    "策略人才領袖", "STL", "人力重新設計", "AI 倫理治理", "人機協作效率",
    "產能回收計畫", "自動化路徑專案", "技能再造", "成功指標", "繼任人才庫"
];

/**
 * 核心渲染引擎：支援黃底高亮 + 階層式文字處理
 */
function renderContent(slide, lines, opts = {}) {
    let safeX = opts.x || 0.6;
    let safeY = opts.y || 1.4;
    let safeW = Math.min(opts.w || 6.2, 9.6 - safeX);
    let safeH = Math.min(opts.h || 3.4, 5.2 - safeY);

    const baseSize = opts.fontSize || 17;
    const content = [];

    lines.forEach(line => {
        let currentLine = line;

        // 1. 注入雙語 (V16 規則)
        Object.keys(BILINGUAL_NOTES).forEach(key => {
            if (currentLine.includes(key) && !currentLine.includes(`${key} (`)) {
                currentLine = currentLine.replace(key, `${key} (${BILINGUAL_NOTES[key]})`);
            }
        });

        // 2. 切分高亮與括號
        const parts = currentLine.split(/(\([^)]+\))/g);
        parts.forEach((part, pIdx) => {
            if (!part) return;
            const isParens = part.startsWith("(") && part.endsWith(")");

            if (!isParens) {
                let subParts = [part];
                HIGHLIGHT_LIST.forEach(term => {
                    let next = [];
                    subParts.forEach(sp => {
                        if (typeof sp === 'string') {
                            const exploded = sp.split(new RegExp(`(${term})`, 'g'));
                            exploded.forEach(e => next.push(e));
                        } else { next.push(sp); }
                    });
                    subParts = next;
                });

                subParts.forEach((sp, sIdx) => {
                    if (!sp) return;
                    const isHighlight = HIGHLIGHT_LIST.includes(sp);
                    content.push({
                        text: sp,
                        options: {
                            fontSize: baseSize,
                            color: isHighlight ? "#000000" : THEME.text,
                            fill: isHighlight ? THEME.highlight : null,
                            fontFace: sp.match(/[a-zA-Z]/) ? FONT_BODY : FONT_TITLE,
                            bold: isHighlight,
                            bullet: (pIdx === 0 && sIdx === 0),
                            breakLine: (pIdx === parts.length - 1 && sIdx === subParts.length - 1)
                        }
                    });
                });
            } else {
                content.push({
                    text: part,
                    options: {
                        fontSize: Math.max(11, baseSize - 5),
                        color: THEME.secondary,
                        fontFace: FONT_BODY,
                        italic: true,
                        breakLine: (pIdx === parts.length - 1)
                    }
                });
            }
        });
    });

    slide.addText(content, {
        x: safeX, y: safeY, w: safeW, h: safeH,
        lineSpacing: opts.spacing || 24,
        valign: "top"
    });
}

pres.defineSlideMaster({
    title: 'MARMOREAL_V20',
    background: { color: THEME.white },
    objects: [
        { rect: { x: 0, y: 0, w: 0.1, h: "100%", fill: { color: THEME.secondary } } },
        { shape: pres.shapes.RECTANGLE, options: { x: 0.5, y: 5.5, w: 9.0, h: 0.1, fill: { color: THEME.line } } }
    ]
});

function applyHeader(slide, title, pageNum) {
    slide.addText(title, { x: 0.5, y: 0.4, w: 9, h: 0.6, fontSize: 30, bold: true, color: THEME.primary, fontFace: FONT_TITLE });
    slide.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.1, w: 1.5, h: 0.05, fill: { color: THEME.secondary } });
    if (pageNum > 0) {
        slide.addText(`MDL | PAGE ${pageNum}`, { x: 8.5, y: 5.2, w: 1.2, h: 0.3, fontSize: 10, color: THEME.secondary, align: "right", fontFace: FONT_BODY, bold: true });
    }
}

let pg = 0;

// 1. Cover
let sTitle = pres.addSlide();
sTitle.background = { color: THEME.white };
if (fs.existsSync(ASSETS.cover)) sTitle.addImage({ path: ASSETS.cover, x: 5.5, y: 0.5, w: 4.5, h: 4.5, sizing: { type: 'contain' } });
sTitle.addText("重塑 HRBP 角色：\n引領企業 AI 轉型的策略實務", { x: 0.5, y: 1.8, w: 5.2, h: 1.5, fontSize: 36, bold: true, color: THEME.primary, fontFace: FONT_TITLE });
sTitle.addText("大理石之生 | 四階連動與黃底標籤版 (v20)", { x: 0.5, y: 3.3, w: 5.5, h: 0.5, fontSize: 16, color: THEME.secondary, fontFace: FONT_TITLE });

// 2-9 Slides with Highlighting
pg++;
let s2 = pres.addSlide({ masterName: 'MARMOREAL_V20' });
applyHeader(s2, "挑戰與契機：打破策略孤島", pg);
renderContent(s2, [
    "數據揭謬：51% 的 HRBP 在 策略參與 方面存在顯著斷層。",
    "困局解鎖：擺脫 行政事務性束縛，將產能釋放於高增值的治理領域。",
    "轉型奇點：AI 轉型 不僅是技術升級，更是 HR 角色邊界的重新詮釋。 "
]);
if (fs.existsSync(ASSETS.paradox)) s2.addImage({ path: ASSETS.paradox, x: 6.8, y: 1.2, w: 2.8, h: 4.2 });

pg++;
let s3 = pres.addSlide({ masterName: 'MARMOREAL_V20' });
applyHeader(s3, "未來角色：策略人才領袖 (STL)", pg);
if (fs.existsSync(ASSETS.stl)) s3.addImage({ path: ASSETS.stl, x: 0.5, y: 1.3, w: 3.5, h: 4.0 });
renderContent(s3, [
    "身份跨越：從執行者演進為真正的 策略人才領袖。",
    "核心責任：主導 AI 倫理治理，在自動化浪潮中捍衛公平價值。",
    "效能賦能：透過重新定義流程，極大化 人機協作效率。 "
], { x: 4.5, w: 5.0 });

pg++;
let s4 = pres.addSlide({ masterName: 'MARMOREAL_V20' });
applyHeader(s4, "任務 I：主導 人力重新設計", pg);
renderContent(s4, [
    "結構調整：基於產出回收預測，執行規模化的 人力重新設計。",
    "人才韌性：將 技能再造 視為企業核心競爭力的二次投資。",
    "職能建模：定義 AI 年代不可替代的人格化高價值領域。"
], { y: 3.9, h: 1.3, w: 9, x: 0.6 });
if (fs.existsSync(ASSETS.workforce)) s4.addImage({ path: ASSETS.workforce, x: 0.5, y: 1.2, w: 10, h: 2.5, sizing: { type: 'cover' } });

pg++;
let s5 = pres.addSlide({ masterName: 'MARMOREAL_V20' });
applyHeader(s5, "任務 II：建立 AI 倫理治理 憲法", pg);
if (fs.existsSync(ASSETS.ethics)) s5.addImage({ path: ASSETS.ethics, x: 6.2, y: 1.2, w: 3.2, h: 3.8 });
renderContent(s5, [
    "偏見監測：建立嚴格的算法審計機制，確保 AI 倫理治理 的公正性。",
    "透明治理：確保每一項 AI 代行的決策皆具備可追溯性與解釋性。",
    "品牌信任：透過負責、透明的技術應用，維護企業僱傭品牌。 "
], { x: 0.6, w: 5.4 });

pg++;
let s6 = pres.addSlide({ masterName: 'MARMOREAL_V20' });
applyHeader(s6, "任務 III：優化 人機協作效率", pg);
renderContent(s6, [
    "共生流設計：設計具備高同步性的人機共生引擊，優化 人機協作效率。",
    "心理賦能：動態評估員工對技術的適應度，降低轉型期的技術焦慮。",
    "價值挖掘：挖掘 AI 產出中的非結構化洞察，輔助 STL 決策。"
], { x: 4.5, w: 5.0 });
if (fs.existsSync(ASSETS.collab)) s6.addImage({ path: ASSETS.collab, x: 0.5, y: 1.4, w: 3.8, h: 3.8 });

const PHASES = [
    { t: "P1 剔除動作：產能回收計畫", c: ["重點核心：制定 自動化路徑專案，徹底回收行政型 產能回收計畫。", "基石建立：對低價值重疊流程進行「斷捨離」。"], img: ASSETS.p1, pg: 7 },
    { t: "P2 強化動作：繼任人才庫", c: ["核心賦能：利用預測模型鎖定 繼任人才庫，確保未來競爭力。", "職能放大：將回收產能轉化為高階策略人才的精準投資。"], img: ASSETS.p2, pg: 8 },
    { t: "P3 開拓動作：策略主導權", c: ["關鍵開拓：在組織大腦中嵌入 STL 策略節點。", "主導轉型：全面執掌 AI 轉型 決策權位。"], img: ASSETS.p3, pg: 9 }
];

PHASES.forEach(p => {
    pg++;
    let s = pres.addSlide({ masterName: 'MARMOREAL_V20' });
    applyHeader(s, p.t, pg);
    renderContent(s, p.c, { x: 0.6, w: 6.0, y: 1.8 });
    if (fs.existsSync(p.img)) s.addImage({ path: p.img, x: 7.0, y: 1.2, w: 2.5, h: 4.0 });
});

// Final Slide: 4-LEVEL MIND MAP (v20 Complete Alignment)
pg++;
let sMM = pres.addSlide();
applyHeader(sMM, "大理石之生：羅馬美學與四階全量對齊 (v20)", pg);

const M_DATA = [
    {
        t: "挑戰與契機：策略孤島", c: [
            { k: "策略參與(51%)缺口", p: 2 },
            { k: "行政事務性束縛", p: 2 },
            { k: "AI 轉型 與角色重塑", p: 2 }
        ]
    },
    {
        t: "STL 定位：決策領袖", c: [
            { k: "人力重新設計 主導", p: 4 },
            { k: "AI 倫理治理 監測", p: 5 },
            { k: "人機協作效率 優化", p: 6 }
        ]
    },
    {
        t: "實務攻略：三階路徑", c: [
            { k: "P1 自動化路徑專案", p: 7 },
            { k: "P2 繼任人才庫 賦能", p: 8 },
            { k: "P3 策略主導權 確立", p: 9 }
        ]
    },
    {
        t: "驗收指標：轉型成效", c: [
            { k: "產能回收計畫 與效率", p: 7 },
            { k: "技能再造 適配指標", p: 4 },
            { k: "繼任人才庫 之準備度", p: 8 }
        ]
    }
];

const MX = 0.5, MY = 2.4;
sMM.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: MX, y: MY, w: 1.4, h: 0.6, fill: { color: THEME.primary }, rectRadius: 0.1 });
sMM.addText("HRBP\nAI 轉型", { x: MX, y: MY, w: 1.4, h: 0.6, color: THEME.white, bold: true, align: "center", fontFace: FONT_TITLE, fontSize: 13 });

M_DATA.forEach((n, i) => {
    let nx = MX + 2.0, ny = 0.82 + (i * 1.15);
    sMM.addShape(pres.shapes.LINE, { x: MX + 1.4, y: MY + 0.3, w: 0.2, h: 0, line: { color: THEME.secondary, width: 1.5 } });
    sMM.addShape(pres.shapes.LINE, { x: MX + 1.6, y: Math.min(MY + 0.3, ny + 0.25), w: 0, h: Math.abs(MY + 0.3 - (ny + 0.25)), line: { color: THEME.secondary, width: 1.5 } });
    sMM.addShape(pres.shapes.LINE, { x: MX + 1.6, y: ny + 0.25, w: 0.4, h: 0, line: { color: THEME.secondary, width: 1.5 } });

    // L2 Node (16pt)
    sMM.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: nx, y: ny, w: 1.8, h: 0.5, fill: { color: THEME.secondary }, rectRadius: 0.1 });
    sMM.addText(n.t, { x: nx, y: ny, w: 1.8, h: 0.5, color: THEME.white, bold: true, align: "center", fontFace: FONT_TITLE, fontSize: 16 });

    n.c.forEach((ch, j) => {
        let cx = nx + 2.1, cy = ny - 0.22 + (j * 0.42);
        sMM.addShape(pres.shapes.LINE, { x: nx + 1.8, y: ny + 0.25, w: 0.1, h: 0, line: { color: THEME.line, width: 1 } });
        sMM.addShape(pres.shapes.LINE, { x: nx + 1.9, y: Math.min(ny + 0.25, cy + 0.15), w: 0, h: Math.abs(ny + 0.25 - (cy + 0.15)), line: { color: THEME.line, width: 1 } });
        sMM.addShape(pres.shapes.LINE, { x: nx + 1.9, y: cy + 0.15, w: 0.2, h: 0, line: { color: THEME.line, width: 1 } });

        // L3 Node (14pt)
        sMM.addText(ch.k, { x: cx, y: cy, w: 2.6, h: 0.35, color: THEME.text, fontSize: 14, fontFace: FONT_TITLE, valign: "middle" });

        // L4 Node (Page Reference, <12pt)
        let px = cx + 2.5;
        sMM.addShape(pres.shapes.LINE, { x: cx + 2.45, y: cy + 0.15, w: 0.05, h: 0, line: { color: THEME.accent, width: 0.5, dashType: "dash" } });
        sMM.addText(`P.${ch.p}`, { x: px, y: cy, w: 0.5, h: 0.35, color: THEME.accent, fontSize: 10, fontFace: FONT_BODY, valign: "middle", bold: true });
    });
});

const outP = path.join(SCRATCH_DIR, "output", "HRBP_AI_Transformation_Ultimate_v20.pptx");
pres.writeFile({ fileName: outP }).then(fn => {
    console.log(`Final Success: v20 Ultimate Roman Generated at ${fn}`);
}).catch(err => console.error(err));
