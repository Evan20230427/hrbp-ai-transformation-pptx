const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

let pres = new pptxgen();
pres.layout = 'LAYOUT_16x9';
pres.author = 'Antigravity AI';
pres.title = 'HRBP AI 轉型最佳實務 - 羅馬美學與四階全量版 v19';

/**
 * THEME: ROMAN MARBLE & AZURE (V19 Aesthetic)
 */
const THEME = {
    primary: "0F172A",
    secondary: "3B82F6",
    text: "1E293B",
    white: "FFFFFF",
    line: "CBD5E1",
    accent: "64748B",
    highlight: "FFFF00" // Yellow Highlight for alignment
};

const FONT_TITLE = "Microsoft JhengHei";
const FONT_BODY = "Arial";

// Resource Management (Using Roman Statue placeholders until quota resets)
const IMG_DIR = "C:\\Users\\TW-Evan.Chen\\.gemini\\antigravity\\brain\\cb95dffe-33bd-4e40-a98b-feaff376ea1a";
const SCRATCH_DIR = "C:\\Users\\TW-Evan.Chen\\.gemini\\antigravity\\scratch\\pdf-xlsx-to-pptx";

// Placeholder assets for Roman series (v19 unique path)
const ASSETS = {
    cover: path.join(IMG_DIR, "hrbp_roman_cover_v19.png"),
    paradox: path.join(IMG_DIR, "hrbp_roman_paradox_v19.png"),
    stl: path.join(IMG_DIR, "hrbp_roman_stl_v19.png"),
    workforce: path.join(IMG_DIR, "hrbp_roman_workforce_v19.png"),
    ethics: path.join(IMG_DIR, "hrbp_roman_ethics_v19.png"),
    collab: path.join(IMG_DIR, "hrbp_roman_collab_v19.png"),
    p1: path.join(IMG_DIR, "hrbp_roman_p1_v19.png"),
    p2: path.join(IMG_DIR, "hrbp_roman_p2_v19.png"),
    p3: path.join(IMG_DIR, "hrbp_roman_p3_v19.png")
};

/**
 * HIGHLIGHT_TERMS: Keywords that must be highlighted in yellow background.
 * These correspond to the Mind Map nodes for 1:1 mapping.
 */
const HIGHLIGHT_TERMS = [
    "策略參與", "挑戰與契機", "行政事務性束縛", "AI 轉型",
    "策略人才領袖", "STL", "人力重新設計", "AI 倫理治理", "人機協作效率",
    "產能回收", "自動化路徑", "技能再造", "成功指標", "繼任人才庫"
];

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
 * Enhanced Engine: Yellow Highlight + Single Bilingual + Boundary Proof.
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

        // 1. Inject Bilingual (V16 Rule)
        Object.keys(BILINGUAL_NOTES).forEach(key => {
            if (currentLine.includes(key) && !currentLine.includes(`${key} (`)) {
                currentLine = currentLine.replace(key, `${key} (${BILINGUAL_NOTES[key]})`);
            }
        });

        // 2. Fragment by Parenthesis and Highlights
        // Regex to split by (parens) OR highlight terms (keeping delimiters)
        const parts = currentLine.split(/(\([^)]+\))/g);

        parts.forEach((part, pIdx) => {
            if (part === "") return;

            const isParens = part.startsWith("(") && part.endsWith(")");

            // Sub-fragment by HIGHLIGHT_TERMS if not a parenthesis
            if (!isParens) {
                let subParts = [part];
                HIGHLIGHT_TERMS.forEach(term => {
                    let nextSubParts = [];
                    subParts.forEach(sp => {
                        if (typeof sp === 'string') {
                            const exploded = sp.split(new RegExp(`(${term})`, 'g'));
                            exploded.forEach(ex => nextSubParts.push(ex));
                        } else {
                            nextSubParts.push(sp); // Already an object
                        }
                    });
                    subParts = nextSubParts;
                });

                subParts.forEach((sp, sIdx) => {
                    if (sp === "") return;
                    const isHighlight = HIGHLIGHT_TERMS.includes(sp);
                    content.push({
                        text: sp,
                        options: {
                            fontSize: baseSize,
                            color: isHighlight ? "#000000" : THEME.text,
                            fill: isHighlight ? THEME.highlight : null,
                            fontFace: sp.match(/[a-zA-Z]/) ? FONT_BODY : FONT_TITLE,
                            bullet: (pIdx === 0 && sIdx === 0),
                            breakLine: (pIdx === parts.length - 1 && sIdx === subParts.length - 1)
                        }
                    });
                });
            } else {
                // Handle Parenthesis (Bilingual Note)
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
    title: 'ROMAN_V19',
    background: { color: THEME.white },
    objects: [
        { rect: { x: 0, y: 0, w: 0.1, h: "100%", fill: { color: THEME.secondary } } }
    ]
});

function applyHeader(slide, title, pageNum) {
    slide.addText(title, { x: 0.5, y: 0.4, w: 9, h: 0.6, fontSize: 28, bold: true, color: THEME.primary, fontFace: FONT_TITLE });
    slide.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.0, w: 2.8, h: 0.04, fill: { color: THEME.secondary } });
    if (pageNum > 0) {
        slide.addText(`PAGE ${pageNum}`, { x: 8.8, y: 5.2, w: 1.0, h: 0.3, fontSize: 10, color: THEME.secondary, align: "right", fontFace: FONT_BODY });
    }
}

let pg = 0;

// 1. Cover
let sTitle = pres.addSlide();
sTitle.background = { color: THEME.white };
if (fs.existsSync(ASSETS.cover)) sTitle.addImage({ path: ASSETS.cover, x: 5.8, y: 0.8, w: 3.8, h: 3.8, sizing: { type: 'contain' } });
sTitle.addText("重塑 HRBP 角色：\n引領企業 AI 轉型的策略實務", { x: 0.5, y: 1.8, w: 5.0, h: 1.5, fontSize: 34, bold: true, color: THEME.primary, fontFace: FONT_TITLE });
sTitle.addText("古羅馬雕像日常美學 | 四階心智圖 & 黃底對齊 (v19)", { x: 0.5, y: 3.3, w: 5.2, h: 0.5, fontSize: 15, color: THEME.secondary, fontFace: FONT_TITLE });

// 2. Challenges
pg++;
let s2 = pres.addSlide({ masterName: 'ROMAN_V19' });
applyHeader(s2, "現狀與契機：策略參與之缺口", pg);
renderContent(s2, [
    "策略參與不足：高達 51% 的核心討論中 HRBP 缺席，形成嚴重的策略孤島。",
    "行政事務性束縛：被低增值的事務性工作淹沒，阻礙了向價值創造者的轉型。",
    "AI 轉型 之潮：這是重塑 HR 邊界的 挑戰與契機，將數據轉化為治理資產。"
]);
if (fs.existsSync(ASSETS.paradox)) s2.addImage({ path: ASSETS.paradox, x: 7.0, y: 1.3, w: 2.6, h: 3.8, sizing: { type: "contain" } });

// 3. New Role STL
pg++;
let s3 = pres.addSlide({ masterName: 'ROMAN_V19' });
applyHeader(s3, "未來定位：策略人才領袖 (STL)", pg);
if (fs.existsSync(ASSETS.stl)) s3.addImage({ path: ASSETS.stl, x: 0.5, y: 1.3, w: 3.5, h: 3.8, sizing: { type: "cover" } });
renderContent(s3, [
    "STL 定位：不僅是翻譯者，而是具備技術深度的人才資本 策略人才領袖。",
    "治理核心：主導 AI 倫理治理，規避數據驅動決策中的隱性偏見。",
    "生產力重構：優化 人機協作效率，實現產出的質性飛躍。"
], { x: 4.5, w: 5.0, fontSize: 18 });

// 4. Responsibility 1: Workforce Redesign
pg++;
let s4 = pres.addSlide({ masterName: 'ROMAN_V19' });
applyHeader(s4, "責任 I：主導 人力重新設計", pg);
renderContent(s4, [
    "人力重新設計：基於技術更迭預測，動態調整高階職能與團隊架構。",
    "技能再造：以前瞻性視角執行 技能再造 決策，解決未來的人才飢渴。",
    "增值區域：識別自動化邊界，釋放人力於不可替代的創造性策略領域。"
], { y: 3.8, h: 1.4, w: 9, x: 0.6 });
if (fs.existsSync(ASSETS.workforce)) s4.addImage({ path: ASSETS.workforce, x: 0.5, y: 1.2, w: 9, h: 2.3, sizing: { type: "cover" } });

// 5. Responsibility 2: Ethics
pg++;
let s5 = pres.addSlide({ masterName: 'ROMAN_V19' });
applyHeader(s5, "責任 II：AI 倫理治理 標準", pg);
if (fs.existsSync(ASSETS.ethics)) s5.addImage({ path: ASSETS.ethics, x: 5.8, y: 1.3, w: 3.7, h: 3.7, sizing: { type: "contain" } });
renderContent(s5, [
    "算法公平：監測人資 AI 中的決策透明度，達成 AI 倫理治理 之標竿。",
    "標準建立：建立負責任的數據化準則，防止技術偏見侵蝕企業合規性。",
    "品牌保護：確保技術篩選過程透明且公正，維護良善之企業雇主形象。"
], { x: 0.6, w: 4.8, y: 1.8 });

// 6. Responsibility 3: Collaboration
pg++;
let s6 = pres.addSlide({ masterName: 'ROMAN_V19' });
applyHeader(s6, "責任 III：優化 人機協作效率", pg);
renderContent(s6, [
    "流程優化：設計具備高擴展性的人機界面與流程，極大化 人機協作效率。",
    "焦慮緩解：動態降解技術轉型帶來的心理負載，維護組織文化韌性。",
    "反饋循環：建立科學的生產力反饋體系，持續迭代人的不可替代價值。"
], { x: 4.5, w: 5.0, y: 1.5 });
if (fs.existsSync(ASSETS.collab)) s6.addImage({ path: ASSETS.collab, x: 0.5, y: 1.5, w: 3.7, h: 3.7, sizing: { type: "contain" } });

// 7-9: Phases
const PHASES = [
    { t: "P1 剔除行動：產能回收 計畫", c: ["剔除冗務：藉由 自動化路徑 標竿，徹底回收低價值之行政產能。", "效率革命：建立視覺化的轉型地圖，明確回收之重心與具體時點。"], img: ASSETS.p1, flip: false, page: 7 },
    { t: "P2 強化行動：賦能 策略職能", c: ["預測賦能：利用預測模型強化繼任人才管理，提升決策之前瞻性。", "資產轉化：將 回收產能 轉化為人才二次開發的資本動能。"], img: ASSETS.p2, flip: true, page: 8 },
    { t: "P3 開拓階段：嵌入 策略主導", c: ["開拓領地：在關鍵變革中深度嵌入 STL 條款，確立治理高度。", "主導轉型：全面主導 AI 轉型 決策，達成 HR 職能的終極躍遷。"], img: ASSETS.p3, flip: false, page: 9 }
];

PHASES.forEach(item => {
    pg++;
    let s = pres.addSlide({ masterName: 'ROMAN_V19' });
    applyHeader(s, item.t, pg);
    if (!item.flip) {
        renderContent(s, item.c, { x: 0.6, w: 5.8, y: 1.8 });
        if (fs.existsSync(item.img)) s.addImage({ path: item.img, x: 6.8, y: 1.2, w: 2.8, h: 4, sizing: { type: "contain" } });
    } else {
        if (fs.existsSync(item.img)) s.addImage({ path: item.img, x: 0.5, y: 1.2, w: 2.8, h: 4, sizing: { type: "contain" } });
        renderContent(s, item.c, { x: 3.6, w: 6.0, y: 1.8 });
    }
});

// Final Slide: 4-LEVEL MIND MAP (Node -> L1 -> L2 -> Pagemap)
pg++;
let sMM = pres.addSlide();
sMM.background = { color: THEME.white };
applyHeader(sMM, "全課精華：四層階梯層級與內容呼應 (v19)", pg);

const MX = 0.5, MY = 2.4;
const M_DATA = [
    {
        t: "挑戰與契機", c: [
            { k: "策略參與(51%)不足", p: 2 },
            { k: "行政事務性束縛", p: 2 },
            { k: "AI 轉型 之衝擊", p: 2 }
        ]
    },
    {
        t: "STL 定位責任", c: [
            { k: "人力重新設計", p: 4 },
            { k: "AI 倫理治理", p: 5 },
            { k: "人機協作效率", p: 6 }
        ]
    },
    {
        t: "轉型攻略路徑", c: [
            { k: "P1 剔除回收產能", p: 7 },
            { k: "P2 強化預測職能", p: 8 },
            { k: "P3 嵌入策略主導", p: 9 }
        ]
    },
    {
        t: "驗收成功指標", c: [
            { k: "產能回收與效率", p: 7 },
            { k: "繼任人才庫開發", p: 8 },
            { k: "技能再造 適配度", p: 4 }
        ]
    }
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

    // L1 Node (Level 2) - L2=16pt
    sMM.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: nx, y: ny, w: 1.8, h: 0.5, fill: { color: THEME.secondary }, rectRadius: 0.1 });
    sMM.addText(n.t, { x: nx, y: ny, w: 1.8, h: 0.5, color: THEME.white, bold: true, align: "center", fontFace: FONT_TITLE, fontSize: 16 });

    // L2 Nodes (Level 3 Child) - L3=14pt
    n.c.forEach((ch, j) => {
        let cx = nx + 2.1, cy = ny - 0.22 + (j * 0.42);
        // L1 -> L2
        sMM.addShape(pres.shapes.LINE, { x: nx + 1.8, y: ny + 0.25, w: 0.1, h: 0, line: { color: THEME.line, width: 1 } });
        sMM.addShape(pres.shapes.LINE, { x: nx + 1.9, y: Math.min(ny + 0.25, cy + 0.15), w: 0, h: Math.abs(ny + 0.25 - (cy + 0.15)), line: { color: THEME.line, width: 1 } });
        sMM.addShape(pres.shapes.LINE, { x: nx + 1.9, y: cy + 0.15, w: 0.2, h: 0, line: { color: THEME.line, width: 1 } });

        sMM.addText(ch.k, { x: cx, y: cy, w: 2.5, h: 0.35, color: THEME.text, fontSize: 14, fontFace: FONT_TITLE, valign: "middle" });

        // L3 -> L4 (Page Number Mapping) - L4=11pt
        let px = cx + 2.4;
        sMM.addShape(pres.shapes.LINE, { x: cx + 2.3, y: cy + 0.15, w: 0.1, h: 0, line: { color: THEME.accent, width: 0.8, dashType: "dash" } });
        sMM.addText(`P.${ch.p}`, { x: px, y: cy, w: 0.6, h: 0.35, color: THEME.accent, fontSize: 11, fontFace: FONT_BODY, valign: "middle", bold: true });
    });
});

const outP = path.join(SCRATCH_DIR, "output", "HRBP_AI_Transformation_Roman_v19.pptx");
pres.writeFile({ fileName: outP }).then(fn => {
    console.log(`Success: Zenith v19 (Roman Integrity) Generated at ${fn}`);
}).catch(err => console.error(err));
