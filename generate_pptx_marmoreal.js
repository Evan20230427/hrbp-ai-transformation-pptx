const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

let pres = new pptxgen();
pres.layout = 'LAYOUT_16x9';
pres.author = 'Antigravity AI';
pres.title = 'HR 招募系統 Chatbox 專案進度更新 - 大理石之生';

/**
 * THEME: MARMOREAL LIFE (Digital Brutalism & Clarity)
 */
const THEME = {
    primary: "334155",     // Cool Gray (陰影處的冷灰色)
    secondary: "0EA5E9",   // Azure Blue (蔚藍色光點)
    text: "0F172A",        // Darkest Slate for main text
    white: "F8FAFC",       // Marble White (純淨的大理石白)
    line: "CBD5E1",        // Light cool gray
    accent: "64748B",
    highlight: "FFFF00"    // Yellow Background for highlights
};

const FONT_TITLE = "Microsoft JhengHei";
const FONT_BODY = "Arial";

const SCRATCH_DIR = "C:\\Users\\TW-Evan.Chen\\.gemini\\antigravity\\scratch\\pdf-xlsx-to-pptx";

const HIGHLIGHT_LIST = [
    "Phase 1", "Part 1", "AI 回答功能", "招募網站", "職前簡介資訊",
    "職缺媒合", "常見問題擴充", "API", "內部共識會議", "報價基準",
    "POC", "遴選階段", "極致", "艾卡拉"
];

const BILINGUAL_NOTES = {
    "POC": "概念驗證",
    "FAQ": "常見問題",
    "API": "應用程式介面",
    "Token": "權杖",
    "WIS": "工作面試系統"
};

/**
 * 核心渲染引擎：支援黃底高亮 + 階層式文字處理
 */
function renderContent(slide, lines, opts = {}) {
    let safeX = opts.x || 0.6;
    let safeY = opts.y || 1.4;
    let safeW = Math.min(opts.w || 8.4, 9.6 - safeX);
    let safeH = Math.min(opts.h || 3.4, 5.2 - safeY);

    const baseSize = opts.fontSize || 16;
    const content = [];

    lines.forEach((line, lineIdx) => {
        let currentLine = line;

        // 1. 注入雙語
        Object.keys(BILINGUAL_NOTES).forEach(key => {
            if (currentLine.includes(key) && !currentLine.includes(`${key} (`)) {
                currentLine = currentLine.replace(new RegExp(`\\b${key}\\b`, 'g'), `${key} (${BILINGUAL_NOTES[key]})`);
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
                        fontSize: Math.max(11, baseSize - 4),
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
        lineSpacing: opts.spacing || 20,
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
sTitle.addText("HR 招募系統 Chatbox：\n進度更新與 Phase 1 範圍確認", { x: 0.5, y: 1.8, w: 8.5, h: 1.5, fontSize: 36, bold: true, color: THEME.primary, fontFace: FONT_TITLE });
sTitle.addText("大理石之生 | 視覺哲學重新演繹 (v21)", { x: 0.5, y: 3.3, w: 5.5, h: 0.5, fontSize: 16, color: THEME.secondary, fontFace: FONT_TITLE });

// 2. 專案範圍釐清
pg++;
let s2 = pres.addSlide({ masterName: 'MARMOREAL_V20' });
applyHeader(s2, "專案範圍釐清：Phase 1 確切範疇", pg);
renderContent(s2, [
    "預算考量：今年預算為 100 萬，目標上線時間 2026/06。",
    "首階段範圍：確認包含 Phase 1 的 Part 1 (AI 回答功能) 作為報價基準。",
    "後續評估：待四家廠商提案與報價取得後，再視狀況評估是否延伸至 Part 2 或 Part 3。"
]);

// 3. 內容覆蓋範圍
pg++;
let s3 = pres.addSlide({ masterName: 'MARMOREAL_V20' });
applyHeader(s3, "雇主品牌與資訊涵蓋範圍", pg);
renderContent(s3, [
    "資訊來源：未來包含但不限於 招募網站 內容。",
    "擴充規劃：需可視招募活動及 常見問題擴充，分階段進行。",
    "新增範疇：除了招募網站，建議一併納入「職前簡介資訊」連結，以提供更完整的資訊覆蓋率以回應求職者對雇主品牌的詢問。"
]);

// 4. 職缺媒合與資料交換
pg++;
let s4 = pres.addSlide({ masterName: 'MARMOREAL_V20' });
applyHeader(s4, "資料庫與介接規劃：職缺媒合與 API", pg);
renderContent(s4, [
    "媒合資料來源：招募網站匯出 + 使用者聊天回覆 (如期望地區與職務)。",
    "API 介接需求：若需增加通勤資訊(如近捷運站)才需介接其他 API。",
    "系統整合：需與科技部確認現有「麥當勞招募管理系統」與「小尖兵 APP」流程，以利與 Chatbox 廠商討論 API 串接方式。",
    "後續行動：預計下週安排跨部門會議(包含 Rooson、Dart)進行流程操作說明。"
]);

// 5. 廠商報價與 POC 評估
pg++;
let s5 = pres.addSlide({ masterName: 'MARMOREAL_V20' });
applyHeader(s5, "廠商評估策略：報價結構與 POC 決策", pg);
renderContent(s5, [
    "POC 決策：基於開發成本與技術複雜度在可控範圍，為節省時程，建議直接進入 遴選階段，不執行 POC。",
    "報價廠商：包含 極致、艾卡拉 及原規劃的另外兩家，共計四家進行評估，不新增第五家。",
    "報價結構要求：廠商需完整涵蓋一次性開發費、雲地端建置費、維護費、更新費(人天)與 Token 費用。",
    "下一步：取得報價單、提案規格與成功範例後，將召開內部共識會議。"
]);

const M_DATA = [
    {
        t: "Phase 1 範圍", c: [
            { k: "聚焦 Part 1 (AI)", p: 2 },
            { k: "作為 報價基準", p: 2 },
            { k: "後延 Part 2 評估", p: 2 }
        ]
    },
    {
        t: "資訊與內容", c: [
            { k: "涵蓋 招募網站", p: 3 },
            { k: "納入 職前簡介資訊", p: 3 },
            { k: "支援 常見問題擴充", p: 3 }
        ]
    },
    {
        t: "系統與資料", c: [
            { k: "整合 職缺媒合 源", p: 4 },
            { k: "確認 API 介接需求", p: 4 },
            { k: "跨部門流程再釐清", p: 4 }
        ]
    },
    {
        t: "決策與時程", c: [
            { k: "免 POC 進 遴選階段", p: 5 },
            { k: "四家廠商(含極致等)報價", p: 5 },
            { k: "擬辦 內部共識會議", p: 5 }
        ]
    }
];

// Final Slide: 4-LEVEL MIND MAP
pg++;
let sMM = pres.addSlide();
applyHeader(sMM, "大理石之生：專案架構與四階全量對齊", pg);

const MX = 0.5, MY = 2.4;
sMM.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: MX, y: MY, w: 1.4, h: 0.6, fill: { color: THEME.primary }, rectRadius: 0.1 });
sMM.addText("HR Chatbox\n專案架構", { x: MX, y: MY, w: 1.4, h: 0.6, color: THEME.white, bold: true, align: "center", fontFace: FONT_TITLE, fontSize: 13 });

M_DATA.forEach((n, i) => {
    let nx = MX + 2.0, ny = 0.82 + (i * 1.15);
    sMM.addShape(pres.shapes.LINE, { x: MX + 1.4, y: MY + 0.3, w: 0.2, h: 0, line: { color: THEME.secondary, width: 1.5 } });
    sMM.addShape(pres.shapes.LINE, { x: MX + 1.6, y: Math.min(MY + 0.3, ny + 0.25), w: 0, h: Math.abs(MY + 0.3 - (ny + 0.25)), line: { color: THEME.secondary, width: 1.5 } });
    sMM.addShape(pres.shapes.LINE, { x: MX + 1.6, y: ny + 0.25, w: 0.4, h: 0, line: { color: THEME.secondary, width: 1.5 } });

    // L2 Node
    sMM.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: nx, y: ny, w: 1.8, h: 0.5, fill: { color: THEME.secondary }, rectRadius: 0.1 });
    sMM.addText(n.t, { x: nx, y: ny, w: 1.8, h: 0.5, color: THEME.white, bold: true, align: "center", fontFace: FONT_TITLE, fontSize: 16 });

    n.c.forEach((ch, j) => {
        let cx = nx + 2.1, cy = ny - 0.22 + (j * 0.42);
        sMM.addShape(pres.shapes.LINE, { x: nx + 1.8, y: ny + 0.25, w: 0.1, h: 0, line: { color: THEME.line, width: 1 } });
        sMM.addShape(pres.shapes.LINE, { x: nx + 1.9, y: Math.min(ny + 0.25, cy + 0.15), w: 0, h: Math.abs(ny + 0.25 - (cy + 0.15)), line: { color: THEME.line, width: 1 } });
        sMM.addShape(pres.shapes.LINE, { x: nx + 1.9, y: cy + 0.15, w: 0.2, h: 0, line: { color: THEME.line, width: 1 } });

        // L3 Node
        sMM.addText(ch.k, { x: cx, y: cy, w: 2.6, h: 0.35, color: THEME.text, fontSize: 14, fontFace: FONT_TITLE, valign: "middle" });

        // L4 Node (Page Reference, <12pt)
        let px = cx + 2.5;
        sMM.addShape(pres.shapes.LINE, { x: cx + 2.45, y: cy + 0.15, w: 0.05, h: 0, line: { color: THEME.accent, width: 0.5, dashType: "dash" } });
        sMM.addText(`P.${ch.p}`, { x: px, y: cy, w: 0.5, h: 0.35, color: THEME.accent, fontSize: 10, fontFace: FONT_BODY, valign: "middle", bold: true });
    });
});

const outP = path.join(SCRATCH_DIR, "output", "HR_Chatbox_Update_Marmoreal_V2.pptx");
pres.writeFile({ fileName: outP }).then(fn => {
    console.log(`Success: Generated ${fn}`);
}).catch(err => console.error(err));
