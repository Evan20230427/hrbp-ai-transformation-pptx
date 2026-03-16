const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

let pres = new pptxgen();
pres.layout = 'LAYOUT_16x9';
pres.author = 'Antigravity AI';
pres.title = 'HRBP AI 轉型最佳實務 - 大理石之生 v21';

/**
 * THEME: MARMOREAL LIFE (V21 - Integrity Edition)
 */
const THEME = {
    primary: "1E293B",    // Slate 800
    secondary: "3B82F6",  // Blue 500
    text: "334155",       // Slate 700
    white: "FFFFFF",
    line: "E2E8F0",       // Slate 200
    accent: "64748B",     // Slate 500
    highlight: "FFFF00"   // Yellow Background for focus points
};

const FONT_TITLE = "Microsoft JhengHei";
const FONT_BODY = "Arial";

// Resource Management
const BRAIN_DIR = "C:\\Users\\TW-Evan.Chen\\.gemini\\antigravity\\brain\\c7d2daf0-cf50-4ec0-af51-75cb91a0c238";
const SCRATCH_DIR = "C:\\Users\\TW-Evan.Chen\\.gemini\\antigravity\\scratch\\pdf-xlsx-to-pptx";

const ASSETS = {
    cover: path.join(BRAIN_DIR, "hrbp_v21_cover_1773679946141.png"),
    stl: path.join(BRAIN_DIR, "hrbp_v21_stl_role_1773679960676.png"),
    ethics: path.join(BRAIN_DIR, "hrbp_v21_ethics_governance_1773679974929.png")
};

const BILINGUAL_NOTES = {
    "HRBP": "人力資源業務夥伴",
    "STL": "策略人才領袖",
    "AI 轉型": "AI Transformation",
    "人力重新設計": "Workforce Redesign",
    "AI 倫理治理": "AI Ethics Governance",
    "人機協作效率": "Human-Machine Collaboration",
    "產能回收計畫": "Cycle Time Recovery",
    "自動化路徑專案": "Automation Roadmap",
    "行政事務性束縛": "Administrative Constraints"
};

const HIGHLIGHT_LIST = [
    "策略參與", "挑戰與契機", "行政事務性束縛", "AI 轉型",
    "策略人才領袖", "STL", "人力重新設計", "AI 倫理治理", "人機協作效率",
    "產能回收計畫", "自動化路徑專案", "生存者", "產品經理", "使命守護者"
];

function renderContent(slide, lines, opts = {}) {
    let safeX = opts.x || 0.6;
    let safeY = opts.y || 1.4;
    let safeW = Math.min(opts.w || 6.2, 9.6 - safeX);
    let safeH = Math.min(opts.h || 3.4, 5.2 - safeY);

    const baseSize = opts.fontSize || 17;
    const content = [];

    lines.forEach(line => {
        let currentLine = line;

        // Bilingual Injection
        Object.keys(BILINGUAL_NOTES).forEach(key => {
            if (currentLine.includes(key) && !currentLine.includes(`${key} (`)) {
                currentLine = currentLine.replace(key, `${key} (${BILINGUAL_NOTES[key]})`);
            }
        });

        // Split highlights and parentheses
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
    title: 'MARMOREAL_V21',
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
if (fs.existsSync(ASSETS.cover)) sTitle.addImage({ path: ASSETS.cover, x: 0, y: 0, w: '100%', h: '100%' });
sTitle.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.5, w: 6.0, h: 2.5, fill: { color: 'FFFFFF', transparency: 15 } });
sTitle.addText("重塑 HRBP 角色：\n引領企業 AI 轉型的策略實務", { x: 0.7, y: 1.8, w: 5.5, h: 1.5, fontSize: 36, bold: true, color: THEME.primary, fontFace: FONT_TITLE });
sTitle.addText("大理石之生 | 羅馬美學與四階全量對齊 (v21)", { x: 0.7, y: 3.3, w: 5.5, h: 0.5, fontSize: 16, color: THEME.secondary, fontFace: FONT_TITLE });

// 2. Overview
pg++;
let s2 = pres.addSlide({ masterName: 'MARMOREAL_V21' });
applyHeader(s2, "挑戰與契機：打破策略孤島", pg);
renderContent(s2, [
    "數據揭謬：51% 的 HRBP 在 策略參與 方面存在顯著斷層。",
    "困局解鎖：擺脫 行政事務性束縛，將產能釋放於高增值的治理領域。",
    "轉型奇點：AI 轉型 不僅是技術升級，更是 HR 角色邊界的重新詮釋。"
]);

// 3. STL Roles
pg++;
let s3 = pres.addSlide({ masterName: 'MARMOREAL_V21' });
applyHeader(s3, "未來角色：策略人才領袖 (STL)", pg);
renderContent(s3, [
    "生存者：人機協同中的「關鍵接口」，負責校準算法偏見與情緒疏導。",
    "產品經理：組織操作系統的設計者，將問題封裝成產品並用數據驅動增長。",
    "使命守護者：企業長期根基的構築者，專注於文化與 AI 倫理治理。 "
], { x: 4.2, w: 5.3 });
if (fs.existsSync(ASSETS.stl)) s3.addImage({ path: ASSETS.stl, x: 0.5, y: 1.2, w: 3.5, h: 4.2, sizing: { type: 'contain' } });

// 4. Workforce Redesign
pg++;
let s4 = pres.addSlide({ masterName: 'MARMOREAL_V21' });
applyHeader(s4, "任務 I：主導 人力重新設計", pg);
renderContent(s4, [
    "結構調整：基於產出回收預測，執行規模化的 人力重新設計。",
    "人才韌性：將 技能再造 視為企業核心競爭力的二次投資。",
    "職能建模：定義 AI 年代不可替代的人格化高價值領域。"
]);

// 5. Ethics
pg++;
let s5 = pres.addSlide({ masterName: 'MARMOREAL_V21' });
applyHeader(s5, "任務 II：建立 AI 倫理治理 憲法", pg);
renderContent(s5, [
    "偏見監測：建立嚴格的算法審計機制，確保 AI 倫理治理 的公正性。",
    "透明治理：確保每一項 AI 決策皆具備可追溯性與解釋性。",
    "品牌信任：透過負責、透明的技術應用，維護企業僱傭品牌。"
], { x: 0.6, w: 5.3 });
if (fs.existsSync(ASSETS.ethics)) s5.addImage({ path: ASSETS.ethics, x: 6.0, y: 1.2, w: 3.5, h: 4.2, sizing: { type: 'contain' } });

// 6. Collaboration
pg++;
let s6 = pres.addSlide({ masterName: 'MARMOREAL_V21' });
applyHeader(s6, "任務 III：優化 人機協作效率", pg);
renderContent(s6, [
    "共生流設計：設計具備高同步性的人機共生引擎，優化 人機協作效率。",
    "心理賦能：動態評估適應度，降低轉型期的技術焦慮。",
    "價值挖掘：挖掘 AI 產出中的非結構化洞察，輔助 STL 決策。"
]);

// 7. Roadmap
pg++;
let s7 = pres.addSlide({ masterName: 'MARMOREAL_V21' });
applyHeader(s7, "實務攻略：三階路徑", pg);
renderContent(s7, [
    "P1 (剔除)：制定 自動化路徑專案，徹底回收行政型 產能回收計畫。",
    "P2 (強化)：利用預測模型鎖定 繼任人才庫，精準投資高階人才。",
    "P3 (開拓)：在組織大腦中嵌入 STL 策略節點，全面主導 AI 轉型。"
]);

// 8. Conclusion
pg++;
let s8 = pres.addSlide();
s8.background = { color: THEME.primary };
s8.addText("大理石之生：\n雕琢 AI 時代的策略傑作", { x: 0.5, y: 2.0, w: 9, h: 1.5, fontSize: 40, bold: true, color: THEME.white, fontFace: FONT_TITLE, align: "center" });
s8.addText("即刻啟動產能回收與角色轉型", { x: 0.5, y: 3.5, w: 9, h: 0.5, fontSize: 20, color: THEME.secondary, fontFace: FONT_TITLE, align: "center" });

const outP = path.join(SCRATCH_DIR, "output", "HRBP_AI_Transformation_v21_Final.pptx");
pres.writeFile({ fileName: outP }).then(fn => {
    console.log(`Success: v21 PPTX Generated at ${fn}`);
}).catch(err => console.error(err));
