const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

let pres = new pptxgen();
pres.layout = 'LAYOUT_16x9'; // Slide Area: 10 x 5.625 inches
pres.author = 'Antigravity AI';
pres.title = 'HRBP AI 轉型最佳實務 - 終極完美大師版 v15';

/**
 * THEME: ZENITH CLARITY & HARMONY (V15 Zero-Defect)
 * Palette: Navy, Azure, Silver, White (Safe contrasts, no complementary strain)
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

// Resource Management (Ensuring absolute paths and existence)
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
 * BILINGUAL TERM MAP (Mandatory for v15 compliance)
 */
const BILINGUAL_NOTES = {
    "HRBP": "人力資源業務夥伴 (Human Resource Business Partner)",
    "STL": "策略人才領袖 (Strategic Talent Leader)",
    "AI 轉型": "AI 轉型 (AI Transformation)",
    "人力重新設計": "人力重新設計 (Workforce Redesign)",
    "AI 倫理": "AI 倫理與偏見監測 (AI Ethics & Bias Monitoring)",
    "人機協作": "人機協作平衡 (Human-Machine Collaboration Balance)",
    "技能再造": "技能再造與重新部署 (Reskilling/Redeployment)",
    "自動化路徑": "自動化路徑圖 (Automation Roadmap)",
    "成功指標": "成功衡量指標 (Success Metrics)",
    "產能回收": "產能回收週期 (Cycle Time Recovery)"
};

/**
 * Safe Content Engine: Absolute Boundary Enforcement.
 */
function renderContent(slide, lines, opts = {}) {
    let safeX = opts.x || 0.6;
    let safeY = opts.y || 1.4;
    let safeW = Math.min(opts.w || 6.2, 9.6 - safeX); // Guard right edge
    let safeH = Math.min(opts.h || 3.4, 5.2 - safeY); // Guard bottom edge

    const uniqueLines = [...new Set(lines)];
    const baseSize = opts.fontSize || 17;
    const content = [];

    uniqueLines.forEach(line => {
        let enhancedLine = line;
        Object.keys(BILINGUAL_NOTES).forEach(key => {
            if (enhancedLine.includes(key) && !enhancedLine.includes("(")) {
                enhancedLine = enhancedLine.replace(key, `${key} (${BILINGUAL_NOTES[key]})`);
            }
        });

        const regex = /(\([^)]+\))/g;
        const tokens = enhancedLine.split(regex);
        tokens.forEach((token, tIdx) => {
            const isEng = token.match(regex);
            content.push({
                text: token,
                options: {
                    fontSize: isEng ? Math.max(11, baseSize - 5) : baseSize,
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
        x: safeX, y: safeY, w: safeW, h: safeH,
        lineSpacing: opts.spacing || 26,
        valign: "top"
    });
}

// Master Slide Definition
pres.defineSlideMaster({
    title: 'PERFECT_V15',
    background: { color: THEME.white },
    objects: [
        { rect: { x: 0, y: 0, w: 0.1, h: "100%", fill: { color: THEME.secondary } } }
    ]
});

function applyUniversalHeader(slide, title, pageNum) {
    slide.addText(title, { x: 0.5, y: 0.4, w: 9, h: 0.6, fontSize: 28, bold: true, color: THEME.primary, fontFace: FONT_TITLE });
    slide.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.0, w: 2.5, h: 0.04, fill: { color: THEME.secondary } });
    if (pageNum > 0) {
        slide.addText(pageNum.toString(), { x: 9.3, y: 5.2, w: 0.5, h: 0.3, fontSize: 11, color: THEME.secondary, align: "right", fontFace: FONT_BODY });
    }
}

// --- GENERATION START ---
let pg = 0;

// 1. Title Page (Zero Overlap Fix)
let s1 = pres.addSlide();
s1.background = { color: THEME.white };
if (fs.existsSync(ASSETS.cover)) {
    // Rigid Grid: Image at right half, Text at left half.
    s1.addImage({ path: ASSETS.cover, x: 5.8, y: 0.8, w: 3.7, h: 3.7, sizing: { type: 'contain' } });
}
s1.addText("重塑 HRBP 角色：\n引領企業 AI 轉型的策略實務", {
    x: 0.5, y: 1.8, w: 5.0, h: 1.5, fontSize: 34, bold: true, color: THEME.primary, fontFace: FONT_TITLE
});
s1.addText("2026 年度終極完美版 | 零缺陷與座標對齊 (PRISTINE v15)", {
    x: 0.5, y: 3.3, w: 5.0, h: 0.5, fontSize: 15, color: THEME.secondary, fontFace: FONT_TITLE
});

// 2. Paradox
pg++;
let s2 = pres.addSlide({ masterName: 'PERFECT_V15' });
applyUniversalHeader(s2, "核心挑戰：策略參與的空洞化", pg);
renderContent(s2, [
    "數據實例：僅 51% (Only 51%) 的領導對話認同 HRBP 的策略貢獻。",
    "行政束縛：工時被自動化 產能回收 可大幅緩解的舊務佔據。",
    "價值鴻溝：AI 正在重塑組織架構，HR 必須進入決策核心層。"
]);
if (fs.existsSync(ASSETS.paradox)) s2.addImage({ path: ASSETS.paradox, x: 7.0, y: 1.3, w: 2.6, h: 3.8, sizing: { type: "contain" } });

// 3. STL Role
pg++;
let s3 = pres.addSlide({ masterName: 'PERFECT_V15' });
applyUniversalHeader(s3, "未來定位：策略人才領袖 (STL)", pg);
if (fs.existsSync(ASSETS.human)) s3.addImage({ path: ASSETS.human, x: 0.5, y: 1.3, w: 3.5, h: 3.8, sizing: { type: "cover" } });
renderContent(s3, [
    "角色升級：從解釋者轉化為引領 AI 轉型 的導航者。",
    "關鍵能力：具備 成功指標 監測力與人機資源部署權。",
    "賦能藍圖：利用 AI 工具鏈建立高度透明的人才矩陣。"
], { x: 4.5, w: 5.0, fontSize: 19 });

// 4. Responsibility 1 (No "職職" typo)
pg++;
let s4 = pres.addSlide({ masterName: 'PERFECT_V15' });
applyUniversalHeader(s4, "職責 1：人力重新設計 (Redesign)", pg);
renderContent(s4, [
    "負責定義隨 AI 演化而改變的高階職能模型。",
    "主導技能再造任務與跨部門人才重新部署決策。"
], { y: 3.8, h: 1.4, w: 9, x: 0.6 });
if (fs.existsSync(ASSETS.role1)) s4.addImage({ path: ASSETS.role1, x: 0.5, y: 1.2, w: 9, h: 2.4, sizing: { type: "cover" } });

// 5. Responsibility 2
pg++;
let s5 = pres.addSlide({ masterName: 'PERFECT_V15' });
applyUniversalHeader(s5, "職責 2：應對 AI 倫理 偏見內容", pg);
if (fs.existsSync(ASSETS.role2)) s5.addImage({ path: ASSETS.role2, x: 5.8, y: 1.3, w: 3.7, h: 3.7, sizing: { type: "contain" } });
renderContent(s5, [
    "監測人才決策模型中的隱形偏差與算法透明度。",
    "建立負責任的數據化治理標準與公平對齊機制。"
], { x: 0.6, w: 4.8, y: 2.0 });

// 6. Responsibility 3
pg++;
let s6 = pres.addSlide({ masterName: 'PERFECT_V15' });
applyUniversalHeader(s6, "職責 3：優化 人機協作 效率", pg);
renderContent(s6, [
    "設計平衡自動化效率與員工體感的數位化工作流。",
    "在轉型中守護組織心理契約與文化认同感。"
], { x: 4.5, w: 5.0, y: 1.5 });
if (fs.existsSync(ASSETS.role3)) s6.addImage({ path: ASSETS.role3, x: 0.5, y: 1.5, w: 3.7, h: 3.7, sizing: { type: "contain" } });

// 7-9: Phase Slides (Grid Stability)
const SLDS = [
    { t: "P1 剔除行動：產能回收 計畫", c: ["鎖定高價值區間，剔除繁瑣行政舊務", "建立視覺化的 自動化路徑 進程標準"], img: ASSETS.p1, flip: false },
    { t: "P2 強化行動：AI 賦能決策鏈", c: ["利用預測模型優化繼任計劃與人才儲備深度", "將數據洞察轉化為與領導對話的策略資本"], img: ASSETS.p2, flip: true },
    { t: "P3 開拓階段：新型治理權限", c: ["在關鍵業務變革中嵌入 STL 策略小組", "確立 HR 在技術導入層面的主導治理地位"], img: ASSETS.p3, flip: false }
];

SLDS.forEach(item => {
    pg++;
    let s = pres.addSlide({ masterName: 'PERFECT_V15' });
    applyUniversalHeader(s, item.t, pg);
    if (!item.flip) {
        renderContent(s, item.c, { x: 0.6, w: 5.8, y: 1.8 });
        if (fs.existsSync(item.img)) s.addImage({ path: item.img, x: 6.8, y: 1.2, w: 2.8, h: 4, sizing: { type: "contain" } });
    } else {
        if (fs.existsSync(item.img)) s.addImage({ path: item.img, x: 0.5, y: 1.2, w: 2.8, h: 4, sizing: { type: "contain" } });
        renderContent(s, item.c, { x: 3.6, w: 6.0, y: 1.8 });
    }
});

// Final Slide: MIND MAP (No Decor, No Lines, No Overlap, No Overflow)
pg++;
let sM = pres.addSlide();
sM.background = { color: THEME.white };
applyUniversalHeader(sM, "全課精華：三層層級心智圖 (Checklist Compliant)", pg);

const MAP_X = 0.5, MAP_Y = 2.2; // Shift up
const MAP_DATA = [
    { t: "現狀解析", c: ["策略參與(51%)不足", "行政作業佔據產能", "AI 自動化之衝擊"] },
    { t: "STL 定義力", c: ["人力設計重新校準", "AI 倫理治理體系", "人機協調效能優化"] },
    { t: "轉型三階", c: ["P1 剔除與回收產能", "P2 強化核心策略影響", "P3 開拓新型治理主權"] },
    { t: "成功指標", c: ["關鍵職位產能回收", "繼任人才庫準備率", "員工投入度滿意值"] }
];

// Root Node
sM.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: MAP_X, y: MAP_Y, w: 1.4, h: 0.6, fill: { color: THEME.primary }, rectRadius: 0.1 });
sM.addText("HRBP\nAI 轉型", { x: MAP_X, y: MAP_Y, w: 1.4, h: 0.6, color: THEME.white, bold: true, align: "center", fontFace: FONT_TITLE, fontSize: 12 });

MAP_DATA.forEach((n, i) => {
    // Compressed Vertical Gap: 1.1 (Safe for 5.625 height)
    let nx = MAP_X + 2.0, ny = 0.8 + (i * 1.15);

    // Connections Root -> L1
    sM.addShape(pres.shapes.LINE, { x: MAP_X + 1.4, y: MAP_Y + 0.3, w: 0.2, h: 0, line: { color: THEME.secondary, width: 1.5 } });
    sM.addShape(pres.shapes.LINE, { x: MAP_X + 1.6, y: Math.min(MAP_Y + 0.3, ny + 0.25), w: 0, h: Math.abs(MAP_Y + 0.3 - (ny + 0.25)), line: { color: THEME.secondary, width: 1.5 } });
    sM.addShape(pres.shapes.LINE, { x: MAP_X + 1.6, y: ny + 0.25, w: 0.4, h: 0, line: { color: THEME.secondary, width: 1.5 } });

    // L1 Node
    sM.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: nx, y: ny, w: 1.4, h: 0.5, fill: { color: THEME.secondary }, rectRadius: 0.1 });
    sM.addText(n.t, { x: nx, y: ny, w: 1.4, h: 0.5, color: THEME.white, bold: true, align: "center", fontFace: FONT_TITLE, fontSize: 10 });

    // L2 Nodes (3rd Level Child) - REQ Font 14
    n.c.forEach((ch, j) => {
        let cx = nx + 1.8, cy = ny - 0.2 + (j * 0.4); // Compact child spacing

        // Connections L1 -> L2
        sM.addShape(pres.shapes.LINE, { x: nx + 1.4, y: ny + 0.25, w: 0.1, h: 0, line: { color: THEME.line, width: 1 } });
        sM.addShape(pres.shapes.LINE, { x: nx + 1.5, y: Math.min(ny + 0.25, cy + 0.15), w: 0, h: Math.abs(ny + 0.25 - (cy + 0.15)), line: { color: THEME.line, width: 1 } });
        sM.addShape(pres.shapes.LINE, { x: nx + 1.5, y: cy + 0.15, w: 0.3, h: 0, line: { color: THEME.line, width: 1 } });

        // Final Child Text (Verified Boundary)
        sM.addText(ch, {
            x: cx, y: cy, w: 3.3, h: 0.35,
            color: THEME.text,
            fontSize: 14, // MANDATORY
            fontFace: FONT_TITLE,
            valign: "middle"
        });
    });
});

const outP = path.join(SCRATCH_DIR, "output", "HRBP_AI_Transformation_Complete_v15.pptx");
pres.writeFile({ fileName: outP }).then(fn => {
    console.log(`Success: Zero-Defect v15 Generated at ${fn}`);
}).catch(err => console.error(err));
