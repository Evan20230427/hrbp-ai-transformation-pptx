const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

let pres = new pptxgen();
pres.layout = 'LAYOUT_16x9';
pres.author = 'Antigravity AI';
pres.title = 'HRBP AI 轉型最佳實務 - 終極巔峰版 v14';

// THEME: ZENITH CLARITY (V14 Refinement)
const THEME = {
    primary: "0F172A",    // Deep Navy
    secondary: "3B82F6",  // Azure Blue
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
    p1: path.join(IMG_DIR, "hrbp_roadmap_v13_simple_1773160725232.png"), // Reused logic
    p2: path.join(IMG_DIR, "hrbp_data_strategy_meeting_v10_1773159591017.png"),
    p3: path.join(IMG_DIR, "hrbp_transformation_journey_v10_1773159569827.png")
};

/**
 * BILINGUAL TERM MAP (CN <-> EN Notes)
 */
const BILINGUAL_NOTES = {
    "HRBP": "人力資源業務夥伴 (Human Resource Business Partner)",
    "STL": "策略人才領袖 (Strategic Talent Leader)",
    "AI 轉型": "AI 轉型 (AI Transformation)",
    "人力設計": "人力重新設計 (Workforce Redesign)",
    "倫理監測": "AI 倫理與偏見監測 (AI Ethics & Bias Monitoring)",
    "人機協作": "人機協作平衡 (Human-Machine Collaboration Balance)",
    "技能再造": "技能再造與重新部署 (Reskilling/Redeployment)",
    "自動化藍圖": "自動化藍圖 (Automation Roadmap)",
    "治理條款": "治理條款 (Governance Clauses)",
    "關鍵指標": "成功衡量指標 (Success Metrics)",
    "產能回收": "產能回收週期 (Cycle Time Recovery)",
    "員工參與度": "員工參與度 (Employee Engagement)"
};

/**
 * Enhanced Text Engine with Boundary Check & Bilingual Formatting.
 */
function renderContent(slide, lines, opts = {}) {
    // 1. BOUNDARY PROTECTION (Ensure x+w <= 10, y+h <= 5.625)
    let safeX = opts.x || 0.6;
    let safeY = opts.y || 1.4;
    let safeW = Math.min(opts.w || 6.2, 9.5 - safeX);
    let safeH = Math.min(opts.h || 3.4, 5.4 - safeY);

    const uniqueLines = [...new Set(lines)];
    const baseSize = opts.fontSize || 17;
    const smallSize = Math.max(11, baseSize - 5);
    const content = [];

    uniqueLines.forEach(line => {
        // Find technical terms to append notes implicitly if they appear
        let enhancedLine = line;
        // Check for bilingual terms (Simple inject if not present)
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
        x: safeX,
        y: safeY,
        w: safeW,
        h: safeH,
        lineSpacing: opts.spacing || 26,
        valign: "top"
    });
}

// MASTER SLIDE
pres.defineSlideMaster({
    title: 'ZENITH_V14',
    background: { color: THEME.white },
    objects: [
        { rect: { x: 0, y: 0, w: 0.1, h: "100%", fill: { color: THEME.secondary } } }
    ]
});

function applyHeader(slide, title, pageNum) {
    slide.addText(title, { x: 0.5, y: 0.4, w: 9, h: 0.6, fontSize: 28, bold: true, color: THEME.primary, fontFace: FONT_TITLE });
    slide.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.0, w: 2.5, h: 0.04, fill: { color: THEME.secondary } });
    if (pageNum > 0) {
        slide.addText(pageNum.toString(), { x: 9.3, y: 5.2, w: 0.5, h: 0.3, fontSize: 11, color: THEME.secondary, align: "right", fontFace: FONT_BODY });
    }
}

// --- SLIDES GENERATION (DYNAMIC LAYOUT) ---
let currentPage = 0;

// 1. Cover (Page 0, hidden)
let s1 = pres.addSlide();
s1.background = { color: THEME.white };
if (fs.existsSync(ASSETS.cover)) s1.addImage({ path: ASSETS.cover, x: 5.2, y: 0.8, w: 4.2, h: 4.2 });
s1.addText("重塑 HRBP 角色：\n引領企業 AI 轉型的策略實務", { x: 0.5, y: 1.8, w: 5.5, h: 1.5, fontSize: 36, bold: true, color: THEME.primary, fontFace: FONT_TITLE });
s1.addText("2026 年度終極巔峰版 | 雙語考證與動態佈局 (ZENITH v14)", { x: 0.5, y: 3.3, w: 5.5, h: 0.5, fontSize: 16, color: THEME.secondary, fontFace: FONT_TITLE });

// 2. Paradox (Layout: Left Text, Right Image)
currentPage++;
let s2 = pres.addSlide({ masterName: 'ZENITH_V14' });
applyHeader(s2, "核心挑戰：策略參與的空洞化", currentPage);
renderContent(s2, [
    "數據實例：僅 51% (Only 51%) 的領導者認同 HRBP 具備足夠策略影響力。",
    "行政束縛：工時被自動化可吸收的舊務佔據。",
    "生存威脅：必須在 AI 轉型浪潮中建立 STL 的核心地位。"
]);
if (fs.existsSync(ASSETS.paradox)) s2.addImage({ path: ASSETS.paradox, x: 7.0, y: 1.2, w: 2.5, h: 3.8, sizing: { type: "contain" } });

// 3. STL Role (Layout: Flipped - Right Text, Left Image)
currentPage++;
let s3 = pres.addSlide({ masterName: 'ZENITH_V14' });
applyHeader(s3, "未來定位：策略人才領袖 (STL)", currentPage);
if (fs.existsSync(ASSETS.human)) s3.addImage({ path: ASSETS.human, x: 0.5, y: 1.2, w: 3.8, h: 4.0, sizing: { type: "cover" } });
renderContent(s3, [
    "定位：主導 AI 轉型下的人員架構設計與優化。",
    "權責：引進領先人才洞察，主導高階策略對話。",
    "轉型：利用 產能回收 支持人才資本的二次投資。"
], { x: 4.6, w: 5.0, fontSize: 20 });

// 4. Role 1 (Layout: Top Split)
currentPage++;
let s4 = pres.addSlide({ masterName: 'ZENITH_V14' });
applyHeader(s4, "職責 1：人力設計 (Redesign)", currentPage);
renderContent(s4, [
    "定義隨 AI 演進的新型職能模型與職權架構。",
    "主導技能再造決策，確保人才適配度與技術領跑。"
], { y: 3.8, h: 1.5, w: 9, x: 0.6 });
if (fs.existsSync(ASSETS.role1)) s4.addImage({ path: ASSETS.role1, x: 0.5, y: 1.2, w: 9, h: 2.4, sizing: { type: "cover" } });

// 5. Role 2 (Layout: Grid Balanced)
currentPage++;
let s5 = pres.addSlide({ masterName: 'ZENITH_V14' });
applyHeader(s5, "職責 2：應對倫理監測 (Governance)", currentPage);
if (fs.existsSync(ASSETS.role2)) s5.addImage({ path: ASSETS.role2, x: 5.5, y: 1.2, w: 4, h: 4, sizing: { type: "cover" } });
renderContent(s5, [
    "監測人才決策中的算法公平性與透明度。",
    "建立負責任的人機協作倫理監督機制體系。"
], { x: 0.6, w: 4.8, y: 2.0 });

// 6. Role 3
currentPage++;
let s6 = pres.addSlide({ masterName: 'ZENITH_V14' });
applyHeader(s6, "職責 3：優化人機協作效率", currentPage);
renderContent(s6, [
    "設計平衡生產力與員工參與度的工作流路徑。",
    "在自動化推進中維護組織文化的完整性。"
], { x: 4.5, w: 5, y: 1.5 });
if (fs.existsSync(ASSETS.role3)) s6.addImage({ path: ASSETS.role3, x: 0.5, y: 1.5, w: 3.8, h: 3.5, sizing: { type: "cover" } });

// 7-9: Phases (Dynamic Layouts)
const PH = [
    { t: "P1 剔除行動：產能回收藍圖", c: ["定義策略重點區域", "建立完整之自動化藍圖與產能釋放標準"], img: ASSETS.p1, orient: "right" },
    { t: "P2 強化行動：AI 賦智高增值", c: ["更新職能模型，實現人才準備度透明化", "利用預測大數據強化領導對應能力"], img: ASSETS.p2, orient: "left" },
    { t: "P3 擴展開拓：新型 治理條款", c: ["啟動 STL Pods 小組領先核心策略決策", "在重大收購決策中嵌入 STL 考量機制"], img: ASSETS.p3, orient: "right" }
];

PH.forEach(p => {
    currentPage++;
    let s = pres.addSlide({ masterName: 'ZENITH_V14' });
    applyHeader(s, p.t, currentPage);
    if (p.orient === "right") {
        renderContent(s, p.c, { x: 0.6, w: 5.5, y: 1.8, fontSize: 20 });
        if (fs.existsSync(p.img)) s.addImage({ path: p.img, x: 6.5, y: 1.2, w: 3.2, h: 4, sizing: { type: "cover" } });
    } else {
        if (fs.existsSync(p.img)) s.addImage({ path: p.img, x: 0.5, y: 1.2, w: 3.2, h: 4, sizing: { type: "cover" } });
        renderContent(s, p.c, { x: 4.2, w: 5.4, y: 1.8, fontSize: 20 });
    }
});

// Final Slide: Mind Map (Pristine v14, Font 14, No Lines/Decorations, Pure title)
currentPage++;
let sMap = pres.addSlide(); // Independent slide
sMap.background = { color: THEME.white };
// Title + Page Number only
sMap.addText("全課精華：三層層級心智圖 (Peak v14)", { x: 0.5, y: 0.4, w: 8.5, h: 0.6, fontSize: 26, bold: true, color: THEME.primary, fontFace: FONT_TITLE });
sMap.addText(currentPage.toString(), { x: 9.3, y: 5.3, w: 0.5, h: 0.3, fontSize: 11, color: THEME.secondary, align: "right", fontFace: FONT_BODY });

const OX = 0.5, OY = 2.4;
const MAP = [
    { t: "現狀解析", c: ["策略參與(51%)不足", "行政作業佔據產能", "AI 自動化之衝擊"] },
    { t: "STL 定義力", c: ["人力設計重新校準", "倫理監測治理體系", "人機協調效能優化"] },
    { t: "轉型三階", c: ["P1 剔除/P2 強化階段", "P3 開拓新型治理權"] },
    { t: "關鍵指標", c: ["產能回收核心週期", "繼任人才庫準備率", "員工參與度滿意值"] }
];

// Draw logic with 14pt children and direct titles
sMap.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: OX, y: OY, w: 1.4, h: 0.6, fill: { color: THEME.primary }, rectRadius: 0.1 });
sMap.addText("HRBP\nAI 轉型", { x: OX, y: OY, w: 1.4, h: 0.6, color: THEME.white, bold: true, align: "center", fontFace: FONT_TITLE, fontSize: 12 });

MAP.forEach((n, i) => {
    let nx = OX + 2.0, ny = 1.0 + (i * 1.35);
    // Connections
    sMap.addShape(pres.shapes.LINE, { x: OX + 1.4, y: OY + 0.3, w: 0.2, h: 0, line: { color: THEME.secondary, width: 1.5 } });
    sMap.addShape(pres.shapes.LINE, { x: OX + 1.6, y: Math.min(OY + 0.3, ny + 0.25), w: 0, h: Math.abs(OY + 0.3 - (ny + 0.25)), line: { color: THEME.secondary, width: 1.5 } });
    sMap.addShape(pres.shapes.LINE, { x: OX + 1.6, y: ny + 0.25, w: 0.4, h: 0, line: { color: THEME.secondary, width: 1.5 } });

    sMap.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: nx, y: ny, w: 1.4, h: 0.5, fill: { color: THEME.secondary }, rectRadius: 0.1 });
    sMap.addText(n.t, { x: nx, y: ny, w: 1.4, h: 0.5, color: THEME.white, bold: true, align: "center", fontFace: FONT_TITLE, fontSize: 10 });

    n.c.forEach((ch, j) => {
        let cx = nx + 1.8, cy = ny - 0.25 + (j * 0.48);
        sMap.addShape(pres.shapes.LINE, { x: nx + 1.4, y: ny + 0.25, w: 0.1, h: 0, line: { color: THEME.line, width: 1 } });
        sMap.addShape(pres.shapes.LINE, { x: nx + 1.5, y: Math.min(ny + 0.25, cy + 0.15), w: 0, h: Math.abs(ny + 0.25 - (cy + 0.15)), line: { color: THEME.line, width: 1 } });
        sMap.addShape(pres.shapes.LINE, { x: nx + 1.5, y: cy + 0.15, w: 0.3, h: 0, line: { color: THEME.line, width: 1 } });

        // USER REQ: 3rd level font size = 14
        sMap.addText(ch, { x: cx, y: cy, w: 3.5, h: 0.35, color: THEME.text, fontSize: 14, fontFace: FONT_TITLE, valign: "middle" });
    });
});

const outP = path.join(SCRATCH_DIR, "output", "HRBP_AI_Transformation_Peak_v14.pptx");
pres.writeFile({ fileName: outP }).then(fn => {
    console.log(`Successfully generated Zenith v14 (Peak Edition) at ${fn}`);
}).catch(err => console.error(err));
