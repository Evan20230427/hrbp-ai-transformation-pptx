const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

let pres = new pptxgen();
pres.layout = 'LAYOUT_16x9';
pres.author = 'Antigravity AI';
pres.title = 'HRBP AI 轉型最佳實務 - 終極大師版 v13';

// THEME: ZEN CLARITY (V13 Refinement)
const THEME = {
    primary: "0F172A",
    secondary: "3B82F6",
    text: "1E293B",
    white: "FFFFFF",
    line: "CBD5E1"
};

const FONT_TITLE = "Microsoft JhengHei";
const FONT_BODY = "Arial";

// Resource Management
const IMG_DIR = "C:\\Users\\TW-Evan.Chen\\.gemini\\antigravity\\brain\\cb95dffe-33bd-4e40-a98b-feaff376ea1a";
const SCRATCH_DIR = "C:\\Users\\TW-Evan.Chen\\.gemini\\antigravity\\scratch\\pdf-xlsx-to-pptx";

// Using verified timestamps from previous successful tools
const ASSETS = {
    cover: path.join(IMG_DIR, "hrbp_simple_minimalist_cover_v12_1773160264391.png"),
    paradox: path.join(IMG_DIR, "hrbp_luxury_dashboard_v9_1773159099389.png"),
    human: path.join(IMG_DIR, "hrbp_professional_human_v12_simple_1773160281938.png"),
    role1: path.join(IMG_DIR, "hrbp_workforce_redesign_v13_simple_1773160725232.png"),
    role2: path.join(IMG_DIR, "hrbp_ai_ethics_v13_simple_1773160703050.png"),
    role3: path.join(IMG_DIR, "hrbp_professional_human_ai_v10_1773159552183.png"),
    p1: path.join(IMG_DIR, "hrbp_luxury_dashboard_v9_1773159099389.png"), // Fallback to Dashboard for P1
    p2: path.join(IMG_DIR, "hrbp_data_strategy_meeting_v10_1773159591017.png"),
    p3: path.join(IMG_DIR, "hrbp_transformation_journey_v10_1773159569827.png")
};

/**
 * Enhanced Text Engine: Uniform, Scaled English, Anti-Overlap.
 */
function renderContent(slide, lines, opts = {}) {
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
        w: opts.w || 6.2,
        h: opts.h || 3.4,
        lineSpacing: opts.spacing || 26,
        valign: "top"
    });
}

// MASTER SLIDE: CLARITY WITH PAGE NUMBER
pres.defineSlideMaster({
    title: 'CLARITY_MASTER_V13',
    background: { color: THEME.white },
    objects: [
        { rect: { x: 0, y: 0, w: 0.1, h: "100%", fill: { color: THEME.secondary } } }
    ]
});

function applyHeader(slide, title, pageNum) {
    slide.addText(title, { x: 0.5, y: 0.4, w: 8.5, h: 0.6, fontSize: 28, bold: true, color: THEME.primary, fontFace: FONT_TITLE });
    slide.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.0, w: 2.5, h: 0.04, fill: { color: THEME.secondary } });
    if (pageNum > 0) {
        slide.addText(pageNum.toString(), { x: 9.3, y: 5.2, w: 0.5, h: 0.3, fontSize: 11, color: THEME.secondary, align: "right", fontFace: FONT_BODY });
    }
}

// --- SLIDES GENERATION ---
let currentPage = 0;

// 1. Cover (Page 0, hidden)
let s1 = pres.addSlide();
s1.background = { color: THEME.white };
if (fs.existsSync(ASSETS.cover)) s1.addImage({ path: ASSETS.cover, x: 5, y: 0.5, w: 4.5, h: 4.5 });
s1.addText("重塑 HRBP 角色：\n引領企業 AI 轉型的策略實務", { x: 0.5, y: 1.8, w: 5.5, h: 1.5, fontSize: 36, bold: true, color: THEME.primary, fontFace: FONT_TITLE });
s1.addText("2026 年度終極大師版 | 全量插圖與對齊 (ULTIMA v13)", { x: 0.5, y: 3.3, w: 5.5, h: 0.5, fontSize: 16, color: THEME.secondary, fontFace: FONT_TITLE });

// 2. Paradox
currentPage++;
let s2 = pres.addSlide({ masterName: 'CLARITY_MASTER_V13' });
applyHeader(s2, "核心挑戰：策略參與的空洞化", currentPage);
renderContent(s2, [
    "數據實例：僅 51% (Only 51%) 的領導者認同 HRBP 具備足夠策略影響力。",
    "行政束縛：大量工時被 AI 正在取代的「事務性舊務」佔據 (Transactional Workload)。",
    "生存威脅：角色價值認知鴻溝導致企業決策層對 HR 支援的邊緣化。"
]);
if (fs.existsSync(ASSETS.paradox)) s2.addImage({ path: ASSETS.paradox, x: 7.0, y: 1.2, w: 2.6, h: 3.8, sizing: { type: "contain" } });

// 3. STL Role
currentPage++;
let s3 = pres.addSlide({ masterName: 'CLARITY_MASTER_V13' });
applyHeader(s3, "未來定位：策略人才領袖 (STL)", currentPage);
renderContent(s3, [
    "目標：主導人工智慧轉型中的人員設計 (Consultative Transformation)。",
    "權責：從解释者進化為「參與轉型對話者」。",
    "進化：透過 AI 賦能回收產能，轉向高邊際效應的策略投資。"
], { fontSize: 20 });
if (fs.existsSync(ASSETS.human)) s3.addImage({ path: ASSETS.human, x: 7.0, y: 1.2, w: 2.6, h: 3.8, sizing: { type: "cover" } });

// 4. Role 1
currentPage++;
let s4 = pres.addSlide({ masterName: 'CLARITY_MASTER_V13' });
applyHeader(s4, "職職 1：人力重新設計 (Workforce Redesign)", currentPage);
renderContent(s4, [
    "主導隨 AI 改變的職能重塑決策。",
    "優化人才培訓與部署 (Reskilling/Deployment)。",
    "確保組織人才矩陣與技術力路徑對齊。"
]);
if (fs.existsSync(ASSETS.role1)) s4.addImage({ path: ASSETS.role1, x: 6.8, y: 1.5, w: 2.8, h: 3.2, sizing: { type: "cover" } });

// 5. Role 2
currentPage++;
let s5 = pres.addSlide({ masterName: 'CLARITY_MASTER_V13' });
applyHeader(s5, "職職 2：應對 AI 倫理與偏見 (Addressing Bias)", currentPage);
renderContent(s5, [
    "監測人才決策中的算法偏見。",
    "確保 AI 洞察之透明度與企業公平倫理邊界。",
    "建立負責任的數據化決策治理框架。"
]);
if (fs.existsSync(ASSETS.role2)) s5.addImage({ path: ASSETS.role2, x: 6.8, y: 1.5, w: 2.8, h: 3.2, sizing: { type: "cover" } });

// 6. Role 3
currentPage++;
let s6 = pres.addSlide({ masterName: 'CLARITY_MASTER_V13' });
applyHeader(s6, "職職 3：優化人機協作效率", currentPage);
renderContent(s6, [
    "設計具生產力的、人機平衡之工作流程。",
    "在提升產出時維護員工認同感與體驗。",
    "引領組織心理契約之數位化轉向。"
]);
if (fs.existsSync(ASSETS.role3)) s6.addImage({ path: ASSETS.role3, x: 6.8, y: 1.5, w: 2.8, h: 3.2, sizing: { type: "cover" } });

// Phase Slides
const PHASES = [
    { t: "P1 剔除行動：移除行政負擔", c: ["精確定義策略優先級 (Strategic Priority)", "建立 1-2 年自動化路線圖 (Roadmap)"], img: ASSETS.p1 },
    { t: "P2 強化行動：AI 賦能高價值", c: ["更新模型使 AI 準備度透明化", "利用預測大數據強化領導對話深度 (Insights)"], img: ASSETS.p2 },
    { t: "P3 擴展開拓：新型策略領域", c: ["試行 STL Pods 小組領航變革計畫", "在重大決策中嵌入強力 STL 治理條款"], img: ASSETS.p3 }
];

PHASES.forEach(p => {
    currentPage++;
    let s = pres.addSlide({ masterName: 'CLARITY_MASTER_V13' });
    applyHeader(s, p.t, currentPage);
    renderContent(s, p.c, { fontSize: 20 });
    if (fs.existsSync(p.img)) s.addImage({ path: p.img, x: 6.8, y: 1.5, w: 2.8, h: 3.2, sizing: { type: "cover" } });
});

// Final Slide: Mind Map (User Req: No other design/lines, font 14)
currentPage++;
let sMap = pres.addSlide(); // Independent slide, no master lines
sMap.background = { color: THEME.white };
// Title only
sMap.addText("全課精華：三層層級心智圖 (Master Edition v13)", { x: 0.5, y: 0.4, w: 8.5, h: 0.6, fontSize: 26, bold: true, color: THEME.primary, fontFace: FONT_TITLE });
sMap.addText(currentPage.toString(), { x: 9.3, y: 5.2, w: 0.5, h: 0.3, fontSize: 11, color: THEME.secondary, align: "right", fontFace: FONT_BODY });

const OX = 0.5, OY = 2.4;
const MAP = [
    { t: "現狀解析", c: ["策略參與(51%)不足", "行政作業負擔過重", "AI 自動化之威脅"] },
    { t: "STL 定義", c: ["人力設計重新優化", "倫理監測治理體系", "人機協調效能提升"] },
    { t: "轉型三階", c: ["P1 剔除/P2 強化階段", "P3 開拓新型競爭力"] },
    { t: "關鍵指標", c: ["產能回收核心週期", "繼任人才庫準備率", "轉型滿意度回饋"] }
];

sMap.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: OX, y: OY, w: 1.4, h: 0.6, fill: { color: THEME.primary } });
sMap.addText("HRBP\nAI 轉型", { x: OX, y: OY, w: 1.4, h: 0.6, color: THEME.white, bold: true, align: "center", fontFace: FONT_TITLE, fontSize: 12 });

MAP.forEach((n, i) => {
    let nx = OX + 2.0, ny = 1.0 + (i * 1.35);
    // Logic Lines
    sMap.addShape(pres.shapes.LINE, { x: OX + 1.4, y: OY + 0.3, w: 0.2, h: 0, line: { color: THEME.secondary, width: 2 } });
    sMap.addShape(pres.shapes.LINE, { x: OX + 1.6, y: Math.min(OY + 0.3, ny + 0.25), w: 0, h: Math.abs(OY + 0.3 - (ny + 0.25)), line: { color: THEME.secondary, width: 2 } });
    sMap.addShape(pres.shapes.LINE, { x: OX + 1.6, y: ny + 0.25, w: 0.4, h: 0, line: { color: THEME.secondary, width: 2 } });

    sMap.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: nx, y: ny, w: 1.4, h: 0.5, fill: { color: THEME.secondary } });
    sMap.addText(n.t, { x: nx, y: ny, w: 1.4, h: 0.5, color: THEME.white, bold: true, align: "center", fontFace: FONT_TITLE, fontSize: 11 });

    n.c.forEach((ch, j) => {
        let cx = nx + 1.8, cy = ny - 0.25 + (j * 0.48);
        // Link to nodes
        sMap.addShape(pres.shapes.LINE, { x: nx + 1.4, y: ny + 0.25, w: 0.15, h: 0, line: { color: THEME.line, width: 1 } });
        sMap.addShape(pres.shapes.LINE, { x: nx + 1.55, y: Math.min(ny + 0.25, cy + 0.15), w: 0, h: Math.abs(ny + 0.25 - (cy + 0.15)), line: { color: THEME.line, width: 1 } });
        sMap.addShape(pres.shapes.LINE, { x: nx + 1.55, y: cy + 0.15, w: 0.25, h: 0, line: { color: THEME.line, width: 1 } });

        // USER REQ: 3rd level font size = 14
        sMap.addText(ch, { x: cx, y: cy, w: 3.5, h: 0.35, color: THEME.text, fontSize: 14, fontFace: FONT_TITLE, valign: "middle" });
    });
});

const outP = path.join(SCRATCH_DIR, "output", "HRBP_AI_Transformation_Final_v13.pptx");
pres.writeFile({ fileName: outP }).then(fn => {
    console.log(`Successfully generated Ultima v13 at ${fn}`);
}).catch(err => console.error(err));
