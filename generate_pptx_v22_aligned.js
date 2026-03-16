const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

let pres = new pptxgen();
pres.layout = 'LAYOUT_16x9';
pres.author = 'Antigravity AI';
pres.title = 'HRBP AI 轉型最佳實務 - v22 最終版';

const THEME = {
    primary: "1E293B",    
    secondary: "3B82F6",  
    text: "334155",       
    white: "FFFFFF",
    line: "E2E8F0",       
    accent: "64748B",     
    highlight: "FFFF00"   
};

const FONT_TITLE = "Microsoft JhengHei";
const FONT_BODY = "Arial";

const BRAIN_DIR = "C:\\Users\\TW-Evan.Chen\\.gemini\\antigravity\\brain\\c7d2daf0-cf50-4ec0-af51-75cb91a0c238";
const SCRATCH_DIR = "C:\\Users\\TW-Evan.Chen\\.gemini\\antigravity\\scratch\\pdf-xlsx-to-pptx";

const ASSETS = {
    cover: path.join(BRAIN_DIR, "hrbp_v21_cover_1773679946141.png"),
    stl: path.join(BRAIN_DIR, "hrbp_v21_stl_role_1773679960676.png"),
    ethics: path.join(BRAIN_DIR, "hrbp_v21_ethics_governance_1773679974929.png"),
    data: path.join(BRAIN_DIR, "hrbp_v21_data_insight_1773680231164.png"),
    pm: path.join(BRAIN_DIR, "hrbp_v21_product_manager_1773680252018.png"),
    roadmap: path.join(BRAIN_DIR, "hrbp_v21_roadmap_roman_1773680269997.png"),
    // Statues
    shock: path.join(BRAIN_DIR, "hrbp_v22_statue_shock_1773680596901.png"),
    laugh: path.join(BRAIN_DIR, "hrbp_v22_statue_laugh_1773680611632.png"),
    think: path.join(BRAIN_DIR, "hrbp_v22_statue_think_1773680624839.png"),
    angry: path.join(BRAIN_DIR, "hrbp_v22_statue_angry_1773680640224.png"),
    sad: path.join(BRAIN_DIR, "hrbp_v22_statue_sad_1773680656123.png"),
    wow: path.join(BRAIN_DIR, "hrbp_v22_statue_wow_1773680679782.png"),
    bored: path.join(BRAIN_DIR, "hrbp_v22_statue_bored_1773680700394.png")
};

const HIGHLIGHT_LIST = ["策略參與", "挑戰與契機", "行政事務性束縛", "AI 轉型", "策略人才領袖", "STL", "人力重新設計", "AI 倫理治理", "人機協作效率", "產品經理", "使命守護者", "生存者", "產能回收", "繼任人才庫", "最後一公里"];

/**
 * 佈局管理器 (嚴格不重疊規則)
 * Grid: 10x10 units
 * Safe Margin: 0.2 inches
 */
const LAYOUT = {
    header: { x: 0.5, y: 0.4, w: 9.0, h: 0.6 },
    line: { x: 0.5, y: 1.1, w: 1.5, h: 0.05 },
    main: {
        split: {
            leftImg: { x: 0.5, y: 1.4, w: 4.0, h: 3.8 },
            rightText: { x: 4.7, y: 1.4, w: 4.8, h: 3.8 }
        },
        fullText: { x: 0.6, y: 1.4, w: 8.8, h: 3.8 }
    },
    footer: { x: 8.5, y: 5.2, w: 1.2, h: 0.3 }
};

function applyHeader(slide, title, pageNum) {
    slide.addText(title, { x: LAYOUT.header.x, y: LAYOUT.header.y, w: LAYOUT.header.w, h: LAYOUT.header.h, fontSize: 28, bold: true, color: THEME.primary, fontFace: FONT_TITLE });
    slide.addShape(pres.shapes.RECTANGLE, { x: LAYOUT.line.x, y: LAYOUT.line.y, w: LAYOUT.line.w, h: LAYOUT.line.h, fill: { color: THEME.secondary } });
    slide.addText(`MDL | PAGE ${pageNum}`, { x: LAYOUT.footer.x, y: LAYOUT.footer.y, w: LAYOUT.footer.w, h: LAYOUT.footer.h, fontSize: 10, color: THEME.secondary, align: "right", fontFace: FONT_BODY, bold: true });
}

function renderContent(slide, lines, opts = {}) {
    let safeX = opts.x || LAYOUT.main.fullText.x;
    let safeY = opts.y || LAYOUT.main.fullText.y;
    let safeW = opts.w || LAYOUT.main.fullText.w;
    let safeH = opts.h || LAYOUT.main.fullText.h;
    
    const content = [];
    lines.forEach(line => {
        let subParts = [line];
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
                    fontSize: opts.fontSize || 16,
                    color: isHighlight ? "#000000" : THEME.text,
                    fill: isHighlight ? THEME.highlight : null,
                    fontFace: sp.match(/[a-zA-Z]/) ? FONT_BODY : FONT_TITLE,
                    bold: isHighlight,
                    bullet: (sIdx === 0),
                    breakLine: (sIdx === subParts.length - 1)
                }
            });
        });
    });
    slide.addText(content, { x: safeX, y: safeY, w: safeW, h: safeH, lineSpacing: 22, valign: "top" });
}

pres.defineSlideMaster({
    title: 'MARMOREAL_V22',
    background: { color: THEME.white },
    objects: [
        { rect: { x: 0, y: 0, w: 0.1, h: "100%", fill: { color: THEME.secondary } } },
        { shape: pres.shapes.RECTANGLE, options: { x: 0.5, y: 5.5, w: 9.0, h: 0.1, fill: { color: THEME.line } } }
    ]
});

let pg = 0;

// 1. Cover
pg++;
let s1 = pres.addSlide();
if (fs.existsSync(ASSETS.cover)) s1.addImage({ path: ASSETS.cover, x: 0, y: 0, w: '100%', h: '100%' });
s1.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.5, w: 6.5, h: 2.5, fill: { color: 'FFFFFF', transparency: 10 } });
s1.addText("重塑 HRBP 角色：\n引領企業 AI 轉型的策略實務", { x: 0.8, y: 1.8, w: 6.0, h: 1.5, fontSize: 34, bold: true, color: THEME.primary, fontFace: FONT_TITLE });
s1.addText("羅馬美學 v22 | 擬真人雕像與三階對齊全量版", { x: 0.8, y: 3.3, w: 6.0, h: 0.5, fontSize: 16, color: THEME.secondary, fontFace: FONT_TITLE });

// 2. TOC
pg++;
let s2 = pres.addSlide({ masterName: 'MARMOREAL_V22' });
applyHeader(s2, "主題概覽與結構", 2);
renderContent(s2, [
    "第一部分：時代變革與角色重定義 (P.3-6)",
    "第二部分：策略人才領袖 (STL) 三位一體 (P.7-9)",
    "第三部分：核心任務與執行策略 (P.10-14)",
    "第四部分：實務攻略與成效指標 (P.15-19)",
    "結尾：三階架構與頁碼對齊圖 (P.20)"
]);

// 3. Shock
pg++;
let s3 = pres.addSlide({ masterName: 'MARMOREAL_V22' });
applyHeader(s3, "數據破局：策略斷層分析", 3);
if (fs.existsSync(ASSETS.shock)) s3.addImage({ path: ASSETS.shock, x: 0.5, y: 1.4, w: 4.0, h: 3.8, sizing: { type: 'contain' } });
renderContent(s3, [
    "數據揭謬：51% 的 HRBP 正在失去 策略參與 的席位。",
    "生存挑戰：如果無視技術斷層，HR 將徹底邊緣化。",
    "需求缺口：企業迫切需要具備 AI 轉型 視野的核心夥伴。"
], { x: 4.7, w: 4.8 });

// 4. Bored
pg++;
let s4 = pres.addSlide({ masterName: 'MARMOREAL_V22' });
applyHeader(s4, "現狀：行政事務的泥淖", 4);
if (fs.existsSync(ASSETS.bored)) s4.addImage({ path: ASSETS.bored, x: 0.5, y: 1.4, w: 4.0, h: 3.8, sizing: { type: 'contain' } });
renderContent(s4, [
    "窒息感：重複性勞動吸乾了所有創造力。",
    "行政事務性束縛 是轉型路上最沈重的枷鎖。",
    "目標：透過機器人流程自動化釋放 深度思考時間。"
], { x: 4.7, w: 4.8 });

// 5. Wow
pg++;
let s5 = pres.addSlide({ masterName: 'MARMOREAL_V22' });
applyHeader(s5, "奇點：AI 浪潮襲來", 5);
if (fs.existsSync(ASSETS.wow)) s5.addImage({ path: ASSETS.wow, x: 5.3, y: 1.4, w: 4.0, h: 3.8, sizing: { type: 'contain' } });
renderContent(s5, [
    "震撼衝擊：AI 不是工具，而是新的組織物種。",
    "轉型奇點：重新定義「人」在數位組織中的絕對價值。",
    "視野躍遷：看見自動化背後的 繼任人才庫 機會。"
], { x: 0.6, w: 4.5 });

// 6. Think
pg++;
let s6 = pres.addSlide({ masterName: 'MARMOREAL_V22' });
applyHeader(s6, "重塑：思考角色邊界", 6);
if (fs.existsSync(ASSETS.think)) s6.addImage({ path: ASSETS.think, x: 0.5, y: 1.4, w: 4.0, h: 3.8, sizing: { type: 'contain' } });
renderContent(s6, [
    "深度反思：我們的核心競爭力是否還在？",
    "角色重塑：從資源管理器向策略領航者過渡。",
    "戰略規劃：將 HR 邏輯嵌入業務技術的中樞。"
], { x: 4.7, w: 4.8 });

// 7. Laugh
pg++;
let s7 = pres.addSlide({ masterName: 'MARMOREAL_V22' });
applyHeader(s7, "生存者：情緒與偏見的平衡", 7);
if (fs.existsSync(ASSETS.laugh)) s7.addImage({ path: ASSETS.laugh, x: 5.3, y: 1.4, w: 4.0, h: 3.8, sizing: { type: 'contain' } });
renderContent(s7, [
    "最後一笑：機器有算力，但人類有共情力。",
    "生存者 角色：負責解決 AI 產生的冷酷與不公。",
    "核心價值：校準演算法偏見，找回 組織溫度。"
], { x: 0.6, w: 4.5 });

// 8. Think (PM)
pg++;
let s8 = pres.addSlide({ masterName: 'MARMOREAL_V22' });
applyHeader(s8, "產品經理：模組化組織設計", 8);
if (fs.existsSync(ASSETS.pm)) s8.addImage({ path: ASSETS.pm, x: 0.5, y: 1.4, w: 4.0, h: 3.8, sizing: { type: 'contain' } });
renderContent(s8, [
    "工程思維：將 HR 服務轉化為高效能「數位產品」。",
    "產品經理 定位：用數據回饋來迭代組織架構。",
    "核心：實現人才解決方案的 規模化與自動化。"
], { x: 4.7, w: 4.8 });

// 9. Wow (Guardian)
pg++;
let s9 = pres.addSlide({ masterName: 'MARMOREAL_V22' });
applyHeader(s9, "使命守護者：文化錨點", 9);
if (fs.existsSync(ASSETS.wow)) s9.addImage({ path: ASSETS.wow, x: 5.3, y: 1.4, w: 4.0, h: 3.8, sizing: { type: 'contain' } });
renderContent(s9, [
    "神聖使命：在數據洪流中守護 企業核心價值。",
    "使命守護者 擔當：確保技術不偏離人性與文化。",
    "長期視野：建立跨越技術週期的 組織韌性。"
], { x: 0.6, w: 4.5 });

// 10. Shock (Workforce)
pg++;
let s10 = pres.addSlide({ masterName: 'MARMOREAL_V22' });
applyHeader(s10, "任務 I：主導人力資源重新設計", 10);
if (fs.existsSync(ASSETS.shock)) s10.addImage({ path: ASSETS.shock, x: 0.5, y: 1.4, w: 4.0, h: 3.8, sizing: { type: 'contain' } });
renderContent(s10, [
    "結構震撼：傳統崗位正在大面積消失與重組。",
    "主導 人力重新設計：基於產能預測的動態調整。",
    "關鍵：找出 產能回收 後的高價值配置點。"
], { x: 4.7, w: 4.8 });

// 11. Laugh (Skill)
pg++;
let s11 = pres.addSlide({ masterName: 'MARMOREAL_V22' });
applyHeader(s11, "任務 II：技能再造實務", 11);
if (fs.existsSync(ASSETS.laugh)) s11.addImage({ path: ASSETS.laugh, x: 5.3, y: 1.4, w: 4.0, h: 3.8, sizing: { type: 'contain' } });
renderContent(s11, [
    "成長喜悅：舊技能凋零，新潛能綻放。",
    "Reskilling 核心：不僅是知識，更是 數位心態。 ",
    "轉型武器：建立具備持續學習能力的 繼任人才庫。"
], { x: 0.6, w: 4.5 });

// 12. Angry (Ethics)
pg++;
let s12 = pres.addSlide({ masterName: 'MARMOREAL_V22' });
applyHeader(s12, "任務 III：建立 AI 倫理憲法", 12);
if (fs.existsSync(ASSETS.angry)) s12.addImage({ path: ASSETS.angry, x: 0.5, y: 1.4, w: 4.0, h: 3.8, sizing: { type: 'contain' } });
renderContent(s12, [
    "正義之怒：絕不容忍數據歧視與暗箱操作。",
    "AI 倫理治理 核心：公開、公正、可追溯。",
    "監督職責：成為 算法公正性 的最終仲裁者。"
], { x: 4.7, w: 4.8 });

// 13. Shock (Audit)
pg++;
let s13 = pres.addSlide({ masterName: 'MARMOREAL_V22' });
applyHeader(s13, "算法審計：防患於未然", 13);
if (fs.existsSync(ASSETS.shock)) s13.addImage({ path: ASSETS.shock, x: 5.3, y: 1.4, w: 4.0, h: 3.8, sizing: { type: 'contain' } });
renderContent(s13, [
    "監測雷達：精確捕捉 AI 決策中的不當偏差。",
    "透明治理 實務：向全體員工揭示 技術黑盒。",
    "風險評估：衡量技術應用造成的品牌與法律風險。"
], { x: 0.6, w: 4.5 });

// 14. Laugh (Collab)
pg++;
let s14 = pres.addSlide({ masterName: 'MARMOREAL_V22' });
applyHeader(s14, "任務 IV：人機共生流程優化", 14);
if (fs.existsSync(ASSETS.laugh)) s14.addImage({ path: ASSETS.laugh, x: 0.5, y: 1.4, w: 4.0, h: 3.8, sizing: { type: 'contain' } });
renderContent(s14, [
    "共生魅力：技術賦能人類，人類指揮技術。",
    "優化 人機協作效率：設計動態回饋的協作流。",
    "KPI：如何顯著降低 決策週期時間。"
], { x: 4.7, w: 4.8 });

// 15. Roadmap
pg++;
let s15 = pres.addSlide({ masterName: 'MARMOREAL_V22' });
applyHeader(s15, "數位羅馬大道 (Roadmap)", 15);
if (fs.existsSync(ASSETS.roadmap)) s15.addImage({ path: ASSETS.roadmap, x: 0.5, y: 1.5, w: 9.0, h: 3.5, sizing: { type: 'cover' } });
s15.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 3.5, w: 9.0, h: 1.2, fill: { color: 'FFFFFF', transparency: 15 } });
s15.addText("啟動計畫 (P1) -> 核心轉型 (P2) -> 主導市場 (P3)", { x: 0.7, y: 3.8, w: 8.6, fontSize: 24, bold: true, color: THEME.secondary, align: "center", fontFace: FONT_TITLE });

// 16. Sad (P1)
pg++;
let s16 = pres.addSlide({ masterName: 'MARMOREAL_V22' });
applyHeader(s16, "P1：徹底回收過時產能", 16);
if (fs.existsSync(ASSETS.sad)) s16.addImage({ path: ASSETS.sad, x: 5.3, y: 1.4, w: 4.0, h: 3.8, sizing: { type: 'contain' } });
renderContent(s16, [
    "告別過去：切斷對低效 行政事務性束縛 的依戀。",
    "自動化路徑專案：透過 RPA 與 AI 回收 30% 工時。",
    "陣痛轉換：舊流程的終點即是 新 STL 的起點。"
], { x: 0.6, w: 4.5 });

// 17. Think (P2)
pg++;
let s17 = pres.addSlide({ masterName: 'MARMOREAL_V22' });
applyHeader(s17, "P2：繼任人才庫精準賦能", 17);
if (fs.existsSync(ASSETS.think)) s17.addImage({ path: ASSETS.think, x: 0.5, y: 1.4, w: 4.0, h: 3.8, sizing: { type: 'contain' } });
renderContent(s17, [
    "精準聚焦：利用預測模型鎖定未來 策略人才。",
    "職能放大：將回收產能精準投資於高階 STL 培訓。",
    "賦能轉型：建立組織內部的 創新實驗室。 "
], { x: 4.7, w: 4.8 });

// 18. Laugh (P3)
pg++;
let s18 = pres.addSlide({ masterName: 'MARMOREAL_V22' });
applyHeader(s18, "P3：確立全局策略主導權", 18);
if (fs.existsSync(ASSETS.laugh)) s18.addImage({ path: ASSETS.laugh, x: 5.3, y: 1.4, w: 4.0, h: 3.8, sizing: { type: 'contain' } });
renderContent(s18, [
    "凱旋回歸：HR 成為業務大腦的核心成員。",
    "主導位：在 業務決策 矩陣中佔據決定性一環。",
    "願景達成：實現 AI 時代 的人才驅動型組織。"
], { x: 0.6, w: 4.5 });

// 19. Think (KPI)
pg++;
let s19 = pres.addSlide({ masterName: 'MARMOREAL_V22' });
applyHeader(s19, "驗收：量化轉型成功", 19);
if (fs.existsSync(ASSETS.think)) s19.addImage({ path: ASSETS.think, x: 0.5, y: 1.4, w: 4.0, h: 3.8, sizing: { type: 'contain' } });
renderContent(s19, [
    "核心指標 I：行政產能回收 的實質比例 (目標 > 35%)。",
    "核心指標 II：STL 策略提案 的業務採納率。",
    "核心指標 III：員工對於 人機協作 滿意度的成長矩陣。"
], { x: 4.7, w: 4.8 });

// 20. Mindmap (3-Level Aligned)
pg++;
let s20 = pres.addSlide({ masterName: 'MARMOREAL_V22' });
applyHeader(s20, "主題架構圖：三階全量頁碼對齊", 20);

const M_DATA = [
    { t: "時代變革", c: [
        { k: "轉型背景", s: ["策略斷層 (P.3)", "事務束縛 (P.4)"] },
        { k: "轉型必然", s: ["AI 奇點 (P.5)", "角色重塑 (P.6)"] }
    ]},
    { t: "核心角色", c: [
        { k: "STL 定位", s: ["生存者 (P.7)", "產品經理 (P.8)", "守護者 (P.9)"] }
    ]},
    { t: "執行任務", c: [
        { k: "人力與技能", s: ["人力設計 (P.10)", "技能再造 (P.11)"] },
        { k: "治理與效率", s: ["倫理治理 (P.12)", "算法審計 (P.13)", "協作優化 (P.14)"] }
    ]},
    { t: "實務路徑", c: [
        { k: "發展階段", s: ["P1 產能回收 (P.16)", "P2 人才賦能 (P.17)", "P3 策略主導 (P.18)"] },
        { k: "成果衡量", s: ["轉型指標 (P.19)"] }
    ]}
];

// Render Logic for 3-Level Mindmap (Strict Separation)
const startX = 0.5, startY = 1.3;
const colWidth = 2.4;

M_DATA.forEach((l1, i) => {
    let x = startX + (i * colWidth);
    // L1 Node
    pres.shapes.ROUNDED_RECTANGLE; 
    s20.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: x, y: startY, w: 2.2, h: 0.4, fill: { color: THEME.primary }, rectRadius: 0.1 });
    s20.addText(l1.t, { x: x, y: startY, w: 2.2, h: 0.4, color: THEME.white, bold: true, align: "center", fontSize: 13, fontFace: FONT_TITLE });

    let currentY = startY + 0.6;
    l1.c.forEach(l2 => {
        // L2 Node
        s20.addShape(pres.shapes.RECTANGLE, { x: x + 0.1, y: currentY, w: 2.0, h: 0.35, fill: { color: THEME.secondary }, transparency: 20 });
        s20.addText(l2.k, { x: x + 0.1, y: currentY, w: 2.0, h: 0.35, color: THEME.white, bold: true, align: "center", fontSize: 11, fontFace: FONT_TITLE });
        
        currentY += 0.45;
        l2.s.forEach(l3 => {
            // L3 Node
            s20.addText("● " + l3, { x: x + 0.2, y: currentY, w: 2.0, h: 0.3, color: THEME.text, fontSize: 9, fontFace: FONT_TITLE });
            currentY += 0.35;
        });
        currentY += 0.2;
    });
});

const outP = path.join(SCRATCH_DIR, "output", "HRBP_AI_Transformation_v22_Final_Aligned.pptx");
pres.writeFile({ fileName: outP }).then(fn => {
    console.log(`Final Success: v22 Aligned PPTX Generated at ${fn}`);
}).catch(err => console.error(err));
