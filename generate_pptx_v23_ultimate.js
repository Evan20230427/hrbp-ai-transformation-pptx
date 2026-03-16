const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

let pres = new pptxgen();
pres.layout = 'LAYOUT_16x9';
pres.author = 'Antigravity AI';
pres.title = 'HRBP AI 轉型最佳實務 - v23 終極一致版';

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

// v23 極致版：22 張絕對唯一不重複圖資
const IMAGES = [
    path.join(BRAIN_DIR, "hrbp_v21_cover_1773679946141.png"),                            // 1. Cover
    path.join(BRAIN_DIR, "hrbp_transformation_mindmap_infographic_1773157697477.png"),     // 2. TOC
    path.join(BRAIN_DIR, "v23_statue_stoic_1773681336943.png"),                          // 3. Intro
    path.join(BRAIN_DIR, "hrbp_v21_data_insight_1773680231164.png"),                    // 4. Data
    path.join(BRAIN_DIR, "hrbp_v22_statue_bored_1773680700394.png"),                    // 5. Bored
    path.join(BRAIN_DIR, "hrbp_v22_statue_shock_1773680596901.png"),                    // 6. Awakening
    path.join(BRAIN_DIR, "hrbp_v22_statue_laugh_1773680611632.png"),                    // 7. Survivor
    path.join(BRAIN_DIR, "hrbp_v21_product_manager_1773680252018.png"),                  // 8. PM
    path.join(BRAIN_DIR, "hrbp_v22_statue_wow_1773680679782.png"),                      // 9. Mission Guardian
    path.join(BRAIN_DIR, "hrbp_ethical_ai_balance_v10_1773159607684.png"),               // 10. Balance
    path.join(BRAIN_DIR, "v23_statue_fear_1773681354253.png"),                          // 11. Redesign
    path.join(BRAIN_DIR, "hrbp_workforce_redesign_v13_simple_1773160725232.png"),        // 12. Rescale
    path.join(BRAIN_DIR, "v23_statue_proud_1773681369773.png"),                          // 13. Reskilling
    path.join(BRAIN_DIR, "hrbp_visionary_ladder_v9_1773159135984.png"),                  // 14. Vision
    path.join(BRAIN_DIR, "hrbp_v21_ethics_governance_1773679974929.png"),                // 15. Constitution
    path.join(BRAIN_DIR, "hrbp_professional_human_ai_v10_1773159552183.png"),            // 16. Audit
    path.join(BRAIN_DIR, "hrbp_v22_statue_think_1773680624839.png"),                    // 17. Collab
    path.join(BRAIN_DIR, "hrbp_v22_statue_angry_1773680640224.png"),                    // 18. Authority
    path.join(BRAIN_DIR, "hrbp_v21_roadmap_roman_1773680269997.png"),                   // 19. Roadmap
    path.join(BRAIN_DIR, "hrbp_transformation_journey_v10_1773159569827.png"),           // 20. Journey Phases
    path.join(BRAIN_DIR, "hrbp_ai_ethics_v13_simple_1773160703050.png"),                // 21. Metrics
    path.join(BRAIN_DIR, "hrbp_v21_stl_role_1773679960676.png")                         // 22. End Mindmap
];

// 核心檢查函式：確保圖資唯一
const usedImages = new Set();
function getUniqueImage(index) {
    const img = IMAGES[index];
    if (usedImages.has(img)) {
        console.error(`ERROR: Duplicate image detected at index ${index}: ${img}`);
        process.exit(1);
    }
    if (!fs.existsSync(img)) {
        console.error(`ERROR: Missing image at ${img}`);
        process.exit(1);
    }
    usedImages.add(img);
    return img;
}

const HIGHLIGHT_LIST = ["策略參與", "挑戰與契機", "行政事務性束縛", "AI 轉型", "策略人才領袖", "STL", "人力重新設計", "AI 倫理治理", "人機協作效率", "產品經理", "使命守護者", "生存者", "產能回收", "繼任人才庫", "最後一公里"];

const LAYOUT = {
    header: { x: 0.5, y: 0.4, w: 9.0, h: 0.6 },
    line: { x: 0.5, y: 1.1, w: 1.5, h: 0.05 },
    main: {
        split: {
            leftImg: { x: 0.5, y: 1.4, w: 4.0, h: 3.8 },
            rightText: { x: 4.7, y: 1.4, w: 4.8, h: 3.8 },
            rightImg: { x: 5.3, y: 1.4, w: 4.0, h: 3.8 },
            leftText: { x: 0.6, y: 1.4, w: 4.5, h: 3.8 }
        },
        fullText: { x: 0.6, y: 1.4, w: 8.8, h: 3.8 }
    },
    footer: { x: 8.5, y: 5.2, w: 1.2, h: 0.3 }
};

function applyHeader(slide, title, pageNum) {
    slide.addText(title, { x: LAYOUT.header.x, y: LAYOUT.header.y, w: LAYOUT.header.w, h: LAYOUT.header.h, fontSize: 26, bold: true, color: THEME.primary, fontFace: FONT_TITLE });
    slide.addShape(pres.shapes.RECTANGLE, { x: LAYOUT.line.x, y: LAYOUT.line.y, w: LAYOUT.line.w, h: LAYOUT.line.h, fill: { color: THEME.secondary } });
    slide.addText(`MDL | v23 | PAGE ${pageNum}`, { x: LAYOUT.footer.x, y: LAYOUT.footer.y, w: LAYOUT.footer.w, h: LAYOUT.footer.h, fontSize: 10, color: THEME.secondary, align: "right", fontFace: FONT_BODY, bold: true });
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
                    fontSize: opts.fontSize || 15,
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
    title: 'MARMOREAL_V23',
    background: { color: THEME.white },
    objects: [
        { rect: { x: 0, y: 0, w: 0.1, h: "100%", fill: { color: THEME.secondary } } },
        { shape: pres.shapes.RECTANGLE, options: { x: 0.5, y: 5.5, w: 9.0, h: 0.05, fill: { color: THEME.line } } }
    ]
});

// 1. Cover
let s1 = pres.addSlide();
s1.addImage({ path: getUniqueImage(0), x: 0, y: 0, w: '100%', h: '100%' });
s1.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.5, w: 6.5, h: 2.5, fill: { color: 'FFFFFF', transparency: 10 } });
s1.addText("重塑 HRBP 角色：\n引領企業 AI 轉型的策略實務", { x: 0.8, y: 1.8, w: 6.0, h: 1.5, fontSize: 34, bold: true, color: THEME.primary, fontFace: FONT_TITLE });
s1.addText("羅馬美學 v23 | 終極一致性排版傑作", { x: 0.8, y: 3.3, w: 6.0, h: 0.5, fontSize: 16, color: THEME.secondary, fontFace: FONT_TITLE });

// 2. TOC
let s2 = pres.addSlide({ masterName: 'MARMOREAL_V23' });
applyHeader(s2, "主題概覽：AI 時代的 HR 藍圖", 2);
s2.addImage({ path: getUniqueImage(1), x: 0.5, y: 1.4, w: 4.0, h: 3.8, sizing: { type: 'contain' } });
renderContent(s2, [
    "第一章：時代變革與角色奇點 (P.3-6)",
    "第二章：策略人才領袖 (STL) 三位一體 (P.7-10)",
    "第三章：四大核心執行任務 (P.11-17)",
    "第四章：實務路徑與量化指標 (P.18-21)",
    "附錄：三階架構全量心智圖 (P.22)"
], { x: 4.7, w: 4.8 });

// 3. Intro (Stoic)
let s3 = pres.addSlide({ masterName: 'MARMOREAL_V23' });
applyHeader(s3, "引言：在技術洪流中保持冷靜", 3);
s3.addImage({ path: getUniqueImage(2), x: 5.3, y: 1.4, w: 4.0, h: 3.8, sizing: { type: 'contain' } });
renderContent(s3, [
    "AI 不是要取代人類，而是要重新定義人類的職能。",
    "面對轉型，HR 需要具備石刻般的冷靜與 策略人才領袖 的勇氣。",
    "關鍵奇點：當技術效率超越行政勞動的邊界。"
], { x: 0.6, w: 4.5 });

// 4. Data
let s4 = pres.addSlide({ masterName: 'MARMOREAL_V23' });
applyHeader(s4, "數據破局：全球 HR 策略斷層", 4);
s4.addImage({ path: getUniqueImage(3), x: 0.5, y: 1.4, w: 4.0, h: 3.8, sizing: { type: 'contain' } });
renderContent(s4, [
    "51% 的 HRBP 目前缺乏與業務深度對話的數據支撐。",
    "碎片化的 策略參與 導致組織轉型遲滯。",
    "數據洞察 是打破行政牢籠的首要武器。"
], { x: 4.7, w: 4.8 });

// 5. Bored
let s5 = pres.addSlide({ masterName: 'MARMOREAL_V23' });
applyHeader(s5, "挑戰：日常行政的毒藥", 5);
s5.addImage({ path: getUniqueImage(4), x: 5.3, y: 1.4, w: 4.0, h: 3.8, sizing: { type: 'contain' } });
renderContent(s5, [
    "重複勞動如何扼殺了 90% 的組織創造力。",
    "行政事務性束縛 是通往 STL 之路的最大敵手。",
    "擺脫「人事代理人」標籤的迫切性。"
], { x: 0.6, w: 4.5 });

// 6. Awakening (Shock)
let s6 = pres.addSlide({ masterName: 'MARMOREAL_V23' });
applyHeader(s6, "覺醒：當 AI 拆除舊有圍牆", 6);
s6.addImage({ path: getUniqueImage(5), x: 0.5, y: 1.4, w: 4.0, h: 3.8, sizing: { type: 'contain' } });
renderContent(s6, [
    "震碎舊思維：技術正以不可思議的速度重構價值鏈。",
    "警鐘響起：不具備技術整合能力的 HRBP 將面臨淘汰。",
    "角色覺醒：從資源守護者轉向 組織架構師。"
], { x: 4.7, w: 4.8 });

// 7. Survivor (Laugh)
let s7 = pres.addSlide({ masterName: 'MARMOREAL_V23' });
applyHeader(s7, "角色 I：生存者 (Survivor)", 7);
s7.addImage({ path: getUniqueImage(6), x: 5.3, y: 1.4, w: 4.0, h: 3.8, sizing: { type: 'contain' } });
renderContent(s7, [
    "笑對機器：保有 AI 無法模擬的情緒厚度。",
    "生存任務：校準偏見，成為技術決策中的 人性燈塔。",
    "價值：在演算法失效時的人工接管與 倫理守望。"
], { x: 0.6, w: 4.5 });

// 8. PM
let s8 = pres.addSlide({ masterName: 'MARMOREAL_V23' });
applyHeader(s8, "角色 II：產品經理 (Product Manager)", 8);
s8.addImage({ path: getUniqueImage(7), x: 0.5, y: 1.4, w: 4.0, h: 3.8, sizing: { type: 'contain' } });
renderContent(s8, [
    "模組化：將 HR 服務封裝成可迭代的「數位產品」。",
    "數據閉環：用用戶思維解決人才 產能回收 問題。",
    "目標：實現 組織效能 的規模化提升。"
], { x: 4.7, w: 4.8 });

// 9. Mission Guardian (Wow)
let s9 = pres.addSlide({ masterName: 'MARMOREAL_V23' });
applyHeader(s9, "角色 III：使命守護者", 9);
s9.addImage({ path: getUniqueImage(8), x: 5.3, y: 1.4, w: 4.0, h: 3.8, sizing: { type: 'contain' } });
renderContent(s9, [
    "敬畏之心：守護企業文化的 靈魂完整性。",
    "長線防禦：防止技術異化導致的人才流失。",
    "守護任務：確保 組織韌性 始終在高位運行。"
], { x: 0.6, w: 4.5 });

// 10. Balance
let s10 = pres.addSlide({ masterName: 'MARMOREAL_V23' });
applyHeader(s10, "三項角色的協同與制衡", 10);
s10.addImage({ path: getUniqueImage(9), x: 0.5, y: 1.4, w: 4.0, h: 3.8, sizing: { type: 'contain' } });
renderContent(s10, [
    "動態平衡：生存、效率與文化的黃金三角。",
    "協同效應：如何讓 STL 的三種面向共同 賦能業務。",
    "核心：實現 1+1+1 > 3 的策略影響力。"
], { x: 4.7, w: 4.8 });

// 11. Redesign (Fear)
let s11 = pres.addSlide({ masterName: 'MARMOREAL_V23' });
applyHeader(s11, "任務 I：人力重新設計", 11);
s11.addImage({ path: getUniqueImage(10), x: 5.3, y: 1.4, w: 4.0, h: 3.8, sizing: { type: 'contain' } });
renderContent(s11, [
    "恐懼的終點：正視崗位消失帶來的組織焦慮。",
    "人力重新設計：基於 AI 替代率的崗位拆解。",
    "核心：將人力資源釋放至 策略性高價值區域。"
], { x: 0.6, w: 4.5 });

// 12. Rescale
let s12 = pres.addSlide({ masterName: 'MARMOREAL_V23' });
applyHeader(s12, "規模化：未來組織的縮放", 12);
s12.addImage({ path: getUniqueImage(11), x: 0.5, y: 1.4, w: 4.0, h: 3.8, sizing: { type: 'contain' } });
renderContent(s12, [
    "數位彈性：如何建立具備快速縮放能力的人才雲。",
    "產能回收 實踐：量化分析自動化對組織的貢獻值。",
    "實務：建立職能流動的 動態監測系統。"
], { x: 4.7, w: 4.8 });

// 13. Reskilling (Proud)
let s13 = pres.addSlide({ masterName: 'MARMOREAL_V23' });
applyHeader(s13, "任務 II：技能再造 Reskilling", 13);
s13.addImage({ path: getUniqueImage(12), x: 5.3, y: 1.4, w: 4.0, h: 3.8, sizing: { type: 'contain' } });
renderContent(s13, [
    "自豪轉型：看見員工習得 數位新技能 的成就感。",
    "Reskilling 策略：精準畫像與個性化發展路徑。",
    "武器：將培訓轉化為 組織競爭力 的持續積累。"
], { x: 0.6, w: 4.5 });

// 14. Vision
let s14 = pres.addSlide({ masterName: 'MARMOREAL_V23' });
applyHeader(s14, "願景梯：建立繼任者高地", 14);
s14.addImage({ path: getUniqueImage(13), x: 0.5, y: 1.4, w: 4.0, h: 3.8, sizing: { type: 'contain' } });
renderContent(s14, [
    "高處視野：透過預測模型識別具有 策略領袖 潛力者。",
    "繼任人才庫：建立具備 AI 轉型 適配度的梯隊。",
    "長遠目標：確保組織大腦的 代際領先。"
], { x: 4.7, w: 4.8 });

// 15. Constitution
let s15 = pres.addSlide({ masterName: 'MARMOREAL_V23' });
applyHeader(s15, "任務 III：AI 倫理治理憲法", 15);
s15.addImage({ path: getUniqueImage(14), x: 5.3, y: 1.4, w: 4.0, h: 3.8, sizing: { type: 'contain' } });
renderContent(s15, [
    "倫理基石：建立不可逾越的 數位行為準則。",
    "治理憲法：界定數據應用與演算法決策的邊界。",
    "任務：保護員工權益，防範 技術濫用 風險。"
], { x: 0.6, w: 4.5 });

// 16. Audit
let s16 = pres.addSlide({ masterName: 'MARMOREAL_V23' });
applyHeader(s16, "審計：確保演算法的透明性", 16);
s16.addImage({ path: getUniqueImage(15), x: 0.5, y: 1.4, w: 4.0, h: 3.8, sizing: { type: 'contain' } });
renderContent(s16, [
    "算法審計：定期的去偏見檢測與邏輯回歸。",
    "透明治理 實作：讓決策黑盒變白，贏得員工信任。",
    "關鍵指標：算法決策的 偏差修正率。 "
], { x: 4.7, w: 4.8 });

// 17. Collab (Think)
let s17 = pres.addSlide({ masterName: 'MARMOREAL_V23' });
applyHeader(s17, "任務 IV：人機協作效率優化", 17);
s17.addImage({ path: getUniqueImage(16), x: 5.3, y: 1.4, w: 4.0, h: 3.8, sizing: { type: 'contain' } });
renderContent(s17, [
    "思維對接：尋找人類直覺與機器算力的 最佳切點。",
    "效率優化：設計具備高擴展性的 人機共生流程。",
    "心理賦能：管理轉型期的心理落差與 效率焦慮。"
], { x: 0.6, w: 4.5 });

// 18. Authority (Angry)
let s18 = pres.addSlide({ masterName: 'MARMOREAL_V23' });
applyHeader(s18, "策略主導：奪回 HR 的發言權", 18);
s18.addImage({ path: getUniqueImage(17), x: 0.5, y: 1.4, w: 4.0, h: 3.8, sizing: { type: 'contain' } });
renderContent(s18, [
    "強勢回歸：不再是配角，而是 業務策略 的制定者。",
    "策略主導權：基於人才洞察向業務發出 轉型挑戰。",
    "決策深度：在董事會中嵌入 STL 價值節點。"
], { x: 4.7, w: 4.8 });

// 19. Roadmap
let s19 = pres.addSlide({ masterName: 'MARMOREAL_V23' });
applyHeader(s19, "數位羅馬大道：2026 行動綱領", 19);
s19.addImage({ path: getUniqueImage(18), x: 0.5, y: 1.5, w: 9.0, h: 3.5, sizing: { type: 'cover' } });
s19.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 3.5, w: 9.0, h: 1.2, fill: { color: 'FFFFFF', transparency: 15 } });
s19.addText("啟動計畫 (Stage 1) -> 職能演化 (Stage 2) -> 策略深耕 (Stage 3)", { x: 0.7, y: 3.8, w: 8.6, fontSize: 22, bold: true, color: THEME.secondary, align: "center", fontFace: FONT_TITLE });

// 20. Journey
let s20 = pres.addSlide({ masterName: 'MARMOREAL_V23' });
applyHeader(s20, "執行三部曲：轉型時間軸", 20);
s20.addImage({ path: getUniqueImage(19), x: 0.5, y: 1.4, w: 4.0, h: 3.8, sizing: { type: 'contain' } });
renderContent(s20, [
    "Stage 1 啟動：3 個月內完成 產能回收 與自動化部署。",
    "Stage 2 演化：6 個月內實現 STL 核心職能轉型。",
    "Stage 3 引領：12 個月內全面對齊 業務策略核心。 "
], { x: 4.7, w: 4.8 });

// 21. Metrics
let s21 = pres.addSlide({ masterName: 'MARMOREAL_V23' });
applyHeader(s21, "驗收：量化轉型的商業價值", 21);
s21.addImage({ path: getUniqueImage(20), x: 5.3, y: 1.4, w: 4.0, h: 3.8, sizing: { type: 'contain' } });
renderContent(s21, [
    "價值指標 I：行政耗時減少 40% 以上。",
    "價值指標 II：繼任人才庫 的 AI 適配能力成長百分比。",
    "價值指標 III：HR 策略引導下的 業務人效成長率。"
], { x: 0.6, w: 4.5 });

// 22. End Mindmap
let s22 = pres.addSlide({ masterName: 'MARMOREAL_V22' });
s22.addImage({ path: getUniqueImage(21), x: 6.5, y: 0.5, w: 3.0, h: 3.5, transparency: 80 });
applyHeader(s22, "三階架構：全量主題與頁碼對齊", 22);

const M_DATA = [
    { t: "時代變革", c: [
        { k: "轉型背景及趨勢", s: ["策略斷層 (P.4)", "行政束縛 (P.5)", "奇點引言 (P.3)", "角色覺醒 (P.6)"] }
    ]},
    { t: "核心角色 (STL)", c: [
        { k: "三位一體定位", s: ["生存者 (P.7)", "產品經理 (P.8)", "守護者 (P.9)", "協同平衡 (P.10)"] }
    ]},
    { t: "執行任務", c: [
        { k: "人才與設計", s: ["人力設計 (P.11-12)", "技能再造 (P.13-14)"] },
        { k: "治理與效率", s: ["倫理憲法 (P.15-16)", "算法審計 (P.16)", "協作優化 (P.17)"] }
    ]},
    { t: "實務路徑", c: [
        { k: "路徑與指標", s: ["策略主導 (P.18)", "數位大道 (P.19)", "轉型階段 (P.20)", "量化指標 (P.21)"] }
    ]}
];

const startX = 0.5, startY = 1.3, colWidth = 2.4;
M_DATA.forEach((l1, i) => {
    let x = startX + (i * colWidth);
    s22.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: x, y: startY, w: 2.2, h: 0.4, fill: { color: THEME.primary }, rectRadius: 0.1 });
    s22.addText(l1.t, { x: x, y: startY, w: 2.2, h: 0.4, color: THEME.white, bold: true, align: "center", fontSize: 13, fontFace: FONT_TITLE });
    let currentY = startY + 0.6;
    l1.c.forEach(l2 => {
        s22.addShape(pres.shapes.RECTANGLE, { x: x + 0.1, y: currentY, w: 2.0, h: 0.35, fill: { color: THEME.secondary }, transparency: 20 });
        s22.addText(l2.k, { x: x + 0.1, y: currentY, w: 2.0, h: 0.35, color: THEME.white, bold: true, align: "center", fontSize: 10, fontFace: FONT_TITLE });
        currentY += 0.45;
        l2.s.forEach(l3 => {
            s22.addText("● " + l3, { x: x + 0.2, y: currentY, w: 2.0, h: 0.25, color: THEME.text, fontSize: 8.5, fontFace: FONT_TITLE });
            currentY += 0.3;
        });
        currentY += 0.15;
    });
});

const outP = path.join(SCRATCH_DIR, "output", "HRBP_AI_Transformation_v23_Ultimate_Final.pptx");
pres.writeFile({ fileName: outP }).then(fn => {
    console.log(`Ultimate Success: v23 Consistent Aligned PPTX Generated at ${fn}`);
}).catch(err => console.error(err));
