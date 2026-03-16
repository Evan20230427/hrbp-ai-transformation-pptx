const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

let pres = new pptxgen();
pres.layout = 'LAYOUT_16x9';
pres.author = 'Antigravity AI';
pres.title = 'HRBP AI 轉型最佳實務 - 20 頁擴充版 (v21)';

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
    roadmap: path.join(BRAIN_DIR, "hrbp_v21_roadmap_roman_1773680269997.png")
};

const HIGHLIGHT_LIST = ["策略參與", "挑戰與契機", "行政事務性束縛", "AI 轉型", "策略人才領袖", "STL", "人力重新設計", "AI 倫理治理", "人機協作效率", "產品經理", "使命守護者", "生存者", "產能回收", "繼任人才庫", "最後一公里"];

function applyHeader(slide, title, pageNum) {
    slide.addText(title, { x: 0.5, y: 0.4, w: 9, h: 0.6, fontSize: 28, bold: true, color: THEME.primary, fontFace: FONT_TITLE });
    slide.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.1, w: 1.5, h: 0.05, fill: { color: THEME.secondary } });
    slide.addText(`MDL | PAGE ${pageNum}`, { x: 8.5, y: 5.2, w: 1.2, h: 0.3, fontSize: 10, color: THEME.secondary, align: "right", fontFace: FONT_BODY, bold: true });
}

function renderContent(slide, lines, opts = {}) {
    let safeX = opts.x || 0.6;
    let safeY = opts.y || 1.4;
    let safeW = opts.w || 8.8;
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
    slide.addText(content, { x: safeX, y: safeY, w: safeW, h: 4.0, lineSpacing: 22, valign: "top" });
}

pres.defineSlideMaster({
    title: 'MARMOREAL_V21_EXT',
    background: { color: THEME.white },
    objects: [
        { rect: { x: 0, y: 0, w: 0.1, h: "100%", fill: { color: THEME.secondary } } },
        { shape: pres.shapes.RECTANGLE, options: { x: 0.5, y: 5.5, w: 9.0, h: 0.1, fill: { color: THEME.line } } }
    ]
});

// --- P1: Cover ---
let s1 = pres.addSlide();
if (fs.existsSync(ASSETS.cover)) s1.addImage({ path: ASSETS.cover, x: 0, y: 0, w: '100%', h: '100%' });
s1.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.5, w: 6.5, h: 2.5, fill: { color: 'FFFFFF', transparency: 10 } });
s1.addText("重塑 HRBP 角色：\n引領企業 AI 轉型的策略實務", { x: 0.8, y: 1.8, w: 6.0, h: 1.5, fontSize: 34, bold: true, color: THEME.primary, fontFace: FONT_TITLE });
s1.addText("20 頁擴充版 | 羅馬美學與四階全量對齊 (v21)", { x: 0.8, y: 3.3, w: 6.0, h: 0.5, fontSize: 16, color: THEME.secondary, fontFace: FONT_TITLE });

// --- P2: Intro ---
let s2 = pres.addSlide({ masterName: 'MARMOREAL_V21_EXT' });
applyHeader(s2, "引言：AI 轉型的「最後一公里」", 2);
renderContent(s2, [
    "HRBP 不僅是流程的執行者，更是技術落地的最後環節。",
    "為什麼只有 HR 才能彌補技術轉向業務價值的 最後一公里？",
    "從行政支援轉型為 策略人才領袖 的必然性。"
]);

// --- P3: Data ---
let s3 = pres.addSlide({ masterName: 'MARMOREAL_V21_EXT' });
applyHeader(s3, "數據洞察：全球 500 強整合現況", 3);
if (fs.existsSync(ASSETS.data)) s3.addImage({ path: ASSETS.data, x: 0.5, y: 1.3, w: 4.5, h: 3.8, sizing: { type: 'contain' } });
renderContent(s3, [
    "51% 的 HR 團隊目前對於 策略參與 感到力不從心。",
    "數據碎片化是實現 AI 轉型 的最大結構性障礙。",
    "成功的轉型企業，其 HRBP 均具備高度的數據洞察力。"
], { x: 5.2, w: 4.2 });

// --- P4: Challenge I ---
let s4 = pres.addSlide({ masterName: 'MARMOREAL_V21_EXT' });
applyHeader(s4, "挑戰 I：行政事務性束縛", 4);
renderContent(s4, [
    "行政事務性束縛 佔據了 HRBP 超過 60% 的工作時間。",
    "「吸塵器效應」：瑣碎事務如何扼殺高價值的策略性思考。",
    "釋放產能是通往 STL 的第一張入門票。"
]);

// --- P5: Challenge II ---
let s5 = pres.addSlide({ masterName: 'MARMOREAL_V21_EXT' });
applyHeader(s5, "挑戰 II：組織慣性與羅馬陷阱", 5);
renderContent(s5, [
    "羅馬不是一天造成的，組織慣性也不是一天能移除的。",
    "舊有架構與新興 AI 技術之間的「斷層線」。",
    "如何在不破壞現有穩定性的前提下進行 角色重塑。"
]);

// --- P6: Position ---
let s6 = pres.addSlide({ masterName: 'MARMOREAL_V21_EXT' });
applyHeader(s6, "角色定位：從保姆到組織架構師", 6);
renderContent(s6, [
    "傳統 HRBP：被動回應業務需求的「救火員」。",
    "未來 STL：主動設計人才流動與價值的「建築師」。",
    "核心轉折點：數據主導與 AI 賦能。"
]);

// --- P7: Survivor ---
let s7 = pres.addSlide({ masterName: 'MARMOREAL_V21_EXT' });
applyHeader(s7, "STL 剖析 I：生存者 (Survivor)", 7);
renderContent(s7, [
    "人機協同中的「關鍵接口」，處理 AI 無法覆蓋的情緒。 ",
    "生存任務：校準偏見，確保 AI 代行時的公平一致性。",
    "最後守門員：在算法失效時的專業人工接管。"
], { x: 4.2, w: 5.3 });
if (fs.existsSync(ASSETS.stl)) s7.addImage({ path: ASSETS.stl, x: 0.5, y: 1.2, w: 3.5, h: 4.2, sizing: { type: 'contain' } });

// --- P8: PM ---
let s8 = pres.addSlide({ masterName: 'MARMOREAL_V21_EXT' });
applyHeader(s8, "STL 剖析 II：產品經理 (Product Manager)", 8);
if (fs.existsSync(ASSETS.pm)) s8.addImage({ path: ASSETS.pm, x: 0.5, y: 1.3, w: 4.5, h: 3.8, sizing: { type: 'contain' } });
renderContent(s8, [
    "將 HR 解決方案封裝成可規模化的「數位產品」。",
    "核心價值：用 產品經理 思維解決人才留任與發展問題。",
    "數據驅動：透過 A/B 測試與數據迭代優化組織效能。"
], { x: 5.2, w: 4.2 });

// --- P9: Guardian ---
let s9 = pres.addSlide({ masterName: 'MARMOREAL_V21_EXT' });
applyHeader(s9, "STL 剖析 III：使命守護者", 9);
renderContent(s9, [
    "使命守護者：捍衛企業文化在 AI 時代的純粹性。",
    "長期價值：不僅關注效率，更關注 技術與人性的融合。",
    "倫理官：全權負責企業內部的 AI 倫理治理 監督。"
]);

// --- P10: Task I ---
let s10 = pres.addSlide({ masterName: 'MARMOREAL_V21_EXT' });
applyHeader(s10, "核心任務 I：人力重新設計", 10);
renderContent(s10, [
    "執行基於 AI 替代率預測的 人力重新設計。",
    "重新平衡自動化工作與人類高價值創造的工作比例。",
    "產能回收 的實質落地策略。"
]);

// --- P11: Reskilling ---
let s11 = pres.addSlide({ masterName: 'MARMOREAL_V21_EXT' });
applyHeader(s11, "技能地圖再造：Reskilling 策略", 11);
renderContent(s11, [
    "將「技能再造」提升至企業策略高度。",
    "建立動態技能雷達，預測未來兩年的 策略人才需求。",
    "不僅是培訓，更是組織競爭力的二次開發。"
]);

// --- P12: Ethics ---
let s12 = pres.addSlide({ masterName: 'MARMOREAL_V21_EXT' });
applyHeader(s12, "核心任務 II：AI 倫理治理憲法", 12);
renderContent(s12, [
    "建立企業內部的透明性準則與 倫理治理 憲法。",
    "確保數據應用符合規範，保護員工隱私。",
    "建立針對 AI 決策的「人類上訴」機制。 "
], { x: 0.6, w: 5.3 });
if (fs.existsSync(ASSETS.ethics)) s12.addImage({ path: ASSETS.ethics, x: 6.0, y: 1.2, w: 3.5, h: 4.2, sizing: { type: 'contain' } });

// --- P13: Audit ---
let s13 = pres.addSlide({ masterName: 'MARMOREAL_V21_EXT' });
applyHeader(s13, "算法審計：建立防禦體系", 13);
renderContent(s13, [
    "對內部招募與績效算法進行定期的 偏見監測。",
    "透明治理 的落地：讓員工理解 AI 是如何輔助決策的。",
    "預算與風險：倫理失當造成的僱傭品牌損害評估。"
]);

// --- P14: Efficiency ---
let s14 = pres.addSlide({ masterName: 'MARMOREAL_V21_EXT' });
applyHeader(s14, "核心任務 III：人機協作效率", 14);
renderContent(s14, [
    "設計高同步性的人機共生流程，優化 人機協作效率。",
    "心理賦能：管理轉型期的技術焦慮，提升員工適應度。",
    "共生指標：如何衡量 1+1 > 2 的協作成果。"
]);

// --- P15: Roadmap ---
let s15 = pres.addSlide({ masterName: 'MARMOREAL_V21_EXT' });
applyHeader(s15, "實務攻略：分階段路徑圖", 15);
if (fs.existsSync(ASSETS.roadmap)) s15.addImage({ path: ASSETS.roadmap, x: 0, y: 1.5, w: '100%', h: 3.8, sizing: { type: 'cover' } });
s15.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 4.0, w: 9.0, h: 1.2, fill: { color: 'FFFFFF', transparency: 15 } });
s15.addText("啟動 (P1) -> 轉型 (P2) -> 引領 (P3)", { x: 0.7, y: 4.3, w: 8.6, fontSize: 24, bold: true, color: THEME.secondary, align: "center", fontFace: FONT_TITLE });

// --- P16: P1 Tech ---
let s16 = pres.addSlide({ masterName: 'MARMOREAL_V21_EXT' });
applyHeader(s16, "P1：自動化路徑專案", 16);
renderContent(s16, [
    "透過 自動化路徑專案 釋放 30% 以上的行政工時。",
    "產能回收計畫：將節省的時間精確投資於 STL 角色訓練。",
    "工具導入：選擇適合 HR 情境的 AI 自動化工具。"
]);

// --- P17: P2 Talent ---
let s17 = pres.addSlide({ masterName: 'MARMOREAL_V21_EXT' });
applyHeader(s17, "P2：繼任人才庫賦能", 17);
renderContent(s17, [
    "利用預測模型識別具有 STL 潛力的 繼任人才庫。",
    "職能放大：將回收後的產能轉化為對高階人才的個性化發展。",
    "建立未來型態的領導力評估模型。"
]);

// --- P18: P3 Power ---
let s18 = pres.addSlide({ masterName: 'MARMOREAL_V21_EXT' });
applyHeader(s18, "P3：策略主導權確立", 18);
renderContent(s18, [
    "在組織大腦中嵌入 STL 策略節點。",
    "全面執掌 AI 轉型 決策權，成為業務技術不可或缺的夥伴。",
    "建立 HR 與業務部門的策略共同體。"
]);

// --- P19: KPIs ---
let s19 = pres.addSlide({ masterName: 'MARMOREAL_V21_EXT' });
applyHeader(s19, "成功指標：如何衡量轉型價值", 19);
renderContent(s19, [
    "產能回收 率：行政時間減少的百分比。",
    "人機協作 滿意度：員工與技術配對後的效能提升速率。",
    "STL 策略貢獻值：HR 提案在業務決策中的採納比例。"
]);

// --- P20: Final ---
let s20 = pres.addSlide();
s20.background = { color: THEME.primary };
s20.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0.5, w: '100%', h: 4.5, fill: { color: 'FFFFFF', transparency: 90 } });
s20.addText("大理石之生：\n在技術洪流中雕琢人類智慧", { x: 0.5, y: 1.8, w: 9, h: 1.5, fontSize: 36, bold: true, color: THEME.white, fontFace: FONT_TITLE, align: "center" });
s20.addText("HRBP AI 轉型最佳實務完結", { x: 0.5, y: 3.3, w: 9, h: 0.5, fontSize: 18, color: THEME.secondary, fontFace: FONT_TITLE, align: "center" });

const outP = path.join(SCRATCH_DIR, "output", "HRBP_AI_Transformation_v21_20Pages.pptx");
pres.writeFile({ fileName: outP }).then(fn => {
    console.log(`Success: 20 Pages v21 PPTX Generated at ${fn}`);
}).catch(err => console.error(err));
