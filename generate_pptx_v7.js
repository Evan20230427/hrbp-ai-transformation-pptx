const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

let pres = new pptxgen();
pres.layout = 'LAYOUT_16x9';
pres.author = 'Antigravity AI';
pres.title = 'AI 時代重新設計 HRBP 角色的最佳實務 - 旗艦 2.0';

// Stable & Professional Theme
const THEME = {
    primary: "0F172A",
    secondary: "2563EB",
    accent: "059669",
    bg: "FFFFFF",
    text: "1E293B",
    subtle: "64748B",
    white: "FFFFFF",
    line: "CBD5E1"
};

const FONT_TCH = "Microsoft JhengHei";
const FONT_BODY = "Arial";

/**
 * Robust Text Formatter: Scales English in brackets and removes duplicates.
 * Returns an array of PPTX text objects for use in slide.addText()
 */
function processText(lines, baseSize = 18) {
    const uniqueLines = [...new Set(lines)]; // Deduplicate
    const smallSize = Math.max(12, baseSize - 4);

    return uniqueLines.map(line => {
        const parts = [];
        // Catch (English Content) or (Numbers%)
        const regex = /(\([^)]+\))/g;
        const tokens = line.split(regex);

        tokens.forEach(token => {
            if (token.match(regex)) {
                parts.push({ text: token, options: { fontSize: smallSize, color: THEME.secondary, fontFace: FONT_BODY, italic: true } });
            } else if (token.trim()) {
                parts.push({ text: token, options: { fontSize: baseSize, color: THEME.text, fontFace: FONT_TCH } });
            }
        });
        return parts;
    });
}

// Master Slide Design
pres.defineSlideMaster({
    title: 'STL_V7_MASTER',
    background: { color: THEME.bg },
    objects: [
        { rect: { x: 0, y: 0, w: "100%", h: 0.1, fill: { color: THEME.secondary } } },
        { text: { text: "© Gartner Insight | HRBP AI Transformation Strategy Guide", options: { x: 0.5, y: 5.3, w: 9, h: 0.25, fontSize: 10, color: THEME.subtle, align: "right", fontFace: FONT_BODY } } }
    ]
});

function addHeader(slide, title) {
    slide.addText(title, { x: 0.5, y: 0.35, w: 9, h: 0.6, fontSize: 26, bold: true, color: THEME.primary, fontFace: FONT_TCH });
    slide.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 0.95, w: 9, h: 0.02, fill: { color: THEME.secondary } });
}

// --- Content Generation ---

// 1. Title
let slide1 = pres.addSlide();
slide1.background = { color: THEME.primary };
slide1.addText("重塑 HRBP 角色：\n引領企業 AI 轉型的策略實務", { x: 0, y: 1.8, w: "100%", h: 1.2, fontSize: 44, bold: true, color: THEME.white, fontFace: FONT_TCH, align: "center" });
slide1.addText("2026 年度旗艦指南 | 全量深度數據版", { x: 0, y: 3.1, w: "100%", h: 0.5, fontSize: 18, color: THEME.secondary, fontFace: FONT_TCH, align: "center" });

// 2. The Paradox
let slide2 = pres.addSlide({ masterName: "STL_V7_MASTER" });
addHeader(slide2, "存續危機：策略貢獻與事務束縛");
processText([
    "關鍵數據揭露：僅 51% (Only 51%) 的領導者同意 HRBP 參與了重大策略討論。",
    "轉型瓶頸：HRBP 仍被鎖定在 AI 正在迅速吸收的工作中 (Transactional Work)。",
    "自動化威脅：包括職務描述、數據摘要、政策回答 (FAQ) 等任務。",
    "核心風險：即使被視為「策略性」的人員也面臨未充分利用的風險 (Underleveraged)。"
]).forEach((p, i) => slide2.addText(p, { x: 0.8, y: 1.5 + (i * 0.75), w: 8.5, bullet: true, lineSpacing: 28 }));

// 3. The Future Role: STL
let slide3 = pres.addSlide({ masterName: "STL_V7_MASTER" });
addHeader(slide3, "未來定位：AI 轉型顧問 (STL)");
processText([
    "未來的 HRBP 角色應定位為「策略人才領袖 (Strategic Talent Leaders)」。",
    "權責擴張：直接引導人工智慧驅動轉型的人員面向 (People side of AI-driven transformation)。",
    "從解釋人員策略進化為「主導轉型對話」積極參與者。"
]).forEach((p, i) => slide3.addText(p, { x: 0.8, y: 1.5 + (i * 0.9), w: 8.5, bullet: true, lineSpacing: 30 }));

// 4. STL Responsibilities (1)
let slide4 = pres.addSlide({ masterName: "STL_V7_MASTER" });
addHeader(slide4, "職職責 1：引導人力重新設計");
processText([
    "隨 AI 改變職務角色，主導職能重塑決策 (Workforce Redesign)。",
    "決定何時進行人員再培訓 (Reskill)、重新部署 (Redeploy) 或逐步淘汰職位。",
    "確保組織架構與 AI 產出能力完美對齊。"
]).forEach((p, i) => slide4.addText(p, { x: 0.8, y: 1.8, w: 8.5, bullet: true }));

// 5. STL Responsibilities (2)
let slide5 = pres.addSlide({ masterName: "STL_V7_MASTER" });
addHeader(slide5, "職責 2：解決 AI 倫理與偏見");
processText([
    "解決 AI 驅動的人才決策中的偏見 (Addressing Bias) 與倫理挑戰。",
    "確保數據推薦算法的透明度與公平性。",
    "維護企業文化的倫理邊界，防止過度依賴非結構化 AI 產出。"
]).forEach((p, i) => slide5.addText(p, { x: 0.8, y: 1.8, w: 8.5, bullet: true }));

// 6. STL Responsibilities (3)
let slide6 = pres.addSlide({ masterName: "STL_V7_MASTER" });
addHeader(slide6, "職責 3：塑造人機協作效益");
processText([
    "在提升生產力的同時，確保不以犧牲參與度為代價 (Human-machine collaboration)。",
    "優化人類直覺與 AI 計算的動態平衡。",
    "重新設計工作流，讓員工感受到 AI 是增益而非威脅。"
]).forEach((p, i) => slide6.addText(p, { x: 0.8, y: 1.8, w: 8.5, bullet: true }));

// 7. Roadmap Steps Overview
let slide7 = pres.addSlide({ masterName: "STL_V7_MASTER" });
addHeader(slide7, "調整任務：轉型三階段路徑概觀");
processText([
    "階段 1：剔除與 AI 重疊的舊有行政負荷 (Strip out legacy work)。",
    "階段 2：透過 AI 強化核心策略責任 (Augment core strategic responsibilities)。",
    "階段 3：擴展至由 AI 推動的新策略領域 (Expand into new strategic work)。"
]).forEach((p, i) => slide7.addText(p, { x: 0.8, y: 1.8 + (i * 0.8), w: 8.5, bullet: { type: "number" } }));

// 8. Phase 1 Actions
let slide8 = pres.addSlide({ masterName: "STL_V7_MASTER" });
addHeader(slide8, "P1 剔除行動：回收策略產能");
processText([
    "事光明確定義策略重點區域 (Define strategic focus)。",
    "建立 12-24 個月的自動化路徑圖 (Automation roadmap)。",
    "實施「停止行動」清單以界定 AI 與 HRBP 的界限 (Stop-doing list)。"
]).forEach((p, i) => slide8.addText(p, { x: 0.8, y: 1.8, w: 8.5, bullet: true }));

// 9. Phase 2 Actions
let slide9 = pres.addSlide({ masterName: "STL_V7_MASTER" });
addHeader(slide9, "P2 強化行動：AI 賦能高價值產出");
processText([
    "更新 HRBP 職能模型，使「AI 準備度」透明化 (AI-readiness)。",
    "利用 AI 生成的預測洞察作為領導對話的基準 (AI-generated inputs)。",
    "定義 HRBP-CoE-AI 工作流，確保專業指導不偏移。"
]).forEach((p, i) => slide9.addText(p, { x: 0.8, y: 1.8, w: 8.5, bullet: true }));

// 10. Phase 3 Actions
let slide10 = pres.addSlide({ masterName: "STL_V7_MASTER" });
addHeader(slide10, "P3 擴展行動：主導新型態策略");
processText([
    "啟動與試行「策略人才領袖 (STL)」 pods 小組。",
    "在轉型決策中嵌入 STL 條款，確保決策需參考 HR 專業判斷。",
    "主動引導組織文化向「AI 原生」轉型。"
]).forEach((p, i) => slide10.addText(p, { x: 0.8, y: 1.8, w: 8.5, bullet: true }));

// 11. Metrics: Efficiency
let slide11 = pres.addSlide({ masterName: "STL_V7_MASTER" });
addHeader(slide11, "成功指標：效率與週期 (Metrics)");
processText([
    "特定行政任務（入職清單、報告請求）回收的工時。",
    "人力/繼任決策週期大幅縮短 (Cycle time reduction)。",
    "目標：回收 20-30% 工時重新投入策略事項。"
]).forEach((p, i) => slide11.addText(p, { x: 0.8, y: 1.8, w: 8.5, bullet: true }));

// 12. Metrics: Quality & Attrition
let slide12 = pres.addSlide({ masterName: "STL_V7_MASTER" });
addHeader(slide12, "成功指標：人才儲備與流失");
processText([
    "18 個月內關鍵職位接班準備率提升 (% Increase in successors)。",
    "AI 標記的高風險角色中，遺憾離職率顯著降低 (Reduction in attrition)。",
    "因 AI 重新設計而被重新部署（而非裁員）的角色比例。"
]).forEach((p, i) => slide12.addText(p, { x: 0.8, y: 1.8, w: 8.5, bullet: true }));

// 13. Gartner 4 Pillars
let slide13 = pres.addSlide({ masterName: "STL_V7_MASTER" });
addHeader(slide13, "Gartner 建議：HR 專業人士專注領務");
processText([
    "1. 增強專業知識，保持對趨勢的即時領略。",
    "2. 使用診斷工具識別開發與變革機會。",
    "3. 利用最佳實務指南加速項目執行流程。",
    "4. 善用數據驅動決策，優化團隊績效產出。"
]).forEach((p, i) => slide13.addText(p, { x: 0.8, y: 1.8, w: 8.5, bullet: true }));

// 14. Action Plan for CHROs
let slide14 = pres.addSlide({ masterName: "STL_V7_MASTER" });
addHeader(slide14, "總結行動清單：給 CHRO 的建議");
processText([
    "停止觀望：立即重新定義 HRBP 角色與職權範圍。",
    "文化建設：建立透明、倫理且具 AI 洞察力的 HR 部門。",
    "技術投資：導入能產生「決策基準數據」的 AI 工具匯總平台。"
]).forEach((p, i) => slide14.addText(p, { x: 0.8, y: 1.8, w: 8.5, bullet: true }));

// 15. Native Mind Map (V7 Stable Tree)
let slideMap = pres.addSlide({ masterName: "STL_V7_MASTER" });
addHeader(slideMap, "全課精華：三層層級心智圖 (Native Stable)");

const ROOT_X = 1.0, ROOT_Y = 2.4;
// Layer 1
const L1_NODES = [
    { t: "挑戰與現狀", c: ["51%策略不足", "行政作業佔據", "AI自動自動威脅"] },
    { t: "STL 顧問定義", c: ["人力重設計", "倫理監測", "人機協作效益"] },
    { t: "轉型三階段", c: ["P1 剔除舊事務", "P2 AI 強化核心", "P3 擴展開拓新領域"] },
    { t: "成功指標", c: ["週期轉短", "繼任準備率提升", "離職率優化"] }
];

// Draw Root
slideMap.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: ROOT_X, y: ROOT_Y, w: 1.6, h: 0.6, fill: { color: THEME.primary }, rectRadius: 0.1 });
slideMap.addText("HRBP AI 轉型", { x: ROOT_X, y: ROOT_Y, w: 1.6, h: 0.6, color: THEME.white, bold: true, align: "center", fontFace: FONT_TCH, fontSize: 13 });

L1_NODES.forEach((node, i) => {
    let nx = ROOT_X + 2.5;
    let ny = 1.0 + (i * 1.2);

    // Connect Root -> L1 (Horiz then Vert)
    slideMap.addShape(pres.shapes.LINE, { x: ROOT_X + 1.6, y: ROOT_Y + 0.3, w: 0.45, h: 0, line: { color: THEME.secondary, width: 2 } });
    slideMap.addShape(pres.shapes.LINE, { x: ROOT_X + 2.05, y: Math.min(ROOT_Y + 0.3, ny + 0.25), w: 0, h: Math.abs(ROOT_Y + 0.3 - (ny + 0.25)), line: { color: THEME.secondary, width: 2 } });
    slideMap.addShape(pres.shapes.LINE, { x: ROOT_X + 2.05, y: ny + 0.25, w: 0.45, h: 0, line: { color: THEME.secondary, width: 2 } });

    // L1 Box
    slideMap.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: nx, y: ny, w: 1.5, h: 0.5, fill: { color: THEME.secondary }, rectRadius: 0.1 });
    slideMap.addText(node.t, { x: nx, y: ny, w: 1.5, h: 0.5, color: THEME.white, bold: true, align: "center", fontFace: FONT_TCH, fontSize: 11 });

    // L2 Children
    node.c.forEach((child, j) => {
        let cx = nx + 2.2;
        let cy = ny - 0.2 + (j * 0.4);

        // Connect L1 -> L2
        slideMap.addShape(pres.shapes.LINE, { x: nx + 1.5, y: ny + 0.25, w: 0.3, h: 0, line: { color: THEME.line, width: 1 } });
        slideMap.addShape(pres.shapes.LINE, { x: nx + 1.8, y: Math.min(ny + 0.25, cy + 0.15), w: 0, h: Math.abs(ny + 0.25 - (cy + 0.15)), line: { color: THEME.line, width: 1 } });
        slideMap.addShape(pres.shapes.LINE, { x: nx + 1.8, y: cy + 0.15, w: 0.4, h: 0, line: { color: THEME.line, width: 1 } });

        // L2 Node
        slideMap.addShape(pres.shapes.RECTANGLE, { x: cx, y: cy, w: 2.0, h: 0.3, fill: { color: "F8FAFC" }, line: { color: THEME.line, width: 1 } });
        slideMap.addText(child, { x: cx, y: cy, w: 2.0, h: 0.3, color: THEME.text, fontSize: 9, fontFace: FONT_TCH, valign: "middle" });
    });
});

// 16. Last Slide
let slideLast = pres.addSlide();
slideLast.background = { color: THEME.primary };
slideLast.addText("啟動您的數據領航之旅", { x: 0, y: 2.3, w: "100%", h: 0.6, bold: true, fontSize: 36, color: THEME.white, align: "center", fontFace: FONT_TCH });
slideLast.addText("Gartner HRBP AI Transformation Deliverable v7", { x: 0, y: 3.1, w: "100%", h: 0.4, fontSize: 14, color: THEME.secondary, align: "center", fontFace: FONT_BODY });

// Final Save Execution
const outDir = path.join(__dirname, "output");
if (!fs.existsSync(outDir)) fs.mkdirSync(outDir);
const finalPath = path.join(outDir, "HRBP_AI_Transformation_Full_v7.pptx");

pres.writeFile({ fileName: finalPath }).then(fn => {
    console.log(`Successfully generated exhaustive PPTX v7 at ${fn}`);
}).catch(err => console.error(err));
