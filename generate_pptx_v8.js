const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

let pres = new pptxgen();
pres.layout = 'LAYOUT_16x9';
pres.author = 'Antigravity AI';
pres.title = 'AI 時代重新設計 HRBP 角色的最佳實務 - 旗艦 2.1';

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
 * Advanced Text Engine: 
 * Converts an array of strings into a single structured array of text objects 
 * for a single addText call to avoid overlap.
 */
function createContentBlock(lines, baseSize = 18) {
    const uniqueLines = [...new Set(lines)];
    const smallSize = Math.max(12, baseSize - 4);
    const finalContent = [];

    uniqueLines.forEach((line, idx) => {
        const regex = /(\([^)]+\))/g;
        const tokens = line.split(regex);

        tokens.forEach((token, tIdx) => {
            const isLastToken = (tIdx === tokens.length - 1);
            const opt = {
                fontSize: token.match(regex) ? smallSize : baseSize,
                color: token.match(regex) ? THEME.secondary : THEME.text,
                fontFace: token.match(regex) ? FONT_BODY : FONT_TCH,
                italic: token.match(regex) ? true : false,
                bullet: (tIdx === 0) ? true : false, // Only first part gets bullet
                breakLine: isLastToken // Break line after the full original line is done
            };
            finalContent.push({ text: token, options: opt });
        });
    });
    return finalContent;
}

// Master Slide
pres.defineSlideMaster({
    title: 'STL_V8_MASTER',
    background: { color: THEME.bg },
    objects: [
        { rect: { x: 0, y: 0, w: "100%", h: 0.1, fill: { color: THEME.secondary } } },
        { text: { text: "© Gartner Insight | HRBP AI Transformation Strategy Guide v8", options: { x: 0.5, y: 5.3, w: 9, h: 0.25, fontSize: 10, color: THEME.subtle, align: "right", fontFace: FONT_BODY } } }
    ]
});

function addHeader(slide, title) {
    slide.addText(title, { x: 0.5, y: 0.35, w: 9, h: 0.6, fontSize: 26, bold: true, color: THEME.primary, fontFace: FONT_TCH });
    slide.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 0.95, w: 9, h: 0.02, fill: { color: THEME.secondary } });
}

// --- CONTENT GENERATION ---

// 1. Title
let slide1 = pres.addSlide();
slide1.background = { color: THEME.primary };
slide1.addText("重塑 HRBP 角色：\n引領企業 AI 轉型的策略實務", { x: 0, y: 1.8, w: "100%", h: 1.2, fontSize: 44, bold: true, color: THEME.white, fontFace: FONT_TCH, align: "center" });
slide1.addText("2026 年度終極旗艦版 | 專家深度內容匯總", { x: 0, y: 3.1, w: "100%", h: 0.5, fontSize: 18, color: THEME.secondary, fontFace: FONT_TCH, align: "center" });

// 2. Paradox
let slide2 = pres.addSlide({ masterName: "STL_V8_MASTER" });
addHeader(slide2, "存續危機：策略貢獻與事務束縛");
slide2.addText(createContentBlock([
    "數據揭露：僅 51% (Only 51%) 的領導者同意 HRBP 參與了重大策略討論。",
    "轉型瓶頸：HRBP 仍被鎖定在 AI 正在快速吸收的工作中 (Transactional Work)。",
    "自動化威脅：包括職務描述、數據摘要、政策回答 (FAQ) 等任務。",
    "核心風險：即使被視後「策略性」的人員也面臨未充分利用的風險 (Underleveraged)。"
]), { x: 0.7, y: 1.4, w: 8.5, h: 3.5, lineSpacing: 24, valign: "top" });

// 3. STL 
let slide3 = pres.addSlide({ masterName: "STL_V8_MASTER" });
addHeader(slide3, "未來定位：AI 轉型顧問 (STL)");
slide3.addText(createContentBlock([
    "未來的 HRBP 角色應定位為「策略人才領袖 (Strategic Talent Leaders)」。",
    "權責擴張：直接引導人工智慧驅動轉型的人員面向 (People side of AI transformation)。",
    "從解釋人員策略進化為「主導轉型對話」積極參與者。"
], 20), { x: 0.7, y: 1.5, w: 8.5, h: 3, lineSpacing: 30 });

// 4. Role 1
let slide4 = pres.addSlide({ masterName: "STL_V8_MASTER" });
addHeader(slide4, "職責 1：引導人力重新設計");
slide4.addText(createContentBlock([
    "隨 AI 改變職務角色，主導職能重塑決策 (Workforce Redesign)。",
    "決定何時進行人員再培訓 (Reskill)、重新部署 (Redeploy) 或逐步淘汰職位。",
    "確保組織架構與 AI 產出能力達成最佳對齊。"
]), { x: 0.7, y: 1.8, w: 8.5, h: 3, lineSpacing: 28 });

// 5. Role 2
let slide5 = pres.addSlide({ masterName: "STL_V8_MASTER" });
addHeader(slide5, "職責 2：解決 AI 倫理與偏見");
slide5.addText(createContentBlock([
    "解決人才決策中的偏見 (Addressing Bias) 與倫理挑戰。",
    "確保數據推薦算法的透明度與公平性 (Transparency)。",
    "維護企業文化的倫理邊界，防止過度依賴黑盒算法。"
]), { x: 0.7, y: 1.8, w: 8.5, h: 3, lineSpacing: 28 });

// 6. Role 3
let slide6 = pres.addSlide({ masterName: "STL_V8_MASTER" });
addHeader(slide6, "職責 3：塑造人機協作效益");
slide6.addText(createContentBlock([
    "在提升生產力的同時，確保參與度不墜 (Engagement)。",
    "優化人類直覺與 AI 計算的平衡 (Human-Machine Collaboration)。",
    "重新設計工作流，讓員工感受到 AI 的賦能。 "
]), { x: 0.7, y: 1.8, w: 8.5, h: 3, lineSpacing: 28 });

// 7. Roadmap
let slide7 = pres.addSlide({ masterName: "STL_V8_MASTER" });
addHeader(slide7, "調整任務：轉型三階段路徑");
slide7.addText(createContentBlock([
    "階段 1：剔除與 AI 重疊的舊有行政負荷 (Strip out legacy work)。",
    "階段 2：透過 AI 強化核心策略責任 (Augment core responsibilities)。",
    "階段 3：擴展至由 AI 推動的新策略領域 (Expand into new strategic work)。"
]), { x: 0.7, y: 1.8, w: 8.5, h: 3, lineSpacing: 32 });

// 8. P1
let slide8 = pres.addSlide({ masterName: "STL_V8_MASTER" });
addHeader(slide8, "P1 剔除行動：回收策略產能");
slide8.addText(createContentBlock([
    "明確定義策略重點區域 (Define strategic focus)。",
    "建立 12-24 個月的自動化路徑圖 (Automation roadmap)。",
    "實施「停止行動」清單以界定 AI 與 HRBP 的界限 (Stop-doing list)。"
]), { x: 0.7, y: 1.8, w: 8.5, h: 3, lineSpacing: 28 });

// 9. P2
let slide9 = pres.addSlide({ masterName: "STL_V8_MASTER" });
addHeader(slide9, "P2 強化行動：AI 賦能高價值產出");
slide9.addText(createContentBlock([
    "更新職能模型，使「AI 準備度」透明化 (AI-readiness)。",
    "利用 AI 生成的預測洞察作為領導對話的基準 (AI inputs)。",
    "定義高效協作流，確保專業指導不偏移。"
]), { x: 0.7, y: 1.8, w: 8.5, h: 3, lineSpacing: 28 });

// 10. P3
let slide10 = pres.addSlide({ masterName: "STL_V8_MASTER" });
addHeader(slide10, "P3 擴展行動：主導新型態策略");
slide10.addText(createContentBlock([
    "啟動與試行「策略人才領袖 (STL)」 pods 小組。",
    "在轉型決策中嵌入 STL 條款，確保決策權限。",
    "主導引導組織文化向「AI 共生」轉型。"
]), { x: 0.7, y: 1.8, w: 8.5, h: 3, lineSpacing: 28 });

// 11. Metrics 1
let slide11 = pres.addSlide({ masterName: "STL_V8_MASTER" });
addHeader(slide11, "成功指標：效率與週期回收");
slide11.addText(createContentBlock([
    "行政任務（入職、報告請求）回收的工時數據。",
    "決策週期大幅縮短 (Cycle time reduction)。",
    "目標：回收 20% 以上產能投入策略事項。"
]), { x: 0.7, y: 1.8, w: 8.5, h: 3, lineSpacing: 28 });

// 12. Metrics 2
let slide12 = pres.addSlide({ masterName: "STL_V8_MASTER" });
addHeader(slide12, "成功指標：人才儲備與流失");
slide12.addText(createContentBlock([
    "關鍵職位接班準備率提升 (% Successor readiness)。",
    "遺憾離職率下降 (Reduction in regrettable attrition)。",
    "角色重新部署而非裁員的成功比例。"
]), { x: 0.7, y: 1.8, w: 8.5, h: 3, lineSpacing: 28 });

// 13. Gartner önerileri
let slide13 = pres.addSlide({ masterName: "STL_V8_MASTER" });
addHeader(slide13, "Gartner 建議：HR 專業提升");
slide13.addText(createContentBlock([
    "1. 增強 AI 趨勢洞察 (Trend Insights)。",
    "2. 識別組織開發與變革契機。",
    "3. 加速項目執行流程管裡。",
    "4. 善用數據驅動決策，優化團隊績效。"
]), { x: 0.7, y: 1.8, w: 8.5, h: 3, lineSpacing: 32 });

// 14. Action for CHRO
let slide14 = pres.addSlide({ masterName: "STL_V8_MASTER" });
addHeader(slide14, "對 CHRO 的最終行動清單");
slide14.addText(createContentBlock([
    "立即重新定義 HRBP 的年度績效目標與職權。",
    "建立透明、具倫理深度且懂 AI 的 HR 部門文化。",
    "投資關鍵 AI 分析工具匯總平台，降低數據獲取成本。"
]), { x: 0.7, y: 1.8, w: 8.5, h: 3, lineSpacing: 32 });

// 15. Mind Map (V8 Stable Tree)
let slideMap = pres.addSlide({ masterName: "STL_V8_MASTER" });
addHeader(slideMap, "全課精華：三層層級心智圖 (V8 Final)");

const RX = 1.0, RY = 2.4;
const L1 = [
    { t: "現狀解析", c: ["51%策略不足", "行政作業佔據", "AI自動自動威脅"] },
    { t: "STL 定義", c: ["人力重設計", "倫理監測", "人機協作效率"] },
    { t: "轉型三階", c: ["P1 剔除行政", "P2 AI 強化", "P3 擴展新域"] },
    { t: "成功衡量", c: ["週期轉短", "接班準備率", "離職優化"] }
];

slideMap.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: RX, y: RY, w: 1.5, h: 0.6, fill: { color: THEME.primary }, rectRadius: 0.1 });
slideMap.addText("HRBP AI 轉型", { x: RX, y: RY, w: 1.5, h: 0.6, color: THEME.white, bold: true, align: "center", fontFace: FONT_TCH, fontSize: 12 });

L1.forEach((node, i) => {
    let nx = RX + 2.4, ny = 1.0 + (i * 1.25);
    // Connect
    slideMap.addShape(pres.shapes.LINE, { x: RX + 1.5, y: RY + 0.3, w: 0.4, h: 0, line: { color: THEME.secondary, width: 2 } });
    slideMap.addShape(pres.shapes.LINE, { x: RX + 1.9, y: Math.min(RY + 0.3, ny + 0.25), w: 0, h: Math.abs(RY + 0.3 - (ny + 0.25)), line: { color: THEME.secondary, width: 2 } });
    slideMap.addShape(pres.shapes.LINE, { x: RX + 1.9, y: ny + 0.25, w: 0.5, h: 0, line: { color: THEME.secondary, width: 2 } });

    slideMap.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: nx, y: ny, w: 1.6, h: 0.5, fill: { color: THEME.secondary }, rectRadius: 0.1 });
    slideMap.addText(node.t, { x: nx, y: ny, w: 1.6, h: 0.5, color: THEME.white, bold: true, align: "center", fontFace: FONT_TCH, fontSize: 11 });

    node.c.forEach((child, j) => {
        let cx = nx + 2.1, cy = ny - 0.2 + (j * 0.45);
        slideMap.addShape(pres.shapes.LINE, { x: nx + 1.6, y: ny + 0.25, w: 0.25, h: 0, line: { color: THEME.line, width: 1 } });
        slideMap.addShape(pres.shapes.LINE, { x: nx + 1.85, y: Math.min(ny + 0.25, cy + 0.15), w: 0, h: Math.abs(ny + 0.25 - (cy + 0.15)), line: { color: THEME.line, width: 1 } });
        slideMap.addShape(pres.shapes.LINE, { x: nx + 1.85, y: cy + 0.15, w: 0.25, h: 0, line: { color: THEME.line, width: 1 } });
        slideMap.addShape(pres.shapes.RECTANGLE, { x: cx, y: cy, w: 1.8, h: 0.35, fill: { color: "F8FAFC" }, line: { color: THEME.line, width: 1 } });
        slideMap.addText(child, { x: cx, y: cy, w: 1.8, h: 0.35, color: THEME.text, fontSize: 9, fontFace: FONT_TCH, valign: "middle" });
    });
});

// 16. Last
let slideLast = pres.addSlide();
slideLast.background = { color: THEME.primary };
slideLast.addText("成功引領組織，重塑未來人才競爭力", { x: 0, y: 2.3, w: "100%", h: 0.6, bold: true, fontSize: 36, color: THEME.white, align: "center", fontFace: FONT_TCH });

// Save
const outDir = path.join("C:\\Users\\TW-Evan.Chen\\.gemini\\antigravity\\scratch\\Skills_Workspace", "Output");
if (!fs.existsSync(outDir)) fs.mkdirSync(outDir);
const finalPath = path.join(outDir, "HRBP_AI_Transformation_Full_v8.pptx");

pres.writeFile({ fileName: finalPath }).then(fn => {
    console.log(`Successfully generated v8 at ${fn}`);
}).catch(err => console.error(err));
