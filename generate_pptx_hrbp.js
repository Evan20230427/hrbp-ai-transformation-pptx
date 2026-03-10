const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

let pres = new pptxgen();
pres.layout = 'LAYOUT_16x9';
pres.author = 'Antigravity AI';
pres.title = '重塑 HRBP 角色以推動 AI 轉型';

// Modern Professional Theme
const THEME = {
    primary: "0F172A",    // Deep Navy
    secondary: "3B82F6",  // Royal Blue
    accent: "10B981",     // Emerald Green
    bg: "F8FAFC",         // Ghost White
    text: "1E293B",       // Slate 800
    white: "FFFFFF"
};

// Define Master Slide
pres.defineSlideMaster({
    title: 'MASTER_SLIDE',
    background: { color: THEME.bg },
    objects: [
        { rect: { x: 0, y: 0, w: "100%", h: 0.1, fill: { color: THEME.secondary } } },
        { text: { text: "Gartner Insight: Redesigning HRBP for AI", options: { x: 0.5, y: 5.3, w: 5, h: 0.3, fontSize: 10, color: "94A3B8" } } }
    ]
});

// Helper for Section Header
function addSectionHeader(slide, title) {
    slide.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 0.5, w: 0.1, h: 0.6, fill: { color: THEME.secondary } });
    slide.addText(title, { x: 0.7, y: 0.5, w: 8, h: 0.6, fontSize: 32, bold: true, color: THEME.primary, fontFace: "Microsoft JhengHei" });
}

// Slide 1: Title Slide
let slide1 = pres.addSlide();
slide1.background = { color: THEME.primary };
slide1.addText("重塑 HRBP 角色：\n成為 AI 轉型顧問", { x: 0.5, y: 1.5, w: 9, h: 1.5, fontSize: 44, bold: true, color: THEME.white, fontFace: "Microsoft JhengHei", align: "center" });
slide1.addText("Redesigning the HRBP Role to Fuel AI Transformation", { x: 0.5, y: 3.0, w: 9, h: 0.5, fontSize: 20, color: THEME.secondary, fontFace: "Arial", align: "center" });
slide1.addShape(pres.shapes.LINE, { x: 3, y: 3.8, w: 4, h: 0, line: { color: THEME.accent, width: 2 } });

// Slide 2: The Challenge
let slide2 = pres.addSlide({ masterName: "MASTER_SLIDE" });
addSectionHeader(slide2, "現況挑戰：策略參與度不足");
slide2.addText([
    { text: "僅 51% 的領導者認為 HRBP 參與了關鍵策略討論\n", options: { bullet: true, fontSize: 20, color: THEME.text, breakLine: true } },
    { text: "多數 HRBP 仍困於「行政事務」與「一般性流程」\n", options: { bullet: true, fontSize: 20, color: THEME.text, breakLine: true } },
    { text: "AI 的出現讓傳統技能迅速貶值，但也提供了轉型契機", options: { bullet: true, fontSize: 20, color: THEME.text } }
], { x: 0.8, y: 1.5, w: 8.5, h: 3, fontFace: "Microsoft JhengHei", lineSpacing: 35 });

// Slide 3: The Target: AI Transformation Consultant
let slide3 = pres.addSlide({ masterName: "MASTER_SLIDE" });
addSectionHeader(slide3, "未來定位：AI 轉型顧問 (STL)");
slide3.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.5, w: 9, h: 3, fill: { color: "FFFFFF" }, shadow: { type: "outer", color: "000000", opacity: 0.05, blur: 10 } });
slide3.addText([
    { text: "1. 引導人力重新設計 (Workforce Redesign)\n", options: { bold: true, color: THEME.secondary, breakLine: true } },
    { text: "   - 主導職務重塑、技能再培訓與人員重新部署策略。\n", options: { fontSize: 16, color: THEME.text, breakLine: true } },
    { text: "2. 解決 AI 倫理與偏見 (AI Bias & Ethics)\n", options: { bold: true, color: THEME.secondary, breakLine: true } },
    { text: "   - 確保 AI 驅動的人才決策透明且符合倫理。\n", options: { fontSize: 16, color: THEME.text, breakLine: true } },
    { text: "3. 塑造人機協作 (Human-Machine Collaboration)\n", options: { bold: true, color: THEME.secondary, breakLine: true } },
    { text: "   - 在提升生產力的同時，兼顧員工參與度與文化認同。", options: { fontSize: 16, color: THEME.text } }
], { x: 0.8, y: 1.7, w: 8.5, h: 2.5, fontFace: "Microsoft JhengHei", lineSpacing: 28 });

// Slide 4: Three Phases of Task Shifts
let slide4 = pres.addSlide({ masterName: "MASTER_SLIDE" });
addSectionHeader(slide4, "轉型三階段實作路徑");
const phases = ["Phase 1: 剔除舊有工作", "Phase 2: 強化核心責任", "Phase 3: 擴展開放領域"];
phases.forEach((p, i) => {
    let y = 1.5 + (i * 1.2);
    slide4.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: y, w: 9, h: 1, fill: { color: (i === 1 ? THEME.secondary : "E2E8F0") }, rectRadius: 0.1 });
    slide4.addText(p, { x: 0.8, y: y, w: 8, h: 1, fontSize: 24, bold: true, color: (i === 1 ? "FFFFFF" : THEME.primary), fontFace: "Microsoft JhengHei", align: "left", valign: "middle" });
});

// Slide 5: Phase 1: Strip Out
let slide5 = pres.addSlide({ masterName: "MASTER_SLIDE" });
addSectionHeader(slide5, "P1：剔除與 AI 重疊的行政工作");
slide5.addText("核心目標：騰出 20-30% 的產能投入策略項目", { x: 0.5, y: 1.3, w: 8, h: 0.4, fontSize: 18, color: THEME.secondary, fontFace: "Microsoft JhengHei" });
slide5.addTable([
    ["關鍵行動", "預期成果"],
    ["建立自動化路徑圖 (Roadmap)", "明確列出 12-24 個月內由 AI 接管的任務"],
    ["定義 AI vs HRBP 邊界", "實施「停止行動」清單，避免重複勞動"]
], { x: 0.5, y: 1.8, w: 9, colW: [4, 5], fill: { color: "FFFFFF" }, border: { pt: 0.5, color: "CBD5E1" }, fontSize: 16, fontFace: "Microsoft JhengHei", margin: 10 });

// Slide 6: Phase 2: Augment
let slide6 = pres.addSlide({ masterName: "MASTER_SLIDE" });
addSectionHeader(slide6, "P2：透過 AI 強化核心策略產出");
slide6.addText([
    { text: "賦能 HRBP 的新武器：\n", options: { bold: true, fontSize: 24, breakLine: true } },
    { text: "• 使用 AI 生成的數據洞察作為業務部門(BU)對話的基礎。\n", options: { breakLine: true } },
    { text: "• 建立 HRBP-CoE-AI 協作流，確保專業知識與 AI 工具完美結合。\n", options: { breakLine: true } },
    { text: "• 更新職能模型：將「AI 準備度」納入 HRBP 的關鍵考核。", options: {} }
], { x: 0.8, y: 1.8, w: 8.5, h: 3, fontSize: 20, color: THEME.text, fontFace: "Microsoft JhengHei" });

// Slide 7: Phase 3: Expand
let slide7 = pres.addSlide({ masterName: "MASTER_SLIDE" });
addSectionHeader(slide7, "P3：擴展至由 AI 推動的新策略領域");
slide7.addText("目標：建立高價值的 STL (Strategic Talent Leader) 模型", { x: 0.5, y: 1.3, w: 9, h: 0.4, fontSize: 18, color: THEME.accent, fontFace: "Microsoft JhengHei" });
slide7.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 2, w: 4.25, h: 2, fill: { color: "FFFFFF" }, line: { color: THEME.secondary, width: 2 } });
slide7.addText("啟動 STL 小組\n選拔具備 AI 洞察力的 HRBP 主導組織變革。", { x: 0.7, y: 2.2, w: 3.8, h: 1.5, fontSize: 16, color: THEME.text, fontFace: "Microsoft JhengHei", align: "center" });
slide7.addShape(pres.shapes.RECTANGLE, { x: 5.25, y: 2, w: 4.25, h: 2, fill: { color: "FFFFFF" }, line: { color: THEME.secondary, width: 2 } });
slide7.addText("嵌入 STL 條款\n在重大 AI 轉型決策中，法律規定必須參考 HRBP 的洞察。", { x: 5.45, y: 2.2, w: 3.8, h: 1.5, fontSize: 16, color: THEME.text, fontFace: "Microsoft JhengHei", align: "center" });

// Slide 8: Final Call to Action
let slide8 = pres.addSlide();
slide8.background = { color: THEME.secondary };
slide8.addText("準備好迎接 AI 時代的 HR 轉型了嗎？", { x: 0.5, y: 2, w: 9, h: 1, fontSize: 36, bold: true, color: THEME.white, fontFace: "Microsoft JhengHei", align: "center" });
slide8.addText("從現在開始，將 HRBP 定位為您的 AI 戰略夥伴。", { x: 0.5, y: 3, w: 9, h: 0.5, fontSize: 18, color: "E0F2FE", fontFace: "Microsoft JhengHei", align: "center" });

// Final Execution Folder Check
const outputDir = "output";
if (!fs.existsSync(outputDir)) {
    fs.mkdirSync(outputDir);
}

const outputPath = path.join(outputDir, "HRBP_AI_Transformation.pptx");
pres.writeFile({ fileName: outputPath }).then(fileName => {
    console.log(`Success: Generated ${fileName}`);
}).catch(err => {
    console.error(`Error: ${err}`);
});
