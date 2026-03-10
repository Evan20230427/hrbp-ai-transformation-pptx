const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

let pres = new pptxgen();
pres.layout = 'LAYOUT_16x9';
pres.author = 'Antigravity AI';
pres.title = 'HRBP AI 轉型完整指南';

// Advanced Professional Theme
const THEME = {
    primary: "0F172A",    // Deep Navy
    secondary: "3B82F6",  // Royal Blue
    accent: "10B981",     // Emerald Green
    bg: "F8FAFC",         // Ghost White
    text: "1E293B",       // Slate 800
    subtle: "94A3B8",     // Slate 400
    white: "FFFFFF",
    highlight: "F59E0B"   // Amber
};

// Define Master Slide
pres.defineSlideMaster({
    title: 'MASTER_SLIDE',
    background: { color: THEME.bg },
    objects: [
        { rect: { x: 0, y: 0, w: "100%", h: 0.1, fill: { color: THEME.secondary } } },
        { text: { text: "© 2026 Gartner Insight | HRBP AI Transformation Guide", options: { x: 0.5, y: 5.3, w: 9, h: 0.3, fontSize: 10, color: THEME.subtle, align: "right" } } }
    ]
});

// Helper for Section Header
function addSectionHeader(slide, title, subtitle = "") {
    slide.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 0.4, w: 0.1, h: 0.7, fill: { color: THEME.secondary } });
    slide.addText(title, { x: 0.7, y: 0.35, w: 8, h: 0.5, fontSize: 28, bold: true, color: THEME.primary, fontFace: "Microsoft JhengHei" });
    if (subtitle) {
        slide.addText(subtitle, { x: 0.7, y: 0.8, w: 8, h: 0.3, fontSize: 14, color: THEME.secondary, fontFace: "Microsoft JhengHei" });
    }
}

// --- Slide 1: Title ---
let slide1 = pres.addSlide();
slide1.background = { color: THEME.primary };
slide1.addText("重塑 HRBP 角色：\n引領 AI 轉型的策略實務", { x: 0.5, y: 1.5, w: 9, h: 1.5, fontSize: 42, bold: true, color: THEME.white, fontFace: "Microsoft JhengHei", align: "center" });
slide1.addText("全面指南：從行政事務到策略人才領袖 (STL)", { x: 0.5, y: 3.0, w: 9, h: 0.5, fontSize: 18, color: THEME.secondary, fontFace: "Microsoft JhengHei", align: "center" });

// --- Slide 2: Challenges ---
let slide2 = pres.addSlide({ masterName: "MASTER_SLIDE" });
addSectionHeader(slide2, "轉型的迫切性：打破行政枷鎖");
slide2.addText([
    { text: "現狀揭露：\n", options: { bold: true, color: THEME.highlight, fontSize: 22, breakLine: true } },
    { text: "• 僅 51% 領導者滿意 HRBP 的策略貢獻。\n", options: { bullet: true, breakLine: true } },
    { text: "• 大量工時消耗在：職缺說明撰寫、基礎數據匯總、政策諮詢。\n", options: { bullet: true, breakLine: true } },
    { text: "• 風險：若不轉型，AI 將重複 HRBP 既有任務，使其失去組織價值。\n", options: { bullet: true, breakLine: true } },
    { text: "關鍵行動：CHRO 必須在 AI 時代重新定義 HRBP 的「策略相關性」。", options: { bullet: true } }
], { x: 0.8, y: 1.5, w: 8.5, h: 3, fontFace: "Microsoft JhengHei", lineSpacing: 32 });

// --- Slide 3: New Role Defined ---
let slide3 = pres.addSlide({ masterName: "MASTER_SLIDE" });
addSectionHeader(slide3, "未來 HRBP：AI 轉型顧問 (STL)");
slide3.addText("不僅是「翻譯」人力策略，而是「主導」轉型方向。", { x: 0.5, y: 1.2, w: 8, h: 0.3, fontSize: 16, color: THEME.subtle, fontFace: "Microsoft JhengHei" });
const roles = [
    { t: "引導人力設計", d: "判斷何時需重塑職能、再培訓或淘汰職位。" },
    { t: "倫理與偏見監測", d: "解決 AI 驅動決策中的透明度與公平性問題。" },
    { t: "人機協作優化", d: "在技術導入時確保員工生產力與心理契約同步提升。" }
];
roles.forEach((r, i) => {
    let y = 1.8 + (i * 1.1);
    slide3.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: y, w: 9, h: 0.9, fill: { color: "FFFFFF" }, shadow: { type: "outer", color: "000000", opacity: 0.05, blur: 5 } });
    slide3.addText(r.t, { x: 0.7, y: y + 0.15, w: 3, h: 0.3, fontSize: 18, bold: true, color: THEME.secondary, fontFace: "Microsoft JhengHei" });
    slide3.addText(r.d, { x: 0.7, y: y + 0.45, w: 8, h: 0.3, fontSize: 14, color: THEME.text, fontFace: "Microsoft JhengHei" });
});

// --- Slide 4: 3-Phase Roadmap ---
let slide4 = pres.addSlide({ masterName: "MASTER_SLIDE" });
addSectionHeader(slide4, "轉型導圖：三階段任務配置");
const navObjs = [
    { n: "P1: 剔除", c: "移除行政負擔" },
    { n: "P2: 強化", c: "AI 賦能決策" },
    { n: "P3: 擴展", c: "主導策略領袖" }
];
navObjs.forEach((o, i) => {
    let x = 0.5 + (i * 3.1);
    slide4.addShape(pres.shapes.RECTANGLE, { x: x, y: 1.5, w: 2.8, h: 3, fill: { color: THEME.primary }, rectRadius: 0.2 });
    slide4.addText(o.n, { x: x, y: 2.0, w: 2.8, h: 0.5, fontSize: 24, bold: true, color: THEME.secondary, align: "center", fontFace: "Microsoft JhengHei" });
    slide4.addText(o.c, { x: x, y: 3.0, w: 2.8, h: 0.5, fontSize: 16, color: THEME.white, align: "center", fontFace: "Microsoft JhengHei" });
});

// --- Slide 5: Phase 1 Detail ---
let slide5 = pres.addSlide({ masterName: "MASTER_SLIDE" });
addSectionHeader(slide5, "第一階段：剔除 (Strip Out)", "移除與 AI 重疊的低價值工作");
slide5.addTable([
    [{ text: "關鍵行動", options: { bold: true, fill: "E2E8F0" } }, { text: "實施策略", options: { bold: true, fill: "E2E8F0" } }],
    ["確立策略重點", "明確 HRBP 應投入時間的目標領域。"],
    ["自動化路線圖", "根據任務準備度，制定 12-24 個月的移交計畫。"],
    ["定義任務邊界", "使用「停止行動」清單，明確 AI 與人力的界限。"]
], { x: 0.5, y: 1.5, w: 9, colW: [3, 6], fontSize: 16, fontFace: "Microsoft JhengHei", margin: 10, border: { color: "CBD5E1" } });

// --- Slide 6: Phase 2 Detail ---
let slide6 = pres.addSlide({ masterName: "MASTER_SLIDE" });
addSectionHeader(slide6, "第二階段：強化 (Augment)", "以 AI 賦能核心責任");
slide6.addText([
    { text: "更新職能模型：\n", options: { bold: true, color: THEME.secondary, breakLine: true } },
    { text: "   • 確保「AI 準備度」成為 HRBP 評估的標配。\n", options: { breakLine: true } },
    { text: "優化對話品質：\n", options: { bold: true, color: THEME.secondary, breakLine: true } },
    { text: "   • 利用 AI 生成的洞察作為與 BU 領導者交談的基線訊息。\n", options: { breakLine: true } },
    { text: "協作流程設計：\n", options: { bold: true, color: THEME.secondary, breakLine: true } },
    { text: "   • 定義 HRBP-CoE-AI 工作流，最大化技術槓桿。", options: {} }
], { x: 0.8, y: 1.5, w: 8.5, h: 3, fontSize: 20, fontFace: "Microsoft JhengHei", lineSpacing: 30 });

// --- Slide 7: Phase 3 Detail ---
let slide7 = pres.addSlide({ masterName: "MASTER_SLIDE" });
addSectionHeader(slide7, "第三階段：擴展 (Expand)", "開拓 AI 驅動的新策略領域");
slide7.addTable([
    ["指標類型", "具體項目"],
    ["效率指標", "人力培訓/決策週期縮短時間。"],
    ["質量指標", "18 個月內關鍵職位接班準備率提升。"],
    ["轉型指標", "STL 小組主導的重大轉型決策件數。"]
], { x: 0.5, y: 1.5, w: 9, fontSize: 16, fontFace: "Microsoft JhengHei", margin: 8, border: { color: "CBD5E1" } });

// --- Slide 8: Success Metrics ---
let slide8 = pres.addSlide({ masterName: "MASTER_SLIDE" });
addSectionHeader(slide8, "成功衡量標準 (Success Metrics)");
const metrics = [
    "✅ 週期時間縮短 (Cycle Time Reduction)",
    "✅ 關鍵職位儲備增加 (% Increase in Successors)",
    "✅ 高風險流失率降低 (Reduction in Regrettable Attrition)",
    "✅ 偏差案例減少 (Reduced AI Bias Cases)"
];
metrics.forEach((m, i) => {
    slide8.addText(m, { x: 1, y: 1.5 + (i * 0.7), w: 8, h: 0.5, fontSize: 20, color: THEME.primary, fontFace: "Microsoft JhengHei" });
});

// --- Slide 9: Mind Map (Final) ---
let slide9 = pres.addSlide({ masterName: "MASTER_SLIDE" });
addSectionHeader(slide9, "課程心智圖 (Concept Map)");

// Center Root
const rootX = 4.0, rootY = 2.5, rootW = 2.0, rootH = 0.8;
slide9.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: rootX, y: rootY, w: rootW, h: rootH, fill: { color: THEME.primary }, rectRadius: 0.1 });
slide9.addText("HRBP AI 轉型", { x: rootX, y: rootY, w: rootW, h: rootH, color: THEME.white, bold: true, align: "center", fontSize: 16, fontFace: "Microsoft JhengHei" });

// Branches
const branches = [
    { t: "現況挑戰", x: 1.0, y: 1.2, color: THEME.highlight },
    { t: "顧問定位", x: 7.0, y: 1.2, color: THEME.secondary },
    { t: "實作階段", x: 1.0, y: 4.0, color: THEME.accent },
    { t: "成功衡量", x: 7.0, y: 4.0, color: THEME.secondary }
];

branches.forEach((b, i) => {
    // Branch node
    slide9.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: b.x, y: b.y, w: 1.5, h: 0.5, fill: { color: b.color }, rectRadius: 0.1 });
    slide9.addText(b.t, { x: b.x, y: b.y, w: 1.5, h: 0.5, color: THEME.white, bold: true, align: "center", fontSize: 12, fontFace: "Microsoft JhengHei" });

    // Connect to leaf node placeholders (visual lines)
    slide9.addShape(pres.shapes.LINE, { x: rootX + 1, y: rootY + 0.4, w: (b.x > rootX ? (b.x - (rootX + 1)) : -(rootX - (b.x + 1.5))), h: (b.y - rootY), line: { color: "CBD5E1", width: 1 } });
});

// Final Slide: Appendix
let slide10 = pres.addSlide();
slide10.background = { color: THEME.primary };
slide10.addText("Q&A 與 結語", { x: 0.5, y: 2, w: 9, h: 1, fontSize: 36, bold: true, color: THEME.white, fontFace: "Microsoft JhengHei", align: "center" });
slide10.addText("準備好定義您的帶領地位了嗎？成果已備妥。", { x: 0.5, y: 3, w: 9, h: 0.5, fontSize: 18, color: THEME.subtle, fontFace: "Microsoft JhengHei", align: "center" });

// Save
const outputDir = path.join(__dirname, "output");
if (!fs.existsSync(outputDir)) fs.mkdirSync(outputDir);
const finalPath = path.join(outputDir, "HRBP_AI_Transformation_Expanded.pptx");

pres.writeFile({ fileName: finalPath }).then(fileName => {
    console.log(`Success: Generated ${fileName}`);
}).catch(err => {
    console.error(`Error: ${err}`);
});
