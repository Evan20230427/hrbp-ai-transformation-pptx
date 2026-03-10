const pptxgen = require("pptxgenjs");

let pres = new pptxgen();
pres.layout = 'LAYOUT_16x9';
pres.author = 'Evan Chen';
pres.title = 'HR招募系統Chatbox 進度更新';

// Defined Theme: Tech Innovation (from theme-factory context)
// Primary: 1E293B (Slate)
// Secondary: 0EA5E9 (Sky Blue)
// Accent: F59E0B (Amber)
// Background: F8FAFC (Light slate)

const THEME = {
    primary: "1E293B",
    secondary: "0EA5E9",
    accent: "F59E0B",
    bg: "F8FAFC",
    text: "334155",
    lightText: "94A3B8"
};

// Master Slide for consistent styling
pres.defineSlideMaster({
    title: 'DEFAULT_SLIDE',
    background: { color: THEME.bg },
    objects: [
        { rect: { x: 0, y: 0, w: "100%", h: 0.15, fill: { color: THEME.secondary } } }, // Top accent bar
        { rect: { x: 0.5, y: 5.2, w: 9, h: 0.05, fill: { color: "E2E8F0" } } }, // Bottom subtle line
        { text: { text: "HR Chatbox Project Update | 2026", options: { x: 0.5, y: 5.3, w: 4, h: 0.2, fontSize: 10, color: THEME.lightText } } }
    ]
});

// Slide 1: Title
let slide1 = pres.addSlide();
slide1.background = { color: THEME.primary };
slide1.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 2.5, w: 0.1, h: 1.2, fill: { color: THEME.secondary } });
slide1.addText("HR招募系統Chatbox", { x: 0.8, y: 2.3, w: 8, h: 0.8, fontSize: 44, bold: true, color: "FFFFFF", fontFace: "Microsoft JhengHei" });
slide1.addText("專案範圍與報價進度更新說明", { x: 0.8, y: 3.1, w: 8, h: 0.5, fontSize: 24, color: "E0F2FE", fontFace: "Microsoft JhengHei" });
slide1.addText("2026年3月10日", { x: 0.8, y: 3.8, w: 8, h: 0.3, fontSize: 14, color: THEME.lightText, fontFace: "Arial" });

// Slide 2: Executive Summary
let slide2 = pres.addSlide({ masterName: "DEFAULT_SLIDE" });
slide2.addText("重點摘要", { x: 0.5, y: 0.5, w: 8, h: 0.6, fontSize: 32, bold: true, color: THEME.primary, fontFace: "Microsoft JhengHei" });
slide2.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.5, w: 4.25, h: 3, fill: { color: "FFFFFF" }, shadow: { type: "outer", color: "000000", blur: 5, offset: 2, angle: 45, opacity: 0.1 } });
slide2.addText("專案現況", { x: 0.7, y: 1.7, w: 3.8, h: 0.4, fontSize: 20, bold: true, color: THEME.secondary, fontFace: "Microsoft JhengHei" });
slide2.addText([
    { text: "開發成本與風險可控\n", options: { bullet: true, breakLine: true } },
    { text: "跳過 POC，直接進入報價遴選\n", options: { bullet: true, breakLine: true } },
    { text: "目標上線：2026 年 6 月\n", options: { bullet: true, breakLine: true } },
    { text: "預算：100 萬", options: { bullet: true } }
], { x: 0.7, y: 2.3, w: 3.8, h: 2, fontSize: 16, color: THEME.text, fontFace: "Microsoft JhengHei", lineSpacing: 24 });

slide2.addShape(pres.shapes.RECTANGLE, { x: 5.25, y: 1.5, w: 4.25, h: 3, fill: { color: "FFFFFF" }, shadow: { type: "outer", color: "000000", blur: 5, offset: 2, angle: 45, opacity: 0.1 } });
slide2.addText("後續行動", { x: 5.45, y: 1.7, w: 3.8, h: 0.4, fontSize: 20, bold: true, color: THEME.accent, fontFace: "Microsoft JhengHei" });
slide2.addText([
    { text: "邀請 4 家廠商提供報價與規格\n", options: { bullet: true, breakLine: true } },
    { text: "包含：極致、艾卡拉等\n", options: { bullet: true, breakLine: true } },
    { text: "不新增第 5 家廠商\n", options: { bullet: true, breakLine: true } },
    { text: "安排會議確認 DB API 串接", options: { bullet: true } }
], { x: 5.45, y: 2.3, w: 3.8, h: 2, fontSize: 16, color: THEME.text, fontFace: "Microsoft JhengHei", lineSpacing: 24 });

// Slide 3: Phase 1 Scope
let slide3 = pres.addSlide({ masterName: "DEFAULT_SLIDE" });
slide3.addText("Phase 1 專案範圍釐清", { x: 0.5, y: 0.5, w: 8, h: 0.6, fontSize: 32, bold: true, color: THEME.primary, fontFace: "Microsoft JhengHei" });

slide3.addTable([
    [
        { text: "項目", options: { fill: { color: THEME.secondary }, color: "FFFFFF", bold: true, fontFace: "Microsoft JhengHei", fontSize: 16, align: "center" } },
        { text: "目前結論", options: { fill: { color: THEME.secondary }, color: "FFFFFF", bold: true, fontFace: "Microsoft JhengHei", fontSize: 16, align: "center" } }
    ],
    [
        { text: "第一階段核心", options: { fontFace: "Microsoft JhengHei", fontSize: 14, color: THEME.text, bold: true } },
        { text: "聚焦於 Part 1（AI 回答功能）作為現階段報價基準。", options: { fontFace: "Microsoft JhengHei", fontSize: 14, color: THEME.text } }
    ],
    [
        { text: "雇主品牌與職缺", options: { fontFace: "Microsoft JhengHei", fontSize: 14, color: THEME.text, bold: true, fill: { color: "F1F5F9" } } },
        { text: "不限於招募網站，將一併評估納入「職前簡介資訊」。未來分階段擴充。", options: { fontFace: "Microsoft JhengHei", fontSize: 14, color: THEME.text, fill: { color: "F1F5F9" } } }
    ],
    [
        { text: "職缺媒合功能", options: { fontFace: "Microsoft JhengHei", fontSize: 14, color: THEME.text, bold: true } },
        { text: "需確認資料庫與 API。交通通勤等額外資訊才需介接新 API。將另開會議說明系統流程。", options: { fontFace: "Microsoft JhengHei", fontSize: 14, color: THEME.text } }
    ],
    [
        { text: "FAQ 後續流程", options: { fontFace: "Microsoft JhengHei", fontSize: 14, color: THEME.text, bold: true, fill: { color: "F1F5F9" } } },
        { text: "目前不在本次專案範圍，但須確認方向一致。", options: { fontFace: "Microsoft JhengHei", fontSize: 14, color: THEME.text, fill: { color: "F1F5F9" } } }
    ]
], { x: 0.5, y: 1.5, w: 9, colW: [2.5, 6.5], border: { pt: 1, color: "CBD5E1" }, valign: "middle", margin: 10 });


// Slide 4: Quotation Requirements
let slide4 = pres.addSlide({ masterName: "DEFAULT_SLIDE" });
slide4.addText("報價單需求項目", { x: 0.5, y: 0.5, w: 8, h: 0.6, fontSize: 32, bold: true, color: THEME.primary, fontFace: "Microsoft JhengHei" });

slide4.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.5, w: 9, h: 0.6, fill: { color: "F1F5F9" }, rectRadius: 0.1 });
slide4.addText("請 4 家廠商於報價單與提案規格書中，必須完整涵蓋以下可能產生的費用，以利後續預算評估：", { x: 0.7, y: 1.6, w: 8.6, h: 0.4, fontSize: 16, color: THEME.text, fontFace: "Microsoft JhengHei" });

const reqs = ["一次性開發費", "雲端或地端建置費 (每月)", "系統維護費", "更新費 (人天計算)", "Token 費用估算"];
reqs.forEach((req, idx) => {
    let xPos = 0.5 + (idx % 3) * 3.1;
    let yPos = 2.5 + Math.floor(idx / 3) * 1.5;
    slide4.addShape(pres.shapes.RECTANGLE, { x: xPos, y: yPos, w: 2.8, h: 1, fill: { color: "FFFFFF" }, shadow: { type: "outer", color: "000000", blur: 4, offset: 2, angle: 45, opacity: 0.1 } });
    slide4.addShape(pres.shapes.RECTANGLE, { x: xPos, y: yPos, w: 0.1, h: 1, fill: { color: THEME.accent } });
    slide4.addText(req, { x: xPos + 0.3, y: yPos + 0.3, w: 2.4, h: 0.4, fontSize: 18, bold: true, color: THEME.primary, fontFace: "Microsoft JhengHei" });
});

// Save presentation
pres.writeFile({ fileName: "output/HR_Chatbox_Update.pptx" }).then(fileName => {
    console.log(`Created presentation: ${fileName}`);
});
