const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

/**
 * [Theme Factory] 整合系統
 * 核心目標：根據內容關鍵字動態調整風格與配色
 */
const THEMES = {
    "tech": {
        name: "Tech Innovation",
        primary: "1E1E1E", secondary: "0066FF", accent: "00FFFF", text: "FFFFFF",
        fontTitle: "DejaVu Sans Bold", fontBody: "DejaVu Sans"
    },
    "ocean": {
        name: "Ocean Depths",
        primary: "1A2332", secondary: "2D8B8B", accent: "A8DADC", text: "F1FAEE",
        fontTitle: "DejaVu Sans Bold", fontBody: "DejaVu Sans"
    },
    "default": {
        name: "Roman Classic v23",
        primary: "F1F5F9", secondary: "B45309", accent: "475569", text: "1E293B",
        fontTitle: "Microsoft JhengHei", fontBody: "Arial"
    }
};

// 1. 載入 OCR 內容並辨識風格需求
const JSON_DATA_PATH = path.join(__dirname, "extracted_content.json");
if (!fs.existsSync(JSON_DATA_PATH)) {
    console.error("[ERROR] extracted_content.json not found.");
    process.exit(1);
}
const rawData = JSON.parse(fs.readFileSync(JSON_DATA_PATH, "utf-8"));

// 智慧辨識主題：若關鍵字包含科技、AI -> 科技風；若包含專業、流程、管理 -> 專業海洋風
const fullText = rawData.map(p => p.text).join(" ");
let activeTheme = THEMES.default;
if (fullText.includes("AI") || fullText.includes("技術") || fullText.includes("科技") || fullText.includes("Microsoft")) {
    activeTheme = THEMES.tech;
} else if (fullText.includes("管理") || fullText.includes("專業") || fullText.includes("商務")) {
    activeTheme = THEMES.ocean;
}

console.log(`[STYLE] Detected Content. Applying Theme: ${activeTheme.name}`);

// 2. 初始化簡報引擎
let pres = new pptxgen();
pres.layout = 'LAYOUT_16x9';
pres.author = 'Antigravity Dynamic Engine';
pres.title = activeTheme.name + " Render";

// 3. 通用視覺組件 (羅馬美學延續)
const IMAGES = [
    "C:\\Users\\TW-Evan.Chen\\.gemini\\antigravity\\brain\\c7d2daf0-cf50-4ec0-af51-75cb91a0c238\\hrbp_v21_cover_1773679946141.png",
    "C:\\Users\\TW-Evan.Chen\\.gemini\\antigravity\\brain\\c7d2daf0-cf50-4ec0-af51-75cb91a0c238\\v23_statue_stoic_1773681336943.png",
    "C:\\Users\\TW-Evan.Chen\\.gemini\\antigravity\\brain\\c7d2daf0-cf50-4ec0-af51-75cb91a0c238\\hrbp_v22_statue_laugh_1773680611632.png",
    "C:\\Users\\TW-Evan.Chen\\.gemini\\antigravity\\brain\\ac6f3712-34eb-4b1e-b0f3-0fb2ebf77cf6\\v23_statue_curious_ac6f3712_png_1773711545138.png"
];

function applySlideFrame(slide, title, pageNum) {
    // 頂部導航條
    slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: '100%', h: 0.8, fill: { color: activeTheme.secondary } });
    slide.addText(title, { x: 0.5, y: 0.2, w: 9, h: 0.4, fontSize: 24, bold: true, color: activeTheme.text, fontFace: activeTheme.fontTitle });
    
    // 頁次 footer
    slide.addText(`MDL | ${activeTheme.name} | Page ${pageNum}`, { x: 8.5, y: 5.2, w: 1.2, h: 0.3, fontSize: 10, color: activeTheme.accent, align: "right" });
}

// 4. 定義 Master Slide (前端美學分層)
pres.defineSlideMaster({
    title: 'DYNAMIC_MASTER',
    background: { color: activeTheme.primary },
    objects: [{ rect: { x: 0, y: 0, w: 0.1, h: "100%", fill: { color: activeTheme.accent } } }]
});

// 5. 簡報生成邏輯
const mainTitle = rawData[0].text.split("\n")[0].substring(0, 40) || "AI 轉型方案";

// Cover
let cover = pres.addSlide({ masterName: 'DYNAMIC_MASTER' });
cover.addText(mainTitle, { x: 0.8, y: 2.2, w: 8.4, fontSize: 44, bold: true, color: activeTheme.secondary, align: "center", fontFace: activeTheme.fontTitle });
cover.addShape(pres.shapes.LINE, { x: 2.5, y: 3.2, w: 5, h: 0, line: { color: activeTheme.accent, width: 2 } });
cover.addText(activeTheme.name + " - High Fidelity AI Report", { x: 0.8, y: 3.5, w: 8.4, fontSize: 18, color: activeTheme.accent, align: "center" });

// Content Pages
rawData.slice(1, 15).forEach((pData, idx) => {
    let slide = pres.addSlide({ masterName: 'DYNAMIC_MASTER' });
    let pageNum = idx + 2;
    
    const lines = pData.text.split(" ").filter(l => l.length > 2);
    const pageTitle = lines[0] ? lines[0].substring(0, 25) : `解構分析 ${pageNum}`;
    const bodyPoints = lines.slice(1, 10).join(" ").substring(0, 400);

    applySlideFrame(slide, pageTitle, pageNum);

    // [Frontend-Design] 左右佈局
    const isOdd = pageNum % 2 !== 0;
    slide.addText(bodyPoints, { 
        x: isOdd ? 5.2 : 0.5, y: 1.2, w: 4.3, h: 3.5, 
        fontSize: 16, color: activeTheme.text, fontFace: activeTheme.fontBody, lineSpacing: 22 
    });

    if (IMAGES[idx % IMAGES.length]) {
        slide.addImage({ 
            path: IMAGES[idx % IMAGES.length], 
            x: isOdd ? 0.5 : 5.2, y: 1.2, w: 4.3, h: 3.5,
            sizing: { type: 'contain' } 
        });
    }
});

// 6. 輸出
const safeFilename = mainTitle.replace(/[\\/:"*?<>|]/g, "_");
const outPath = path.join(__dirname, "output", `${safeFilename}_Dynamic_${activeTheme.name.replace(/ /g, "_")}.pptx`);

pres.writeFile({ fileName: outPath }).then(fn => {
    console.log(`[SUCCESS] Generated Artifact at: ${fn}`);
}).catch(err => {
    console.error(`[ERROR] Render failed: ${err.message}`);
    process.exit(1);
});
