const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

// 動態讀取數據備份
const JSON_DATA_PATH = path.join(__dirname, "extracted_content.json");
if (!fs.existsSync(JSON_DATA_PATH)) {
    console.error("[ERROR] extracted_content.json not found. Please run OCR first.");
    process.exit(1);
}

const rawData = JSON.parse(fs.readFileSync(JSON_DATA_PATH, "utf-8"));

// 提取主標題 (從第一頁的前幾行提取)
const firstPageLines = rawData[0].text.split("\n").filter(l => l.trim().length > 2);
const mainTitle = firstPageLines[0] || "AI 轉型實務簡報";
const subTitle = firstPageLines[1] || "羅馬美學 v23 動態驅動版";

let pres = new pptxgen();
pres.layout = 'LAYOUT_16x9';
pres.author = 'Antigravity AI';
pres.title = mainTitle;

const THEME = {
    primary: "F1F5F9", secondary: "B45309", text: "1E293B",
    white: "FFFFFF", line: "D1D5DB", accent: "475569", highlight: "FEF3C7"
};

const FONT_TITLE = "Microsoft JhengHei";
const FONT_BODY = "Arial";

// 25 張唯一圖資庫 (與先前一致)
const IMAGES = [
    "C:\\Users\\TW-Evan.Chen\\.gemini\\antigravity\\brain\\c7d2daf0-cf50-4ec0-af51-75cb91a0c238\\hrbp_v21_cover_1773679946141.png",
    "C:\\Users\\TW-Evan.Chen\\.gemini\\antigravity\\brain\\c7d2daf0-cf50-4ec0-af51-75cb91a0c238\\hrbp_transformation_mindmap_infographic_1773157697477.png",
    "C:\\Users\\TW-Evan.Chen\\.gemini\\antigravity\\brain\\c7d2daf0-cf50-4ec0-af51-75cb91a0c238\\v23_statue_stoic_1773681336943.png",
    "C:\\Users\\TW-Evan.Chen\\.gemini\\antigravity\\brain\\c7d2daf0-cf50-4ec0-af51-75cb91a0c238\\hrbp_v21_data_insight_1773680231164.png",
    "C:\\Users\\TW-Evan.Chen\\.gemini\\antigravity\\brain\\c7d2daf0-cf50-4ec0-af51-75cb91a0c238\\hrbp_v22_statue_bored_1773680700394.png",
    "C:\\Users\\TW-Evan.Chen\\.gemini\\antigravity\\brain\\c7d2daf0-cf50-4ec0-af51-75cb91a0c238\\hrbp_v22_statue_shock_1773680596901.png",
    "C:\\Users\\TW-Evan.Chen\\.gemini\\antigravity\\brain\\c7d2daf0-cf50-4ec0-af51-75cb91a0c238\\hrbp_v22_statue_laugh_1773680611632.png",
    "C:\\Users\\TW-Evan.Chen\\.gemini\\antigravity\\brain\\c7d2daf0-cf50-4ec0-af51-75cb91a0c238\\hrbp_v21_product_manager_1773680252018.png",
    "C:\\Users\\TW-Evan.Chen\\.gemini\\antigravity\\brain\\c7d2daf0-cf50-4ec0-af51-75cb91a0c238\\hrbp_v22_statue_wow_1773680679782.png",
    "C:\\Users\\TW-Evan.Chen\\.gemini\\antigravity\\brain\\c7d2daf0-cf50-4ec0-af51-75cb91a0c238\\hrbp_ethical_ai_balance_v10_1773159607684.png",
    "C:\\Users\\TW-Evan.Chen\\.gemini\\antigravity\\brain\\c7d2daf0-cf50-4ec0-af51-75cb91a0c238\\v23_statue_fear_1773681354253.png",
    "C:\\Users\\TW-Evan.Chen\\.gemini\\antigravity\\brain\\c7d2daf0-cf50-4ec0-af51-75cb91a0c238\\hrbp_workforce_redesign_v13_simple_1773160725232.png",
    "C:\\Users\\TW-Evan.Chen\\.gemini\\antigravity\\brain\\c7d2daf0-cf50-4ec0-af51-75cb91a0c238\\v23_statue_proud_1773681369773.png",
    "C:\\Users\\TW-Evan.Chen\\.gemini\\antigravity\\brain\\c7d2daf0-cf50-4ec0-af51-75cb91a0c238\\hrbp_visionary_ladder_v9_1773159135984.png",
    "C:\\Users\\TW-Evan.Chen\\.gemini\\antigravity\\brain\\c7d2daf0-cf50-4ec0-af51-75cb91a0c238\\hrbp_v21_ethics_governance_1773679974929.png",
    "C:\\Users\\TW-Evan.Chen\\.gemini\\antigravity\\brain\\c7d2daf0-cf50-4ec0-af51-75cb91a0c238\\hrbp_professional_human_ai_v10_1773159552183.png",
    "C:\\Users\\TW-Evan.Chen\\.gemini\\antigravity\\brain\\c7d2daf0-cf50-4ec0-af51-75cb91a0c238\\hrbp_v22_statue_think_1773680624839.png",
    "C:\\Users\\TW-Evan.Chen\\.gemini\\antigravity\\brain\\c7d2daf0-cf50-4ec0-af51-75cb91a0c238\\hrbp_v22_statue_angry_1773680640224.png",
    "C:\\Users\\TW-Evan.Chen\\.gemini\\antigravity\\brain\\c7d2daf0-cf50-4ec0-af51-75cb91a0c238\\hrbp_v21_roadmap_roman_1773680269997.png",
    "C:\\Users\\TW-Evan.Chen\\.gemini\\antigravity\\brain\\c7d2daf0-cf50-4ec0-af51-75cb91a0c238\\hrbp_transformation_journey_v10_1773159569827.png",
    "C:\\Users\\TW-Evan.Chen\\.gemini\\antigravity\\brain\\c7d2daf0-cf50-4ec0-af51-75cb91a0c238\\hrbp_ai_ethics_v13_simple_1773160703050.png",
    "C:\\Users\\TW-Evan.Chen\\.gemini\\antigravity\\brain\\ac6f3712-34eb-4b1e-b0f3-0fb2ebf77cf6\\v23_statue_curious_ac6f3712_png_1773711545138.png",
    "C:\\Users\\TW-Evan.Chen\\.gemini\\antigravity\\brain\\ac6f3712-34eb-4b1e-b0f3-0fb2ebf77cf6\\v23_statue_excitement_ac6f3712_png_1773711561255.png",
    "C:\\Users\\TW-Evan.Chen\\.gemini\\antigravity\\brain\\c7d2daf0-cf50-4ec0-af51-75cb91a0c238\\hrbp_v21_stl_role_1773679960676.png",
    "C:\\Users\\TW-Evan.Chen\\.gemini\\antigravity\\brain\\c7d2daf0-cf50-4ec0-af51-75cb91a0c238\\hrbp_v22_statue_sad_1773680656123.png"
];

function getImageForSlide(idx) {
    const img = IMAGES[idx % IMAGES.length];
    return fs.existsSync(img) ? img : null;
}

const LAYOUT = {
    header: { x: 0.5, y: 0.4, w: 9.0, h: 0.6 },
    line: { x: 0.5, y: 1.1, w: 1.5, h: 0.05 },
    main: {
        leftImg: { x: 0.5, y: 1.4, w: 4.2, h: 3.8 },
        rightText: { x: 5.0, y: 1.4, w: 4.5, h: 3.8 },
        rightImg: { x: 5.3, y: 1.4, w: 4.2, h: 3.8 },
        leftText: { x: 0.6, y: 1.4, w: 4.5, h: 3.8 }
    },
    footer: { x: 8.5, y: 5.2, w: 1.2, h: 0.3 }
};

function applySlideFrame(slide, title, pageNum) {
    slide.addText(title, { x: LAYOUT.header.x, y: LAYOUT.header.y, w: LAYOUT.header.w, h: LAYOUT.header.h, fontSize: 22, bold: true, color: THEME.secondary, fontFace: FONT_TITLE });
    slide.addShape(pres.shapes.RECTANGLE, { x: LAYOUT.line.x, y: LAYOUT.line.y, w: LAYOUT.line.w, h: LAYOUT.line.h, fill: { color: THEME.secondary } });
    slide.addText(`MDL | v23 Dynamic | Slide ${pageNum}`, { x: LAYOUT.footer.x, y: LAYOUT.footer.y, w: LAYOUT.footer.w, h: LAYOUT.footer.h, fontSize: 10, color: THEME.accent, align: "right", fontFace: FONT_BODY });
}

pres.defineSlideMaster({
    title: 'ROMAN_V23',
    background: { color: THEME.primary },
    objects: [{ rect: { x: 0, y: 0, w: 0.08, h: "100%", fill: { color: THEME.secondary } } }]
});

// 1. Cover
let cover = pres.addSlide({ masterName: 'ROMAN_V23' });
cover.addImage({ path: getImageForSlide(0), x: 0, y: 0, w: '100%', h: '100%' });
cover.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 3.5, w: 9.0, h: 1.6, fill: { color: 'FFFFFF', transparency: 15 } });
cover.addText(mainTitle, { x: 0.8, y: 3.7, w: 8.4, fontSize: 32, bold: true, color: THEME.secondary, fontFace: FONT_TITLE, align: "center" });
cover.addText(subTitle, { x: 0.8, y: 4.5, w: 8.4, fontSize: 18, color: THEME.text, fontFace: FONT_TITLE, align: "center" });

// 2. 動態生成各頁
rawData.slice(1).forEach((pData, idx) => {
    let slide = pres.addSlide({ masterName: 'ROMAN_V23' });
    let pageNum = idx + 2;
    
    // 簡單提取該頁第一行作為標題
    const lines = pData.text.split("\n").filter(l => l.trim().length > 3);
    const pageTitle = lines[0] ? lines[0].substring(0, 30) : `關鍵解析 (頁 ${pData.page})`;
    const bodyPoints = lines.slice(1).slice(0, 5); // 取 5 點以內避免溢出

    applySlideFrame(slide, pageTitle, pageNum);
    
    // 左右交替排版
    const imgLeft = (pageNum % 2 === 0);
    const imgX = imgLeft ? LAYOUT.main.leftImg.x : LAYOUT.main.rightImg.x;
    const txtX = imgLeft ? LAYOUT.main.rightText.x : LAYOUT.main.leftText.x;
    
    const imgPath = getImageForSlide(pageNum % IMAGES.length);
    if (imgPath) {
        slide.addImage({ path: imgPath, x: imgX, y: LAYOUT.main.leftImg.y, w: LAYOUT.main.leftImg.w, h: LAYOUT.main.leftImg.h, sizing: { type: 'contain' } });
    }
    
    const bulletPoints = bodyPoints.map(line => ({
        text: line.trim(),
        options: { bullet: true, fontSize: 16, color: THEME.text, fontFace: FONT_TITLE, lineSpacing: 24, breakLine: true }
    }));
    
    slide.addText(bulletPoints, { x: txtX, y: LAYOUT.main.rightText.y, w: LAYOUT.main.rightText.w, h: LAYOUT.main.rightText.h, valign: "top" });
});

// 輸出
const safeFilename = mainTitle.replace(/[\\/:"*?<>|]/g, "_").substring(0, 40);
const outputDir = path.join(__dirname, "output");
if (!fs.existsSync(outputDir)) fs.mkdirSync(outputDir, { recursive: true });
const outP = path.join(outputDir, `${safeFilename}_v23_Dynamic.pptx`);

pres.writeFile({ fileName: outP }).then(fn => {
    console.log(`[SUCCEESS] Dynamic PPTX Generated at: ${fn}`);
}).catch(err => {
    console.error(`[ERROR] Failed to write file: ${err.message}`);
    process.exit(1);
});
