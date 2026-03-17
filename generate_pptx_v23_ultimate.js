const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

let pres = new pptxgen();
pres.layout = 'LAYOUT_16x9';
pres.author = 'Antigravity AI';
pres.title = 'HRBP AI 轉型最佳實務 - v23 Ultimate Edition';

const THEME = {
    primary: "F1F5F9",    // 大理石紋理底色 (亮色調)
    secondary: "B45309",  // 古典金 (強調色)
    text: "1E293B",       // 深岩藍 (文本色)
    white: "FFFFFF",
    line: "D1D5DB",       
    accent: "475569",     
    highlight: "FEF3C7"   
};

const FONT_TITLE = "Microsoft JhengHei";
const FONT_BODY = "Arial";

// 包含 25 張絕對不重複的圖資
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

// 核心檢查與唯一性追蹤
const usedImages = new Set();
function getUniqueImage(index) {
    const img = IMAGES[index];
    if (usedImages.has(img)) throw new Error(`Duplicate image: ${img}`);
    if (!fs.existsSync(img)) throw new Error(`Missing image: ${img}`);
    usedImages.add(img);
    return img;
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
    slide.addText(title, { x: LAYOUT.header.x, y: LAYOUT.header.y, w: LAYOUT.header.w, h: LAYOUT.header.h, fontSize: 24, bold: true, color: THEME.secondary, fontFace: FONT_TITLE });
    slide.addShape(pres.shapes.RECTANGLE, { x: LAYOUT.line.x, y: LAYOUT.line.y, w: LAYOUT.line.w, h: LAYOUT.line.h, fill: { color: THEME.secondary } });
    slide.addText(`Gartner HRBP | v23 Ultimate | Slide ${pageNum}`, { x: LAYOUT.footer.x, y: LAYOUT.footer.y, w: LAYOUT.footer.w, h: LAYOUT.footer.h, fontSize: 10, color: THEME.accent, align: "right", fontFace: FONT_BODY });
}

pres.defineSlideMaster({
    title: 'ROMAN_V23',
    background: { color: THEME.primary },
    objects: [
        { rect: { x: 0, y: 0, w: 0.08, h: "100%", fill: { color: THEME.secondary } } }
    ]
});

// 24 頁內容矩陣
const SLIDES_CONTENT = [
    { title: "AI 時代下的 HRBP 角色重塑", text: ["引領企業 AI 轉型的策略實務", "羅馬美學 v23 終極版", "未來 HRBP：AI 轉型顧問"], isCover: true },
    { title: "Agenda：簡報大綱與學習路徑", text: ["Section 1: 轉型背景與願景", "Section 2: 三階段任務轉移策略", "Section 3: 支援與結語"], imgRight: true },
    { title: "轉折點：AI 如何重塑 HRBP", text: ["AI 正在改變 HR 為企業創造價值的模式", "傳統事務性工作正被 AI 吸收", "HRBP 面臨被低估的風險"], imgLeft: true },
    { title: "現狀揭示：策略參與的斷層", text: ["僅 51% 的管理者同意 HRBP 參與策略討論", "其餘仍被交易性工作綁架", "AI 時代需要更深層的業務對話"], imgRight: true },
    { title: "消失的優勢：行政工作的毒藥", text: ["當 AI 重複您擁有的任務時...", "工作若不影響業務成果，則失去相關性", "剔除舊有作品是邁向策略的第一步"], imgLeft: true },
    { title: "未來願景：AI 轉型顧問", text: ["從 HRBP 進化為 AI Transformation Consultant", "擴大職權至引導人員面向的轉型", "成為不可或缺的策略夥伴"], imgRight: true },
    { title: "核心：策略人才領袖 (STL)", text: ["識別最迫切的人才機會", "掌控業務單位的策略人才規劃", "引導 AI 驅動的勞動力重新設計"], imgLeft: true },
    { title: "三階段任務轉變藍圖概要", text: ["Phase 1: 剔除舊有工作 (Strip legacy)", "Phase 2: 強化策略核心 (Augment)", "Phase 3: 擴展新策略空間 (Expand)"], imgRight: true },
    { title: "Phase 1: 剔除舊有行政壓力", text: ["明確定義 HRBP 應該投入的時間分配", "解決行政負荷導致的信譽風險", "利用 AI 實現流程自動化"], imgLeft: true },
    { title: "自動化路線圖：從現在到未來", text: ["劃分自動化準備度：12-24 個月路徑", "哪些保留為人類主導？", "建立透明的技術導入進程"], imgRight: true },
    { title: "定義界線：AI 執行 vs. HRBP 主導", text: ["落實 'Stop-doing' 清單", "清晰傳達角色邊界", "避免陷入低階數據摘錄工作"], imgLeft: true },
    { title: "成功衡量：回收高價值時間", text: ["回收特定任務工時 (e.g. 入職檢查、報表)", "重新分配至策略優先事項的百分比", "指標：產能回收轉化率"], imgRight: true },
    { title: "Phase 2: 強化核心策略責任", text: ["以 AI 強化人力規劃與接班決策", "提供領袖無法在其他地方獲得的洞見", " evidence-based 的未來預測"], imgLeft: true },
    { title: "能力更新：AI 時代的 HRBP", text: ["確保 AI 準備度成為角色評核的一部分", "掌握 AI 生成輸入的使用方法", "培養與 AI 工具協作的數位敏銳度"], imgRight: true },
    { title: "深度洞察：利用 AI 促進對話", text: ["以 AI 訊號為基準，促進高品質業務對談", "避免重複工作：HRBP-CoE-AI 工作流", "聚焦於業務領袖可用的決策訊號"], imgLeft: true },
    { title: "成功衡量：提升人才產出成果", text: ["縮短人力/接班決策週期", "繼任者就緒率增加百分比", "降低 AI 標記的高風險職位流失率"], imgRight: true },
    { title: "Phase 3: 主導企業 AI 變革", text: ["擴展至由 AI 推動的新策略工作", "啟動策略人才領袖 (STL) 小組", "直接參與 AI 轉型決策過程"], imgLeft: true },
    { title: "STL 小組：選拔與變革領導", text: ["挑選具備判斷力與變革能力的成員", "處理 AI 轉型中的人性面向", "試點計畫：從小規模成功擴散"], imgRight: true },
    { title: "影響決策：在決策前處理風險", text: ["在最終決策前，預判人才風險與機會", "確保 HR 在 AI 轉型中的角色明確化", "嵌入 STL 條款於企業核心專案"], imgLeft: true },
    { title: "成功衡量：區域級轉型影響", text: ["重新部署 vs. 裁員的比例優化", "標記與減輕 AI 決策中的偏見案件", "縮短關鍵角色的能力實現時間"], imgRight: true },
    { title: "Gartner 資源：強化專業知識", text: ["獲取前沿趨勢與 cutting-edge 洞察", "利用 Competency Model 進行自我診斷", "建立與同儕對標的專業高地"], imgLeft: true },
    { title: "工具箱：DataHub 與引導指南", text: ["利用 Ignition Guides 避開常見錯誤", "DataHub：將數據轉化為人才決策", "加速專案執行並確保最佳實務"], imgRight: true },
    { title: "結語：為 AI 時代做好準備", text: ["不要讓 HRBP 的角色停留於過去", "從現在開始，推動企業的人才進化", "Ready for what's next? We help."], imgLeft: true },
    { title: "三階架構：全量主題與頁碼對齊", isMindmap: true }
];

SLIDES_CONTENT.forEach((sData, idx) => {
    let slide = pres.addSlide({ masterName: 'ROMAN_V23' });
    let pageNum = idx + 1;

    if (sData.isCover) {
        slide.addImage({ path: getUniqueImage(idx), x: 0, y: 0, w: '100%', h: '100%' });
        slide.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 3.5, w: 9.0, h: 1.5, fill: { color: 'FFFFFF', transparency: 15 } });
        slide.addText(sData.title, { x: 0.8, y: 3.7, w: 8.4, fontSize: 36, bold: true, color: THEME.secondary, fontFace: FONT_TITLE, align: "center" });
        slide.addText(sData.text.join(" | "), { x: 0.8, y: 4.5, w: 8.4, fontSize: 18, color: THEME.text, fontFace: FONT_TITLE, align: "center" });
    } else if (sData.isMindmap) {
        applySlideFrame(slide, sData.title, pageNum);
        slide.addImage({ path: getUniqueImage(idx), x: 6.5, y: 0.5, w: 3.5, h: 3.5, transparency: 85 });
        const M_DATA = [
            { t: "轉型願景", c: [{ k: "AI 轉型顧問", s: ["角色重塑 (P.1-3)", "現狀分析 (P.4-5)", "願景目標 (P.6-7)"] }] },
            { t: "三階段策略", c: [{ k: "任務轉移", s: ["Phase 1 (P.9-12)", "Phase 2 (P.13-16)", "Phase 3 (P.17-20)"] }] },
            { t: "支援工具", c: [{ k: "資源與績效", s: ["成功指標 (P.16,20)", "Gartner 資源 (P.21-22)"] }] }
        ];
        const startX = 0.5, startY = 1.4, colWidth = 3.1;
        M_DATA.forEach((l1, i) => {
            let x = startX + (i * colWidth);
            slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: x, y: startY, w: 2.9, h: 0.45, fill: { color: THEME.secondary }, rectRadius: 0.1 });
            slide.addText(l1.t, { x: x, y: startY, w: 2.9, h: 0.45, color: THEME.white, bold: true, align: "center", fontSize: 14, fontFace: FONT_TITLE });
            let currentY = startY + 0.65;
            l1.c.forEach(l2 => {
                slide.addText(l2.k, { x: x + 0.2, y: currentY, w: 2.5, h: 0.4, color: THEME.secondary, bold: true, fontSize: 12, fontFace: FONT_TITLE });
                currentY += 0.45;
                l2.s.forEach(l3 => {
                    slide.addText("▶ " + l3, { x: x + 0.3, y: currentY, w: 2.6, h: 0.3, color: THEME.text, fontSize: 10, fontFace: FONT_TITLE });
                    currentY += 0.35;
                });
            });
        });
    } else {
        applySlideFrame(slide, sData.title, pageNum);
        let imgX = sData.imgLeft ? LAYOUT.main.leftImg.x : LAYOUT.main.rightImg.x;
        let txtX = sData.imgLeft ? LAYOUT.main.rightText.x : LAYOUT.main.leftText.x;
        
        slide.addImage({ path: getUniqueImage(idx), x: imgX, y: LAYOUT.main.leftImg.y, w: LAYOUT.main.leftImg.w, h: LAYOUT.main.leftImg.h, sizing: { type: 'contain' } });
        
        const bulletPoints = sData.text.map(line => ({ text: line, options: { bullet: true, fontSize: 18, color: THEME.text, fontFace: FONT_TITLE, lineSpacing: 28, breakLine: true } }));
        slide.addText(bulletPoints, { x: txtX, y: LAYOUT.main.rightText.y, w: LAYOUT.main.rightText.w, h: LAYOUT.main.rightText.h, valign: "top" });
    }
});

const outP = path.join("C:\\Users\\TW-Evan.Chen\\.gemini\\antigravity\\scratch\\pdf-xlsx-to-pptx\\output", "HRBP_AI_Transformation_v23_Ultimate.pptx");
pres.writeFile({ fileName: outP }).then(fn => {
    console.log(`Success: v23 PPTX Generated at ${fn}`);
}).catch(err => console.error(err));
