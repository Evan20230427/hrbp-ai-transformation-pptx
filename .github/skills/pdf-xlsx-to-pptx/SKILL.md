---
name: pdf-xlsx-to-pptx
description: 從 PDF 和 XLSX 檔案擷取內容，自動生成專業簡報（PPTX）。同時可產生社群宣傳用的 Slack GIF。當使用者提到將 PDF 或 Excel 資料轉為簡報、製作報告簡報、或需要生成搭配的宣傳動畫時，請使用此技能。此技能整合了 pdf、xlsx、pptx、theme-factory、doc-coauthoring、slack-gif-creator 六個子技能。
---

# PDF+XLSX → PPTX 組合技能

將 PDF 和/或 XLSX 來源檔案的內容，自動擷取、結構化、並生成專業簡報。

## 工作流程

依照以下四個階段執行：

### 階段一：擷取（Extract）

根據輸入檔案類型，使用對應的子技能讀取內容：

- **PDF 檔案**：使用 `pdf` 技能讀取文字、表格、圖片
- **XLSX 檔案**：使用 `xlsx` 技能讀取工作表資料、圖表

> 將所有擷取到的內容整理為結構化的中間格式，包含：標題、章節、重點摘要、數據表格、關鍵圖表。

### 階段二：結構化（Structure）

使用 `doc-coauthoring` 技能的方法論來組織內容：

1. 分析擷取的內容，識別核心主題與邏輯架構
2. 建立簡報大綱，包含：
   - 封面頁（標題、副標題、日期）
   - 目錄/議程頁
   - 各章節內容頁
   - 數據/圖表頁
   - 總結/結論頁
3. 確保每頁有明確的標題和 3-5 個重點
4. 控制總頁數在使用者指定範圍內

### 階段三：生成（Generate）

使用 `pptx` 技能建立 PowerPoint 簡報：

1. 使用 PptxGenJS 建立幻燈片
2. 每頁包含適當的排版：標題、正文、列表、表格
3. 數據以表格或圖表形式呈現
4. 確保文字精簡、重點突出

### 階段四：美化（Style）

使用 `theme-factory` 技能套用專業主題：

1. 展示可用的 10 個預設主題給使用者選擇
2. 若使用者未指定，預設使用「Modern Minimalist」主題
3. 套用統一的色彩、字體、排版風格
4. 確保視覺一致性和專業感

### 階段五：宣傳素材（Promote）

若使用者需要分享簡報生成結果：
使用 `slack-gif-creator` 生成吸睛的宣傳 GIF 動畫，供使用者在 Slack 或內部群組發布。

## 使用規範

- 簡報語言與來源檔案一致（繁體中文/英文）
- 每頁內容精簡，避免文字牆
- 數據優先以視覺化方式呈現（表格、圖表）
- 頁數依使用者要求控制，預設不超過 20 頁
- 輸出檔案存放在 `output/` 目錄

## 子技能參考

以下子技能已安裝在同層的 skills 目錄中：

| 子技能 | 用途 | 參考 |
|--------|------|------|
| pdf | 讀取 PDF 內容 | `../pdf/SKILL.md` |
| xlsx | 讀取 Excel 內容 | `../xlsx/SKILL.md` |
| pptx | 生成 PPTX 簡報 | `../pptx/SKILL.md` |
| theme-factory | 套用主題樣式 | `../theme-factory/SKILL.md` |
| doc-coauthoring | 內容結構化 | `../doc-coauthoring/SKILL.md` |
| slack-gif-creator | 製作 Slack 宣傳動畫 | `../slack-gif-creator/SKILL.md` |
