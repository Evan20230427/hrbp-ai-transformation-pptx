# pdf-xlsx-to-pptx 使用說明 (v23 Ultimate)

本文件詳細說明如何透過終端機執行簡報生成流程。本專案已整合 OCR 擷取技術與羅馬美學 v23 渲染引擎。

## 前置準備
- **Python 環境**：需安裝 `pytesseract` 與 `PyMuPDF (fitz)`，並確保系統中已安裝 `Tesseract-OCR`。
- **Node.js 環境**：需安裝 `pptxgenjs` 依賴。

---

## 方法一：PowerShell 一鍵自動化 (推薦)
此方式會自動按順序執行「內容擷取」與「簡報生成」，避免手動輸錯指令。

1. **開啟 PowerShell**。
2. **切換至專案目錄**：
   ```powershell
   cd "C:\Users\TW-Evan.Chen\.gemini\antigravity\scratch\pdf-xlsx-to-pptx"
   ```
3. **執行啟動腳本**：
   ```powershell
   ./start_flow.ps1
   ```

---

## 方法二：手動分步執行
如果您需要單獨階段的輸出結果或進行排錯，請依序執行以下指令：

### 階段 1：PDF 內容擷取 (OCR)
將 PDF 掃描件轉換為結構化 JSON 內容。
```powershell
python extract_ocr.py
```
- **產出物**：`extracted_content.json`

### 階段 2：羅馬美學簡報渲染 (v23)
根據擷取內容生成具備「零重疊」與「雕像唯一性」的 PPTX。
```powershell
node generate_pptx_v23_ultimate.js
```
- **產出物**：`output/HRBP_AI_Transformation_v23_Ultimate.pptx`

---

## 輸出結果說明
- **簡報路徑**：`./output/`
- **視覺規範**：羅馬美學 v23 (Unique Statue System + Zero-Overlap Layout)
- **支援頁數**：20 頁以上 (當前版本預設 24 頁)
