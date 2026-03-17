# PDF+XLSX → PPTX 簡報生成專案

將 PDF 和 XLSX 檔案中的內容自動擷取、結構化，並生成專業 PowerPoint 簡報。

## 使用方式

1. 將來源 PDF / XLSX 檔案放入 `input/` 目錄
2. 開啟 PowerShell 並切換至此目錄
3. 執行指令：`./start_flow.ps1` (一鍵完成 OCR 與生成)
4. 或手動按序執行：`python extract_ocr.py` 接著 `node generate_pptx_v23_ultimate.js`
5. 生成的簡報將輸出到 `output/` 目錄

## 目前任務

- 正處理：`AI 時代重新設計人力資源業務夥伴（HRBP）角色的最佳實務.pdf`
- 目標版本：v23 (Roman Aesthetics Ultimate Edition)

## 已安裝技能

| 技能 | 功能 |
|------|------|
| **pdf-xlsx-to-pptx** | 組合流程技能（orchestrator） |
| **pdf** | 讀取 PDF 內容 |
| **xlsx** | 讀取 Excel 內容 |
| **pptx** | 生成 PowerPoint 簡報 |
| **theme-factory** | 套用專業主題樣式 |
| **doc-coauthoring** | 內容結構化與大綱撰寫 |
| **slack-gif-creator** | 製作社群發布用的 GIF 動畫 |

## 專案現況
- **當前版本**：v21 (Marmoreal Integrity Edition)。
- **任務目標**：處理 HRBP AI 轉型最佳實務 PDF。
- **技術進展**：導入 PaddleOCR 以應對掃描件擷取。

## 支援平台

- Antigravity（`.agent/skills/`）
- GitHub Copilot CLI（`.github/skills/`）
- OpenCode（`.opencode/skills/`）
