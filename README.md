# PDF+XLSX → PPTX 簡報生成專案

將 PDF 和 XLSX 檔案中的內容自動擷取、結構化，並生成專業 PowerPoint 簡報。

## 使用方式

1. 將來源 PDF / XLSX 檔案放入 `input/` 目錄
2. 在任一支援平台（Antigravity / Copilot CLI / OpenCode）中開啟此專案
3. 告訴 AI：「請將 input 中的檔案生成簡報」
4. 生成的簡報將輸出到 `output/` 目錄

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

## 支援平台

- Antigravity（`.agent/skills/`）
- GitHub Copilot CLI（`.github/skills/`）
- OpenCode（`.opencode/skills/`）
