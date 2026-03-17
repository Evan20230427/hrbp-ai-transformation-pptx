# pdf-xlsx-to-pptx 啟動腳本
$ErrorActionPreference = "Continue"

Write-Host "--- 階段 1: 執行 PDF 內容擷取 (OCR) ---" -ForegroundColor Cyan
python extract_ocr.py
if ($LASTEXITCODE -ne 0) { 
    Write-Host "!! OCR 擷取失敗 !!" -ForegroundColor Red
    exit 1 
}

Write-Host "--- 階段 2: 執行簡報渲染生成 (Node.js) ---" -ForegroundColor Cyan
node generate_pptx_v23_ultimate.js
if ($LASTEXITCODE -ne 0) { 
    Write-Host "!! 簡報生成失敗 !!" -ForegroundColor Red
    exit 1 
}

Write-Host "--- 任務圓滿完成！簡報已產出至 output 目錄 ---" -ForegroundColor Green
