# PDF 工具箱

離線 PDF 工具，無需安裝 Adobe。支援 Windows，介面使用繁體中文。

## 功能

- **PDF 合併** — 多個 PDF 合為一份，支援自動補空白頁（雙面列印用）
- **Word → PDF** — 批次將 `.docx` / `.doc` 轉為 PDF（需安裝 Microsoft Word）
- **PDF → JPG** — 每頁匯出為高解析 JPG（300 DPI）

## 下載使用

前往 [Releases](../../releases) 頁面下載 `PDF工具箱.exe`，雙擊即可執行，**不需安裝 Python**。

> 首次啟動約需 5–10 秒（解壓中），之後正常。

## 自行執行（需 Python 3.10+）

```bash
pip install pypdf pymupdf pywin32
python pdf_toolbox.py
```

## 注意事項

- **Word → PDF** 功能需要目標電腦已安裝 Microsoft Word
- 其餘功能完全離線，無需額外安裝

## 授權

[MIT License](LICENSE) © 2026 陳冠廷
