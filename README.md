# 資料填報小幫手

填報輔助工具（純前端網頁）。

- 所有檔案只在使用者的瀏覽器中處理，**不會上傳到任何伺服器**；關閉分頁即清空。
- 文字辨識（Tesseract.js、PaddleOCR）於本機瀏覽器執行，模型檔放在 `lib/`。
- 本儲存庫**不含任何資料檔**（`.gitignore` 排除 Excel、PDF、Word 與 `data/`）。

本機測試：`python tools/serve.py`，開啟 http://127.0.0.1:8000/
