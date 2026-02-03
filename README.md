# Event Extract Tools

 給某日本原神用的 全服通用 

## 環境需求

- Python 3.x
- 依賴庫：`pandas`, `openpyxl`

安裝依賴：
```bash
pip install pandas openpyxl
```

## 使用方法

1. 運行 `event.py` 腳本。
2. 按照提示輸入：
   - **輸出類型**：
     - `1`: 簡要描述 (txt)
     - `2`: 完整信息 (txt)
     - `3`: Excel 導出 (xlsx)
   - **事件資料夾路徑**：支援直接貼上包含引號的路徑 (例如 `"D:\Path\To\Event"` )。
   - **輸出文件名稱**：輸入檔名即可，程式會自動加上對應的副檔名 (`.txt` 或 `.xlsx`)。

## 示例

假設你的事件資料夾路徑是 ./event，並且你想要輸出到 output.txt 文件中：

```bash
python event.py
```

然後按照提示輸入：

```text
請輸入輸出類型(1: 簡要描述, 2: 完整信息, 3: Excel 導出): 1
請輸入事件資料夾的路徑: "./event"
請輸入輸出文件的名稱(txt): output
```

這將在當前目錄下生成 `output` 文件。
