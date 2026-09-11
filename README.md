# Excel 中文姓名遮罩工具（mask_names.py）

把 Excel 檔裡的中文姓名**第二個字改成 O**（例：王小明 → 王O明、李明 → 李O、歐陽小花 → 歐O小花），
結果**另存成同一個檔案裡的新工作表**（原表名＋「_遮罩」），原工作表完全不更動。

## 🔒 隱私保證

- **完全離線**：程式只讀寫你電腦上的那個 Excel 檔，不連網、不上傳、不傳送任何資料。
- 程式碼僅 100 多行、只使用 `openpyxl` 讀寫本機檔案，歡迎自行檢視原始碼確認。
- 適合處理含個資的名單（點名單、成績冊）後再分享遮罩版工作表。

## 安裝（一次即可）

1. 安裝 [Python 3.8+](https://www.python.org/downloads/)（安裝時勾選 Add to PATH）
2. 開啟命令提示字元執行：`pip install openpyxl`
3. 下載本工具：點本頁綠色 **Code → Download ZIP**，或直接抓
   [mask_names.py](https://raw.githubusercontent.com/rlirdo/excel-name-masker/main/mask_names.py)

## 使用

```bash
python mask_names.py 你的檔案.xlsx              # 自動找標題含「姓名／名字／Name」的欄
python mask_names.py 你的檔案.xlsx --sheet 名單  # 指定工作表
python mask_names.py 你的檔案.xlsx --col C      # 指定姓名在 C 欄
python mask_names.py 你的檔案.xlsx --all        # 掃描所有欄位遮罩所有中文姓名
```

Windows 使用者也可以把 Excel 檔**直接拖曳到 `遮罩姓名.bat`** 上執行。

## 規則說明

- 只改「連續中文字」的第 2 個字；英文、數字、括號附註都保留（`林柏亨(Aaron)` → `林O亨(Aaron)`）。
- 標題列（姓名/名字/Name）不遮罩。
- 公式儲存格在新工作表中會存為計算後的值；若檔案從未用 Excel 開啟過，計算值可能是空白。
- 重複執行會建立 `_遮罩2`、`_遮罩3`⋯不會覆蓋舊結果。

## 授權

MIT License — 歡迎自由使用與修改。
