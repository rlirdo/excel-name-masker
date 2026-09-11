#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
mask_names.py — 將 Excel 中文姓名的第二個字改為 O，另存成同檔案的新工作表
=========================================================================
用法：
    python mask_names.py 檔案.xlsx                     # 自動尋找標題含「姓名/名字/Name」的欄位
    python mask_names.py 檔案.xlsx --sheet 學生名單     # 指定工作表
    python mask_names.py 檔案.xlsx --col C             # 指定姓名所在欄（欄字母）
    python mask_names.py 檔案.xlsx --all               # 掃描所有欄，凡像中文姓名的儲存格都遮罩

隱私聲明：
    本程式「完全離線」執行——只讀寫你電腦上的這個 Excel 檔，
    不連網、不上傳、不寫入任何其他位置。原工作表內容不會被更動，
    遮罩結果存成同一檔案內的新工作表（原表名＋「_遮罩」）。

需求：Python 3.8+ 與 openpyxl（安裝：pip install openpyxl）
"""
import argparse
import re
import sys

try:
    import openpyxl
    from openpyxl.utils import column_index_from_string
except ImportError:
    sys.exit("缺少 openpyxl，請先執行：pip install openpyxl")

# 中日韓統一表意文字的基本判定（涵蓋一般中文姓名用字與擴充區）
CJK = r"一-鿿㐀-䶿"
NAME_RUN = re.compile(rf"[{CJK}]{{2,}}")          # 連續 2 個以上中文字
HEADER_HINTS = ("姓名", "名字", "name", "Name", "NAME")


def mask_name(text):
    """把字串中第一段連續中文字（視為姓名）的第二個字改為 O。

    「王小明」→「王O明」、「李明」→「李O」、「歐陽小花」→「歐O小花」
    「林柏亨(Aaron)」→「林O亨(Aaron)」；不含中文的字串原樣返回。
    """
    if not isinstance(text, str):
        return text, False
    m = NAME_RUN.search(text)
    if not m:
        return text, False
    s, run = m.start(), m.group(0)
    masked = run[0] + "O" + run[2:]
    return text[:s] + masked + text[s + len(run):], True


def find_name_columns(ws):
    """在前 3 列尋找標題含「姓名/名字/Name」的欄，回傳欄索引清單（1-based）。"""
    cols = []
    for row in ws.iter_rows(min_row=1, max_row=min(3, ws.max_row)):
        for cell in row:
            v = cell.value
            if isinstance(v, str) and any(h in v for h in HEADER_HINTS):
                if cell.column not in cols:
                    cols.append(cell.column)
    return cols


def unique_sheet_name(wb, base):
    name = base[:28] + "_遮罩"
    cand, i = name, 2
    while cand in wb.sheetnames:
        cand = f"{name}{i}"
        i += 1
    return cand


def main():
    ap = argparse.ArgumentParser(description="將 Excel 中文姓名第二字改為 O，另存新工作表（純本機、不連網）")
    ap.add_argument("file", help="Excel 檔案路徑（.xlsx）")
    ap.add_argument("--sheet", help="要處理的工作表名稱（預設：第一個工作表）")
    ap.add_argument("--col", help="姓名所在欄字母，如 C（預設：自動偵測標題）")
    ap.add_argument("--all", action="store_true", help="掃描所有欄位遮罩所有像中文姓名的儲存格")
    a = ap.parse_args()

    wb = openpyxl.load_workbook(a.file)          # 保留原表不動
    wbv = openpyxl.load_workbook(a.file, data_only=True)  # 取公式的計算值
    src_name = a.sheet or wb.sheetnames[0]
    if src_name not in wb.sheetnames:
        sys.exit(f"找不到工作表：{src_name}（現有：{', '.join(wb.sheetnames)}）")
    ws, wsv = wb[src_name], wbv[src_name]

    if a.all:
        target_cols = None                        # None = 全部欄
    elif a.col:
        target_cols = [column_index_from_string(a.col.upper())]
    else:
        target_cols = find_name_columns(ws)
        if not target_cols:
            sys.exit("自動偵測不到「姓名/名字/Name」標題欄，請改用 --col 指定欄位或 --all 掃描全部。")

    new_name = unique_sheet_name(wb, src_name)
    out = wb.create_sheet(new_name)
    masked_count = 0
    for row in wsv.iter_rows():
        for cell in row:
            v = cell.value
            if v is not None and (target_cols is None or cell.column in target_cols):
                if isinstance(v, str):
                    v2, hit = mask_name(v)
                    if hit and cell.row > (1 if target_cols is not None else 0):
                        # 標題列不遮罩（指定欄模式下跳過第 1 列）
                        header = isinstance(wsv.cell(1, cell.column).value, str) and \
                                 any(h in str(wsv.cell(1, cell.column).value) for h in HEADER_HINTS)
                        if not (cell.row == 1 and header):
                            v = v2
                            masked_count += 1
            out.cell(cell.row, cell.column).value = v

    wb.save(a.file)
    print(f"完成：已遮罩 {masked_count} 個姓名 → 新工作表「{new_name}」（原表未更動，檔案未離開你的電腦）")


if __name__ == "__main__":
    main()
