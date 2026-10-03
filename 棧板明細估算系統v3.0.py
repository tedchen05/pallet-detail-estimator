# -*- coding: utf-8 -*-
"""
Created on Thu Feb 20 05:26:10 2025

@author: 陳皓均

依包材規則把出貨件數拆成棧板明細，寫入輸出 Excel 的工作表2
"""
import os
import tkinter as tk
from datetime import datetime
from tkinter import filedialog

import pandas as pd

RULE_COLUMNS = ["包材", "分裝", "包裝", "單一包裝"]
ITEM_COLUMNS = ["物料", "批次", "件數"]
OUTPUT_COLUMNS = ["品號", "批號", "數量", "包材"]
OUTPUT_SHEET = "工作表2"
CODE_LENGTH = 5  # 物料代碼取最後幾碼比對包材


# 處理過程中可預期的錯誤，訊息直接顯示在介面上
class PalletError(Exception):
    pass


# 讀取 Excel 並確認必要欄位都存在
def read_excel_checked(path, columns, label, **kwargs):
    name = os.path.basename(path)
    if not os.path.exists(path):
        raise PalletError(f"{name} {label}的檔案不存在，請檢查檔名")
    try:
        df = pd.read_excel(path, engine="openpyxl", **kwargs)
    except Exception:
        raise PalletError(f"{name} {label}的檔案讀取失敗，請檢查檔案格式")

    missing = [col for col in columns if col not in df.columns]
    if missing:
        raise PalletError(f"{name} 缺少欄位：{'、'.join(missing)}，請檢查第一列名稱")
    return df


# 同一物料、批次的件數加總，保留第一次出現的順序
def merge_items(df):
    df["批次"] = df["批次"].fillna("")
    return df.groupby(["物料", "批次"], sort=False, as_index=False)["件數"].sum()


# 依物料代碼最後幾碼找出對應的包材規則，查無回傳 None
def find_rule(rules_df, item_code):
    tail = item_code[-CODE_LENGTH:]
    codes = rules_df["包材"].astype(str)
    matched = rules_df[codes.apply(lambda code: code in tail)]
    if matched.empty:
        return None
    return matched.iloc[0]


# 把每個品項依分裝數量拆成多筆明細
def split_items(items_df, rules_df, file_name):
    result = []
    rows = items_df[ITEM_COLUMNS].itertuples(index=False)
    for item_code, batch, qty in rows:
        if not isinstance(item_code, str):
            raise PalletError(f"{file_name} 物料欄位有空白或非文字，請檢查物料欄位")

        rule = find_rule(rules_df, item_code)
        if rule is None:
            result.append([item_code, batch, qty, ""])  # 查無規則仍輸出，包材留空
            continue

        try:
            quotient = int(qty // rule["分裝"])
            remainder = int(qty % rule["分裝"])
        except (TypeError, ValueError, ZeroDivisionError):
            raise PalletError(f"{file_name} 件數或分裝有誤，請檢查件數欄位")

        result.extend([[item_code, batch, rule["分裝"], rule["包裝"]]] * quotient)
        if remainder > 1:
            result.append([item_code, batch, remainder, rule["包裝"]])
        elif remainder == 1:
            result.append([item_code, batch, remainder, rule["單一包裝"]])

    return pd.DataFrame(result, columns=OUTPUT_COLUMNS)


# 把明細寫入輸出檔的工作表2，不影響其他工作表
def write_output(output_df, path):
    name = os.path.basename(path)
    if not os.path.exists(path):
        raise PalletError(f"{name} 輸出結果的檔案不存在，請檢查檔名")
    try:
        with pd.ExcelWriter(path, engine="openpyxl", mode="a",
                            if_sheet_exists="overlay") as writer:
            output_df.to_excel(writer, sheet_name=OUTPUT_SHEET, index=False)
    except PermissionError:
        raise PalletError(f"{name} 輸出結果的檔案被占用，請先關閉檔案再試")
    except Exception:
        raise PalletError(f"{name} 輸出結果的檔案寫入失敗，請檢查檔案格式")


# 從讀檔、拆分到寫檔的完整流程
def build_pallet(rule_path, item_path, output_path):
    rules_df = read_excel_checked(rule_path, RULE_COLUMNS, "條件規則")
    items_df = read_excel_checked(item_path, ITEM_COLUMNS, "需要整合",
                                  dtype={"批次": str})
    merged_df = merge_items(items_df)
    output_df = split_items(merged_df, rules_df, os.path.basename(item_path))
    write_output(output_df, output_path)


# 棧板明細估算系統的視窗介面
class PalletApp:

    # 建立視窗與三組檔案選擇欄位
    def __init__(self, root):
        self.root = root
        root.geometry("350x200")
        root.resizable(False, False)
        root.title("棧板明細估算系統")

        self.rule_entry = self.add_file_row(0, "請選擇條件規則的 Excel 檔案：")
        self.item_entry = self.add_file_row(1, "請選擇需要整合的 Excel 檔案：")
        self.output_entry = self.add_file_row(2, "請選擇輸出結果的 Excel 檔案：")

        tk.Button(root, text="提交", command=self.submit).grid(
            row=3, column=0, columnspan=3, pady=10)
        self.result_label = tk.Label(root, text="")
        self.result_label.grid(row=4, column=0, columnspan=3, pady=5)

        root.bind("<Return>", lambda event: self.submit())

    # 新增一列「說明＋輸入框＋選擇按鈕」，回傳輸入框
    def add_file_row(self, row, text):
        tk.Label(self.root, text=text).grid(
            row=row, column=0, padx=5, pady=5, sticky="e")
        entry = tk.Entry(self.root, width=15)
        entry.grid(row=row, column=1, padx=5, pady=5)
        tk.Button(self.root, text="選擇",
                  command=lambda: self.browse(entry)).grid(
            row=row, column=2, padx=5, pady=5)
        return entry

    # 開啟檔案選擇視窗，把路徑填入輸入框
    def browse(self, entry):
        path = filedialog.askopenfilename(
            filetypes=[("Excel Files", "*.xlsx;*.xls")])
        if path:
            entry.delete(0, tk.END)
            entry.insert(0, path)

    # 執行估算並在下方顯示結果
    def submit(self):
        try:
            build_pallet(self.rule_entry.get().strip(),
                         self.item_entry.get().strip(),
                         self.output_entry.get().strip())
        except PalletError as error:
            self.result_label.config(text=str(error), fg="red")
            return

        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.result_label.config(text=f"處理完成，已生成棧板明細 {now}",
                                 fg="green")


if __name__ == "__main__":
    window = tk.Tk()
    PalletApp(window)
    window.mainloop()
