# -*- coding: utf-8 -*-
"""
Created on Thu Feb 20 05:26:10 2025

@author: user
"""
import pandas as pd
import tkinter as tk
from tkinter import filedialog
from datetime import datetime
import os  

def browse_file1():
    file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx;*.xls")])
    if file_path:
        file1_entry.delete(0, tk.END)
        file1_entry.insert(0, file_path)

def browse_file2():
    file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx;*.xls")])
    if file_path:
        file2_entry.delete(0, tk.END)
        file2_entry.insert(0, file_path)

def browse_package():
    file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx;*.xls")])
    if file_path:
        package_entry.delete(0, tk.END)
        package_entry.insert(0, file_path)

def finish():
    package_file = package_entry.get().strip()
    
    if not os.path.exists(package_file):
        result_label.config(text=f" {os.path.basename(package_file)} 條件規則的檔案不存在，請檢查檔名", fg="red")
        return

    try:
        rules_df = pd.read_excel(package_file)  # 讀取選擇的包材檔案
    except:
        result_label.config(text=f"{os.path.basename(package_file)} 條件規則的檔案失敗，請檢查檔案格式", fg="red")
        return
    
    file1_name = file1_entry.get().strip()
    
    if not os.path.exists(file1_name):
        result_label.config(text=f"{os.path.basename(file1_name)} 需要整合的檔案不存在，請檢查檔名", fg="red")
        return

    try:
        df = pd.read_excel(file1_name, engine="openpyxl", dtype={"批次": str})
        df["批次"] = df["批次"].fillna("")
        df_unique = df.drop_duplicates(subset=["物料", "批次"], keep="first").drop(columns=["件數"])
        sum_df = df.groupby(["物料", "批次"], as_index=False)["件數"].sum()
        final_df = pd.merge(df_unique, sum_df, on=["物料", "批次"], how="left")
        print(df)
        print(df_unique)
        print(sum_df)
        print(final_df)
    except:
        result_label.config(text=f" {os.path.basename(file1_name)} 需要整合的檔案失敗", fg="red")
        return
    
    result = []
    
    for i in range(len(final_df)):
        try:
            item_code = final_df["物料"][i]
            batch = final_df["批次"][i]
            qty = final_df["件數"][i]
        except:
            result_label.config(text=f"{os.path.basename(file1_name)} 條件規則有錯誤，請檢查第一列名稱", fg="red")
            return
    
        # 取得倒數第二個字母
        try:
            second_last_char = item_code[-5:]  # 取最後 5 個字元
        except:
            result_label.config(text=f"{os.path.basename(file1_name)} 條件規則有錯誤，請檢查物料欄位", fg="red")
            return
    
        # 查詢對應的包裝規則
        try:
            rule_row = rules_df[rules_df["包材"].astype(str).apply(lambda x: x in second_last_char)]
            if not rule_row.empty:
                divisor = rule_row.iloc[0]["分裝"]
                packaging = rule_row.iloc[0]["包裝"]
                single_packaging = rule_row.iloc[0]["單一包裝"]
            else:
                result.append([item_code, batch, qty, ""])  
                continue
        except:
            result_label.config(text=f"{os.path.basename(package_file)} 欄位失敗，請確認 Excel 欄位名稱", fg="red")
            return
    
        # 計算商數與餘數
        try:
            quotient = int(qty // divisor)  # 確保 quotient 是 int
            remainder = int(qty % divisor)  # 確保 remainder 是 int
        except:
            result_label.config(text=f"{os.path.basename(file1_name)} 條件規則有錯誤，請檢查件數欄位", fg="red")
            return
    
        # 先寫入完整商數的資料
        for j in range(quotient):
            result.append([item_code, batch, divisor, packaging])
    
        # 處理餘數
        if remainder > 1:
            result.append([item_code, batch, remainder, packaging])  # 將餘數資料直接加到 result 裡
        elif remainder == 1:
            result.append([item_code, batch, remainder, single_packaging])  # 如果餘數是 1，使用單一包裝
    
    # 轉換成 DataFrame，寫回 Excel
    output_df = pd.DataFrame(result, columns=["品號", "批號", "數量", "包材"])
    file_path2 = file2_entry.get().strip()
    
    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    try:
        with pd.ExcelWriter(file_path2, engine="openpyxl", mode="a", if_sheet_exists="overlay") as writer:
            output_df.to_excel(writer, sheet_name="工作表2", index=False)
        result_label.config(text=f"處理完成，已生成棧板明細 {now}", fg="green")  
    except PermissionError:
        result_label.config(text=f"{os.path.basename(file_path2)} 輸出結果的檔案被占用，請先關閉檔案再試", fg="red")
    except:
        result_label.config(text=f" {os.path.basename(file_path2)} 輸出結果的檔案有錯誤，請檢查檔案是否存在", fg="red")

pallet = tk.Tk()
pallet.geometry("350x200")
pallet.resizable(False, False) 
pallet.title("棧板明細估算系統")

# 包材檔案
package_label = tk.Label(pallet, text="請選擇條件規則的 Excel 檔案：")
package_entry = tk.Entry(pallet, width=15)
package_button = tk.Button(pallet, text="選擇", command=browse_package)

# 讀取 Excel 檔案
file1_label = tk.Label(pallet, text="請選擇需要整合的 Excel 檔案：")
file1_entry = tk.Entry(pallet, width=15)
file1_button = tk.Button(pallet, text="選擇", command=browse_file1)

# 寫入 Excel 檔案
file2_label = tk.Label(pallet, text="請選擇輸出結果的 Excel 檔案：")
file2_entry = tk.Entry(pallet, width=15)
file2_button = tk.Button(pallet, text="選擇", command=browse_file2)

confirm = tk.Button(pallet, text="提交", command=finish)
result_label = tk.Label(pallet, text="")

# 佈局設定
package_label.grid(row=0, column=0, padx=5, pady=5, sticky="e")
package_entry.grid(row=0, column=1, padx=5, pady=5)
package_button.grid(row=0, column=2, padx=5, pady=5)

file1_label.grid(row=1, column=0, padx=5, pady=5, sticky="e")
file1_entry.grid(row=1, column=1, padx=5, pady=5)
file1_button.grid(row=1, column=2, padx=5, pady=5)

file2_label.grid(row=2, column=0, padx=5, pady=5, sticky="e")
file2_entry.grid(row=2, column=1, padx=5, pady=5)
file2_button.grid(row=2, column=2, padx=5, pady=5)

confirm.grid(row=3, column=0, columnspan=3, pady=10)
result_label.grid(row=4, column=0, columnspan=3, pady=5)  

pallet.bind("<Return>", lambda event: finish())

pallet.mainloop()
