# -*- coding: utf-8 -*-
"""
Created on Thu Feb 20 05:26:10 2025

@author: 陳皓均

棧板明細估算系統視窗介面
"""
import os
import tkinter as tk
from datetime import datetime
from tkinter import filedialog, ttk

from pallet_core import (DEFAULT_ITEM, DEFAULT_OUTPUT, DEFAULT_RULE,
                         OUTPUT_COLUMNS, OUTPUT_SHEET, REMAINDER, UNMATCHED,
                         PalletError, app_dir, build_pallet, check_output,
                         initial_path, load_items, load_rules, load_settings,
                         merge_items, output_path_for, save_settings)

APP_TITLE = "棧板明細估算系統 v4.0"
FONT = "微軟正黑體"

PRIMARY = "#1F5FAD"
MUTED = "#666666"
SUCCESS = "#1E7E34"
WARN = "#D26900"
ERROR = "#C0392B"
CARD_BG = "#F3F7FC"
CARD_BORDER = "#C9D8EC"

# 明細預覽的底色與圖例
ROW_COLORS = {
    REMAINDER: ("#D3E4FA", "未滿分裝 (餘數)"),
    UNMATCHED: ("#FFE8CC", "找不到包材規則，包材欄留空"),
}

EXCEL_TYPES = [("Excel 檔案", "*.xlsx")]
PREVIEW_ROWS = 12  # 明細預覽顯示幾列
PREVIEW_COLUMNS = ["序號"] + OUTPUT_COLUMNS
COLUMN_WIDTHS = [50, 150, 100, 70, 90]  # 序號、品號、批號、數量、包材
RESUME_DELAY = 300  # 估算完幾毫秒後才開放按鈕


# 數量轉成整數文字
def format_qty(value):
    try:
        number = float(value)
    except (TypeError, ValueError):
        return str(value)
    return str(int(number)) if number.is_integer() else str(number)


# 介面

# 主視窗
class PalletApp:

    # 建立視窗
    def __init__(self, root):
        self.root = root
        self.output_path = ""  # 上次寫入的檔案
        self.keep_name = tk.BooleanVar(value=False)  # 是否使用原檔名
        self.busy = False

        root.title(APP_TITLE)
        root.resizable(False, False)
        self._setup_style()

        self._build_header()
        body = ttk.Frame(root, padding=(16, 0, 16, 16))
        body.pack(fill="both", expand=True)
        self._build_steps(body)
        self._build_side(body)

        root.bind("<Return>", lambda event: self.submit())
        self._load_defaults()

    # 設定字型與樣式
    def _setup_style(self):
        style = ttk.Style()
        style.configure(".", font=(FONT, 11))
        style.configure("TButton", padding=4)
        style.configure("TLabelframe.Label", font=(FONT, 11, "bold"))
        style.configure("Muted.TLabel", foreground=MUTED, font=(FONT, 9))
        style.configure("Accent.TButton", font=(FONT, 12, "bold"),
                        padding=(0, 8))
        style.configure("Treeview", font=(FONT, 10), rowheight=24)
        style.configure("Treeview.Heading", font=(FONT, 10, "bold"))

    # 版面

    # 標題列
    def _build_header(self):
        header = ttk.Frame(self.root, padding=(16, 14, 16, 10))
        header.pack(fill="x")
        tk.Label(header, text="棧板明細估算系統", font=(FONT, 20, "bold"),
                 fg=PRIMARY).pack(anchor="w")
        ttk.Label(header, text="依包材規則把出貨件數拆成棧板明細，寫入輸出檔的"
                  f"「{OUTPUT_SHEET}」", style="Muted.TLabel").pack(anchor="w")

    # 左側三個步驟與執行按鈕
    def _build_steps(self, parent):
        steps = ttk.Frame(parent)
        steps.grid(row=0, column=0, sticky="ns", padx=(0, 16))

        self.steps = {}  # {步驟: (路徑變數, 輸入框, 檢查結果)}
        self._step(steps, "rule", "① 條件規則",
                   "欄位：包材、分裝、包裝、單一包裝")
        self._step(steps, "item", "② 需要整合", "欄位：物料、批次、件數")
        output_box = self._step(steps, "output", "③ 輸出檔案",
                                f"明細寫入「{OUTPUT_SHEET}」的 A～D 欄")
        ttk.Checkbutton(output_box, text="使用原檔名 (直接寫入原檔，不另存新檔)",
                        variable=self.keep_name,
                        command=self._refresh_output).pack(anchor="w",
                                                           pady=(4, 0))

        self.run_button = ttk.Button(steps, text="開始估算 (Enter)",
                                     style="Accent.TButton",
                                     command=self.submit)
        self.run_button.pack(fill="x", pady=(4, 0))

    # 建立一個選檔步驟
    def _step(self, parent, key, title, hint):
        box = ttk.LabelFrame(parent, text=f" {title} ", padding=(10, 6))
        box.pack(fill="x", pady=(0, 10))

        var = tk.StringVar()
        row = ttk.Frame(box)
        row.pack(fill="x")
        entry = ttk.Entry(row, textvariable=var, width=30, state="readonly")
        entry.pack(side="left")
        ttk.Button(row, text="選擇…",
                   command=lambda: self.choose_file(key)).pack(
            side="left", padx=(6, 0))

        ttk.Label(box, text=hint, style="Muted.TLabel").pack(
            anchor="w", pady=(4, 0))
        status = tk.Label(box, text="尚未選擇檔案", fg=MUTED, font=(FONT, 10),
                          anchor="w", justify="left", wraplength=330)
        status.pack(fill="x")
        self.steps[key] = (var, entry, status)
        return box

    # 右側結果卡片與明細預覽
    def _build_side(self, parent):
        side = ttk.Frame(parent)
        side.grid(row=0, column=1, sticky="nsew")

        card = tk.Frame(side, bg=CARD_BG, highlightthickness=1,
                        highlightbackground=CARD_BORDER, padx=18, pady=12)
        card.pack(fill="x")
        tk.Label(card, text="估算結果", bg=CARD_BG, fg=MUTED,
                 font=(FONT, 11)).pack(anchor="w")
        self.count_label = tk.Label(card, text="—", bg=CARD_BG, fg=PRIMARY,
                                    font=(FONT, 26, "bold"))
        self.count_label.pack(anchor="w")
        self.detail_label = tk.Label(card, text="選好三個檔案後按「開始估算」",
                                     bg=CARD_BG, justify="left",
                                     font=(FONT, 11))
        self.detail_label.pack(anchor="w")
        self.hint_label = tk.Label(card, text="", bg=CARD_BG, fg=WARN,
                                   justify="left", wraplength=400,
                                   font=(FONT, 10, "bold"))
        self.hint_label.pack(anchor="w", pady=(4, 0))
        self.open_button = ttk.Button(card, text="開啟輸出檔",
                                      command=self.open_output,
                                      state=tk.DISABLED)
        self.open_button.pack(anchor="e")

        preview = ttk.LabelFrame(side, text=" 明細預覽 ", padding=(8, 6))
        preview.pack(fill="both", expand=True, pady=(12, 0))
        self.tree = ttk.Treeview(preview, columns=PREVIEW_COLUMNS,
                                 show="headings", height=PREVIEW_ROWS)
        for name, width in zip(PREVIEW_COLUMNS, COLUMN_WIDTHS):
            self.tree.heading(name, text=name)
            self.tree.column(name, width=width,
                             anchor="w" if name == "品號" else "center")
        scroll = ttk.Scrollbar(preview, orient="vertical",
                               command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll.set)
        self.tree.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")

        legend = ttk.Frame(side)
        legend.pack(anchor="w", pady=(4, 0))
        for kind, (color, text) in ROW_COLORS.items():
            self.tree.tag_configure(kind, background=color)
            tk.Label(legend, width=2, bg=color, relief="solid",
                     borderwidth=1).pack(side="left")
            ttk.Label(legend, text=text, style="Muted.TLabel").pack(
                side="left", padx=(4, 14))

    # 選擇檔案

    # 帶入上次的選擇或預設檔案
    def _load_defaults(self):
        settings = load_settings()
        self.keep_name.set(settings.get("keep_name") is True)
        for key, name in (("rule", DEFAULT_RULE), ("item", DEFAULT_ITEM),
                          ("output", DEFAULT_OUTPUT)):
            path = initial_path(settings.get(key), name)
            if path:
                self.set_file(key, path)

    # 記住這次的選擇
    def _save_choices(self):
        settings = {key: self.path_of(key) for key in self.steps}
        settings["keep_name"] = self.keep_name.get()
        save_settings(settings)

    # 開啟檔案選擇視窗
    def choose_file(self, key):
        current = self.path_of(key)
        start = os.path.dirname(current) if current else app_dir()
        path = filedialog.askopenfilename(initialdir=start,
                                          filetypes=EXCEL_TYPES)
        if path:
            self.set_file(key, path)

    # 填入路徑並檢查檔案
    def set_file(self, key, path):
        var, entry, status = self.steps[key]
        var.set(path)
        entry.xview_moveto(1)  # 捲到最右邊顯示檔名
        try:
            text = self._check(key, path)
        except PalletError as error:
            status.config(text=f"✗ {error}", fg=ERROR)
            return
        status.config(text=f"✓ {text}", fg=SUCCESS)

    # 檢查檔案內容，回傳說明文字
    def _check(self, key, path):
        if key == "rule":
            return f"{len(load_rules(path))} 條包材規則"
        if key == "item":
            items = load_items(path)
            return f"{len(items)} 筆資料，合併後 {len(merge_items(items))} 個品項"
        check_output(path)
        name = os.path.basename(output_path_for(path, self.keep_name.get()))
        if self.keep_name.get():
            return f"將直接寫入原檔 {name}"
        return f"將另存為 {name}"

    # 切換使用原檔名時更新提示
    def _refresh_output(self):
        if self.path_of("output"):
            self.set_file("output", self.path_of("output"))

    # 取得某個步驟目前的路徑
    def path_of(self, key):
        return self.steps[key][0].get()

    # 估算與顯示

    # 執行估算並顯示結果
    def submit(self):
        if self.busy:
            return
        self.busy = True
        self.run_button.config(state=tk.DISABLED)
        self.root.config(cursor="watch")
        self.root.update_idletasks()

        try:
            for key in self.steps:
                if self.path_of(key):
                    self.set_file(key, self.path_of(key))
            result = build_pallet(self.path_of("rule"), self.path_of("item"),
                                  self.path_of("output"), self.keep_name.get())
        except PalletError as error:
            self._show_error(str(error))
        else:
            self._show_result(result)
        finally:
            self.root.after(RESUME_DELAY, self._resume)

    # 估算結束後開放按鈕
    def _resume(self):
        self.busy = False
        self.run_button.config(state=tk.NORMAL)
        self.root.config(cursor="")

    # 顯示估算結果
    def _show_result(self, result):
        self._save_choices()
        self.output_path = result.output_path
        action = "已寫入原檔" if self.keep_name.get() else "已另存"
        now = datetime.now().strftime("%H:%M:%S")
        self.count_label.config(text=f"{result.row_count} 筆明細", fg=PRIMARY)
        self.detail_label.config(
            text=(f"{result.item_count} 個品項拆分完成\n"
                  f"{action} {os.path.basename(self.output_path)}"
                  f"「{OUTPUT_SHEET}」 {now}"))
        if result.unmatched:
            self.hint_label.config(
                text=(f"{len(result.unmatched)} 個品項找不到包材規則：\n"
                      + "、".join(result.unmatched)), fg=WARN)
        else:
            self.hint_label.config(text="")
        self.open_button.config(state=tk.NORMAL)
        self._fill_preview(result.rows)
        self.set_file("output", self.path_of("output"))  # 更新下一次的檔名

    # 顯示錯誤訊息
    def _show_error(self, text):
        self.count_label.config(text="無法估算", fg=ERROR)
        self.detail_label.config(text="請依下方提示修正後再試一次")
        self.hint_label.config(text=text, fg=ERROR)
        self.open_button.config(state=tk.DISABLED)
        self._fill_preview(None)

    # 更新明細預覽並依類型上底色
    def _fill_preview(self, rows):
        self.tree.delete(*self.tree.get_children())
        if rows is None:
            return
        rows = rows.itertuples(index=False)
        for no, (item_code, batch, qty, package, kind) in enumerate(rows,
                                                                    start=1):
            self.tree.insert("", "end", tags=(kind,),
                             values=(no, item_code, batch, format_qty(qty),
                                     package))

    # 開啟寫入的檔案
    def open_output(self):
        if self.output_path and os.path.exists(self.output_path):
            os.startfile(self.output_path)


# 啟動程式
def main():
    root = tk.Tk()
    PalletApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
