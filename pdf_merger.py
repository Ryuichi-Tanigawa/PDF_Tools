# 作成者: Ryuichi Tanigawa
# 本コードは Copilot により 2025-11-27 に作成されました
# 機能: 複数のPDFをGUIで選択し、Listboxで順番を編集して結合します
# 保存は「Save As ダイアログ」でフォルダとファイル名を一度に指定可能
# 既存ファイルがある場合は上書き確認を行い、初期ファイル名は先頭のPDF名を使用

import os
import re
import tkinter as tk
from tkinter import filedialog, messagebox
from PyPDF2 import PdfReader, PdfWriter

def merge_pdfs_gui():
    root = tk.Tk()
    root.withdraw()

    # PDFファイル選択
    file_paths = filedialog.askopenfilenames(
        title="結合するPDFを選択してください（複数選択可）",
        filetypes=[("PDF files", "*.pdf")]
    )
    if not file_paths:
        print("ファイルが選択されませんでした")
        return

    # 並べ替え用ウィンドウ
    sort_win = tk.Toplevel()
    sort_win.title("結合順序編集（ドラッグ＆ドロップ＋上下ボタン対応）")

    listbox = tk.Listbox(sort_win, selectmode=tk.SINGLE, width=80)
    listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    for f in file_paths:
        listbox.insert(tk.END, f)

    drag_data = {"item": None, "index": None}

    def on_start_drag(event):
        idx = listbox.nearest(event.y)
        if idx >= 0:
            drag_data["item"] = listbox.get(idx)
            drag_data["index"] = idx

    def on_drag_motion(event):
        idx = listbox.nearest(event.y)
        listbox.selection_clear(0, tk.END)
        listbox.selection_set(idx)

    def on_drop(event):
        idx = listbox.nearest(event.y)
        if drag_data["item"] is not None and drag_data["index"] is not None:
            listbox.delete(drag_data["index"])
            listbox.insert(idx, drag_data["item"])
            listbox.selection_clear(0, tk.END)
            listbox.selection_set(idx)
        drag_data["item"] = None
        drag_data["index"] = None

    listbox.bind("<Button-1>", on_start_drag)
    listbox.bind("<B1-Motion>", on_drag_motion)
    listbox.bind("<ButtonRelease-1>", on_drop)

    btn_frame = tk.Frame(sort_win)
    btn_frame.pack(side=tk.RIGHT, fill=tk.Y)

    def move_up():
        sel = listbox.curselection()
        if not sel:
            return
        idx = sel[0]
        if idx == 0:
            return
        text = listbox.get(idx)
        listbox.delete(idx)
        listbox.insert(idx-1, text)
        listbox.selection_set(idx-1)

    def move_down():
        sel = listbox.curselection()
        if not sel:
            return
        idx = sel[0]
        if idx == listbox.size()-1:
            return
        text = listbox.get(idx)
        listbox.delete(idx)
        listbox.insert(idx+1, text)
        listbox.selection_set(idx+1)

    ordered_files = []

    def confirm_order():
        nonlocal ordered_files
        ordered_files = list(listbox.get(0, tk.END))
        sort_win.destroy()

    tk.Button(btn_frame, text="↑ 上へ", command=move_up).pack(fill=tk.X)
    tk.Button(btn_frame, text="↓ 下へ", command=move_down).pack(fill=tk.X)
    tk.Button(btn_frame, text="確定", command=confirm_order).pack(fill=tk.X)

    sort_win.wait_window()

    # Save As ダイアログ（初期ファイル名は先頭のPDF名）
    first_name = os.path.splitext(os.path.basename(ordered_files[0]))[0]
    save_path = filedialog.asksaveasfilename(
        title="保存先とファイル名を指定してください",
        defaultextension=".pdf",
        initialfile=first_name + ".pdf",
        filetypes=[("PDF files", "*.pdf")]
    )
    if not save_path:
        print("保存先が指定されませんでした")
        return

    # 既存ファイルがある場合は上書き確認
    if os.path.exists(save_path):
        overwrite = messagebox.askyesno("確認", f"{save_path} は既に存在します。上書きしますか？")
        if not overwrite:
            print("保存をキャンセルしました")
            return

    try:
        writer = PdfWriter()
        for file_path in ordered_files:
            reader = PdfReader(file_path)
            for page in reader.pages:
                writer.add_page(page)

        with open(save_path, "wb") as f:
            writer.write(f)

        messagebox.showinfo("完了", f"{len(ordered_files)} 個のPDFを結合し、\n{save_path} に保存しました")

    except Exception as e:
        messagebox.showerror("エラー", f"エラーが発生しました: {e}")

if __name__ == "__main__":
    merge_pdfs_gui()