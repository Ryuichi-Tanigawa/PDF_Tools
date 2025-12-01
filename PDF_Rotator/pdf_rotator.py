# 作成者: Ryuichi Tanigawa
# 本コードは Copilot により 2025-12-01 に作成されました
# 機能: 複数のPDFをGUIで選択し、Listboxで順番を編集して結合
#       さらに指定ページをGUIで任意角度回転可能
# 保存は「Save As ダイアログ」でフォルダとファイル名を一度に指定可能
# 既存ファイルがある場合は上書き確認を行い、初期ファイル名は先頭のPDF名を使用

import os
import tkinter as tk
from tkinter import filedialog, messagebox
from PyPDF2 import PdfReader, PdfWriter

def merge_and_rotate_pdfs_gui():
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
    sort_win.title("結合順序編集")

    listbox = tk.Listbox(sort_win, selectmode=tk.SINGLE, width=80)
    listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    for f in file_paths:
        listbox.insert(tk.END, f)

    btn_frame = tk.Frame(sort_win)
    btn_frame.pack(side=tk.RIGHT, fill=tk.Y)

    def move_up():
        sel = listbox.curselection()
        if not sel: return
        idx = sel[0]
        if idx == 0: return
        text = listbox.get(idx)
        listbox.delete(idx)
        listbox.insert(idx-1, text)
        listbox.selection_set(idx-1)

    def move_down():
        sel = listbox.curselection()
        if not sel: return
        idx = sel[0]
        if idx == listbox.size()-1: return
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

    # 回転設定ウィンドウ
    rotate_win = tk.Toplevel()
    rotate_win.title("ページ回転設定")

    tk.Label(rotate_win, text="ページ番号（0始まり）:").grid(row=0, column=0)
    page_entry = tk.Entry(rotate_win)
    page_entry.grid(row=0, column=1)

    tk.Label(rotate_win, text="回転角度（時計回り）:").grid(row=1, column=0)
    angle_entry = tk.Entry(rotate_win)
    angle_entry.grid(row=1, column=1)

    rotate_listbox = tk.Listbox(rotate_win, width=40)
    rotate_listbox.grid(row=2, column=0, columnspan=2)

    rotations = []

    def add_rotation():
        try:
            page_num = int(page_entry.get())
            angle = int(angle_entry.get())
            rotations.append((page_num, angle))
            rotate_listbox.insert(tk.END, f"Page {page_num} → {angle}°")
            page_entry.delete(0, tk.END)
            angle_entry.delete(0, tk.END)
        except ValueError:
            messagebox.showerror("入力エラー", "ページ番号と角度は整数で入力してください")

    tk.Button(rotate_win, text="追加", command=add_rotation).grid(row=3, column=0, columnspan=2)
    tk.Button(rotate_win, text="確定", command=rotate_win.destroy).grid(row=4, column=0, columnspan=2)

    rotate_win.wait_window()

    # Save As ダイアログ
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

    if os.path.exists(save_path):
        overwrite = messagebox.askyesno("確認", f"{save_path} は既に存在します。上書きしますか？")
        if not overwrite:
            print("保存をキャンセルしました")
            return

    try:
        writer = PdfWriter()
        page_index = 0
        for file_path in ordered_files:
            reader = PdfReader(file_path)
            for page in reader.pages:
                # 回転対象なら回転
                for target_page, angle in rotations:
                    if page_index == target_page:
                        page = page.rotate(angle)
                writer.add_page(page)
                page_index += 1

        with open(save_path, "wb") as f:
            writer.write(f)

        messagebox.showinfo("完了", f"{len(ordered_files)} 個のPDFを結合し、指定ページを回転して\n{save_path} に保存しました")

    except Exception as e:
        messagebox.showerror("エラー", f"エラーが発生しました: {e}")

if __name__ == "__main__":
    merge_and_rotate_pdfs_gui()
