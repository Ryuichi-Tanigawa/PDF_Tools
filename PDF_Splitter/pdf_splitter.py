# 作成者: Ryuichi Tanigawa
# 本コードは Copilot により 2025-11-27 に作成されました
# 機能: PDFをページごとに分割し、GUIで入力ファイルと出力先を指定可能
#       ページ範囲を指定しない場合は全ページを分割します
#       ページ範囲は「1-5」や「3」や「1-3,7-9」のように指定可能です
# 出力ファイル名: 0001-元ファイル名.pdf の形式

import os
from tkinter import Tk, filedialog, simpledialog
from PyPDF2 import PdfReader, PdfWriter

def split_pdf_gui():
    # 入力ファイル選択
    root = Tk()
    root.withdraw()  # Tkウィンドウを非表示
    file_path = filedialog.askopenfilename(
        title="分割するPDFを選択してください",
        filetypes=[("PDF files", "*.pdf")]
    )
    if not file_path:
        print("ファイルが選択されませんでした")
        return

    # 出力先フォルダ選択
    output_dir = filedialog.askdirectory(
        title="出力先フォルダを選択してください"
    )
    if not output_dir:
        print("出力先フォルダが選択されませんでした")
        return

    try:
        base_name = os.path.basename(file_path)  # 元ファイル名のみ
        reader = PdfReader(file_path)
        total_pages = len(reader.pages)

        # ページ範囲を入力（未入力なら全ページ）
        page_range = simpledialog.askstring(
            "ページ範囲指定",
            f"分割するページ範囲を入力してください (例: 1-{total_pages}, 3, 1-3,7-9)\n"
            "未入力の場合は全ページを出力します"
        )

        # ページ範囲を解析
        ranges = []
        if page_range:
            try:
                parts = page_range.split(",")
                for part in parts:
                    part = part.strip()
                    if "-" in part:  # 範囲指定
                        start_str, end_str = part.split("-")
                        start_page = int(start_str)
                        end_page = int(end_str)
                        if start_page < 1 or end_page > total_pages or start_page > end_page:
                            raise ValueError
                        ranges.extend(range(start_page, end_page + 1))
                    else:  # 単ページ指定
                        page = int(part)
                        if page < 1 or page > total_pages:
                            raise ValueError
                        ranges.append(page)
            except Exception:
                print("ページ範囲の入力形式が不正です。例: 1-5, 3, 1-3,7-9")
                return
        else:
            ranges = list(range(1, total_pages + 1))

        # 指定範囲で分割
        for page_num in ranges:
            writer = PdfWriter()
            writer.add_page(reader.pages[page_num - 1])

            # ファイル名構成: 0001-元ファイル名.pdf
            output_file = os.path.join(output_dir, f"{page_num:04d}-{base_name}")
            with open(output_file, "wb") as f:
                writer.write(f)

        print(f"{file_path} の {ranges} ページを分割し、{output_dir} に保存しました")

    except Exception as e:
        print(f"エラーが発生しました: {e}")

if __name__ == "__main__":
    split_pdf_gui()