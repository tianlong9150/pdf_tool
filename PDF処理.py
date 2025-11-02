import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from PyPDF2 import PdfReader, PdfWriter

# ------------------- PDF操作関数 -------------------


def extract_pages(input_pdf, pages, output_path):
    """指定されたページを抽出"""
    reader = PdfReader(input_pdf)
    writer = PdfWriter()
    total_pages = len(reader.pages)

    for p in pages:
        if 1 <= p <= total_pages:
            writer.add_page(reader.pages[p - 1])
        else:
            raise ValueError(f"ページ {p} は範囲外です (1-{total_pages})")

    with open(output_path, "wb") as f:
        writer.write(f)


def merge_pdfs(file_list, output_path):
    """複数のPDFを結合"""
    writer = PdfWriter()
    for f in file_list:
        reader = PdfReader(f)
        for page in reader.pages:
            writer.add_page(page)

    with open(output_path, "wb") as out_file:
        writer.write(out_file)


def alternate_merge(pdf1_path, pdf2_path, output_path):
    """2つのPDFを交互に結合"""
    reader1 = PdfReader(pdf1_path)
    reader2 = PdfReader(pdf2_path)
    writer = PdfWriter()
    max_pages = max(len(reader1.pages), len(reader2.pages))

    for i in range(max_pages):
        if i < len(reader1.pages):
            writer.add_page(reader1.pages[i])
        if i < len(reader2.pages):
            writer.add_page(reader2.pages[i])

    with open(output_path, "wb") as f:
        writer.write(f)


def reverse_pdf(input_pdf, output_path):
    """PDFのページを逆順にする"""
    reader = PdfReader(input_pdf)
    writer = PdfWriter()

    for page in reversed(reader.pages):
        writer.add_page(page)

    with open(output_path, "wb") as f:
        writer.write(f)


# ------------------- GUI関数 -------------------


def select_file(entry):
    """PDFファイルを選択"""
    path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    if path:
        entry.delete(0, tk.END)
        entry.insert(0, path)


def select_output_dir():
    """出力フォルダを選択"""
    path = filedialog.askdirectory()
    if path:
        output_entry.delete(0, tk.END)
        output_entry.insert(0, path)


def parse_pages(pages_str):
    """ページ指定文字列をパース (例: "1,3-5,7")"""
    pages = []
    try:
        for part in pages_str.strip().split(","):
            part = part.strip()
            if "-" in part:
                start, end = part.split("-")
                start, end = int(start.strip()), int(end.strip())
                if start > end:
                    raise ValueError(f"範囲指定が不正です: {start}-{end}")
                pages.extend(range(start, end + 1))
            else:
                pages.append(int(part))
        return pages
    except ValueError as e:
        raise ValueError(f"ページ指定が不正です: {pages_str}\n正しい形式: 1,3-5,7")


def update_inputs(event=None):
    """選択された機能に応じて入力欄を表示"""
    mode = mode_combobox.get()

    # すべての入力欄を非表示
    for label, entry, button in pdf_entries_buttons:
        label.grid_forget()
        entry.grid_forget()
        button.grid_forget()
    pages_label.grid_forget()
    pages_entry.grid_forget()
    merge_count_label.grid_forget()
    merge_count_combobox.grid_forget()

    # 出力ファイル名のデフォルト値を設定
    if mode == "PDF抽出":
        pdf_entries_buttons[0][0].config(text="PDF:")
        pdf_entries_buttons[0][0].grid(row=1, column=0, padx=5, pady=5, sticky="e")
        pdf_entries_buttons[0][1].grid(row=1, column=1, padx=5, pady=5, sticky="ew")
        pdf_entries_buttons[0][2].grid(row=1, column=2, padx=5, pady=5, sticky="w")
        pages_label.grid(row=2, column=0, padx=5, pady=5, sticky="e")
        pages_entry.grid(row=2, column=1, padx=5, pady=5, sticky="ew")
        output_filename_entry.delete(0, tk.END)
        output_filename_entry.insert(0, "extracted.pdf")

    elif mode == "PDF逆順":
        pdf_entries_buttons[0][0].config(text="PDF:")
        pdf_entries_buttons[0][0].grid(row=1, column=0, padx=5, pady=5, sticky="e")
        pdf_entries_buttons[0][1].grid(row=1, column=1, padx=5, pady=5, sticky="ew")
        pdf_entries_buttons[0][2].grid(row=1, column=2, padx=5, pady=5, sticky="w")
        output_filename_entry.delete(0, tk.END)
        output_filename_entry.insert(0, "reversed.pdf")

    elif mode == "PDF結合":
        merge_count_label.grid(row=1, column=0, padx=5, pady=5, sticky="e")
        merge_count_combobox.grid(row=1, column=1, padx=5, pady=5, sticky="w")
        show_pdf_entries()
        output_filename_entry.delete(0, tk.END)
        output_filename_entry.insert(0, "merged.pdf")

    elif mode == "PDF交互結合":
        pdf_entries_buttons[0][0].config(text="奇数ページ:")
        pdf_entries_buttons[0][0].grid(row=1, column=0, padx=5, pady=5, sticky="e")
        pdf_entries_buttons[0][1].grid(row=1, column=1, padx=5, pady=5, sticky="ew")
        pdf_entries_buttons[0][2].grid(row=1, column=2, padx=5, pady=5, sticky="w")
        pdf_entries_buttons[1][0].config(text="偶数ページ:")
        pdf_entries_buttons[1][0].grid(row=2, column=0, padx=5, pady=5, sticky="e")
        pdf_entries_buttons[1][1].grid(row=2, column=1, padx=5, pady=5, sticky="ew")
        pdf_entries_buttons[1][2].grid(row=2, column=2, padx=5, pady=5, sticky="w")
        output_filename_entry.delete(0, tk.END)
        output_filename_entry.insert(0, "alternate_merged.pdf")


def show_pdf_entries(event=None):
    """結合するPDFの数に応じてエントリを表示"""
    # まず全て非表示
    for label, entry, button in pdf_entries_buttons:
        label.grid_forget()
        entry.grid_forget()
        button.grid_forget()

    # 指定された数だけ表示（PDF1, PDF2...のラベルに変更）
    count = int(merge_count_combobox.get())
    for i in range(count):
        pdf_entries_buttons[i][0].config(text=f"PDF{i+1}:")
        pdf_entries_buttons[i][0].grid(row=2 + i, column=0, padx=5, pady=5, sticky="e")
        pdf_entries_buttons[i][1].grid(row=2 + i, column=1, padx=5, pady=5, sticky="ew")
        pdf_entries_buttons[i][2].grid(row=2 + i, column=2, padx=5, pady=5, sticky="w")


def execute():
    """選択された機能を実行"""
    mode = mode_combobox.get()

    if not mode:
        messagebox.showwarning("警告", "機能を選択してください")
        return

    output_dir = output_entry.get()
    if not output_dir:
        messagebox.showwarning("警告", "出力フォルダを選択してください")
        return

    if not os.path.isdir(output_dir):
        messagebox.showwarning("警告", "出力フォルダが存在しません")
        return

    output_filename = output_filename_entry.get().strip()
    if not output_filename:
        messagebox.showwarning("警告", "出力ファイル名を入力してください")
        return

    if not output_filename.lower().endswith(".pdf"):
        output_filename += ".pdf"

    output_path = os.path.join(output_dir, output_filename)

    try:
        if mode == "PDF抽出":
            input_pdf = pdf_entries_buttons[0][1].get()
            if not input_pdf:
                messagebox.showwarning("警告", "PDFファイルを選択してください")
                return

            pages_str = pages_entry.get().strip()
            if not pages_str:
                messagebox.showwarning("警告", "ページ指定を入力してください")
                return

            pages = parse_pages(pages_str)
            extract_pages(input_pdf, pages, output_path)

        elif mode == "PDF逆順":
            input_pdf = pdf_entries_buttons[0][1].get()
            if not input_pdf:
                messagebox.showwarning("警告", "PDFファイルを選択してください")
                return

            reverse_pdf(input_pdf, output_path)

        elif mode == "PDF結合":
            count = int(merge_count_combobox.get())
            files = [pdf_entries_buttons[i][1].get() for i in range(count)]

            if any(f == "" for f in files):
                messagebox.showwarning("警告", "すべてのPDFを選択してください")
                return

            merge_pdfs(files, output_path)

        elif mode == "PDF交互結合":
            pdf1 = pdf_entries_buttons[0][1].get()
            pdf2 = pdf_entries_buttons[1][1].get()

            if not pdf1 or not pdf2:
                messagebox.showwarning("警告", "両方のPDFを選択してください")
                return

            alternate_merge(pdf1, pdf2, output_path)

        messagebox.showinfo("完了", f"{mode} が完了しました\n出力: {output_path}")

    except FileNotFoundError as e:
        messagebox.showerror("エラー", f"ファイルが見つかりません:\n{e}")
    except ValueError as e:
        messagebox.showerror("エラー", str(e))
    except Exception as e:
        messagebox.showerror("エラー", f"予期しないエラーが発生しました:\n{e}")


# ------------------- GUI構築 -------------------

root = tk.Tk()
root.title("PDFツール統合")
root.geometry("700x450")

# グリッドの列の重み設定（中央列を伸縮可能に）
root.columnconfigure(1, weight=1)
root.rowconfigure(9, weight=1)  # 出力設定の上の行に重みを設定

# 機能選択
tk.Label(root, text="機能選択:", font=("", 10), width=12, anchor="e").grid(
    row=0, column=0, padx=5, pady=10, sticky="e"
)
mode_combobox = ttk.Combobox(
    root,
    values=["PDF抽出", "PDF逆順", "PDF結合", "PDF交互結合"],
    state="readonly",
    font=("", 10),
)
mode_combobox.grid(row=0, column=1, padx=5, pady=10, sticky="w")
mode_combobox.bind("<<ComboboxSelected>>", update_inputs)

# PDF選択用エントリとボタン（最大5つ）
pdf_entries_buttons = []
for i in range(5):
    # PDF1つの機能用はラベルを「PDF:」に、複数の場合は「PDF1:」「PDF2:」等に
    label = tk.Label(root, text="PDF:", font=("", 10), width=12, anchor="e")
    entry = tk.Entry(root, font=("", 10))
    button = tk.Button(
        root,
        text="参照",
        command=lambda e=entry: select_file(e),
        font=("", 10),
        width=8,
    )
    pdf_entries_buttons.append((label, entry, button))

# ページ指定（抽出用）
pages_label = tk.Label(root, text="ページ指定:", font=("", 10), width=12, anchor="e")
pages_entry = tk.Entry(root, font=("", 10))

# 結合PDF数選択（2～5）
merge_count_label = tk.Label(root, text="PDF数:", font=("", 10), width=12, anchor="e")
merge_count_combobox = ttk.Combobox(
    root, values=[2, 3, 4, 5], state="readonly", font=("", 10), width=5
)
merge_count_combobox.set(2)
merge_count_combobox.bind("<<ComboboxSelected>>", show_pdf_entries)

# 出力フォルダ（下部に固定）
output_frame = tk.Frame(root)
output_frame.grid(row=10, column=0, columnspan=3, padx=5, pady=10, sticky="sew")
root.rowconfigure(10, weight=0)

tk.Label(output_frame, text="出力フォルダ:", font=("", 10), width=12, anchor="e").grid(
    row=0, column=0, padx=5, sticky="e"
)
output_entry = tk.Entry(output_frame, font=("", 10))
output_entry.grid(row=0, column=1, padx=5, sticky="ew")
output_button = tk.Button(
    output_frame, text="参照", command=select_output_dir, font=("", 10), width=8
)
output_button.grid(row=0, column=2, padx=5, sticky="w")

tk.Label(
    output_frame, text="出力ファイル名:", font=("", 10), width=12, anchor="e"
).grid(row=1, column=0, padx=5, pady=(10, 0), sticky="e")
output_filename_entry = tk.Entry(output_frame, font=("", 10))
output_filename_entry.grid(row=1, column=1, padx=5, pady=(10, 0), sticky="ew")
output_filename_entry.insert(0, "output.pdf")

# 出力フレームの列の重み設定（エントリを伸縮可能に）
output_frame.columnconfigure(1, weight=1)

# 実行ボタン（下部に固定）
button_frame = tk.Frame(root)
button_frame.grid(row=11, column=0, columnspan=3, sticky="sew", pady=10)
root.rowconfigure(11, weight=0)

execute_button = tk.Button(
    button_frame,
    text="実行",
    command=execute,
    bg="#4CAF50",
    fg="white",
    font=("", 12, "bold"),
    padx=20,
    pady=10,
)
execute_button.pack()

# 使い方の説明
help_text = "ページ指定例: 1,3-5,7 → 1,3,4,5,7ページを抽出"
tk.Label(button_frame, text=help_text, font=("", 9), fg="gray").pack(pady=5)

root.mainloop()
