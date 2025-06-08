
import tkinter as tk
from tkinter import messagebox, ttk, filedialog
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors
from datetime import datetime
import csv
import os
import pyperclip
import re

def get_next_estimate_number():
    filename = "quote_history.csv"
    if not os.path.exists(filename):
        return "Q-00001"
    with open(filename, "r", encoding="utf-8-sig") as f:
        lines = f.readlines()
        if len(lines) < 2:
            return "Q-00001"
        last_line = lines[-1]
        last_no = last_line.split(",")[0].strip().replace("Q-", "")
        next_no = int(last_no) + 1
        return f"Q-{next_no:05d}"

def paste_from_excel():
    try:
        clipboard = pyperclip.paste()
        rows = clipboard.strip().split("\n")
        for i, row in enumerate(rows):
            if i >= len(item_entries):
                break
            cols = row.split("\t")
            if len(cols) >= 4:
                item_entries[i][0].delete(0, tk.END)
                item_entries[i][0].insert(0, cols[0])
                item_entries[i][1].delete(0, tk.END)
                item_entries[i][1].insert(0, cols[1])
                item_entries[i][2].delete(0, tk.END)
                item_entries[i][2].insert(0, cols[2])
                item_entries[i][4].delete(0, tk.END)
                item_entries[i][4].insert(0, cols[3])
        update_total()
    except Exception as e:
        messagebox.showerror("エラー", f"貼り付け失敗：{e}")

def update_total():
    subtotal = 0
    for row in item_entries:
        try:
            qty = int(row[1].get())
            unit_price = int(row[2].get())
            total = qty * unit_price
            row[3].config(state='normal')
            row[3].delete(0, tk.END)
            row[3].insert(0, str(total))
            row[3].config(state='readonly')
            subtotal += total
        except:
            row[3].config(state='normal')
            row[3].delete(0, tk.END)
            row[3].insert(0, "")
            row[3].config(state='readonly')
            continue
    tax = int(subtotal * 0.10)
    total = subtotal + tax
    subtotal_var.set(f"小計：{subtotal:,} 円")
    tax_var.set(f"消費税：{tax:,} 円")
    total_var.set(f"合計：{total:,} 円")

def generate_pdf():
    global last_pdf_path
    customer = customer_entry.get()
    title = title_entry.get()
    estimate_no = estimate_no_entry.get()
    today = datetime.today().strftime('%Y-%m-%d')
    remarks = remarks_entry.get("1.0", tk.END).strip()

    items = []
    for row in item_entries:
        name = row[0].get()
        qty = row[1].get()
        unit_price = row[2].get()
        note = row[4].get()
        if name and qty and unit_price:
            try:
                total = int(qty) * int(unit_price)
                items.append([name, str(qty), f"¥{int(unit_price):,}", f"¥{total:,}", note])
            except:
                messagebox.showerror("エラー", "数量と単価は数字で入力してください")
                return

    if not items:
        messagebox.showwarning("未入力", "出力対象の品目がありません。")
        return

    subtotal = sum([int(row[3].replace("¥", "").replace(",", "")) for row in items])
    tax = int(subtotal * 0.10)
    total = subtotal + tax

    folder = filedialog.askdirectory(title="保存先フォルダを選択してください")
    if not folder:
        return

    safe_customer = re.sub(r'[\\/*?:"<>|]', "_", customer)
    filename = os.path.join(folder, f"見積_{today}_{safe_customer}.pdf")
    last_pdf_path = filename

    styles = getSampleStyleSheet()
    doc = SimpleDocTemplate(filename, pagesize=A4)
    elements = []

    elements.append(Paragraph(f"見積日：{today}　見積番号：{estimate_no}", styles["Normal"]))
    elements.append(Spacer(1, 6))
    elements.append(Paragraph("K3ソリューションズ株式会社", styles["Normal"]))
    elements.append(Paragraph("〒511-1112 三重県桑名市長島町大倉1番地408", styles["Normal"]))
    elements.append(Paragraph("TEL/FAX：(0594) 84-6019", styles["Normal"]))
    elements.append(Spacer(1, 6))
    elements.append(Paragraph(f"御中：{customer}", styles["Normal"]))
    elements.append(Paragraph(f"件名：{title}", styles["Normal"]))
    elements.append(Spacer(1, 12))

    table_data = [["品名", "数量", "単価", "金額", "摘要"]] + items
    table_data += [["", "", "", "", ""]]
    table_data += [["小計", "", "", f"¥{subtotal:,}", ""]]
    table_data += [["消費税", "", "", f"¥{tax:,}", ""]]
    table_data += [["合計", "", "", f"¥{total:,}", ""]]

    table = Table(table_data, colWidths=[100, 50, 70, 80, 120])
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
        ("ALIGN", (1, 1), (-2, -4), "RIGHT"),
        ("ALIGN", (-2, -3), (-2, -1), "RIGHT"),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTNAME", (0, 1), (-1, -1), "Helvetica"),
        ("FONTSIZE", (0, 0), (-1, -1), 10),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
    ]))
    elements.append(table)

    elements.append(Spacer(1, 12))
    elements.append(Paragraph("備考：", styles["Normal"]))
    for line in remarks.split("\n"):
        elements.append(Paragraph(line, styles["Normal"]))

    doc.build(elements)
    messagebox.showinfo("完了", f"PDFを保存しました：\n{filename}")
    print_button.config(state="normal")

    # CSVも出力
    csv_path = os.path.join(folder, f"見積_{today}_{safe_customer}.csv")
    with open(csv_path, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(["品名", "数量", "単価", "金額", "摘要"])
        writer.writerows(items)
        writer.writerow([])
        writer.writerow(["小計", "", "", subtotal, ""])
        writer.writerow(["消費税", "", "", tax, ""])
        writer.writerow(["合計", "", "", total, ""])

def print_pdf():
    if last_pdf_path and os.path.exists(last_pdf_path):
        try:
            os.startfile(last_pdf_path, "print")
        except Exception as e:
            messagebox.showerror("印刷エラー", f"PDFの印刷に失敗しました：\n{e}")

root = tk.Tk()
root.title("K3見積書作成ツール")
root.geometry("980x720")
root.configure(bg="#fff0e6")

style = ttk.Style()
style.configure("TLabel", background="#fff0e6", font=("Arial", 10))
style.configure("TButton", font=("Arial", 10))
style.configure("TFrame", background="#fff0e6")

frame_top = ttk.Frame(root, padding=10)
frame_top.pack(fill="x")
tk.Label(frame_top, text="顧客名", bg="#fff0e6").grid(row=0, column=0)
customer_entry = tk.Entry(frame_top, width=30, bg="#fff0e6")
customer_entry.grid(row=0, column=1)
tk.Label(frame_top, text="件名", bg="#fff0e6").grid(row=0, column=2)
title_entry = tk.Entry(frame_top, width=30, bg="#fff0e6")
title_entry.grid(row=0, column=3)
tk.Label(frame_top, text="見積番号", bg="#fff0e6").grid(row=0, column=4)
estimate_no_entry = tk.Entry(frame_top, width=15, bg="#fff0e6")
estimate_no_entry.grid(row=0, column=5)
estimate_no_entry.insert(0, get_next_estimate_number())

frame_items = ttk.Frame(root)
frame_items.pack(fill="both", expand=True, padx=10, pady=5)
tk_canvas = tk.Canvas(frame_items, height=400, bg="#fff0e6")
scrollbar = ttk.Scrollbar(frame_items, orient="vertical", command=tk_canvas.yview)
scrollable_frame = ttk.Frame(tk_canvas)
scrollable_frame.bind("<Configure>", lambda e: tk_canvas.configure(scrollregion=tk_canvas.bbox("all")))
tk_canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
tk_canvas.configure(yscrollcommand=scrollbar.set)
tk_canvas.pack(side="left", fill="both", expand=True)
scrollbar.pack(side="right", fill="y")

cols = ["品名", "数量", "単価", "合計", "摘要"]
item_entries = []
for i, col in enumerate(cols):
    tk.Label(scrollable_frame, text=col, bg="#fff0e6").grid(row=0, column=i)

for i in range(30):
    row = []
    for j in range(5):
        width = [30, 8, 10, 12, 20][j]
        state = "readonly" if j == 3 else "normal"
        e = tk.Entry(scrollable_frame, width=width, state=state, bg="#fff0e6")
        e.grid(row=i+1, column=j, padx=2, pady=2)
        if j != 3:
            e.bind("<KeyRelease>", lambda event: update_total())
        row.append(e)
    item_entries.append(row)

frame_bottom = ttk.Frame(root, padding=10)
frame_bottom.pack(fill="x")
tk.Label(frame_bottom, text="備考", bg="#fff0e6").grid(row=0, column=0)
remarks_entry = tk.Text(frame_bottom, height=3, width=80, bg="#fff0e6")
remarks_entry.grid(row=0, column=1, columnspan=4)

subtotal_var = tk.StringVar()
tax_var = tk.StringVar()
total_var = tk.StringVar()
tk.Label(frame_bottom, textvariable=subtotal_var, bg="#fff0e6").grid(row=1, column=1, sticky="e")
tk.Label(frame_bottom, textvariable=tax_var, bg="#fff0e6").grid(row=2, column=1, sticky="e")
tk.Label(frame_bottom, textvariable=total_var, bg="#fff0e6").grid(row=3, column=1, sticky="e")

frame_buttons = ttk.Frame(root, padding=10)
frame_buttons.pack()
ttk.Button(frame_buttons, text="PDF出力", command=generate_pdf).grid(row=0, column=0, padx=5)
ttk.Button(frame_buttons, text="Excel貼り付け", command=paste_from_excel).grid(row=0, column=1, padx=5)
print_button = ttk.Button(frame_buttons, text="印刷", command=print_pdf, state="disabled")
print_button.grid(row=0, column=2, padx=5)

last_pdf_path = None
update_total()

if __name__ == "__main__":
    try:
        root.mainloop()
    except Exception as e:
        import traceback
        traceback.print_exc()
        input("エラーが発生しました。Enterキーで終了します。")
