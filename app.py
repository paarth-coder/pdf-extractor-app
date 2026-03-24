
import pdfplumber
import pandas as pd
import re
from collections import defaultdict
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinterdnd2 import DND_FILES, TkinterDnD

def extract_pdf(pdf_path):
    model_pattern = re.compile(r'([A-Z0-9]+)\s*\[ASE1\]')
    number_pattern = re.compile(r'\b\d{6,8}\b')
    qty_pattern = re.compile(r'\s(\d+)\s+\d{1,3},')

    model_data = defaultdict(list)
    model_qty = {}

    current_model = None
    inside_particulars = False

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                lines = text.split("\n")
                for line in lines:
                    if "Particulars" in line:
                        inside_particulars = True
                        continue
                    if "Total" in line:
                        inside_particulars = False
                    if not inside_particulars:
                        continue

                    model_match = model_pattern.search(line)
                    if model_match:
                        current_model = model_match.group(1)
                        qty_match = qty_pattern.search(line)
                        if qty_match:
                            model_qty[current_model] = int(qty_match.group(1))
                        else:
                            model_qty[current_model] = None
                        continue

                    if current_model:
                        numbers = number_pattern.findall(line)
                        model_data[current_model].extend(numbers)

    for key in model_data:
        model_data[key] = list(dict.fromkeys(model_data[key]))

    df = pd.DataFrame(dict([(k, pd.Series(v)) for k, v in model_data.items()]))
    df.insert(0, "S.No", range(1, len(df) + 1))

    summary = []
    for model in model_data:
        extracted = len(model_data[model])
        expected = model_qty.get(model)
        summary.append({
            "Model": model,
            "Expected Qty": expected,
            "Extracted Count": extracted,
            "Match": "OK" if expected == extracted else "Mismatch"
        })

    summary_df = pd.DataFrame(summary)

    output_path = pdf_path.replace(".pdf", "_output.xlsx")
    with pd.ExcelWriter(output_path) as writer:
        df.to_excel(writer, sheet_name="Data", index=False)
        summary_df.to_excel(writer, sheet_name="Qty_Check", index=False)

    return output_path

def browse_file():
    file_path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
    if file_path:
        process_file(file_path)

def process_file(file_path):
    try:
        output = extract_pdf(file_path)
        messagebox.showinfo("Success", f"File processed!\nSaved at:\n{output}")
    except Exception as e:
        messagebox.showerror("Error", str(e))

def drop_file(event):
    file_path = event.data.strip("{}")
    process_file(file_path)

app = TkinterDnD.Tk()
app.title("PDF Extractor")
app.geometry("400x250")

label = tk.Label(app, text="Drag & Drop PDF Here\nor Click Button", font=("Arial", 14))
label.pack(pady=40)

btn = tk.Button(app, text="Select PDF", command=browse_file, bg="black", fg="white")
btn.pack(pady=10)

app.drop_target_register(DND_FILES)
app.dnd_bind('<<Drop>>', drop_file)

app.mainloop()
