import pdfplumber
import re
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import os


# ----- PDF TEXT EXTRACTION --------
def extract_text_from_pdf(pdf_path):
    """Extract text from all pages of a given PDF file"""
    text = ""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text() or ""
            text += page_text + "\n"
    return text



#------------ SUPPLIER DETECTION -------------
def detect_supplier(text):
    """Return supplier name based on unique keywords in the PDF"""
    text_lower = text.lower()
    if "laxmico" in text_lower:
        return "Colorama"
    elif "aah" in text_lower:
        return "AAH"
    elif "alliance" in text_lower:
        return "Alliance"
    elif "lexon" in text_lower:
        return "Lexon"
    else: 
        return "Unknown"

#------------------ LINE ITEM EXTRACTION ------------
def extract_line_items(text):
    lines = text.splitlines()
    items = []

    pattern = re.compile(
    r"^(?P<description>[A-Za-z0-9%()\-/.,\s]+?)\s+"
    r"(?P<pack_size>\d+\s*[a-zA-Z]*)\s+"
    r"(?P<qty>\d+)\s+"
    r"(?P<unit_price>\d+(?:\.\d+)?)\s+"
    r"(?P<vat_code>[A-Za-z0-9\-]+)\s+"
    r"(?P<net_amount>\d+(?:\.\d+)?)$"
    )

    for line in lines:
        line = line.strip()
        match = pattern.search(line)
        if match:
            items.append(match.groupdict())

    return items




#---------------- INVOICE HEADER EXTRACTION -------
def extract_invoice_fields(text):
    """Extract invoice number, date and total using regex patterns"""
    invoice_no = re.search(r"Invoice No : (\S+)", text, re.IGNORECASE)
    date = re.search(r"Order Date : (\d{1,2}[/-]\d{1,2}[/-]\d{2,4})", text, re.IGNORECASE)
    total = re.search(r"Total: £ (\d+\.\d{2})", text, re.IGNORECASE)

    return {
        "invoice_no": invoice_no.group(1) if invoice_no else None,
        "date": date.group(1) if date else None,
        "total": total.group(1) if total else None
    }

# --------- GUI APPLICATION ----------------
class InvoiceApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Invoice Robot: Made by Nim")
        self.file_list = []

        # Select files button
        self.select_btn = tk.Button(root, text="select pdf files", command=self.select_files)
        self.select_btn.pack(pady=10)

        # Listbox for selected files
        self.listbox = tk.Listbox(root, width=80, height=10)
        self.listbox.pack(padx=10)

        # Process button
        self.process_btn = tk.Button(root, text="process and save", command=self.process_files)
        self.process_btn.pack(pady=10)

    def select_files(self):
        files = filedialog.askopenfilenames(filetypes=[("PDF Files", "*.pdf")])
        if files:
            self.file_list = files
            self.listbox.delete(0, tk.END)
            for f in files:
                self.listbox.insert(tk.END, f)
        else:
            messagebox.showwarning("no files selected", "please select pdf files first!!!!")        


    def process_files(self):
        if not self.file_list:
            messagebox.showwarning("no files selected", "please select pdf files first!!!!")
            return
        
        all_data = []
        for file_path in self.file_list:
            try:
                text = extract_text_from_pdf(file_path)
                supplier = detect_supplier(text)
                header = extract_invoice_fields(text)
                line_items = extract_line_items(text)


                for item in line_items:
                    item.update({
                        "Supplier": supplier,
                        "Invoice No": header.get("invoice_no"),
                        "Date": header.get("date"),
                        "Total": header.get("total"),
                        "Filename": os.path.basename(file_path)
                    })
                    all_data.append(item)
                    
            except Exception as e:
                messagebox.showerror("error", f"failed to process {file_path}.\nError: {e}")
        if not all_data:
            messagebox.showwarning("no data!", "no line items were extracted from the selected files.")

        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv")],
            title="save as ...")
        if not save_path:
            return

        df = pd.DataFrame(all_data)

        try:
            if save_path.endswith(".csv"):
                df.to_csv(save_path, index=False)
            else:
                df.to_excel(save_path, index=False)
            messagebox.showinfo("congratulations", f"invoice data saved to {save_path}")
        except Exception as e:
            messagebox.showwarning("save failed!", str(e))


if __name__ == "__main__":
    root = tk.Tk()
    app = InvoiceApp(root)
    root.mainloop()




