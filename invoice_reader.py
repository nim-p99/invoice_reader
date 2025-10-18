import pdfplumber
import re
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
from PIL import Image, ImageTk 
import tkinter.font as tkfont


# PDF TEXT EXTRACTION
def extract_text_from_pdf(pdf_path):
    """extracts the text from the whole pdf and returns as a string"""
    text = ""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text() or ""
            text += page_text + "\n"
    return text



# DETECT SUPPLIER
def detect_supplier(text):
    """searches text for supplier name - returns supplier"""
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

# EXTRACTS LINE ITEMS
def extract_line_items(text, supplier):
    """directs to supplier-specific line item extraction function"""
    if supplier == "AAH":
        return extract_aah_line_items(text)
    elif supplier == "Colorama":
        return extract_colorama_line_items(text)
    elif supplier == "Alliance":
        return extract_alliance_line_items(text)
    elif supplier == "Lexon":
        return extract_lexon_line_items(text)
    else:
        return []   



# EXTRACT INVOICE HEADER
def extract_invoice_fields(text, supplier):
    """directs to supplier-specific invoice header extraction function"""
    if supplier == "AAH":
        return extract_aah_fields(text)
    elif supplier == "Colorama":
        return extract_colorama_fields(text)
    elif supplier == "Alliance":
        return extract_alliance_fields(text)
    elif supplier == "Lexon":
        return extract_lexon_fields(text)
    else:
        return {"invoice_no": None, "date": None, "total": None}

# COLORAMA EXTRACTION
def extract_colorama_fields(text):
    """Extract invoice number, date and total using regex patterns"""
    invoice_no = re.search(r"Invoice No : (\S+)", text, re.IGNORECASE)
    date = re.search(r"Order Date : (\d{1,2}[/-]\d{1,2}[/-]\d{2,4})", text, re.IGNORECASE)
    total = re.search(r"Total: £ (\d+\.\d{2})", text, re.IGNORECASE)

    return {
        "invoice_no": invoice_no.group(1) if invoice_no else None,
        "date": date.group(1) if date else None,
        "total": total.group(1) if total else None
    }

def extract_colorama_line_items(text):
    """"Extract line items from colorama invoice"""
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

# AAH EXTRACTION
def extract_aah_fields(text):
    invoice_no = re.search(r"Invoice\s*Ref[:\s]*([A-Z0-9]+)(?=\s|$)", text, re.IGNORECASE)
    date = re.search(r"Invoice\s*Date[:\s]*(\d{1,2}[/-]\d{1,2}[/-]\d{2,4})", text)
    total = re.search(r"Total\s*amt\s*due\s*\(GBP\)[:\s]*([\d,]+\.\d{2})", text, re.IGNORECASE)

    return {
        "invoice_no": invoice_no.group(1) if invoice_no else None,
        "date": date.group(1) if date else None,
        "total": total.group(1) if total else None
    }

def extract_aah_line_items(text):
    pattern = re.compile(
        r"(?P<product_code>[A-Z]{3}\d{4,5}[A-Z]?)\s*"
        r"(?P<pack_size>\d{1,3}[A-Z]?)\s+"               
        r"(?P<description>[A-Za-z0-9%/\[\]\-\s]+?)\s+"    
        r"(?P<qty>\d+)\s+"                                
        r"(?P<unit_price>\d+\.\d{2})\s+"                  
        r"(?P<net_price>\d+\.\d{2})\s+"                  
        r"(?P<vat>\d+%)\s+"                               
        r"(?P<total>\d+\.\d{2})",                         
        re.MULTILINE
    )

    items = [m.groupdict() for m in pattern.finditer(text)]
    return items

# ALLIANCE EXTRACTION

def extract_alliance_fields(text):
    invoice_no = re.search(r"\bE\d[A-Z]\d{5,6}\b", text)
    date = re.search(r"(\d{1,2}[A-Z]{3}\d{2}|\d{1,2}[/-]\d{1,2}[/-]\d{2,4})", text)
    total = re.search(r"INVOICE\s+TOTAL\s+([\d,]+\.\d{2})", text, re.IGNORECASE)

    return {
        "invoice_no": invoice_no.group(0) if invoice_no else None,
        "date": date.group(1) if date else None,
        "total": total.group(1) if total else None
    }

def extract_alliance_line_items(text):
    """
    Extracts line items from Alliance invoices with structure like:
    1  1 PRODUCT NAME ... 7.11 Z 7.11 20.00 1.42
       0362-475   CDSCH3
    """
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    items = []

    # Regex for the main item line
    pattern_main = re.compile(
        r'^\s*(\d+)\s+\d+\s+([A-Z0-9%/\-\s()]+?)\s+1\s+([\d.]+)\s+([A-Z])\s+([\d.]+)\s+[\d.]+\s+([\d.]+)',
        re.MULTILINE
    )

    # Regex for the product code on the next line
    pattern_code = re.compile(r'^\s*(\d{4,5}-\d{3})')

    i = 0
    while i < len(lines):
        match = pattern_main.match(lines[i])
        if match:
            qty = match.group(1)
            description = match.group(2).strip()
            unit_price = match.group(3)
            vat_code = match.group(4)
            net_amount = match.group(5)
            vat_amount = match.group(6)

            # Look ahead for product code
            product_code = None
            if i + 1 < len(lines):
                code_match = pattern_code.match(lines[i + 1])
                if code_match:
                    product_code = code_match.group(1)
                    i += 1  # skip the product code line

            items.append({
                "qty": qty,
                "description": description,
                "unit_price": unit_price,
                "vat_code": vat_code,
                "net_amount": net_amount,
                "vat_amount": vat_amount,
                "product_code": product_code
            })

        i += 1

    return items


# GUI APPLICATION 
class InvoiceApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Invoice Wizard: Made by Nim")
        self.root.geometry("800x600")
        self.root.configure(bg="#f3f4f6")
        
        # Fonts
        self.title_font = tkfont.Font(family="Segoe UI", size=20, weight="bold")
        self.button_font = tkfont.Font(family="Segoe UI", size = 11)
        self.list_font = tkfont.Font(family="Consolas", size=10)


        # Frame layout
        top_frame = tk.Frame(root, bg="#ffffff", highlightthickness=0)
        top_frame.place(relx=0.5, rely=0.1, anchor="center")

        mid_frame = tk.Frame(root, bg="#ffffff", bd=0)
        mid_frame.place(relx=0.5, rely=0.45, anchor="center")

        bottom_frame = tk.Frame(root, bg="#ffffff", bd=0)
        bottom_frame.place(relx=0.5, rely=0.85, anchor="center")

        btn_frame = tk.Frame(root, bg="#f3f4f6")
        btn_frame.pack(pady=10)

        # Title
        self.title_label = tk.Label(
            root,
            text="Invoice Wizard (made by nim)",
            font=self.title_font,
            bg="#f3f4f6",
            fg="#2b2b2b",
        )
        self.title_label.pack(pady=20)


        # Select files button
        self.select_btn = tk.Button(
            btn_frame,
            text="select pdf files",
            font=self.button_font,
            bg="#4a90e2",
            fg="white",
            relief="flat",
            padx=15,
            pady=5,
            command=self.select_files
        )
        self.select_btn.grid(row=0, column=0, padx=5)

        # Listbox for selected files
        self.listbox = tk.Listbox(
            root,
            width=90,
            height=15,
            font=self.list_font,
            bg="white",
            fg="#333",
            selectbackground="#4a90e2",
            selectforeground="white",
            relief="flat",
        )
        self.listbox.pack(padx=20, pady=10)

        # Remove button
        self.remove_btn = tk.Button(
            btn_frame,
            text="remove selected",
            font=self.button_font,
            bg="#e63946",
            fg="white",
            relief="flat",
            padx=15,
            pady=5,
            command=self.remove_selected,
        )
        self.remove_btn.grid(row=0, column=1, padx=5)

        # File count label
        self.count_label = tk.Label(mid_frame, text="no files selected", fg="gray")
        self.count_label.pack(pady=5)

        # Progress bar
        self.progress = ttk.Progressbar(mid_frame, length=400, mode="determinate")
        self.progress.pack(pady=5)

        # Process button
        self.process_btn = tk.Button(
            btn_frame,
            text="process and save",
            font=self.button_font,
            bg="#50c878",
            fg="white",
            relief="flat",
            padx=15,
            pady=5,
            command=self.process_files,
        )
        self.process_btn.grid(row=0, column=2, padx=5)

        self.file_list = []

    def set_background(self, image_path):
        # load and resize image
        img = Image.open(image_path)
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        img = img.resize((screen_width, screen_height), Image.LANCZOS)
        self.bg_image = ImageTk.PhotoImage(img)

        # create and place the label
        bg_label = tk.Label(self.root, image=self.bg_image)
        bg_label.place(x=0, y=0, relwidth=1, relheight=1)

    
    def select_files(self):
        files = filedialog.askopenfilenames(filetypes=[("PDF Files", "*.pdf")])
        if files:
            for f in files:
                if f not in self.file_list:
                    self.file_list.append(f)
                    self.listbox.insert(tk.END, f)
            self.update_count_label()
        else:
            messagebox.showwarning("no files selected", "please select pdf files first!!!!")        

    def remove_selected(self):
        selected_index = self.listbox.curselection()
        if not selected_index:
            messagebox.showwarning("no selection", "please select a file to remove!")
            return
        index = selected_index[0]
        removed_file = self.listbox.get(index)
        self.listbox.delete(index)
        self.file_list.remove(removed_file)
        self.update_count_label()

    def update_count_label(self):
        count = len(self.file_list)
        if count == 0:
            self.count_label.config(text="no files selected", fg="gray")
        elif count == 1:
            self.count_label.config(text="1 file selected", fg="green")
        else:
            self.count_label.config(text=f"{count} files selected", fg="green")



    def process_files(self):
        if not self.file_list:
            messagebox.showwarning("no files selected", "please select pdf files first!!!!")
            return
        
        all_data = []
        processed_invoices = 0
        total_files = len(self.file_list)

        self.progress["value"] = 0
        self.progress["maximum"] = total_files
        self.root.update_idletasks()


        for i, file_path in enumerate(self.file_list):  
            try:
                text = extract_text_from_pdf(file_path)
                supplier = detect_supplier(text)
                header = extract_invoice_fields(text, supplier)
                line_items = extract_line_items(text, supplier)

                items = []
                for item in line_items:
                    item.update({
                        "Supplier": supplier,
                        "Invoice No": header.get("invoice_no"),
                        "Date": header.get("date"),
                        "Total": header.get("total"),
                        "Filename": os.path.basename(file_path)
                    })
                    items.append(item)

                if items:
                    processed_invoices += 1
                    all_data.extend(items)
                            
            except Exception as e:
                messagebox.showerror("error", f"failed to process {file_path}.\nError: {e}")

            # update progress bar
            self.progress["value"] = i + 1
            self.root.update_idletasks()


        if not all_data:
            messagebox.showwarning("no data!", "no line items were extracted from the selected files.")
            return

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
            messagebox.showinfo(
                "congratulations",
                "thank you for using nim's invoice wizard.\n\n"
                f"successfully processed {processed_invoices} invoices.\n"
                f"extracted {len(df)} total line items.\n\ninvoice data saved to:\n{save_path}")
        except Exception as e:
            messagebox.showwarning("save failed!", str(e))


if __name__ == "__main__":
    root = tk.Tk()
    app = InvoiceApp(root)
    root.mainloop()




