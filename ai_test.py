import json
import os
import time
import base64
from pathlib import Path
from typing import List, Optional
import tkinter as tk
from tkinter import messagebox, scrolledtext, ttk
import pandas as pd
from PIL import Image, ImageTk

from pydantic import BaseModel
import ollama
from ollama import chat
from ollama._types import ResponseError


# =========================
# 1. JSON Schema
# =========================

class InvoiceItem(BaseModel):
    name: Optional[str]
    quantity: Optional[float]
    unit_price: Optional[float]
    total_price: Optional[float]


class InvoiceData(BaseModel):
    invoice_number: Optional[str]
    phone: Optional[str]
    email: Optional[str]
    address: Optional[str]
    date: Optional[str]
    items: List[InvoiceItem]
    subtotal: Optional[float]
    tax: Optional[float]
    total_price: Optional[float]
    currency: Optional[str]


# =========================
# 2. Image Processing
# =========================

def encode_image_to_base64(image_path: str) -> str:
    """Encode image file to base64 string for sending to Ollama."""
    with open(image_path, "rb") as image_file:
        return base64.b64encode(image_file.read()).decode("utf-8")


# =========================
# 3. Data Editing GUI
# =========================

def edit_invoice_data_gui(invoice_data: InvoiceData, image_path: str) -> InvoiceData:
    """Show a GUI for verifying and editing invoice data with image preview."""
    root = tk.Tk()
    root.title("Verify and Edit Invoice Data")
    root.geometry("1200x800")

    # Main frames
    left_frame = tk.Frame(root, width=400, height=800)
    left_frame.pack(side="left", fill="y", padx=10, pady=10)
    left_frame.pack_propagate(False)

    right_frame = tk.Frame(root)
    right_frame.pack(side="right", fill="both", expand=True, padx=10, pady=10)

    # Image preview
    try:
        img = Image.open(image_path)
        img.thumbnail((380, 600))  # Resize to fit
        img_tk = ImageTk.PhotoImage(img)
        img_label = tk.Label(left_frame, image=img_tk)
        img_label.image = img_tk  # Keep reference
        img_label.pack()
    except Exception as e:
        tk.Label(left_frame, text=f"Image preview failed:\n{str(e)}", wraplength=380).pack()

    # Scrollable frame for data
    canvas = tk.Canvas(right_frame)
    scrollbar = tk.Scrollbar(right_frame, orient="vertical", command=canvas.yview)
    scrollable_frame = tk.Frame(canvas)

    scrollable_frame.bind(
        "<Configure>",
        lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
    )

    canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)

    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    # Fields
    fields = {}
    row = 0

    def add_field(label_text, initial_value, row_idx):
        tk.Label(scrollable_frame, text=label_text).grid(row=row_idx, column=0, sticky="w", padx=10, pady=5)
        entry = tk.Entry(scrollable_frame, width=50)
        entry.insert(0, str(initial_value or ""))
        entry.grid(row=row_idx, column=1, padx=10, pady=5)
        fields[label_text] = entry

    add_field("Invoice Number", invoice_data.invoice_number, row); row += 1
    add_field("Phone", invoice_data.phone, row); row += 1
    add_field("Email", invoice_data.email, row); row += 1
    add_field("Address", invoice_data.address, row); row += 1
    add_field("Date", invoice_data.date, row); row += 1
    add_field("Subtotal", invoice_data.subtotal, row); row += 1
    add_field("Tax", invoice_data.tax, row); row += 1
    add_field("Total Price", invoice_data.total_price, row); row += 1
    add_field("Currency", invoice_data.currency, row); row += 1

    # Items section
    tk.Label(scrollable_frame, text="Items").grid(row=row, column=0, columnspan=2, sticky="w", padx=10, pady=5)
    row += 1

    # Treeview for items
    columns = ("Name", "Quantity", "Unit Price", "Total Price", "Edit", "Delete")
    tree = ttk.Treeview(scrollable_frame, columns=columns, show="headings", height=10)
    for i, col in enumerate(columns):
        tree.heading(col, text=col)
        if col in ["Edit", "Delete"]:
            tree.column(col, width=60, anchor="center")
        else:
            tree.column(col, width=120)

    # Populate tree
    for item in invoice_data.items:
        tree.insert("", "end", values=(item.name or "", item.quantity or "", item.unit_price or "", item.total_price or "", "Edit", "Delete"))

    tree.grid(row=row, column=0, columnspan=2, padx=10, pady=5)
    row += 1

    def on_tree_click(event):
        item_id = tree.identify_row(event.y)
        column = tree.identify_column(event.x)
        if not item_id or column not in ["#5", "#6"]:  # Edit is #5, Delete is #6
            return
        if column == "#5":  # Edit
            edit_item_dialog(item_id)
        elif column == "#6":  # Delete
            if messagebox.askyesno("Delete Item", "Are you sure you want to delete this item?"):
                tree.delete(item_id)

    tree.bind("<Button-1>", on_tree_click)

    def edit_item_dialog(item_id):
        values = tree.item(item_id, "values")
        dialog = tk.Toplevel(root)
        dialog.title("Edit Item")
        dialog.geometry("300x200")

        fields = {}
        labels = ["Name", "Quantity", "Unit Price", "Total Price"]
        for i, label in enumerate(labels):
            tk.Label(dialog, text=label).grid(row=i, column=0, padx=10, pady=5)
            entry = tk.Entry(dialog)
            entry.insert(0, values[i] or "")
            entry.grid(row=i, column=1, padx=10, pady=5)
            fields[label] = entry

        def save():
            new_values = [fields[label].get() for label in labels] + ["Edit", "Delete"]
            tree.item(item_id, values=new_values)
            dialog.destroy()

        tk.Button(dialog, text="Save", command=save).grid(row=len(labels), column=0, pady=10)
        tk.Button(dialog, text="Cancel", command=dialog.destroy).grid(row=len(labels), column=1, pady=10)

    # Add Item button below
    def add_item_dialog():
        dialog = tk.Toplevel(root)
        dialog.title("Add Item")
        dialog.geometry("300x200")

        fields = {}
        labels = ["Name", "Quantity", "Unit Price", "Total Price"]
        for i, label in enumerate(labels):
            tk.Label(dialog, text=label).grid(row=i, column=0, padx=10, pady=5)
            entry = tk.Entry(dialog)
            entry.grid(row=i, column=1, padx=10, pady=5)
            fields[label] = entry

        def save():
            new_values = [fields[label].get() for label in labels] + ["Edit", "Delete"]
            tree.insert("", "end", values=new_values)
            dialog.destroy()

        tk.Button(dialog, text="Add", command=save).grid(row=len(labels), column=0, pady=10)
        tk.Button(dialog, text="Cancel", command=dialog.destroy).grid(row=len(labels), column=1, pady=10)

    tk.Button(scrollable_frame, text="Add Item", command=add_item_dialog).grid(row=row, column=0, columnspan=2, pady=10)
    row += 1

    def save_and_close():
        try:
            # Collect main data
            updated_data = {
                "invoice_number": fields["Invoice Number"].get() or None,
                "phone": fields["Phone"].get() or None,
                "email": fields["Email"].get() or None,
                "address": fields["Address"].get() or None,
                "date": fields["Date"].get() or None,
                "subtotal": float(fields["Subtotal"].get()) if fields["Subtotal"].get() else None,
                "tax": float(fields["Tax"].get()) if fields["Tax"].get() else None,
                "total_price": float(fields["Total Price"].get()) if fields["Total Price"].get() else None,
                "currency": fields["Currency"].get() or None,
                "items": []
            }

            # Collect items from tree
            for child in tree.get_children():
                values = tree.item(child, "values")[:4]  # Skip Edit and Delete columns
                item_dict = {
                    "name": values[0] or None,
                    "quantity": float(values[1]) if values[1] else None,
                    "unit_price": float(values[2]) if values[2] else None,
                    "total_price": float(values[3]) if values[3] else None,
                }
                updated_data["items"].append(item_dict)

            # Validate with Pydantic
            validated_data = InvoiceData(**updated_data)
            root.edited_data = validated_data
            root.destroy()
        except Exception as e:
            messagebox.showerror("Error", f"Invalid data: {str(e)}")

    def cancel():
        root.edited_data = invoice_data
        root.destroy()

    tk.Button(scrollable_frame, text="Save", command=save_and_close).grid(row=row, column=0, pady=10)
    tk.Button(scrollable_frame, text="Cancel", command=cancel).grid(row=row, column=1, pady=10)

    root.mainloop()
    return getattr(root, 'edited_data', invoice_data)


# =========================
# 4. Save to Excel
# =========================

def save_to_excel(invoice_data: InvoiceData, filename: str = "invoice_output.xlsx"):
    """Save invoice data to Excel file, preserving existing rows."""
    data = invoice_data.model_dump()

    # Flatten items
    items_df = pd.DataFrame(data['items'])
    main_data = pd.DataFrame([{k: v for k, v in data.items() if k != 'items'}])

    if Path(filename).exists():
        try:
            existing_main = pd.read_excel(filename, sheet_name='Invoice', engine='openpyxl')
        except Exception:
            existing_main = pd.DataFrame()
        try:
            existing_items = pd.read_excel(filename, sheet_name='Items', engine='openpyxl')
        except Exception:
            existing_items = pd.DataFrame()

        combined_main = pd.concat([existing_main, main_data], ignore_index=True, sort=False) if not existing_main.empty else main_data
        combined_items = pd.concat([existing_items, items_df], ignore_index=True, sort=False) if not existing_items.empty else items_df

        with pd.ExcelWriter(filename, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            combined_main.to_excel(writer, sheet_name='Invoice', index=False)
            combined_items.to_excel(writer, sheet_name='Items', index=False)
    else:
        with pd.ExcelWriter(filename, engine='openpyxl') as writer:
            main_data.to_excel(writer, sheet_name='Invoice', index=False)
            items_df.to_excel(writer, sheet_name='Items', index=False)

    print(f"Data appended to {filename}")


# =========================
# 5. Extract Invoice Data
# =========================

def get_ollama_model() -> str:
    env_model = os.environ.get("OLLAMA_MODEL")
    if env_model:
        return env_model

    try:
        client = ollama.Client()
        response = client.list()
        if response.models:
            return response.models[0].model
    except Exception as exc:
        print("Warning: unable to auto-detect Ollama model:", exc)

    return "gemma4"


def extract_invoice_data(file_path: str) -> InvoiceData:
    """Send image directly to Ollama model and extract invoice data."""
    image_path = Path(file_path)
    if not image_path.exists():
        raise FileNotFoundError(f"Image file not found: {file_path}")

    # Encode image to base64
    image_base64 = encode_image_to_base64(str(image_path))

    prompt = """
Analyze this invoice image and extract the following data in JSON format:

Rules:
1. The invoice may be in English, Chinese, or other languages. Understand the content in its original language, but extract the data as-is without translation.
2. Do not guess missing values. If a value is missing, use null.
3. Extract the following information:
   - Invoice number (invoice_number)
   - Phone (phone)
   - Email (email)
   - Address (address)
   - Date (date)
4. Extract all purchased items with the following fields:
   - Name (name)
   - Quantity (quantity)
   - Unit price (unit_price)
   - Total price (total_price)
5. Extract summary information:
   - Subtotal (subtotal)
   - Tax (tax)
   - Total price (total_price)
   - Currency (currency, if available)

Return ONLY a JSON object, do not add any other text.
"""

    model_name = get_ollama_model()
    print(f"Using Ollama model: {model_name}")
    print(f"Processing image: {image_path.name}")

    try:
        response = chat(
            model=model_name,
            messages=[
                {
                    "role": "user",
                    "content": prompt,
                    "images": [image_base64]
                }
            ],
            format=InvoiceData.model_json_schema(),
            options={
                "temperature": 0
            }
        )
    except ResponseError as exc:
        raise RuntimeError(
            f"Ollama request failed for model '{model_name}'. "
            "Set OLLAMA_MODEL to a valid local model name or install the model in Ollama."
        ) from exc

    return InvoiceData.model_validate_json(response.message.content)


# =========================
# 5. Main Flow
# =========================

def process_invoice(file_path: str):
    file_path = Path(file_path)

    if not file_path.exists():
        raise FileNotFoundError(f"File not found: {file_path}")

    print(f"Extracting invoice data from {file_path.name} using Ollama vision...")
    invoice_data = extract_invoice_data(str(file_path))

    # Show GUI for verification and editing
    print("Opening verification window...")
    edited_data = edit_invoice_data_gui(invoice_data, str(file_path))

    # Save to JSON
    output = edited_data.model_dump()
    with open("invoice_output.json", "w", encoding="utf-8") as f:
        json.dump(output, f, indent=2, ensure_ascii=False)

    # Save to Excel
    save_to_excel(edited_data)

    print("\nExtracted Data:")
    print(json.dumps(output, indent=2, ensure_ascii=False))

    print("\nSaved to invoice_output.json and invoice_output.xlsx")


# =========================
# 6. Run
# =========================

if __name__ == "__main__":
    start_time = time.perf_counter()

    process_invoice("WhatsApp Image 2026-05-08 at 16.32.38.jpeg")

    elapsed_time = time.perf_counter() - start_time
    print(f"Total runtime: {elapsed_time:.2f} seconds")
    # or:
    # process_invoice("invoice.pdf")