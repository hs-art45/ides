import json
import os
import time
import base64
import re
import cv2
import shutil
import pandas as pd
from pathlib import Path
from typing import List, Optional, Dict, Any, Literal
import tkinter as tk
from tkinter import messagebox, ttk, scrolledtext, filedialog
from PIL import Image, ImageTk

from pydantic import BaseModel, Field
import ollama
from ollama import chat
from ollama._types import ResponseError

try:
    from pyzbar import pyzbar
    PYZBAR_AVAILABLE = True
except ImportError:
    PYZBAR_AVAILABLE = False

# =========================
# 1. JSON Schemas
# =========================

class InvoiceItem(BaseModel):
    invoice_number: Optional[str] = Field(None, description="The parent invoice number this item belongs to")
    item_number: Optional[str] = None
    description: Optional[str] = None 
    quantity: Optional[float] = None
    unit_price: Optional[float] = None
    total_price: Optional[float] = None

class InvoiceData(BaseModel):
    invoice_number: Optional[str] = None
    sender_name: Optional[str] = None
    sender_phone_number: Optional[str] = None
    sender_email: Optional[str] = None
    sender_address: Optional[str] = None
    date: Optional[str] = None
    duedate: Optional[str] = None
    description: Optional[str] = None
    items: List[InvoiceItem] = Field(default_factory=list)
    subtotal: Optional[float] = None
    tax: Optional[float] = None
    total_price: Optional[float] = None

class DeliveryOrderItem(BaseModel):
    do_number: Optional[str] = Field(None, description="The parent DO number this item belongs to")
    item_number: Optional[str] = None
    description: Optional[str] = None
    quantity: Optional[float] = None
    uom: Optional[str] = Field(None, description="Unit of Measure, e.g., 'kg', 'pcs'")

class DeliveryOrderData(BaseModel):
    do_number: Optional[str] = None
    po_reference: Optional[str] = Field(None, description="Purchase Order number")
    delivery_date: Optional[str] = None
    recipient_name: Optional[str] = None
    shipping_address: Optional[str] = None
    description: Optional[str] = None
    received_by_signature: Optional[bool] = Field(None, description="True if a signature is detected on the document")
    items: List[DeliveryOrderItem] = Field(default_factory=list)

class DocumentData(BaseModel):
    document_type: Literal["Invoice", "Delivery Order"]
    invoice: Optional[InvoiceData] = None
    delivery_order: Optional[DeliveryOrderData] = None


# =========================
# 2. Logic Functions
# =========================

def encode_image_to_base64(image_path: str) -> str:
    with open(image_path, "rb") as image_file:
        return base64.b64encode(image_file.read()).decode("utf-8")

def save_to_excel(document_data: DocumentData, filename: Optional[str] = None, source_path: Optional[str] = None, json_path: Optional[str] = None, replace_source: Optional[str] = None, old_number: Optional[str] = None):
    if document_data.document_type == "Invoice":
        invoice = document_data.invoice or InvoiceData()
        filename = filename or "invoice_output.xlsx"
        sheet_name = "Invoice"
        if invoice.invoice_number:
            for item in invoice.items:
                item.invoice_number = invoice.invoice_number
        main_data = pd.DataFrame([invoice.model_dump(exclude={"items"})])
        items_df = pd.DataFrame([item.model_dump() for item in invoice.items]) if invoice.items else pd.DataFrame()
    else:
        delivery_order = document_data.delivery_order or DeliveryOrderData()
        filename = filename or "delivery_order_output.xlsx"
        sheet_name = "Delivery Order"
        if delivery_order.do_number:
            for item in delivery_order.items:
                item.do_number = delivery_order.do_number
        main_data = pd.DataFrame([delivery_order.model_dump(exclude={"items"})])
        items_df = pd.DataFrame([item.model_dump() for item in delivery_order.items]) if delivery_order.items else pd.DataFrame()
    main_data["source_file"] = source_path if source_path else None
    main_data["json_file"] = json_path if json_path else None
    if Path(filename).exists():
        try:
            existing_main = pd.read_excel(filename, sheet_name=sheet_name, engine='openpyxl')
        except Exception:
            existing_main = pd.DataFrame()
        try:
            existing_items = pd.read_excel(filename, sheet_name='Items', engine='openpyxl')
        except Exception:
            existing_items = pd.DataFrame()

        if replace_source and not existing_main.empty and 'source_file' in existing_main.columns:
            existing_main = existing_main[existing_main['source_file'] != replace_source]
        if replace_source and not existing_items.empty and old_number:
            if document_data.document_type == 'Invoice' and 'invoice_number' in existing_items.columns:
                existing_items = existing_items[existing_items['invoice_number'] != old_number]
            if document_data.document_type == 'Delivery Order' and 'do_number' in existing_items.columns:
                existing_items = existing_items[existing_items['do_number'] != old_number]

        combined_main = pd.concat([existing_main, main_data], ignore_index=True, sort=False) if not existing_main.empty else main_data
        combined_items = pd.concat([existing_items, items_df], ignore_index=True, sort=False) if not existing_items.empty else items_df

        with pd.ExcelWriter(filename, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            combined_main.to_excel(writer, sheet_name=sheet_name, index=False)
            combined_items.to_excel(writer, sheet_name='Items', index=False)
    else:
        with pd.ExcelWriter(filename, engine='openpyxl') as writer:
            main_data.to_excel(writer, sheet_name=sheet_name, index=False)
            items_df.to_excel(writer, sheet_name='Items', index=False)

def save_raw_capture(source_path: Optional[str] = None, image_array: Optional[Any] = None) -> Optional[str]:
    """Save a raw image (uploaded or captured) into the rawcapture folder and return full path."""
    try:
        base_dir = os.path.dirname(os.path.abspath(__file__))
        raw_dir = os.path.join(base_dir, "rawcapture")
        os.makedirs(raw_dir, exist_ok=True)

        timestamp = time.strftime("%Y%m%d_%H%M%S", time.localtime())
        if image_array is not None:
            # image_array expected to be a CV2 BGR frame
            dest_name = f"capture_{timestamp}.jpg"
            dest_path = os.path.join(raw_dir, dest_name)
            cv2.imwrite(dest_path, image_array)
            return dest_path

        if source_path and os.path.exists(source_path):
            ext = os.path.splitext(source_path)[1] or ".jpg"
            dest_name = f"{timestamp}_{os.path.basename(source_path)}"
            dest_path = os.path.join(raw_dir, dest_name)
            shutil.copy2(source_path, dest_path)
            return dest_path
    except Exception as e:
        print(f"Failed to save raw capture: {e}")
    return None

def get_ollama_model() -> str:
    env_model = os.environ.get("OLLAMA_MODEL")
    if env_model: return env_model
    try:
        client = ollama.Client()
        response = client.list()
        if response.models: return response.models[0].model
    except: pass
    return "llama3-vision"

def extract_barcode_qr_data(image_path: str) -> Optional[Dict[str, Any]]:
    if not PYZBAR_AVAILABLE:
        return None
    try:
        image = cv2.imread(image_path)
        if image is None: return None
        decoded_objects = pyzbar.decode(image)
        if not decoded_objects: return None
        
        barcode_data_list = []
        for obj in decoded_objects:
            data = obj.data.decode('utf-8')
            try:
                parsed_data = json.loads(data)
                barcode_data_list.append({"type": obj.type, "data": parsed_data})
            except json.JSONDecodeError:
                barcode_data_list.append({"type": obj.type, "data": data})
        return {"barcodes": barcode_data_list} if barcode_data_list else None
    except Exception as e:
        print(f"Barcode/QR extraction error: {e}")
        return None

def extract_document_data(file_path: str) -> tuple[DocumentData, Optional[Dict[str, Any]]]:
    image_base64 = encode_image_to_base64(file_path)
    model_name = get_ollama_model()
    start_time = time.time()

    rules_text = """Rules:
1. Extract key names and fields accurately.
2. Use null for missing data.
3. Do not translate text.
4. Do not add any extra fields.
5. Extract exact information; do not simplify texts.
6. Extract tables or line items carefully into correct JSON matching the schema.
7. Always include the items array field. If empty, output [].
8. Return dates in YYYY-MM-DD format.
9. Include the payment details in the description field in 'InvoiceData' and 'DeliveryOrderData' if available.
"""

    classifier_prompt = (
        'Classify this document image as either "Invoice" or "Delivery Order". '
        'Return exactly one word: Invoice or Delivery Order. Do not return extra text.'
    )
    try:
        cls_resp = chat(
            model=model_name,
            messages=[{"role": "user", "content": classifier_prompt, "images": [image_base64]}],
            options={"temperature": 0}
        )
        doc_type_raw = (cls_resp.message.content or "").strip().strip('"').strip()
        doc_type = "Invoice" if doc_type_raw.lower().startswith("invoice") else "Delivery Order"
    except Exception:
        doc_type = "Invoice"
    
    end_time1 = time.time()
    print(f"Ollama whole runtime ({doc_type}): {end_time1 - start_time:.2f} seconds")


    try:
        if doc_type == "Invoice":
            invoice_schema_text = """Invoice Schema:
{
  "invoice_number": string | null,
  "sender_name": string | null,
  "sender_phone_number": string | null,
  "sender_email": string | null,
  "sender_address": string | null,
  "date": string | null,
  "duedate": string | null,
  "description": string | null,
  "items": [
    {
      "invoice_number": string | null,
      "item_number": string | null,
      "description": string | null,
      "quantity": number | null,
      "unit_price": number | null,
      "total_price": number | null
    }
  ],
  "subtotal": number | null,
  "tax": number | null,
  "total_price": number | null
}"""
            detail_prompt = f"{rules_text}\nReturn ONLY the JSON object matching the Invoice schema.\n{invoice_schema_text}"
            resp = chat(
                model=model_name,
                messages=[{"role": "user", "content": detail_prompt, "images": [image_base64]}],
                format=InvoiceData.model_json_schema(),
                options={"temperature": 0}
            )
            # save_raw_ollama_response(file_path, resp)
            invoice = InvoiceData.model_validate_json(resp.message.content)
            if invoice and invoice.invoice_number:
                for item in invoice.items: item.invoice_number = invoice.invoice_number
            document_data = DocumentData(document_type="Invoice", invoice=invoice, delivery_order=None)
            
        else:
            delivery_schema_text = """Delivery Order Schema:
{
  "do_number": string | null,
  "po_reference": string | null,
  "delivery_date": string | null,
  "recipient_name": string | null,
  "shipping_address": string | null,
  "description": string | null,
  "received_by_signature": boolean | null,
  "items": [
    {
      "do_number": string | null,
      "item_number": string | null,
      "description": string | null,
      "quantity": number | null,
      "uom": string | null
    }
  ]
}"""
            detail_prompt = f"{rules_text}\nReturn ONLY the JSON object matching the Delivery Order schema.\n{delivery_schema_text}"
            resp = chat(
                model=model_name,
                messages=[{"role": "user", "content": detail_prompt, "images": [image_base64]}],
                format=DeliveryOrderData.model_json_schema(),
                options={"temperature": 0}
            )
            # save_raw_ollama_response(file_path, resp)
            delivery = DeliveryOrderData.model_validate_json(resp.message.content)
            if delivery and delivery.do_number:
                for item in delivery.items: item.do_number = delivery.do_number
            document_data = DocumentData(document_type="Delivery Order", invoice=None, delivery_order=delivery)
            
    except Exception as e:
        raise

    end_time = time.time()
    print(f"Ollama whole runtime ({doc_type}): {end_time - start_time:.2f} seconds")

    barcode_data = extract_barcode_qr_data(file_path)
    if barcode_data:
        barcode_text = json.dumps(barcode_data, indent=2, ensure_ascii=False)
        if document_data.document_type == "Invoice" and document_data.invoice:
            document_data.invoice.description = barcode_text
        elif document_data.document_type == "Delivery Order" and document_data.delivery_order:
            document_data.delivery_order.description = barcode_text

    return document_data, barcode_data

def compare_invoice_data(document_data: DocumentData, barcode_data: Dict[str, Any]) -> bool:
    if not barcode_data or "barcodes" not in barcode_data: return False

    if document_data.document_type == "Invoice":
        model_items = [item.model_dump() for item in (document_data.invoice.items if document_data.invoice else [])]
        compare_keys = ["item_number", "description", "quantity", "unit_price", "total_price"]
    else:
        model_items = [item.model_dump() for item in (document_data.delivery_order.items if document_data.delivery_order else [])]
        compare_keys = ["item_number", "description", "quantity", "uom"]

    for barcode in barcode_data["barcodes"]:
        data = barcode.get("data")
        if isinstance(data, dict):
            barcode_items = data.get("items", [])
            barcode_items_normalized = [{k: v for k, v in item.items() if k in compare_keys} for item in barcode_items]
            model_items_normalized = [{k: v for k, v in item.items() if k in compare_keys} for item in model_items]
            if barcode_items_normalized == model_items_normalized: return True
    return False

def save_barcode_qr_json(file_path: str, barcode_data: Dict[str, Any]):
    json_file = file_path.rsplit('.', 1)[0] + '_barcode_qr.json'
    with open(json_file, 'w', encoding='utf-8') as f:
        json.dump(barcode_data, f, indent=2, ensure_ascii=False)
    return json_file

def save_document_json(file_path: str, document_data: DocumentData, barcode_data: Optional[Dict[str, Any]] = None):
    base_dir = os.path.dirname(os.path.abspath(__file__))
    json_dir = os.path.join(base_dir, "capturejson")
    os.makedirs(json_dir, exist_ok=True)
    base_name = os.path.splitext(os.path.basename(file_path))[0]
    doc_type = "invoice" if document_data.document_type == "Invoice" else "delivery_order"
    timestamp = time.strftime("%Y%m%d_%H%M%S", time.localtime())
    json_file = os.path.join(json_dir, f"{base_name}_{doc_type}_{timestamp}.json")
    out = document_data.model_dump()
    if barcode_data:
        out["barcode_data"] = barcode_data
    with open(json_file, 'w', encoding='utf-8') as f:
        json.dump(out, f, indent=2, ensure_ascii=False)
    return json_file

def save_raw_ollama_response(file_path: str, resp: Any):
    # function for inspecting ollama response content during development - saves the raw text response into a file for debugging
    if not file_path: return None
    raw_file = file_path.rsplit('.', 1)[0] + '_raw_response.txt'
    try:
        response_text = resp.message.content if hasattr(resp, 'message') else str(resp)
        with open(raw_file, 'w', encoding='utf-8') as f:
            f.write(response_text)
        return raw_file
    except: return None

def append_to_master_json(source_path: str, document_data: DocumentData, barcode_data: Optional[Dict[str, Any]] = None, json_path: Optional[str] = None, master_file: str = "all_extracted_documents.json"):
    try:
        try:
            with open(master_file, 'r', encoding='utf-8') as f: existing = json.load(f)
            if not isinstance(existing, list): existing = []
        except FileNotFoundError: existing = []
        
        record = {
            "timestamp": time.strftime("%Y-%m-%dT%H:%M:%S", time.localtime()),
            "source": source_path,
            "json_file": json_path,
            "document_type": document_data.document_type if document_data else None,
            "data": document_data.model_dump() if document_data else None,
            "barcode_data": barcode_data or None
        }
        existing.append(record)
        with open(master_file, 'w', encoding='utf-8') as f: json.dump(existing, f, indent=2, ensure_ascii=False)
    except Exception as e: print(f"Master file sync alert: {e}")


# =========================
# 3. Integrated GUI Class
# =========================

class DocumentProcessingApp:
    def __init__(self, root):
        self.root = root
        self.root.title("AI Unified Document Processor (Invoices & DO)")
        self.root.geometry("1400x700")
        self.center_window(self.root, 1400, 700)

        self.cap = cv2.VideoCapture(1)
        self.current_frame = None
        # map tree item ids to metadata (source file path, json path, document type, number)
        self.history_map: Dict[str, Dict[str, Any]] = {}
        self.history_entries: List[Dict[str, Any]] = []

        self.setup_ui()
        self.update_camera()
        self.load_history()

    def center_window(self, window, width, height):
        window.update_idletasks()
        screen_width = window.winfo_screenwidth()
        screen_height = window.winfo_screenheight()
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2
        window.geometry(f"{width}x{height}+{x}+{y}")

    def setup_ui(self):
        top_pane = tk.PanedWindow(self.root, orient="horizontal", sashrelief="raised")
        top_pane.pack(fill="both", expand=True, padx=5, pady=5)

        # Left View
        cam_frame = tk.LabelFrame(top_pane, text="Live Processing Hub Feed")
        top_pane.add(cam_frame, width=650)
        
        self.cam_label = tk.Label(cam_frame, bg="black")
        self.cam_label.pack(fill="both", expand=True)

        btn_frame = tk.Frame(cam_frame)
        btn_frame.pack(fill="x")
        tk.Button(btn_frame, text="CAPTURE & PROCESS DOCUMENT", bg="#4CAF50", fg="white", 
                  font=("Arial", 12, "bold"), command=self.process_capture).pack(fill="x", pady=5)
        tk.Button(btn_frame, text="UPLOAD IMAGE / PDF FILE", bg="#2196F3", fg="white", 
                  font=("Arial", 12, "bold"), command=self.upload_and_process).pack(fill="x", pady=5)

        # Right Log
        hist_frame = tk.LabelFrame(top_pane, text="History Records")
        top_pane.add(hist_frame, width=600)

        self.history_filter = tk.StringVar(value="All")
        self.history_search_text = tk.StringVar()
        self.history_sort_column: Optional[str] = None
        self.history_sort_reverse = False

        filter_frame = tk.Frame(hist_frame)
        filter_frame.pack(fill="x", padx=5, pady=5)
        tk.Label(filter_frame, text="Log Filter:").pack(side="left")
        tk.Radiobutton(filter_frame, text="All Logs", variable=self.history_filter, value="All", command=self.load_history).pack(side="left", padx=3)
        tk.Radiobutton(filter_frame, text="Invoices Only", variable=self.history_filter, value="Invoice", command=self.load_history).pack(side="left", padx=3)
        tk.Radiobutton(filter_frame, text="DO Only", variable=self.history_filter, value="Delivery Order", command=self.load_history).pack(side="left", padx=3)

        search_frame = tk.Frame(hist_frame)
        search_frame.pack(fill="x", padx=5, pady=(0, 5))
        tk.Label(search_frame, text="Search:").pack(side="left")
        search_entry = tk.Entry(search_frame, textvariable=self.history_search_text)
        search_entry.pack(side="left", fill="x", expand=True, padx=5)
        search_entry.bind('<KeyRelease>', lambda event: self.refresh_history_tree())
        tk.Button(search_frame, text="Clear", command=self.clear_history_search).pack(side="left")

        self.tree = ttk.Treeview(hist_frame, columns=("Type", "No", "Date", "Name", "Amount"), show="headings")
        self.tree.tag_configure("highlight", background="#fff2a8")
        self.tree.tag_configure("normal", background="")
        for col in ("Type", "No", "Date", "Name", "Amount"):
            if col in ("No", "Date", "Name", "Amount"):
                self.tree.heading(col, text=col, command=lambda c=col: self.sort_history_column(c))
            else:
                self.tree.heading(col, text=col)
            if col == "Type":
                self.tree.column(col, width=30, minwidth=30, anchor="center")
            elif col == "No":
                self.tree.column(col, width=100, minwidth=100, anchor="center")
            elif col == "Amount":
                self.tree.column(col, width=100, minwidth=80, anchor="e")
            elif col == "Name":
                self.tree.column(col, width=230, minwidth=100, anchor="center")
            else:
                self.tree.column(col, width=40, anchor="center")
        self.tree.pack(side="left", fill="both", expand=True)
        # double-click opens history actions (view/edit/delete)
        self.tree.bind("<Double-1>", self.on_history_double_click)

        def show_history_context_menu(event):
            item_id = self.tree.identify_row(event.y)
            if not item_id:
                return
            self.tree.selection_set(item_id)
            meta = self.history_map.get(item_id, {})
            menu = tk.Menu(self.tree, tearoff=0)
            menu.add_command(label="Edit record", command=lambda: self.open_record_for_edit(meta))
            menu.add_command(label="Delete record", command=lambda: self.delete_history_record(meta, item_id=item_id))
            try:
                menu.tk_popup(event.x_root, event.y_root)
            finally:
                menu.grab_release()

        self.tree.bind("<Button-3>", show_history_context_menu)
        
        sb = ttk.Scrollbar(hist_frame, command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
    def update_camera(self):
        ret, frame = self.cap.read()
        if ret:
            self.current_frame = frame
            cv2_rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
            img = Image.fromarray(cv2_rgb)
            img.thumbnail((640, 480))
            imgtk = ImageTk.PhotoImage(img)
            self.cam_label.imgtk = imgtk
            self.cam_label.configure(image=imgtk)
        self.root.after(10, self.update_camera)

    def process_capture(self):
        if self.current_frame is not None:
            # save a temporary capture then copy to rawcapture folder for persistent storage
            temp_path = "last_capture.jpg"
            cv2.imwrite(temp_path, self.current_frame)
            saved_raw = save_raw_capture(source_path=temp_path, image_array=self.current_frame)
            try:
                use_path = saved_raw or temp_path
                document_data, barcode_data = extract_document_data(use_path)
                if barcode_data and compare_invoice_data(document_data, barcode_data):
                    save_barcode_qr_json(use_path, barcode_data)
                append_to_master_json(use_path, document_data, barcode_data)
                # open editor with the saved raw path so edits map to the stored file
                self.open_edit_dialog(document_data, use_path, barcode_data)
            except Exception as e:
                messagebox.showerror("Error", f"Extraction failed: {e}")

    def upload_and_process(self):
        file_path = filedialog.askopenfilename(
            title="Open Document File",
            filetypes=[("Supported Documents", "*.jpg *.jpeg *.png *.bmp *.tiff *.pdf"), ("All files", "*.*")]
        )
        if not file_path: return

        preview_path = file_path
        is_temp = False
        temp_preview_created = False
        if file_path.lower().endswith('.pdf'):
            try:
                from pdf2image import convert_from_path
                images = convert_from_path(file_path)
                if not images: raise Exception("Empty PDF Pages")
                preview_path = "upload_preview.jpg"
                images[0].save(preview_path, 'JPEG')
                is_temp = True
                temp_preview_created = True
            except Exception as e:
                messagebox.showerror("Error", f"PDF converter failure: {e}")
                return

        # Save the image that will be fed to the model into rawcapture for persistent history
        saved_raw = None
        try:
            if is_temp and temp_preview_created:
                saved_raw = save_raw_capture(source_path=preview_path)
            else:
                saved_raw = save_raw_capture(source_path=file_path)
        except Exception:
            saved_raw = None
        if saved_raw:
            model_input_path = saved_raw
        else:
            model_input_path = preview_path

        preview_win = tk.Toplevel(self.root)
        preview_win.title("Preview Window Confirmation")
        self.center_window(preview_win, 800, 760)

        file_name = os.path.basename(file_path)
        file_format = os.path.splitext(file_path)[1].lstrip('.').upper() or 'Unknown'

        info_frame = tk.Frame(preview_win)
        info_frame.pack(fill="x", padx=10, pady=(10, 0))
        tk.Label(info_frame, text=f"File: {file_name}", font=("Arial", 10, "bold"), anchor="w").pack(fill="x")
        tk.Label(info_frame, text=f"Format: {file_format}", font=("Arial", 9), anchor="w").pack(fill="x")

        try:
            img = Image.open(model_input_path)
            img.thumbnail((760, 560))
            imgtk = ImageTk.PhotoImage(img)
            img_label = tk.Label(preview_win, image=imgtk)
            img_label.image = imgtk
            img_label.pack(padx=10, pady=10)
        except Exception as e:
            tk.Label(preview_win, text=f"Preview unavailable: {e}").pack()

        def do_process():
            try:
                # prefer saved raw copy as input for extraction
                input_path = model_input_path
                document_data, barcode_data = extract_document_data(input_path)
                if barcode_data and compare_invoice_data(document_data, barcode_data):
                    save_barcode_qr_json(input_path, barcode_data)
                preview_win.destroy()
                # cleanup any temporary preview created by pdf conversion (not the saved raw copy)
                if temp_preview_created and os.path.exists(preview_path):
                    try: os.remove(preview_path)
                    except: pass
                # open editor with saved raw path (or original if saving failed)
                self.open_edit_dialog(document_data, input_path, barcode_data)
            except Exception as e:
                preview_win.destroy()
                if temp_preview_created and os.path.exists(preview_path):
                    try: os.remove(preview_path)
                    except: pass
                messagebox.showerror("Processing Failed", str(e))

        def change_document():
            preview_win.destroy()
            if is_temp and os.path.exists(preview_path):
                os.remove(preview_path)
            self.upload_and_process()

        btns = tk.Frame(preview_win)
        btns.pack(pady=10)
        tk.Button(btns, text="Scan & Classify Document", bg="#4CAF50", fg="white", command=do_process).pack(side="left", padx=10)
        tk.Button(btns, text="Change Document", command=change_document).pack(side="left", padx=10)
        tk.Button(btns, text="Cancel", command=preview_win.destroy).pack(side="left", padx=10)

    def load_history(self):
        # refresh tree and metadata map
        for i in self.tree.get_children(): self.tree.delete(i)
        self.history_map.clear()
        current_filter = self.history_filter.get()

        self.history_entries.clear()

        if current_filter in ("All", "Invoice") and Path("invoice_output.xlsx").exists():
            try:
                df = pd.read_excel("invoice_output.xlsx", sheet_name="Invoice", engine='openpyxl')
                for _, row in df.tail(50).iterrows():
                    self.history_entries.append({
                        'values': ("INV", row.get('invoice_number',''), row.get('date',''), row.get('sender_name',''), row.get('total_price','')),
                        'meta': {
                            'source': row.get('source_file') if 'source_file' in row.index else None,
                            'json': row.get('json_file') if 'json_file' in row.index else None,
                            'doc_type': 'Invoice',
                            'number': row.get('invoice_number',''),
                            'extra': self._extract_metadata_values(row)
                        }
                    })
            except Exception:
                pass

        if current_filter in ("All", "Delivery Order") and Path("delivery_order_output.xlsx").exists():
            try:
                df = pd.read_excel("delivery_order_output.xlsx", sheet_name="Delivery Order", engine='openpyxl')
                for _, row in df.tail(50).iterrows():
                    self.history_entries.append({
                        'values': ("DO", row.get('do_number',''), row.get('delivery_date',''), row.get('recipient_name',''), "-"),
                        'meta': {
                            'source': row.get('source_file') if 'source_file' in row.index else None,
                            'json': row.get('json_file') if 'json_file' in row.index else None,
                            'doc_type': 'Delivery Order',
                            'number': row.get('do_number',''),
                            'extra': self._extract_metadata_values(row)
                        }
                    })
            except Exception:
                pass

        self.refresh_history_tree()

    def _extract_metadata_values(self, row):
        values = []
        if hasattr(row, 'to_dict'):
            row = row.to_dict()
        if isinstance(row, dict):
            for v in row.values():
                if isinstance(v, str):
                    values.append(v)
                elif v is not None:
                    values.append(str(v))
        else:
            try:
                for v in row:
                    if isinstance(v, str):
                        values.append(v)
                    elif v is not None:
                        values.append(str(v))
            except Exception:
                pass
        return " ".join(values)

    def clear_history_search(self):
        self.history_search_text.set("")
        self.history_filter.set("All")
        self.history_sort_column = None
        self.history_sort_reverse = False
        self.load_history()

    def refresh_history_tree(self):
        query_text = self.history_search_text.get().strip().lower()
        tokens = re.findall(r"\w+", query_text)
        self._update_history_headers()

        for i in self.tree.get_children():
            self.tree.delete(i)
        self.history_map.clear()

        sorted_entries = list(self.history_entries)
        if self.history_sort_column:
            index_map = {"Type": 0, "No": 1, "Date": 2, "Name": 3, "Amount": 4}
            sort_index = index_map.get(self.history_sort_column, 0)
            def sort_key(entry):
                value = entry['values'][sort_index] if sort_index < len(entry['values']) else ""
                if value is None:
                    return ""
                    if self.history_sort_column == "No":
                        return str(value).lower()
                if self.history_sort_column == "Date":
                    try:
                        return time.strptime(str(value), "%Y-%m-%d")
                    except Exception:
                        return str(value).lower()
                if self.history_sort_column == "Amount":
                    v = str(value).strip()
                    # treat '-' or empty as missing; place missing values at the end
                    if v == "-" or v == "":
                        return float('inf') if not self.history_sort_reverse else float('-inf')
                    try:
                        return float(v)
                    except Exception:
                        return str(value).lower()
                return str(value).lower()
            sorted_entries = sorted(sorted_entries, key=sort_key, reverse=self.history_sort_reverse)

        for entry in sorted_entries:
            values = entry['values']
            meta = entry['meta']
            combined = " ".join(
                [str(v) for v in values if v is not None] +
                [meta.get('extra','') or '', meta.get('source','') or '', meta.get('json','') or '']
            ).lower()

            if tokens and not all(token in combined for token in tokens):
                continue

            tag = "highlight" if tokens and any(token in combined for token in tokens) else "normal"
            item_id = self.tree.insert("", "end", values=values, tags=(tag,))
            self.history_map[item_id] = meta

    def sort_history_column(self, column: str):
        if self.history_sort_column == column:
            self.history_sort_reverse = not self.history_sort_reverse
        else:
            self.history_sort_column = column
            self.history_sort_reverse = False
        self.refresh_history_tree()

    def _update_history_headers(self):
        for col in ("Type", "No", "Date", "Name", "Amount"):
            label = col
            if col == self.history_sort_column:
                label += " ▲" if not self.history_sort_reverse else " ▼"
            if col in ("Date", "Name", "Amount"):
                self.tree.heading(col, text=label, command=lambda c=col: self.sort_history_column(c))
            else:
                self.tree.heading(col, text=label)

    def on_history_double_click(self, event):
        sel = self.tree.selection()
        if not sel:
            return
        item_id = sel[0]
        meta = self.history_map.get(item_id, {})
        if not meta:
            messagebox.showwarning("Missing metadata", "This history record has no saved metadata.")
            return
        self.open_record_for_edit(meta)

    def view_image_popup(self, image_path: Optional[str]):
        if not image_path or not os.path.exists(image_path):
            messagebox.showwarning("Not Found", "Saved image not available.")
            return
        win = tk.Toplevel(self.root)
        win.title("View Saved Image")
        self.center_window(win, 800, 600)
        try:
            img = Image.open(image_path)
            img.thumbnail((780, 560))
            imgtk = ImageTk.PhotoImage(img)
            lbl = tk.Label(win, image=imgtk)
            lbl.image = imgtk
            lbl.pack(padx=10, pady=10)
        except Exception as e:
            tk.Label(win, text=f"Unable to open image: {e}").pack()

    def open_record_for_edit(self, meta: Dict[str, Any]):
        json_path = meta.get('json')
        source = meta.get('source')
        if json_path and os.path.exists(json_path):
            try:
                with open(json_path, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                doc = DocumentData.model_validate(data)
                self.open_edit_dialog(doc, source or json_path, data.get('barcode_data'), history_meta=meta)
                return
            except Exception as e:
                messagebox.showerror("Open Failed", f"Cannot open saved JSON: {e}")
                return
        if source and os.path.exists(source):
            try:
                document_data, barcode_data = extract_document_data(source)
                self.open_edit_dialog(document_data, source, barcode_data, history_meta=meta)
                return
            except Exception as e:
                messagebox.showerror("Re-extraction Failed", f"Cannot re-extract document: {e}")
                return
        messagebox.showinfo("Not available", "No saved JSON or source image is available for this record.")

    def delete_history_record(self, meta: Dict[str, Any], item_id: Optional[str] = None, edit_window: Optional[tk.Toplevel] = None):
        if not meta:
            messagebox.showwarning("No record", "No metadata available for this row.")
            return
        if not messagebox.askyesno("Confirm Delete", "Delete saved image, JSON, and Excel record? This cannot be undone."):
            return
        source = meta.get('source')
        jsonp = meta.get('json')
        doc_type = meta.get('doc_type')
        number = meta.get('number')
        # delete image and json files
        try:
            if source and os.path.exists(source):
                os.remove(source)
        except Exception:
            pass
        try:
            if jsonp and os.path.exists(jsonp):
                os.remove(jsonp)
        except Exception:
            pass
        # remove rows from corresponding excel
        try:
            if doc_type == 'Invoice' and Path('invoice_output.xlsx').exists():
                df_main = pd.read_excel('invoice_output.xlsx', sheet_name='Invoice', engine='openpyxl')
                keep = df_main[df_main.get('source_file','') != source]
                with pd.ExcelWriter('invoice_output.xlsx', engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                    keep.to_excel(writer, sheet_name='Invoice', index=False)
                    try:
                        items_df = pd.read_excel('invoice_output.xlsx', sheet_name='Items', engine='openpyxl')
                        if number and 'invoice_number' in items_df.columns:
                            items_df = items_df[items_df.get('invoice_number') != number]
                        items_df.to_excel(writer, sheet_name='Items', index=False)
                    except Exception:
                        pass
            if doc_type == 'Delivery Order' and Path('delivery_order_output.xlsx').exists():
                df_main = pd.read_excel('delivery_order_output.xlsx', sheet_name='Delivery Order', engine='openpyxl')
                keep = df_main[df_main.get('source_file','') != source]
                with pd.ExcelWriter('delivery_order_output.xlsx', engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                    keep.to_excel(writer, sheet_name='Delivery Order', index=False)
                    try:
                        items_df = pd.read_excel('delivery_order_output.xlsx', sheet_name='Items', engine='openpyxl')
                        if number and 'do_number' in items_df.columns:
                            items_df = items_df[items_df.get('do_number') != number]
                        items_df.to_excel(writer, sheet_name='Items', index=False)
                    except Exception:
                        pass
        except Exception as e:
            print(f"Error pruning excel: {e}")
        # also remove from master json
        try:
            mf = 'all_extracted_documents.json'
            if Path(mf).exists():
                with open(mf, 'r', encoding='utf-8') as f:
                    arr = json.load(f)
                arr = [r for r in arr if r.get('source') != (os.path.basename(source) if source else None)]
                with open(mf, 'w', encoding='utf-8') as f:
                    json.dump(arr, f, indent=2, ensure_ascii=False)
        except Exception:
            pass
        if item_id and item_id in self.history_map:
            del self.history_map[item_id]
        if edit_window:
            edit_window.destroy()
        self.load_history()

    # ===================================================
    # 4. Smart Dynamic Editing GUI (Fixed Zoom Logic)
    # ===================================================
    def open_edit_dialog(self, document_data: DocumentData, image_path: str, barcode_data=None, history_meta: Optional[Dict[str, Any]] = None):
        history_meta = history_meta or {}
        is_invoice = (document_data.document_type == "Invoice")
        
        edit_win = tk.Toplevel(self.root)
        edit_win.title(f"Validate Extracted Structured {document_data.document_type}")
        edit_win.geometry("1400x700")
        self.center_window(edit_win, 1400, 700)

        # Layout Panes
        left_frame = tk.Frame(edit_win, width=600, height=700)
        left_frame.pack(side="left", fill="both", expand=True, padx=10, pady=10)
        left_frame.pack_propagate(False)

        right_frame = tk.Frame(edit_win, width=600, height=700)
        right_frame.pack(side="left", fill="both", expand=True, padx=10, pady=10)
        right_frame.pack_propagate(False)

        # Dictionary configuration container to safely store variable state values locally across nested scopes
        state = {
            'zoom': None,
            'image_tk': None,
            'fit_done': False
        }

        try:
            img_original = Image.open(image_path)
            orig_w, orig_h = img_original.size
            
            img_canvas = tk.Canvas(left_frame, bg="white")
            img_canvas.pack(fill="both", expand=True, pady=5)

            def redraw_image(event=None):
                img_canvas.update_idletasks()
                canvas_w = max(1, img_canvas.winfo_width())
                canvas_h = max(1, img_canvas.winfo_height())
                if not state['fit_done'] and canvas_w > 1 and canvas_h > 1:
                    state['zoom'] = min(10.0, max(0.05, min(canvas_w / orig_w, canvas_h / orig_h)))
                    state['fit_done'] = True
                if state['zoom'] is None:
                    state['zoom'] = 0.35
                dw = max(1, int(orig_w * state['zoom']))
                dh = max(1, int(orig_h * state['zoom']))
                resized_img = img_original.resize((dw, dh), Image.LANCZOS)
                state['image_tk'] = ImageTk.PhotoImage(resized_img)
                img_canvas.delete("all")
                if dw <= canvas_w and dh <= canvas_h:
                    img_canvas.create_image(canvas_w // 2, canvas_h // 2, anchor="center", image=state['image_tk'])
                    img_canvas.config(scrollregion=(0, 0, canvas_w, canvas_h))
                else:
                    img_canvas.create_image(0, 0, anchor="nw", image=state['image_tk'])
                    img_canvas.config(scrollregion=(0, 0, dw, dh))
                    img_canvas.xview_moveto(max(0, (dw - canvas_w) / 2 / dw))
                    img_canvas.yview_moveto(max(0, (dh - canvas_h) / 2 / dh))

            def run_zoom_in():
                state['zoom'] = min(10.0, state['zoom'] * 1.25)
                redraw_image()

            def run_zoom_out():
                state['zoom'] = max(0.05, state['zoom'] * 0.8)
                redraw_image()

            tk.Label(left_frame, text="Scroll and drag to zoom and pan the preview.", fg="gray").pack(pady=4)
            img_canvas.bind("<Configure>", redraw_image)
            img_canvas.bind("<MouseWheel>", lambda e: run_zoom_in() if getattr(e, 'delta', 0) > 0 else run_zoom_out())
            img_canvas.bind("<ButtonPress-1>", lambda e: img_canvas.scan_mark(e.x, e.y))
            img_canvas.bind("<B1-Motion>", lambda e: img_canvas.scan_dragto(e.x, e.y, gain=1))
            redraw_image()
        except Exception as err:
            tk.Label(left_frame, text=f"Preview unavailable: {err}").pack()

        # Notebook tabs
        notebook = ttk.Notebook(right_frame)
        doc_tab = tk.Frame(notebook)
        items_tab = tk.Frame(notebook)
        notebook.add(doc_tab, text="Document Data")
        notebook.add(items_tab, text="Item Line Grid Data")
        notebook.pack(fill="both", expand=True)

        # Scroll Setup
        canvas = tk.Canvas(doc_tab)
        scrollable_frame = tk.Frame(canvas)
        sb = tk.Scrollbar(doc_tab, orient="vertical", command=canvas.yview)
        scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=sb.set)
        canvas.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        fields = {}
        row_idx = 0

        def append_field(lbl, value):
            nonlocal row_idx
            tk.Label(scrollable_frame, text=lbl).grid(row=row_idx, column=0, sticky="w", padx=10, pady=4)
            if lbl == "Description":
                ent = scrolledtext.ScrolledText(scrollable_frame, width=50, height=4, wrap=tk.WORD)
                ent.insert("1.0", str(value or ""))
            else:
                ent = tk.Entry(scrollable_frame, width=50)
                ent.insert(0, str(value or ""))
            ent.grid(row=row_idx, column=1, padx=10, pady=4, sticky="ew")
            fields[lbl] = ent
            row_idx += 1

        tk.Label(scrollable_frame, text="Document Type detected:").grid(row=row_idx, column=0, sticky="w", padx=10, pady=4)
        tk.Label(scrollable_frame, text=document_data.document_type, font=("Arial", 10, "bold")).grid(row=row_idx, column=1, sticky="w", padx=10, pady=4)
        row_idx += 1

        if is_invoice:
            inv = document_data.invoice or InvoiceData()
            append_field("Invoice Number", inv.invoice_number)
            append_field("Sender Name", inv.sender_name)
            append_field("Sender Phone", inv.sender_phone_number)
            append_field("Sender Email", inv.sender_email)
            append_field("Sender Address", inv.sender_address)
            append_field("Date", inv.date)
            append_field("Due Date", inv.duedate)
            append_field("Description", inv.description)
            append_field("Subtotal", inv.subtotal)
            append_field("Tax", inv.tax)
            append_field("Total Price", inv.total_price)
            
            item_labels = ["Item Number", "Description", "Quantity", "Unit Price", "Total Price"]
            initial_rows = [(i.item_number or "", i.description or "", i.quantity or "", i.unit_price or "", i.total_price or "") for i in inv.items]
        else:
            do = document_data.delivery_order or DeliveryOrderData()
            append_field("DO Number", do.do_number)
            append_field("PO Reference", do.po_reference)
            append_field("Delivery Date", do.delivery_date)
            append_field("Recipient Name", do.recipient_name)
            append_field("Shipping Address", do.shipping_address)
            append_field("Description", do.description)
            append_field("Received By Signature", "Yes" if do.received_by_signature else "No")
            
            item_labels = ["Item Number", "Description", "Quantity", "UOM"]
            initial_rows = [(i.item_number or "", i.description or "", i.quantity or "", i.uom or "") for i in do.items]

        columns = tuple(item_labels)
        tree = ttk.Treeview(items_tab, columns=columns, show="headings", height=12)
        for col in columns:
            tree.heading(col, text=col)
            tree.column(col, width=120, anchor="w")
        for r_vals in initial_rows:
            tree.insert("", "end", values=r_vals)
        tree.pack(fill="both", expand=True, padx=10, pady=5)
        tk.Label(items_tab, text="Double-click or right-click a row for more actions.", fg="gray").pack(pady=2)

        def edit_row_popup(item_id):
            vals = tree.item(item_id, "values")
            pop = tk.Toplevel(edit_win)
            pop.title("Modify Item Entry")
            pop.geometry("750x380")
            self.center_window(pop, 750, 380)
            pop.grid_columnconfigure(1, weight=1)
            
            item_entries = {}
            for idx, label in enumerate(item_labels):
                tk.Label(pop, text=label).grid(row=idx, column=0, padx=10, pady=5, sticky="w")
                if label == "Description":
                    e = scrolledtext.ScrolledText(pop, width=80, height=5, wrap=tk.WORD)
                    e.insert("1.0", vals[idx])
                else:
                    e = tk.Entry(pop, width=55)
                    e.insert(0, vals[idx])
                e.grid(row=idx, column=1, padx=10, pady=5, sticky="ew")
                item_entries[label] = e

            def save_row():
                new_vals = []
                for l in item_labels:
                    if l == "Description":
                        new_vals.append(item_entries[l].get("1.0", tk.END).strip())
                    else:
                        new_vals.append(item_entries[l].get())
                tree.item(item_id, values=new_vals)
                pop.destroy()

            def delete_row():
                if messagebox.askyesno("Delete row", "Delete this row?"):
                    tree.delete(item_id)
                    pop.destroy()

            btn_frame = tk.Frame(pop)
            btn_frame.grid(row=len(item_labels), column=0, columnspan=2, pady=10)
            tk.Button(btn_frame, text="Apply changes", command=save_row, width=16).pack(side="left", padx=10)
            tk.Button(btn_frame, text="Delete row", command=delete_row, fg="red", width=16).pack(side="left", padx=10)

        def append_row_popup():
            pop = tk.Toplevel(edit_win)
            pop.title("Add New Row Item")
            pop.geometry("750x380")
            self.center_window(pop, 750, 380)
            pop.grid_columnconfigure(1, weight=1)
            item_entries = {}
            for idx, label in enumerate(item_labels):
                tk.Label(pop, text=label).grid(row=idx, column=0, padx=10, pady=5, sticky="w")
                if label == "Description":
                    e = scrolledtext.ScrolledText(pop, width=80, height=5, wrap=tk.WORD)
                else:
                    e = tk.Entry(pop, width=55)
                e.grid(row=idx, column=1, padx=10, pady=5, sticky="ew")
                item_entries[label] = e

            def add_row():
                new_vals = []
                for l in item_labels:
                    if l == "Description":
                        new_vals.append(item_entries[l].get("1.0", tk.END).strip())
                    else:
                        new_vals.append(item_entries[l].get())
                tree.insert("", "end", values=new_vals)
                pop.destroy()
            tk.Button(pop, text="Append Row", command=add_row).grid(row=len(item_labels), column=0, columnspan=2, pady=10)

        tk.Button(items_tab, text="+ Add Item Line Row", command=append_row_popup).pack(pady=5)

        def on_item_double_click(event):
            item_id = tree.identify_row(event.y)
            if not item_id:
                return
            edit_row_popup(item_id)

        def delete_row_by_id(item_id):
            if item_id and messagebox.askyesno("Delete", "Remove selected row entry?"):
                tree.delete(item_id)

        def delete_selected_row(event=None):
            delete_row_by_id(tree.focus())

        def show_item_context_menu(event):
            item_id = tree.identify_row(event.y)
            if not item_id:
                return
            tree.selection_set(item_id)
            menu = tk.Menu(tree, tearoff=0)
            menu.add_command(label="Edit record", command=lambda iid=item_id: edit_row_popup(iid))
            menu.add_command(label="Delete record", command=lambda iid=item_id: delete_row_by_id(iid))
            try:
                menu.tk_popup(event.x_root, event.y_root)
            finally:
                menu.grab_release()

        tree.bind("<Double-1>", on_item_double_click)
        tree.bind("<Delete>", delete_selected_row)
        tree.bind("<Button-3>", show_item_context_menu)

        def save_compiled_dataset():
            try:
                if is_invoice:
                    inv_map = {
                        "invoice_number": fields["Invoice Number"].get() or None,
                        "sender_name": fields["Sender Name"].get() or None,
                        "sender_phone_number": fields["Sender Phone"].get() or None,
                        "sender_email": fields["Sender Email"].get() or None,
                        "sender_address": fields["Sender Address"].get() or None,
                        "date": fields["Date"].get() or None,
                        "duedate": fields["Due Date"].get() or None,
                        "description": fields["Description"].get("1.0", tk.END).strip() or None,
                        "subtotal": float(fields["Subtotal"].get()) if fields["Subtotal"].get() else None,
                        "tax": float(fields["Tax"].get()) if fields["Tax"].get() else None,
                        "total_price": float(fields["Total Price"].get()) if fields["Total Price"].get() else None,
                        "items": []
                    }
                    for node in tree.get_children():
                        v = tree.item(node, "values")
                        inv_map["items"].append({
                            "invoice_number": inv_map["invoice_number"],
                            "item_number": v[0] or None, "description": v[1] or None,
                            "quantity": float(v[2]) if v[2] else None, "unit_price": float(v[3]) if v[3] else None, "total_price": float(v[4]) if v[4] else None
                        })
                    out_data = DocumentData(document_type="Invoice", invoice=InvoiceData(**inv_map))
                else:
                    do_map = {
                        "do_number": fields["DO Number"].get() or None,
                        "po_reference": fields["PO Reference"].get() or None,
                        "delivery_date": fields["Delivery Date"].get() or None,
                        "recipient_name": fields["Recipient Name"].get() or None,
                        "shipping_address": fields["Shipping Address"].get() or None,
                        "description": fields["Description"].get("1.0", tk.END).strip() or None,
                        "received_by_signature": fields["Received By Signature"].get().lower() in ("yes", "true", "1"),
                        "items": []
                    }
                    for node in tree.get_children():
                        v = tree.item(node, "values")
                        do_map["items"].append({
                            "do_number": do_map["do_number"],
                            "item_number": v[0] or None, "description": v[1] or None,
                            "quantity": float(v[2]) if v[2] else None, "uom": v[3] or None
                        })
                    out_data = DocumentData(document_type="Delivery Order", delivery_order=DeliveryOrderData(**do_map))

                # Persist document JSON first (so we have the exact file created), then write Excel with metadata
                json_file = save_document_json(image_path, out_data, barcode_data)
                replace_source = history_meta.get('source') if history_meta else None
                old_number = history_meta.get('number') if history_meta else None
                save_to_excel(out_data, source_path=image_path, json_path=json_file, replace_source=replace_source, old_number=old_number)
                append_to_master_json(image_path, out_data, barcode_data, json_path=json_file)
                self.load_history()
                edit_win.destroy()
                messagebox.showinfo("Success", f"{out_data.document_type} successfully updated and recorded!")
            except Exception as ex:
                messagebox.showerror("Validation Field Error", f"Review field structural formats: {ex}")

        btn_frame = tk.Frame(scrollable_frame)
        btn_frame.grid(row=row_idx, column=0, columnspan=2, pady=15)
        tk.Button(btn_frame, text="Save Document", font=("Arial", 10, "bold"), bg="#4CAF50", fg="white", command=save_compiled_dataset).pack(side="left", padx=10)
        if history_meta:
            tk.Button(btn_frame, text="Delete Record", font=("Arial", 10, "bold"), bg="#F44336", fg="white", command=lambda: self.delete_history_record(history_meta, None, edit_win)).pack(side="left", padx=10)
        tk.Button(btn_frame, text="Cancel", command=edit_win.destroy).pack(side="left", padx=10)

    def on_closing(self):
        self.cap.release()
        self.root.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    app = DocumentProcessingApp(root)
    root.protocol("WM_DELETE_WINDOW", app.on_closing)
    root.mainloop()