import json
import os
import time
import base64
import cv2
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
# 1. JSON Schema
# =========================

class InvoiceItem(BaseModel):
    # This now matches the parent field name exactly for easier relational mapping
    invoice_number: Optional[str] = Field(
        None, description="The parent invoice number this item belongs to"
    )
    item_number: Optional[str] = None
    description: Optional[str] = None 
    quantity: Optional[float] = None
    unit_price: Optional[float] = None
    total_price: Optional[float] = None

class DeliveryOrderItem(BaseModel):
    do_number: Optional[str] = Field(
        None, description="The parent DO number this item belongs to"
    )
    item_number: Optional[str] = None
    description: Optional[str] = None
    quantity: Optional[float] = None
    uom: Optional[str] = Field(None, description="Unit of Measure, e.g., 'kg', 'pcs'")

class InvoiceData(BaseModel):
    """Main class for Invoice extraction"""
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

class DeliveryOrderData(BaseModel):
    """Main class for Delivery Order extraction"""
    do_number: Optional[str] = None
    po_reference: Optional[str] = Field(None, description="Purchase Order number")
    delivery_date: Optional[str] = None
    recipient_name: Optional[str] = None
    shipping_address: Optional[str] = None
    vehicle_number: Optional[str] = None
    items: List[DeliveryOrderItem] = Field(default_factory=list)
    received_by_signature: Optional[bool] = Field(
        None, description="True if a signature is detected on the document"
    )

class DocumentData(BaseModel):
    document_type: Literal["Invoice", "Delivery Order"]
    invoice: Optional[InvoiceData] = None
    delivery_order: Optional[DeliveryOrderData] = None


# =========================
# 2. Logic Functions
# =========================

def encode_image_to_base64(image_path: str) -> str:
    """Encode image file to base64 string for sending to Ollama."""
    with open(image_path, "rb") as image_file:
        return base64.b64encode(image_file.read()).decode("utf-8")

def save_to_excel(document_data: DocumentData, filename: Optional[str] = None):
    """Save document data to Excel, using a separate file for delivery orders."""
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

    if Path(filename).exists():
        try:
            existing_main = pd.read_excel(filename, sheet_name=sheet_name, engine='openpyxl')
        except Exception:
            existing_main = pd.DataFrame()
        try:
            existing_items = pd.read_excel(filename, sheet_name='Items', engine='openpyxl')
        except Exception:
            existing_items = pd.DataFrame()

        combined_main = pd.concat([existing_main, main_data], ignore_index=True, sort=False) if not existing_main.empty else main_data
        combined_items = pd.concat([existing_items, items_df], ignore_index=True, sort=False) if not existing_items.empty else items_df

        with pd.ExcelWriter(filename, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            combined_main.to_excel(writer, sheet_name=sheet_name, index=False)
            combined_items.to_excel(writer, sheet_name='Items', index=False)
    else:
        with pd.ExcelWriter(filename, engine='openpyxl') as writer:
            main_data.to_excel(writer, sheet_name=sheet_name, index=False)
            items_df.to_excel(writer, sheet_name='Items', index=False)

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
    """Extract data from barcodes and QR codes in the image."""
    if not PYZBAR_AVAILABLE:
        return None
    
    try:
        image = cv2.imread(image_path)
        if image is None:
            return None
        
        decoded_objects = pyzbar.decode(image)
        
        if not decoded_objects:
            return None
        
        barcode_data_list = []
        for obj in decoded_objects:
            data = obj.data.decode('utf-8')
            try:
                # Try to parse as JSON
                parsed_data = json.loads(data)
                barcode_data_list.append({
                    "type": obj.type,
                    "data": parsed_data
                })
            except json.JSONDecodeError:
                # If not JSON, store as plain text
                barcode_data_list.append({
                    "type": obj.type,
                    "data": data
                })
        
        return {"barcodes": barcode_data_list} if barcode_data_list else None
    except Exception as e:
        print(f"Barcode/QR extraction error: {e}")
        return None

def extract_document_data(file_path: str) -> tuple[DocumentData, Optional[Dict[str, Any]]]:
    """Two-step extraction:
    1) Lightweight classification to determine `Invoice` vs `Delivery Order`.
    2) Detailed extraction using a schema specific to the determined type.
    Returns a fully populated `DocumentData` and optional `barcode_data`.
    """
    image_base64 = encode_image_to_base64(file_path)
    model_name = get_ollama_model()

    import time
    start_time = time.time()

    rules_text = """Rules:
1. If the page has payment terms, tax, Total Price, or Total Amount Due, set document_type to "Invoice".
2. Otherwise set document_type to "Delivery Order".
3. If Invoice, fill the invoice object and set delivery_order to null.
4. If Delivery Order, fill the delivery_order object and set invoice to null.
5. sender_name is the company/enterprise that issued the document.
6. Use null for missing data.
7. Do not translate text.
8. Do not add any extra fields.
"""

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
}
"""

    delivery_schema_text = """Delivery Order Schema:
{
  "do_number": string | null,
  "po_reference": string | null,
  "delivery_date": string | null,
  "recipient_name": string | null,
  "shipping_address": string | null,
  "vehicle_number": string | null,
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
}
"""

    # --- Step 1: Classify document type (very small prompt, returns only the word) ---
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
        # Fallback to original heuristic if classification fails
        doc_type = "Invoice"

    # --- Step 2: Detailed extraction using a schema specific to the type ---
    try:
        if doc_type == "Invoice":
            detail_prompt = (
                f"{rules_text}\n"
                "Return ONLY the JSON object for an Invoice matching the Invoice schema below. "
                "Use null for missing fields and do not add extra fields. "
                "Do not return any commentary.\n"
                f"{invoice_schema_text}"
            )
            resp = chat(
                model=model_name,
                messages=[{"role": "user", "content": detail_prompt, "images": [image_base64]}],
                format=InvoiceData.model_json_schema(),
                options={"temperature": 0}
            )
            invoice = InvoiceData.model_validate_json(resp.message.content)
            # ensure child items carry parent number
            if invoice and invoice.invoice_number:
                for item in invoice.items:
                    item.invoice_number = invoice.invoice_number
            document_data = DocumentData(document_type="Invoice", invoice=invoice, delivery_order=None)
        else:
            detail_prompt = (
                f"{rules_text}\n"
                "Return ONLY the JSON object for a Delivery Order matching the Delivery Order schema below. "
                "Use null for missing fields and do not add extra fields. "
                "Do not return any commentary.\n"
                f"{delivery_schema_text}"
            )
            resp = chat(
                model=model_name,
                messages=[{"role": "user", "content": detail_prompt, "images": [image_base64]}],
                format=DeliveryOrderData.model_json_schema(),
                options={"temperature": 0}
            )
            delivery = DeliveryOrderData.model_validate_json(resp.message.content)
            if delivery and delivery.do_number:
                for item in delivery.items:
                    item.do_number = delivery.do_number
            document_data = DocumentData(document_type="Delivery Order", invoice=None, delivery_order=delivery)
    except Exception as e:
        # On any extraction error, raise to caller
        raise

    end_time = time.time()
    runtime = end_time - start_time
    print(f"Ollama model ({model_name}) runtime: {runtime:.2f} seconds")

    # Barcode extraction + document description override (preserve previous behavior)
    barcode_data = extract_barcode_qr_data(file_path)
    if barcode_data:
        barcode_text = json.dumps(barcode_data, indent=2, ensure_ascii=False)
        if document_data.document_type == "Invoice" and document_data.invoice:
            document_data.invoice.description = barcode_text
        elif document_data.document_type == "Delivery Order" and document_data.delivery_order:
            document_data.delivery_order.description = barcode_text

    return document_data, barcode_data

def compare_invoice_data(document_data: DocumentData, barcode_data: Dict[str, Any]) -> bool:
    """Compare image-extracted data with barcode/QR extracted data."""
    if not barcode_data or "barcodes" not in barcode_data:
        return False

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
            if barcode_items_normalized == model_items_normalized:
                return True

    return False

def save_barcode_qr_json(file_path: str, barcode_data: Dict[str, Any]):
    """Save barcode/QR extracted data to JSON file."""
    json_file = file_path.rsplit('.', 1)[0] + '_barcode_qr.json'
    with open(json_file, 'w', encoding='utf-8') as f:
        json.dump(barcode_data, f, indent=2, ensure_ascii=False)
    return json_file

# =========================
# 3. Integrated GUI Class
# =========================

class InvoiceApp:
    def __init__(self, root):
        self.root = root
        self.root.title("AI Document Camera & History")
        self.root.geometry("1400x700")
        self.center_window(self.root, 1400, 700)

        self.cap = cv2.VideoCapture(1)
        self.current_frame = None

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
        try:
            window.transient(self.root)
        except Exception:
            pass

    def setup_ui(self):
        # --- TOP: Camera and List ---
        top_pane = tk.PanedWindow(self.root, orient="horizontal", sashrelief="raised")
        top_pane.pack(fill="both", expand=True, padx=5, pady=5)

        # Left: Camera Feed
        cam_frame = tk.LabelFrame(top_pane, text="Live Camera")
        top_pane.add(cam_frame, width=650)
        
        self.cam_label = tk.Label(cam_frame, bg="black")
        self.cam_label.pack(fill="both", expand=True)

        btn_frame = tk.Frame(cam_frame)
        btn_frame.pack(fill="x")
        tk.Button(btn_frame, text="CAPTURE & PROCESS", bg="#4CAF50", fg="white", 
                  font=("Arial", 12, "bold"), command=self.process_capture).pack(fill="x", pady=5)
        tk.Button(btn_frame, text="UPLOAD FILE & PROCESS", bg="#2196F3", fg="white", 
                  font=("Arial", 12, "bold"), command=self.upload_and_process).pack(fill="x", pady=5)

        # Right: History
        hist_frame = tk.LabelFrame(top_pane, text="Recent Records")
        top_pane.add(hist_frame, width=600)

        self.history_filter = tk.StringVar(value="Invoice")
        filter_frame = tk.Frame(hist_frame)
        filter_frame.pack(fill="x", padx=5, pady=5)
        tk.Label(filter_frame, text="Show:").pack(side="left")
        tk.Radiobutton(filter_frame, text="Invoice", variable=self.history_filter, value="Invoice", command=self.load_history).pack(side="left", padx=3)
        tk.Radiobutton(filter_frame, text="Delivery Order", variable=self.history_filter, value="Delivery Order", command=self.load_history).pack(side="left", padx=3)
        tk.Radiobutton(filter_frame, text="All", variable=self.history_filter, value="All", command=self.load_history).pack(side="left", padx=3)

        self.tree = ttk.Treeview(hist_frame, columns=("Type", "No", "Date", "Address", "Total"), show="headings")
        for col in ("Type", "No", "Date", "Address", "Total"):
            self.tree.heading(col, text=col)
            self.tree.column(col, width=100)
        self.tree.pack(side="left", fill="both", expand=True)
        
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
            path = "last_capture.jpg"
            cv2.imwrite(path, self.current_frame)
            try:
                document_data, barcode_data = extract_document_data(path)
                # Check if barcode/QR data matches image data
                if barcode_data and compare_invoice_data(document_data, barcode_data):
                    save_barcode_qr_json(path, barcode_data)
                self.open_edit_dialog(document_data, path, barcode_data)
            except Exception as e:
                messagebox.showerror("Error", f"Extraction failed: {e}")

    def upload_and_process(self):
        file_path = filedialog.askopenfilename(
            title="Select Image or PDF File",
            filetypes=[
                ("Image files", "*.jpg *.jpeg *.png *.bmp *.tiff"),
                ("PDF files", "*.pdf"),
                ("All files", "*.*")
            ]
        )
        if not file_path:
            return

        preview_path = None
        is_temp = False

        try:
            # Prepare a preview image for PDFs (first page) or use the image directly
            if file_path.lower().endswith('.pdf'):
                try:
                    from pdf2image import convert_from_path
                except ImportError:
                    messagebox.showerror("Error", "pdf2image library not found. Install with: pip install pdf2image")
                    return

                images = convert_from_path(file_path)
                if not images:
                    messagebox.showerror("Error", "No pages found in PDF")
                    return
                preview_path = "upload_preview.jpg"
                images[0].save(preview_path, 'JPEG')
                is_temp = True
            else:
                preview_path = file_path

            # Show preview and ask for confirmation before processing
            preview_win = tk.Toplevel(self.root)
            preview_win.title("Preview & Confirm")
            self.center_window(preview_win, 800, 700)

            # Image display
            try:
                img = Image.open(preview_path)
                img.thumbnail((760, 560))
                imgtk = ImageTk.PhotoImage(img)
                img_label = tk.Label(preview_win, image=imgtk)
                img_label.image = imgtk
                img_label.pack(padx=10, pady=10)
            except Exception as e:
                tk.Label(preview_win, text=f"Preview failed: {e}").pack(padx=10, pady=10)

            info_frame = tk.Frame(preview_win)
            info_frame.pack(fill="x", padx=10)
            tk.Label(info_frame, text=f"File: {os.path.basename(file_path)}").pack(anchor="w")
            tk.Label(info_frame, text=f"Type: {'PDF' if file_path.lower().endswith('.pdf') else 'Image'}").pack(anchor="w")

            def do_process():
                try:
                    # Perform extraction and open edit dialog
                    if file_path.lower().endswith('.pdf'):
                        # Use the temp preview image for extraction
                        document_data, barcode_data = extract_document_data(preview_path)
                    else:
                        document_data, barcode_data = extract_document_data(file_path)

                    if barcode_data and compare_invoice_data(document_data, barcode_data):
                        save_barcode_qr_json(preview_path if is_temp else file_path, barcode_data)
                    preview_win.destroy()
                    # Clean up temp preview before opening editor to avoid file locks
                    if is_temp and os.path.exists(preview_path):
                        try:
                            os.remove(preview_path)
                        except Exception:
                            pass
                    self.open_edit_dialog(document_data, file_path if not is_temp else preview_path, barcode_data)
                except Exception as e:
                    preview_win.destroy()
                    if is_temp and os.path.exists(preview_path):
                        try:
                            os.remove(preview_path)
                        except Exception:
                            pass
                    messagebox.showerror("Error", f"Processing failed: {e}")

            def do_cancel():
                preview_win.destroy()
                if is_temp and preview_path and os.path.exists(preview_path):
                    try:
                        os.remove(preview_path)
                    except Exception:
                        pass

            btns = tk.Frame(preview_win)
            btns.pack(pady=10)
            tk.Button(btns, text="Confirm & Process", bg="#4CAF50", fg="white", command=do_process).pack(side="left", padx=10)
            tk.Button(btns, text="Cancel", command=do_cancel).pack(side="left", padx=10)

        except Exception as e:
            # Cleanup any temp preview
            if is_temp and preview_path and os.path.exists(preview_path):
                try:
                    os.remove(preview_path)
                except Exception:
                    pass
            messagebox.showerror("Error", f"Processing failed: {e}")

    def process_pdf(self, pdf_path):
        try:
            from pdf2image import convert_from_path
        except ImportError:
            messagebox.showerror("Error", "pdf2image library not found. Please install it with: pip install pdf2image")
            return

        try:
            # Convert PDF to images
            images = convert_from_path(pdf_path)
            if not images:
                messagebox.showerror("Error", "No pages found in PDF")
                return

            # Use the first page for extraction
            temp_path = "temp_pdf_page.jpg"
            images[0].save(temp_path, 'JPEG')

            document_data, barcode_data = extract_document_data(temp_path)
            # Check if barcode/QR data matches image data
            if barcode_data and compare_invoice_data(document_data, barcode_data):
                save_barcode_qr_json(temp_path, barcode_data)
            self.open_edit_dialog(document_data, temp_path, barcode_data)

            # Clean up temp file
            os.remove(temp_path)

        except Exception as e:
            messagebox.showerror("Error", f"PDF processing failed: {e}")

    def load_history(self):
        for i in self.tree.get_children(): self.tree.delete(i)
        selected_filter = getattr(self, 'history_filter', tk.StringVar(value='All')).get()
        history_files = ["invoice_output.xlsx", "delivery_order_output.xlsx"]
        for file in history_files:
            if not Path(file).exists():
                continue
            sheet_name = 'Invoice' if file == 'invoice_output.xlsx' else 'Delivery Order'
            if selected_filter != 'All' and selected_filter != sheet_name:
                continue
            try:
                df = pd.read_excel(file, sheet_name=sheet_name)
                for _, row in df.tail(15).iterrows():
                    if sheet_name == 'Invoice':
                        number = row.get('invoice_number','')
                        date = row.get('date','')
                        address = row.get('sender_address','')
                        total = row.get('total_price','')
                    else:
                        number = row.get('do_number','')
                        date = row.get('delivery_date','')
                        address = row.get('shipping_address','')
                        total = ''
                    self.tree.insert("", 0, values=(sheet_name, number, date, address, total))
            except: pass

    # =========================
    # 4. Detailed Editing GUI (Advanced with per-row edit/delete)
    # =========================
    def open_edit_dialog(self, document_data, image_path, barcode_data=None):
        edit_win = tk.Toplevel(self.root)
        edit_win.title(f"Verify and Edit {document_data.document_type} Data")
        edit_win.geometry("1400x700")
        self.center_window(edit_win, 1400, 700)

        fullscreen_state = {'on': False}
        def toggle_fullscreen(window, button):
            fullscreen_state['on'] = not fullscreen_state['on']
            window.attributes('-fullscreen', fullscreen_state['on'])
            button.config(text='Exit Fullscreen' if fullscreen_state['on'] else 'Fullscreen')

        control_bar = tk.Frame(edit_win)
        control_bar.pack(fill='x', padx=10, pady=(10, 0))
        fs_button = tk.Button(control_bar, text='Fullscreen', command=lambda: toggle_fullscreen(edit_win, fs_button))
        fs_button.pack(side='right')

        # Main frames
        left_frame = tk.Frame(edit_win, width=600, height=700)
        left_frame.pack(side="left", fill="both", expand=True, padx=10, pady=10)
        left_frame.pack_propagate(False)

        right_frame = tk.Frame(edit_win, width=600, height=700)
        right_frame.pack(side="left", fill="both", expand=True, padx=10, pady=10)
        right_frame.pack_propagate(False)

        # Image preview inline with zoom controls (no popup)
        description_override = None
        try:
            img_original = Image.open(image_path)
            original_width, original_height = img_original.size

            img_zoom = 0.35
            img_canvas = tk.Canvas(left_frame, bg="white")
            img_canvas.pack(fill="both", expand=True, pady=5)

            def update_image(center=None):
                display_width = max(1, int(original_width * img_zoom))
                display_height = max(1, int(original_height * img_zoom))
                display_img = img_original.resize((display_width, display_height), Image.LANCZOS)
                img_tk = ImageTk.PhotoImage(display_img)
                img_canvas.delete("all")
                img_canvas.create_image(0, 0, anchor="nw", image=img_tk)
                img_canvas.image = img_tk
                img_canvas.config(scrollregion=(0, 0, display_width, display_height))

            def zoom_in(event=None):
                nonlocal img_zoom
                img_zoom = min(10.0, img_zoom * 1.25)
                update_image()

            def zoom_out(event=None):
                nonlocal img_zoom
                img_zoom = max(0.05, img_zoom * 0.8)
                update_image()

            control_frame = tk.Frame(left_frame)
            control_frame.pack(pady=4)
            tk.Button(control_frame, text="Zoom In", width=8, command=zoom_in).pack(side="left", padx=5)
            tk.Button(control_frame, text="Zoom Out", width=8, command=zoom_out).pack(side="left", padx=5)

            img_canvas.bind("<MouseWheel>", lambda e: zoom_in() if getattr(e, 'delta', 0) > 0 else zoom_out())
            img_canvas.bind("<ButtonPress-1>", lambda e: img_canvas.scan_mark(e.x, e.y))
            img_canvas.bind("<B1-Motion>", lambda e: img_canvas.scan_dragto(e.x, e.y, gain=1))
            update_image()
        except Exception as e:
            tk.Label(left_frame, text=f"Image preview failed:\n{str(e)}", wraplength=380).pack(fill="both", expand=True)

        if barcode_data:
            # Determine primary barcode description to prioritize over detected description
            try:
                first = (barcode_data.get('barcodes') or [None])[0]
                if first:
                    data = first.get('data') if isinstance(first, dict) else first
                    if isinstance(data, dict):
                        description_override = data.get('description') or json.dumps(data, ensure_ascii=False)
                    else:
                        description_override = str(data)
            except Exception:
                description_override = None

        def open_barcode_popup():
            if not barcode_data:
                messagebox.showinfo("Barcode/QR Data", "No barcode/QR data available.")
                return
            popup = tk.Toplevel(edit_win)
            popup.title("Barcode/QR Data")
            popup.geometry("600x450")
            self.center_window(popup, 600, 450)
            text_widget = scrolledtext.ScrolledText(popup, wrap=tk.WORD)
            text_widget.pack(fill="both", expand=True, padx=10, pady=10)
            barcode_json_str = json.dumps(barcode_data, indent=2, ensure_ascii=False)
            text_widget.insert(tk.END, barcode_json_str)
            text_widget.config(state=tk.DISABLED)
            tk.Button(popup, text="Close", command=popup.destroy).pack(pady=5)

        # Tabbed interface: Document Data / Item Data
        notebook = ttk.Notebook(right_frame)
        doc_tab = tk.Frame(notebook)
        items_tab = tk.Frame(notebook)
        notebook.add(doc_tab, text="Document Data")
        notebook.add(items_tab, text="Item Data")
        notebook.pack(fill="both", expand=True)

        # Document tab: scrollable area for document fields
        canvas = tk.Canvas(doc_tab)
        scrollbar = tk.Scrollbar(doc_tab, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        scrollable_frame.columnconfigure(1, weight=1)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Fields
        fields = {}
        row = 0
        is_invoice = document_data.document_type == "Invoice"
        invoice = document_data.invoice or InvoiceData()
        delivery_order = document_data.delivery_order or DeliveryOrderData()

        def add_field(label_text, initial_value, row_idx):
            tk.Label(scrollable_frame, text=label_text).grid(row=row_idx, column=0, sticky="w", padx=10, pady=5)
            if label_text == "Description":
                entry = scrolledtext.ScrolledText(scrollable_frame, width=50, height=5, wrap=tk.WORD)
                entry.insert("1.0", str(initial_value or ""))
                entry.grid(row=row_idx, column=1, padx=10, pady=5, sticky="ew")
            else:
                entry = tk.Entry(scrollable_frame, width=50)
                entry.insert(0, str(initial_value or ""))
                entry.grid(row=row_idx, column=1, padx=10, pady=5, sticky="ew")
            fields[label_text] = entry

        def get_field_value(label_text):
            widget = fields.get(label_text)
            if isinstance(widget, tk.Text):
                return widget.get("1.0", tk.END).strip()
            return widget.get()

        tk.Label(scrollable_frame, text="Document Type").grid(row=row, column=0, sticky="w", padx=10, pady=5)
        tk.Label(scrollable_frame, text=document_data.document_type or "").grid(row=row, column=1, sticky="w", padx=10, pady=5)
        row += 1

        if is_invoice:
            add_field("Invoice Number", invoice.invoice_number, row); row += 1
            add_field("Sender Name", invoice.sender_name, row); row += 1
            add_field("Sender Phone", invoice.sender_phone_number, row); row += 1
            add_field("Sender Email", invoice.sender_email, row); row += 1
            add_field("Sender Address", invoice.sender_address, row); row += 1
            add_field("Date", invoice.date, row); row += 1
            add_field("Due Date", invoice.duedate, row); row += 1
            add_field("Description", invoice.description, row); row += 1
            if barcode_data:
                tk.Button(scrollable_frame, text="Show Barcode/QR Data", command=open_barcode_popup).grid(row=row, column=0, columnspan=2, pady=8)
                row += 1
            add_field("Subtotal", invoice.subtotal, row); row += 1
            add_field("Tax", invoice.tax, row); row += 1
            add_field("Total Price", invoice.total_price, row); row += 1
            item_labels = ["Item Number", "Description", "Quantity", "Unit Price", "Total Price"]
            columns = ("Item Number", "Description", "Quantity", "Unit Price", "Total Price", "Edit", "Delete")
            initial_rows = [
                (
                    item.item_number or "",
                    item.description or "",
                    item.quantity or "",
                    item.unit_price or "",
                    item.total_price or "",
                    "Edit",
                    "Delete"
                )
                for item in invoice.items
            ]
        else:
            add_field("DO Number", delivery_order.do_number, row); row += 1
            add_field("PO Reference", delivery_order.po_reference, row); row += 1
            add_field("Delivery Date", delivery_order.delivery_date, row); row += 1
            add_field("Recipient Name", delivery_order.recipient_name, row); row += 1
            add_field("Shipping Address", delivery_order.shipping_address, row); row += 1
            add_field("Vehicle Number", delivery_order.vehicle_number, row); row += 1
            add_field("Description", delivery_order.description, row); row += 1
            if barcode_data:
                tk.Button(scrollable_frame, text="Show Barcode/QR Data", command=open_barcode_popup).grid(row=row, column=0, columnspan=2, pady=8)
                row += 1
            add_field("Received By Signature", "Yes" if delivery_order.received_by_signature else "No", row); row += 1
            item_labels = ["Item Number", "Description", "Quantity", "UOM"]
            columns = ("Item Number", "Description", "Quantity", "UOM", "Edit", "Delete")
            initial_rows = [
                (
                    item.item_number or "",
                    item.description or "",
                    item.quantity or "",
                    item.uom or "",
                    "Edit",
                    "Delete"
                )
                for item in delivery_order.items
            ]

        # Items tab: place the tree in the items_tab
        tk.Label(items_tab, text="Items").pack(anchor="w", padx=10, pady=(10, 0))

        tree_frame = tk.Frame(items_tab)
        tree_frame.pack(fill="both", expand=True, padx=10, pady=5)

        tree = ttk.Treeview(tree_frame, columns=columns, show="headings", height=10)
        for i, col in enumerate(columns):
            tree.heading(col, text=col)
            if col in ["Edit", "Delete"]:
                tree.column(col, width=60, anchor="center")
            else:
                tree.column(col, width=120)

        for row_values in initial_rows:
            tree.insert("", "end", values=row_values)

        tree.pack(fill="both", expand=True)

        tree.tag_configure("hover", background="#e8f0ff")
        hovered_item = [None]

        def on_tree_motion(event):
            item_id = tree.identify_row(event.y)
            column = tree.identify_column(event.x)
            if item_id:
                try:
                    col_index = int(column.lstrip("#")) - 1
                    col_name = columns[col_index]
                except (ValueError, IndexError):
                    col_name = None
                if col_name in ("Edit", "Delete"):
                    tree.configure(cursor="hand2")
                    if hovered_item[0] != item_id:
                        if hovered_item[0]:
                            tree.item(hovered_item[0], tags=())
                        tree.item(item_id, tags=("hover",))
                        hovered_item[0] = item_id
                    return
            tree.configure(cursor="")
            if hovered_item[0]:
                tree.item(hovered_item[0], tags=())
                hovered_item[0] = None

        def on_tree_leave(event):
            tree.configure(cursor="")
            if hovered_item[0]:
                tree.item(hovered_item[0], tags=())
                hovered_item[0] = None

        def on_tree_click(event):
            item_id = tree.identify_row(event.y)
            column = tree.identify_column(event.x)
            if not item_id:
                return
            try:
                col_index = int(column.lstrip("#")) - 1
                col_name = columns[col_index]
            except (ValueError, IndexError):
                return
            if col_name == "Edit":
                edit_item_dialog(item_id)
            elif col_name == "Delete":
                if messagebox.askyesno("Delete Item", "Are you sure you want to delete this item?"):
                    tree.delete(item_id)

        tree.bind("<Motion>", on_tree_motion)
        tree.bind("<Leave>", on_tree_leave)
        tree.bind("<Button-1>", on_tree_click)

        def edit_item_dialog(item_id):
            values = tree.item(item_id, "values")
            dialog = tk.Toplevel(edit_win)
            dialog.title("Edit Item")
            dialog.geometry("500x320")
            self.center_window(dialog, 500, 320)

            item_fields = {}
            for i, label in enumerate(item_labels):
                tk.Label(dialog, text=label).grid(row=i, column=0, padx=10, pady=5, sticky="w")
                if label == "Description":
                    entry = scrolledtext.ScrolledText(dialog, width=45, height=4, wrap=tk.WORD)
                    entry.insert("1.0", values[i] or "")
                    entry.grid(row=i, column=1, padx=10, pady=5, sticky="ew")
                else:
                    entry = tk.Entry(dialog, width=45)
                    entry.insert(0, values[i] or "")
                    entry.grid(row=i, column=1, padx=10, pady=5, sticky="ew")
                item_fields[label] = entry

            if is_invoice:
                def calculate_total(*args):
                    try:
                        qty = float(item_fields["Quantity"].get() or 0)
                        unit_price = float(item_fields["Unit Price"].get() or 0)
                        total = qty * unit_price
                        item_fields["Total Price"].delete(0, tk.END)
                        item_fields["Total Price"].insert(0, str(total))
                    except ValueError:
                        pass
                item_fields["Quantity"].bind("<KeyRelease>", calculate_total)
                item_fields["Unit Price"].bind("<KeyRelease>", calculate_total)

            def get_item_value(widget):
                if isinstance(widget, tk.Text):
                    return widget.get("1.0", tk.END).strip()
                return widget.get()

            def save():
                new_values = [get_item_value(item_fields[label]) for label in item_labels] + ["Edit", "Delete"]
                tree.item(item_id, values=new_values)
                dialog.destroy()

            tk.Button(dialog, text="Save", command=save).grid(row=len(item_labels), column=0, pady=10)
            tk.Button(dialog, text="Cancel", command=dialog.destroy).grid(row=len(item_labels), column=1, pady=10)

        def add_item_dialog():
            dialog = tk.Toplevel(edit_win)
            dialog.title("Add Item")
            dialog.geometry("500x320")
            self.center_window(dialog, 500, 320)

            item_fields = {}
            for i, label in enumerate(item_labels):
                tk.Label(dialog, text=label).grid(row=i, column=0, padx=10, pady=5, sticky="w")
                if label == "Description":
                    entry = scrolledtext.ScrolledText(dialog, width=45, height=4, wrap=tk.WORD)
                    entry.grid(row=i, column=1, padx=10, pady=5, sticky="ew")
                else:
                    entry = tk.Entry(dialog, width=45)
                    entry.grid(row=i, column=1, padx=10, pady=5, sticky="ew")
                item_fields[label] = entry

            if is_invoice:
                def calculate_total(*args):
                    try:
                        qty = float(item_fields["Quantity"].get() or 0)
                        unit_price = float(item_fields["Unit Price"].get() or 0)
                        total = qty * unit_price
                        item_fields["Total Price"].delete(0, tk.END)
                        item_fields["Total Price"].insert(0, str(total))
                    except ValueError:
                        pass
                item_fields["Quantity"].bind("<KeyRelease>", calculate_total)
                item_fields["Unit Price"].bind("<KeyRelease>", calculate_total)

            def get_item_value(widget):
                if isinstance(widget, tk.Text):
                    return widget.get("1.0", tk.END).strip()
                return widget.get()

            def save():
                new_values = [get_item_value(item_fields[label]) for label in item_labels] + ["Edit", "Delete"]
                tree.insert("", "end", values=new_values)
                dialog.destroy()

            tk.Button(dialog, text="Add", command=save).grid(row=len(item_labels), column=0, pady=10)
            tk.Button(dialog, text="Cancel", command=dialog.destroy).grid(row=len(item_labels), column=1, pady=10)

        tk.Button(items_tab, text="Add Item", command=add_item_dialog).pack(pady=8)

        def save_final():
            try:
                if is_invoice:
                    invoice_data = {
                        "invoice_number": fields["Invoice Number"].get() or None,
                        "sender_name": fields["Sender Name"].get() or None,
                        "sender_phone_number": fields["Sender Phone"].get() or None,
                        "sender_email": fields["Sender Email"].get() or None,
                        "sender_address": fields["Sender Address"].get() or None,
                        "date": fields["Date"].get() or None,
                        "duedate": fields["Due Date"].get() or None,
                        "description": get_field_value("Description") or None,
                        "subtotal": float(fields["Subtotal"].get()) if fields["Subtotal"].get() else None,
                        "tax": float(fields["Tax"].get()) if fields["Tax"].get() else None,
                        "total_price": float(fields["Total Price"].get()) if fields["Total Price"].get() else None,
                        "items": []
                    }
                    for child in tree.get_children():
                        values = tree.item(child, "values")[:len(item_labels)]
                        invoice_data["items"].append({
                            "invoice_number": invoice_data["invoice_number"],
                            "item_number": values[0] or None,
                            "description": values[1] or None,
                            "quantity": float(values[2]) if values[2] else None,
                            "unit_price": float(values[3]) if values[3] else None,
                            "total_price": float(values[4]) if values[4] else None,
                        })
                    validated_data = DocumentData(
                        document_type="Invoice",
                        invoice=InvoiceData(**invoice_data),
                        delivery_order=None
                    )
                else:
                    do_data = {
                        "do_number": fields["DO Number"].get() or None,
                        "po_reference": fields["PO Reference"].get() or None,
                        "delivery_date": fields["Delivery Date"].get() or None,
                        "recipient_name": fields["Recipient Name"].get() or None,
                        "shipping_address": fields["Shipping Address"].get() or None,
                        "vehicle_number": fields["Vehicle Number"].get() or None,
                        "description": get_field_value("Description") or None,
                        "received_by_signature": str(fields["Received By Signature"].get()).strip().lower() in ("yes", "true", "1"),
                        "items": []
                    }
                    for child in tree.get_children():
                        values = tree.item(child, "values")[:len(item_labels)]
                        do_data["items"].append({
                            "do_number": do_data["do_number"],
                            "item_number": values[0] or None,
                            "description": values[1] or None,
                            "quantity": float(values[2]) if values[2] else None,
                            "uom": values[3] or None,
                        })
                    validated_data = DocumentData(
                        document_type="Delivery Order",
                        invoice=None,
                        delivery_order=DeliveryOrderData(**do_data)
                    )

                save_to_excel(validated_data)
                self.load_history()
                edit_win.destroy()
                messagebox.showinfo("Saved", "Data appended to Excel!")
            except Exception as e:
                messagebox.showerror("Error", f"Invalid data: {str(e)}")

        def cancel():
            edit_win.destroy()

        # Save/Cancel buttons live in the Document tab (they act on both tabs' data)
        btn_frame = tk.Frame(scrollable_frame)
        btn_frame.grid(row=row, column=0, columnspan=2, pady=10)
        tk.Button(btn_frame, text="Save", command=save_final).pack(side="left", padx=6)
        tk.Button(btn_frame, text="Cancel", command=cancel).pack(side="left", padx=6)

    def on_closing(self):
        self.cap.release()
        self.root.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    app = InvoiceApp(root)
    root.protocol("WM_DELETE_WINDOW", app.on_closing)
    root.mainloop()