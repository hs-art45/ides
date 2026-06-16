import json
import os
import time
import base64
import re
import cv2
import shutil
import queue
import threading
import pandas as pd
import serial
from serial.tools import list_ports
from pathlib import Path
from typing import List, Optional, Dict, Any, Literal
import tkinter as tk
from tkinter import messagebox, scrolledtext, filedialog
import tkinter.ttk as ttk
import customtkinter as ctk
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
# UART Configuration
# =========================
ARDUINO_PORT = 'COM3'
ARDUINO_BAUDRATE = 9600
ARDUINO_START_BYTE = 0xAA
ARDUINO_END_BYTE = 0x55

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
# 2. Arduino Serial Communication Functions
# =========================

def verify_packet(packet):
    """Verify packet structure and checksum"""
    if len(packet) < 4:
        return False
    if packet[0] != ARDUINO_START_BYTE:
        return False
    if packet[-1] != ARDUINO_END_BYTE:
        return False
    data = packet[1:-2]
    received_checksum = packet[-2]
    calculated_checksum = sum(data) & 0xFF
    return calculated_checksum == received_checksum

def send_to_arduino(ser, msg: str):
    """Send message to Arduino via serial"""
    if ser is None or not ser.is_open:
        print(f"Cannot send {msg}: serial port not open")
        return False
    try:
        ser.write(msg.encode())
        ser.flush()
        print(f"Sent '{msg}' to Arduino")
        return True
    except Exception as e:
        print(f"Serial write failed: {e}")
        return False


# =========================
# 2. Logic Functions (from invdo_v3)
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
    try:
        base_dir = os.path.dirname(os.path.abspath(__file__))
        raw_dir = os.path.join(base_dir, "rawcapture")
        os.makedirs(raw_dir, exist_ok=True)

        timestamp = time.strftime("%Y%m%d_%H%M%S", time.localtime())
        if image_array is not None:
            dest_name = f"capture_{timestamp}.jpg"
            dest_path = os.path.join(raw_dir, dest_name)
            cv2.imwrite(dest_path, image_array, [cv2.IMWRITE_JPEG_QUALITY, 95])
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
    if env_model:
        return env_model
    return "gemma4:31b-cloud"


def _extract_json_block(text: str) -> str:
    stack = []
    in_string = False
    escaped = False
    start_index = None
    for index, char in enumerate(text):
        if in_string:
            if escaped:
                escaped = False
            elif char == "\\":
                escaped = True
            elif char == '"':
                in_string = False
            continue

        if char == '"':
            in_string = True
            continue

        if start_index is None and char in '{[':
            start_index = index
            stack.append(char)
            continue

        if start_index is not None:
            if char in '{[':
                stack.append(char)
            elif char in '}]':
                if not stack:
                    break
                opener = stack.pop()
                if (opener == '{' and char != '}') or (opener == '[' and char != ']'):
                    raise ValueError(f"Mismatched JSON brackets: {opener} vs {char}")
                if not stack:
                    return text[start_index:index + 1]

    if start_index is not None and not stack:
        return text[start_index:]
    raise ValueError("Could not extract a complete JSON block from the response text.")


def _normalize_ollama_response_text(raw_text: str) -> str:
    if raw_text is None:
        raise ValueError("Empty Ollama response content")
    text = str(raw_text).strip()

    fence_match = re.search(r'```(?:json)?\s*(.*?)\s*```', text, re.S)
    if fence_match:
        text = fence_match.group(1).strip()

    if text.startswith('"') and text.endswith('"'):
        try:
            text = json.loads(text)
        except json.JSONDecodeError:
            pass
        else:
            text = str(text).strip()

    json_start = min([idx for idx in (text.find('{'), text.find('[')) if idx != -1], default=-1)
    if json_start > 0:
        text = text[json_start:]

    if not text:
        raise ValueError("No usable JSON content found in Ollama response")

    try:
        json.loads(text)
        return text
    except json.JSONDecodeError:
        pass

    return _extract_json_block(text)


def validate_model_response(model, raw_response: str):
    normalized = _normalize_ollama_response_text(raw_response)
    parsed = json.loads(normalized)
    return model.model_validate(parsed)


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
            messages=[{"role": "user", "content": classifier_prompt, "images": [file_path]}],
            options={"temperature": 0}
        )
        doc_type_raw = (cls_resp.message.content or "").strip().strip('"').strip()
        doc_type = "Invoice" if doc_type_raw.lower().startswith("invoice") else "Delivery Order"
    except Exception:
        doc_type = "Invoice"
    
    end_time1 = time.time()
    print(f"Ollama classifier runtime ({doc_type}): {end_time1 - start_time:.2f} seconds")

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
                messages=[{"role": "user", "content": detail_prompt, "images": [file_path]}],
                format=InvoiceData.model_json_schema(),
                options={"temperature": 0}
            )
            invoice = validate_model_response(InvoiceData, resp.message.content)
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
                messages=[{"role": "user", "content": detail_prompt, "images": [file_path]}],
                format=DeliveryOrderData.model_json_schema(),
                options={"temperature": 0}
            )
            delivery = validate_model_response(DeliveryOrderData, resp.message.content)
            if delivery and delivery.do_number:
                for item in delivery.items: item.do_number = delivery.do_number
            document_data = DocumentData(document_type="Delivery Order", invoice=None, delivery_order=delivery)
            
    except Exception as e:
        raise

    end_time = time.time()
    print(f"Ollama detail runtime ({doc_type}): {end_time - end_time1:.2f} seconds")

    barcode_data = extract_barcode_qr_data(file_path)
    if barcode_data:
        barcode_text = json.dumps(barcode_data, indent=2, ensure_ascii=False)
        if document_data.document_type == "Invoice" and document_data.invoice:
            existing = document_data.invoice.description or ""
            document_data.invoice.description = f"{barcode_text}\n\n{existing}" if existing else barcode_text
        elif document_data.document_type == "Delivery Order" and document_data.delivery_order:
            existing = document_data.delivery_order.description or ""
            document_data.delivery_order.description = f"{barcode_text}\n\n{existing}" if existing else barcode_text

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
# 3. Integrated GUI Class with Arduino Support
# =========================

class DocumentProcessingApp:
    def __init__(self, root):
        self.root = root
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("dark-blue")
        self.root.title("AI Unified Document Processor (Invoices & DO) + Arduino")
        self.root.geometry("1400x750")
        self.center_window(self.root, 1400, 750)

        self.camera_index = 0
        self.cap = cv2.VideoCapture(self.camera_index)
        self.current_frame = None
        self.camera_rotation_deg = 90
        self.current_com_port = ARDUINO_PORT
        self.history_map: Dict[str, Dict[str, Any]] = {}
        self.history_entries: List[Dict[str, Any]] = []
        self.current_processing_path: Optional[str] = None
        self.available_camera_indices: List[int] = []

        # Arduino/UART Communication
        self.arduino_ser = None
        self.arduino_running = True
        self.waiting_for_22 = False
        self.confirm_dialog_open = False

        # Threading Queue Processing System Setup
        self.processing_queue = queue.Queue()
        self.is_running = True
        
        self.setup_ui()
        self.open_arduino_port()
        self.set_camera_resolution()
        self.refresh_camera_options()
        self.update_camera()
        self.load_history()

        # Start background threads
        self.worker_thread = threading.Thread(target=self.queue_consumer_worker, daemon=True)
        self.worker_thread.start()
        
        self.arduino_thread = threading.Thread(target=self.arduino_listener, daemon=True)
        self.arduino_thread.start()

    def center_window(self, window, width, height):
        window.update_idletasks()
        screen_width = window.winfo_screenwidth()
        screen_height = window.winfo_screenheight()
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2
        window.geometry(f"{width}x{height}+{x}+{y}")

    # =========================
    # Arduino Port Setup
    # =========================
    def open_arduino_port(self):
        """Try to open Arduino serial port"""
        requested_port = getattr(self, "current_com_port", ARDUINO_PORT) or ARDUINO_PORT
        if self.arduino_ser is not None and self.arduino_ser.is_open:
            try:
                self.arduino_ser.close()
            except Exception:
                pass

        attempts = 3
        for i in range(attempts):
            try:
                self.arduino_ser = serial.Serial(
                    requested_port,
                    ARDUINO_BAUDRATE,
                    timeout=0.1
                )
                print(f"Arduino connected on {requested_port}")
                self.arduino_status_var.set(f"Arduino: Connected ({requested_port})")
                self.current_com_port = requested_port
                self.com_port_var.set(requested_port)
                return
            except Exception as e:
                print(f"Arduino connection attempt {i+1} failed: {e}")
                self.arduino_ser = None
                time.sleep(1)
        
        self.arduino_status_var.set(f"Arduino: Connection Failed ({requested_port})")

    def ask_confirm_dialog(self, title: str, message: str, parent: Optional[tk.Toplevel] = None) -> bool:
        dialog_parent = parent or self.root
        dlg = ctk.CTkToplevel(dialog_parent)
        dlg.title(title)
        dlg.geometry("440x180")
        dlg.transient(dialog_parent)
        dlg.grab_set()
        dlg.attributes('-topmost', True)
        self.center_window(dlg, 440, 180)

        ctk.CTkLabel(dlg, text=message, wraplength=400, justify="left", font=("Arial", 13)).pack(fill="x", padx=20, pady=(20, 10))
        result = {"value": False}

        def on_yes():
            result["value"] = True
            dlg.destroy()

        def on_no():
            dlg.destroy()

        btn_frame = ctk.CTkFrame(dlg, fg_color="#1e1e1e")
        btn_frame.pack(fill="x", padx=20, pady=(0, 20))
        center = ctk.CTkFrame(btn_frame, fg_color="#1e1e1e")
        center.pack(expand=True)
        ctk.CTkButton(center, text="Yes", fg_color="#4CAF50", text_color="white", width=100, command=on_yes).pack(side="left", padx=(0, 10))
        ctk.CTkButton(center, text="No", fg_color="#9E9E9E", text_color="white", width=100, command=on_no).pack(side="left")

        dlg.protocol("WM_DELETE_WINDOW", on_no)
        dlg.lift()
        dlg.focus_set()
        dlg.wait_window()
        return result["value"]

    def show_info_dialog(self, title: str, message: str, parent: Optional[tk.Toplevel] = None):
        parent_win = parent or self.root
        dlg = ctk.CTkToplevel(parent_win)
        dlg.title(title)
        dlg.geometry("420x140")
        dlg.transient(parent_win)
        dlg.attributes('-topmost', True)
        self.center_window(dlg, 420, 140)
        ctk.CTkLabel(dlg, text=message, wraplength=380, justify="left", font=("Arial", 12)).pack(fill="x", padx=20, pady=(20, 10))
        def on_ok():
            dlg.destroy()
        btn = ctk.CTkButton(dlg, text="OK", fg_color="#4CAF50", width=100, command=on_ok)
        btn.pack(pady=(0, 12))
        dlg.lift()
        dlg.focus_set()
        dlg.wait_window()

    def show_error_dialog(self, title: str, message: str, parent: Optional[tk.Toplevel] = None):
        parent_win = parent or self.root
        dlg = ctk.CTkToplevel(parent_win)
        dlg.title(title)
        dlg.geometry("460x160")
        dlg.transient(parent_win)
        dlg.attributes('-topmost', True)
        self.center_window(dlg, 460, 160)
        ctk.CTkLabel(dlg, text=message, wraplength=420, justify="left", font=("Arial", 12), text_color="#FF6B6B").pack(fill="x", padx=20, pady=(20, 10))
        def on_ok():
            dlg.destroy()
        btn = ctk.CTkButton(dlg, text="OK", fg_color="#F44336", width=100, command=on_ok)
        btn.pack(pady=(0, 12))
        dlg.lift()
        dlg.focus_set()
        dlg.wait_window()

    def show_warning_dialog(self, title: str, message: str, parent: Optional[tk.Toplevel] = None):
        parent_win = parent or self.root
        dlg = ctk.CTkToplevel(parent_win)
        dlg.title(title)
        dlg.geometry("440x150")
        dlg.transient(parent_win)
        dlg.attributes('-topmost', True)
        self.center_window(dlg, 440, 150)
        ctk.CTkLabel(dlg, text=message, wraplength=400, justify="left", font=("Arial", 12)).pack(fill="x", padx=20, pady=(20, 10))
        def on_ok():
            dlg.destroy()
        btn = ctk.CTkButton(dlg, text="OK", fg_color="#FFC107", width=100, command=on_ok)
        btn.pack(pady=(0, 12))
        dlg.lift()
        dlg.focus_set()
        dlg.wait_window()

    def on_com_port_selected(self, selected_port: str):
        self.current_com_port = selected_port
        self.arduino_status_var.set(f"Arduino: Selected port {selected_port}")

    def refresh_com_ports(self):
        ports = [port.device for port in list_ports.comports()]
        if not ports:
            ports = [ARDUINO_PORT]
        self.com_port_selector.configure(values=ports)
        if self.current_com_port in ports:
            self.com_port_var.set(self.current_com_port)
        else:
            self.com_port_var.set(ports[0])
            self.current_com_port = ports[0]
        self.arduino_status_var.set(f"Arduino: Available ports {', '.join(ports)}")

    def arduino_listener(self):
        """Background thread listening for Arduino packets"""
        buffer = bytearray()
        was_connected = False
        
        while self.arduino_running:
            is_connected = self.arduino_ser is not None and self.arduino_ser.is_open
            
            # Detect connection loss
            if was_connected and not is_connected:
                print("Arduino connection lost!")
                self.root.after(0, lambda: self.arduino_status_var.set("Arduino: Connection Lost"))
            
            was_connected = is_connected
            
            if not is_connected:
                time.sleep(1)
                continue
                
            try:
                if self.arduino_ser.in_waiting:
                    byte = self.arduino_ser.read(1)
                    if byte:
                        buffer.extend(byte)
                        
                        while ARDUINO_START_BYTE in buffer:
                            start = buffer.index(ARDUINO_START_BYTE)
                            try:
                                end = buffer.index(ARDUINO_END_BYTE, start + 1)
                            except ValueError:
                                break
                            
                            packet = buffer[start:end + 1]
                            del buffer[:end + 1]
                            
                            if verify_packet(packet):
                                payload = packet[2:-2]
                                try:
                                    payload_str = payload.decode('ascii')
                                except Exception:
                                    payload_str = ''.join(f"[{b:02X}]" for b in payload)
                                
                                print(f"[VALID] Arduino: {payload_str}")
                                self.root.after(0, self.handle_arduino_payload, payload_str)
                            else:
                                packet_hex = packet.hex(' ').upper()
                                print(f"[CHECKSUM ERROR] {packet_hex}")
            except Exception as e:
                print(f"Arduino listener error: {e}")
                self.root.after(0, lambda err=str(e): self.arduino_status_var.set(f"Arduino: Error - {err}"))
                time.sleep(0.5)

    def handle_arduino_payload(self, payload_str: str):
        """Handle incoming Arduino packet"""
        if payload_str == "1":
            print("Sensor A triggered")
            self.arduino_status_var.set("Arduino: Sensor A triggered - awaiting confirmation")
            self.show_confirm_popup()

        elif payload_str == "2":
            print("Sensor A released")
            self.arduino_status_var.set("Arduino: Sensor A released")

        elif payload_str == "11":
            print("Arduino: Servo heartbeat detected - triggering photo")
            self.arduino_status_var.set("Arduino: Taking photo and starting Servo B...")
            time.sleep(1)
            # Use the same capture process as the button
            self.process_capture()
            # Set flag and send Servo B command
            self.waiting_for_22 = True
            send_to_arduino(self.arduino_ser, "B")
            self.arduino_status_var.set("Arduino: Sent 'B' - waiting for Servo B completion (22)")

        elif payload_str == "22":
            print("Arduino: Servo B completed")
            self.waiting_for_22 = False
            self.arduino_status_var.set("Arduino: Servo B completed - ready for next sequence")

        else:
            print(f"Arduino payload: {payload_str}")
            self.arduino_status_var.set(f"Arduino: Payload '{payload_str}' received")

    def show_confirm_popup(self):
        """Show confirmation dialog for Sensor A trigger"""
        if self.waiting_for_22:
            print("Still waiting for Servo B - ignoring trigger")
            self.arduino_status_var.set("Arduino: Waiting for Servo B - ignoring sensor trigger")
            return
        
        if self.confirm_dialog_open:
            return
        
        self.confirm_dialog_open = True
        ok = self.ask_confirm_dialog("Arduino Sensor Triggered", "Sensor A triggered. Send command 'A' to Arduino?", self.root)
        if ok:
            send_to_arduino(self.arduino_ser, "A")
            self.arduino_status_var.set("Arduino: Sent 'A' - waiting for heartbeat (11)")
        else:
            print("User declined to send 'A'")
            self.arduino_status_var.set("Arduino: Ready for next sensor trigger")

        self.confirm_dialog_open = False

    def setup_ui(self):
        top_pane = tk.PanedWindow(self.root, orient="horizontal", sashrelief="raised", bg="#212121")
        top_pane.pack(fill="both", expand=True, padx=5, pady=5)

        # Left View
        cam_frame = ctk.CTkFrame(top_pane, fg_color="#1e1e1e")
        top_pane.add(cam_frame, width=650)
        
        # Camera label frame
        cam_label_frame = ctk.CTkFrame(cam_frame, fg_color="#0d0d0d")
        cam_label_frame.pack(fill="both", expand=True)
        self.cam_label = tk.Label(cam_label_frame, bg="black")
        self.cam_label.pack(fill="both", expand=True)

        # COM Port Selection Row
        self.com_port_var = tk.StringVar(value=self.current_com_port)
        com_frame = ctk.CTkFrame(cam_frame, fg_color="#1e1e1e")
        com_frame.pack(fill="x", side="bottom", padx=5, pady=(5, 0))
        ctk.CTkLabel(com_frame, text="COM Port:", font=("Arial", 12, "bold")).pack(side="left", padx=(0, 5))
        self.com_port_selector = ctk.CTkOptionMenu(com_frame, values=[], variable=self.com_port_var, width=140, font=("Arial", 12), command=self.on_com_port_selected)
        self.com_port_selector.pack(side="left", padx=(0, 5))
        ctk.CTkButton(com_frame, text="REFRESH", fg_color="#FFC107", text_color="black",
                      font=("Arial", 12, "bold"), width=90, height=28, command=self.refresh_com_ports).pack(side="left", padx=(5, 0))
        ctk.CTkButton(com_frame, text="CONNECT", fg_color="#4CAF50", text_color="white",
                      font=("Arial", 12, "bold"), width=90, height=28, command=self.open_arduino_port).pack(side="left", padx=(5, 0))

        # Arduino Status Bar
        self.arduino_status_var = tk.StringVar(value="Arduino: Initializing...")
        arduino_status_label = ctk.CTkLabel(cam_frame, textvariable=self.arduino_status_var, font=("Arial", 12), text_color="#FF6B6B")
        arduino_status_label.pack(fill="x", side="bottom", padx=5, pady=2)

        # Background Queue Status Bar Line
        self.status_var = tk.StringVar(value="Status: Idle")
        status_label = ctk.CTkLabel(cam_frame, textvariable=self.status_var, font=("Arial", 13), text_color="#CCCCCC")
        status_label.pack(fill="x", side="bottom", padx=5, pady=2)

        btn_frame = ctk.CTkFrame(cam_frame, fg_color="#1e1e1e")
        btn_frame.pack(fill="x", side="bottom", padx=5, pady=5)

        camera_select_frame = ctk.CTkFrame(btn_frame, fg_color="#1e1e1e")
        camera_select_frame.pack(fill="x", pady=5)
        ctk.CTkLabel(camera_select_frame, text="Camera:", font=("Arial", 13, "bold")).pack(side="left", padx=(0, 5))
        self.camera_selector = ctk.CTkOptionMenu(camera_select_frame, values=[], command=self.apply_camera_selection, width=120, font=("Arial", 13))
        self.camera_selector.pack(side="left", padx=(0, 5))
        ctk.CTkButton(camera_select_frame, text="REFRESH", fg_color="#FFC107", text_color="black",
                  font=("Arial", 13, "bold"), width=40, height=40, command=self.refresh_camera_options).pack(side="left", padx=(5, 0))

        icon_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "icon")
        self.rotate_cw_icon = None
        self.rotate_ccw_icon = None
        try:
            cw_image = Image.open(os.path.join(icon_dir, "rotate_cw.png")).resize((24, 24), Image.LANCZOS)
            ccw_image = Image.open(os.path.join(icon_dir, "rotate_anticw.png")).resize((24, 24), Image.LANCZOS)
            self.rotate_cw_icon = ImageTk.PhotoImage(cw_image)
            self.rotate_ccw_icon = ImageTk.PhotoImage(ccw_image)
        except Exception as e:
            print(f"Failed to load rotate icons: {e}")

        ctk.CTkButton(camera_select_frame, text="", image=self.rotate_cw_icon, fg_color="#03A9F4", text_color="white",
                  width=40, height=40, command=lambda: self.rotate_camera("cw")).pack(side="left", padx=(5, 0))
        ctk.CTkButton(camera_select_frame, text="", image=self.rotate_ccw_icon, fg_color="#03A9F4", text_color="white",
                  width=40, height=40, command=lambda: self.rotate_camera("ccw")).pack(side="left", padx=(5, 0))

        ctk.CTkButton(btn_frame, text="CAPTURE & PROCESS DOCUMENT", fg_color="#4CAF50", text_color="white", 
                  font=("Arial", 15, "bold"), command=self.process_capture).pack(fill="x", pady=5)
        ctk.CTkButton(btn_frame, text="UPLOAD IMAGE / PDF FILE", fg_color="#2196F3", text_color="white", 
                  font=("Arial", 15, "bold"), command=self.upload_and_process).pack(fill="x", pady=5)

        # Right Log
        hist_frame = ctk.CTkFrame(top_pane, fg_color="#1e1e1e")
        top_pane.add(hist_frame, width=600)

        self.history_filter = tk.StringVar(value="All")
        self.history_search_text = tk.StringVar()
        self.history_sort_column: Optional[str] = None
        self.history_sort_reverse = False

        filter_frame = ctk.CTkFrame(hist_frame, fg_color="#1e1e1e")
        filter_frame.pack(fill="x", padx=5, pady=5)
        ctk.CTkLabel(filter_frame, text="Log Filter:").pack(side="left")
        ctk.CTkRadioButton(filter_frame, text="All Logs", variable=self.history_filter, value="All", command=self.load_history).pack(side="left", padx=3)
        ctk.CTkRadioButton(filter_frame, text="Invoices Only", variable=self.history_filter, value="Invoice", command=self.load_history).pack(side="left", padx=3)
        ctk.CTkRadioButton(filter_frame, text="DO Only", variable=self.history_filter, value="Delivery Order", command=self.load_history).pack(side="left", padx=3)

        search_frame = ctk.CTkFrame(hist_frame, fg_color="#1e1e1e")
        search_frame.pack(fill="x", padx=5, pady=(0, 5))
        ctk.CTkLabel(search_frame, text="Search:").pack(side="left")
        search_entry = ctk.CTkEntry(search_frame, textvariable=self.history_search_text)
        search_entry.pack(side="left", fill="x", expand=True, padx=5)
        search_entry.bind('<KeyRelease>', lambda event: self.refresh_history_tree())
        ctk.CTkButton(search_frame, text="Clear", command=self.clear_history_search).pack(side="left")

        tree_container = ctk.CTkFrame(hist_frame, fg_color="#1e1e1e")
        tree_container.pack(fill="both", expand=True, padx=5, pady=5)

        # Configure treeview style for dark mode
        style = tk.ttk.Style()
        style.theme_use('clam')
        style.configure('Dark.Treeview',
            background='#161616',
            foreground='#f5f5f5',
            fieldbackground='#161616',
            font=("Arial", 13),
            rowheight=30,
            bordercolor='#161616',
            borderwidth=0,
            relief='flat',
            highlightthickness=0
        )
        style.configure('Dark.Treeview.Heading',
            background='#161616',
            foreground='#f5f5f5',
            relief='raised',
            borderwidth=1,
            bordercolor='#ffffff',
            padding=(10, 8),
            font=("Arial", 13, "bold")
        )
        style.map('Dark.Treeview',
            background=[('selected', '#1f6feb')],
            foreground=[('selected', '#ffffff')],
            fieldbackground=[('selected', '#1f6feb')]
        )
        style.layout('Dark.Treeview', [
            ('Treeview.field', {'sticky': 'nswe', 'children': [
                ('Treeview.padding', {'sticky': 'nswe', 'children': [
                    ('Treeview.treearea', {'sticky': 'nswe'})
                ]})
            ]})
        ])
        
        self.tree = tk.ttk.Treeview(tree_container, columns=("Type", "No", "Date", "Name", "Amount"), show="headings", style="Dark.Treeview")
        self.tree.tag_configure("highlight", background="#3a3a3a", foreground="#ffffff")
        self.tree.tag_configure("normal", background="#2a2a2a", foreground="#ffffff")
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
                self.tree.column(col, width=100, minwidth=80, anchor="center")
            elif col == "Name":
                self.tree.column(col, width=230, minwidth=100, anchor="center")
            else:
                self.tree.column(col, width=40, anchor="center")
        self.tree.pack(side="left", fill="both", expand=True)
        self.tree.bind("<Double-1>", self.on_history_double_click)

        def show_history_context_menu(event):
            item_id = self.tree.identify_row(event.y)
            if not item_id:
                return
            self.tree.selection_set(item_id)
            meta = self.history_map.get(item_id, {})
            menu = tk.Menu(self.tree, tearoff=0, bg="#2a2a2a", fg="#ffffff")
            menu.add_command(label="Edit record", command=lambda: self.open_record_for_edit(meta))
            menu.add_command(label="Delete record", command=lambda: self.delete_history_record(meta, item_id=item_id))
            try:
                menu.tk_popup(event.x_root, event.y_root)
            finally:
                menu.grab_release()

        self.tree.bind("<Button-3>", show_history_context_menu)
        
        self.tree_scrollbar = ctk.CTkScrollbar(tree_container, command=self.tree.yview)
        self.tree.configure(yscrollcommand=self.tree_scrollbar.set)
        self.tree_scrollbar.pack(side="right", fill="y")

        queue_frame = ctk.CTkFrame(hist_frame, fg_color="#1e1e1e")
        queue_frame.pack(fill="x", padx=5, pady=(0, 5))

        queue_label = ctk.CTkLabel(queue_frame, text="Processing Queue", font=("Arial", 13, "bold"))
        queue_label.pack(fill="x", padx=5, pady=(5, 0))

        self.queue_status_var = tk.StringVar(value="Queued items: 0")
        queue_status_label = ctk.CTkLabel(queue_frame, textvariable=self.queue_status_var, text_color="#CCCCCC", font=("Arial", 12))
        queue_status_label.pack(fill="x", padx=5)

        self.queue_thumbnail_canvas = tk.Canvas(queue_frame, height=140, background="#2a2a2a", highlightthickness=0)
        self.queue_scrollbar = ctk.CTkScrollbar(queue_frame, orientation="horizontal", command=self.queue_thumbnail_canvas.xview)
        self.queue_thumbnail_canvas.configure(xscrollcommand=self.queue_scrollbar.set)
        self.queue_thumbnail_canvas.pack(fill="both", expand=True, padx=5, pady=(2, 5))
        self.queue_scrollbar.pack(fill="x", padx=5)

        self.queue_thumbnail_container = tk.Frame(self.queue_thumbnail_canvas, background="#2a2a2a")
        self.queue_thumbnail_canvas.create_window((0, 0), window=self.queue_thumbnail_container, anchor="nw")
        self.queue_thumbnail_container.bind(
            "<Configure>",
            lambda e: self.queue_thumbnail_canvas.configure(scrollregion=self.queue_thumbnail_canvas.bbox("all"))
        )
        self.queue_image_refs = []

    def update_camera(self):
        ret, frame = self.cap.read()
        if ret:
            self.current_frame = frame
            if self.camera_rotation_deg == 90:
                display_frame = cv2.rotate(frame, cv2.ROTATE_90_CLOCKWISE)
            elif self.camera_rotation_deg == 180:
                display_frame = cv2.rotate(frame, cv2.ROTATE_180)
            elif self.camera_rotation_deg == 270:
                display_frame = cv2.rotate(frame, cv2.ROTATE_90_COUNTERCLOCKWISE)
            else:
                display_frame = frame
            cv2_rgb = cv2.cvtColor(display_frame, cv2.COLOR_BGR2RGB)
            img = Image.fromarray(cv2_rgb)
            img.thumbnail((480, 640))
            imgtk = ImageTk.PhotoImage(img)
            self.cam_label.imgtk = imgtk
            self.cam_label.configure(image=imgtk)
        self.root.after(10, self.update_camera)

    # ===================================================
    # Background Thread Processing Consumer Loop System
    # ===================================================
    def queue_consumer_worker(self):
        """Asynchronous execution container loop operating completely outside the main GUI thread context."""
        while self.is_running:
            try:
                use_path = self.processing_queue.get(timeout=1.0)
                
                total_in_queue = self.processing_queue.qsize() + 1
                self.root.after(0, lambda: self.status_var.set(f"Status: Processing ({total_in_queue} in queue)..."))
                self.root.after(0, lambda p=use_path: self.set_current_processing_path(p))
                
                try:
                    document_data, barcode_data = extract_document_data(use_path)
                    if barcode_data and compare_invoice_data(document_data, barcode_data):
                        save_barcode_qr_json(use_path, barcode_data)
                    append_to_master_json(use_path, document_data, barcode_data)
                    
                    self.root.after(0, lambda p=use_path, d=document_data, b=barcode_data: self.on_extraction_complete(d, p, b))
                except Exception as e:
                    self.root.after(0, lambda err=str(e): self.show_error_dialog("Extraction Failed", f"Background data extraction process aborted: {err}"))
                finally:
                    self.processing_queue.task_done()
                    self.root.after(0, self.clear_current_processing)
                    self.root.after(0, self.refresh_queue_list)
                    if self.processing_queue.empty():
                        self.root.after(0, lambda: self.status_var.set("Status: Idle"))
                    
            except queue.Empty:
                continue

    def on_extraction_complete(self, document_data, use_path, barcode_data):
        """Main thread callback triggered securely upon background processing extraction fulfillment."""
        self.load_history()
        self.open_edit_dialog(document_data, use_path, barcode_data)

    def process_capture(self):
        if self.current_frame is not None:
            temp_path = "last_capture.jpg"
            cv2.imwrite(temp_path, self.current_frame)
            saved_raw = save_raw_capture(source_path=temp_path, image_array=self.current_frame)
            use_path = saved_raw or temp_path
            
            self.processing_queue.put(use_path)
        total_in_queue = self.processing_queue.qsize()
        self.status_var.set(f"Status: Queued new capture ({total_in_queue} waiting)...")
        self.refresh_queue_list()

    def pick_camera(self):
        self.apply_camera_selection()

    def apply_camera_selection(self, selection=None):
        if selection is None:
            selection = self.camera_selector.get() if hasattr(self, 'camera_selector') else str(self.camera_index)
        try:
            new_index = int(selection)
        except Exception:
            self.show_warning_dialog("Camera Selection", "Select a valid camera index from the list first.")
            return

        if new_index == self.camera_index:
            self.status_var.set(f"Status: Already using camera {self.camera_index}")
            return

        self.open_camera_index(new_index)

    def rotate_camera(self, direction: str = "cw"):
        if direction == "cw":
            self.camera_rotation_deg = (self.camera_rotation_deg + 90) % 360
        elif direction == "ccw":
            self.camera_rotation_deg = (self.camera_rotation_deg - 90) % 360
        elif direction == "180":
            self.camera_rotation_deg = (self.camera_rotation_deg + 180) % 360
        else:
            return
        self.status_var.set(f"Status: Camera rotated {self.camera_rotation_deg}°")

    def open_camera_index(self, new_index: int):
        new_cap = cv2.VideoCapture(new_index)
        if not new_cap.isOpened():
            self.show_error_dialog("Camera Error", f"Could not open camera index {new_index}.")
            if hasattr(self, 'camera_selector'):
                self.camera_selector.set(self.camera_index)
            new_cap.release()
            return

        if self.cap and self.cap.isOpened():
            self.cap.release()
        self.cap = new_cap
        self.camera_index = new_index
        self.set_camera_resolution()
        if hasattr(self, 'camera_selector'):
            self.camera_selector.set(self.camera_index)
        self.status_var.set(f"Status: Camera switched to index {self.camera_index}")
        self.refresh_queue_list()

    def get_available_cameras(self, max_devices: int = 8) -> List[int]:
        available = []
        for idx in range(max_devices):
            cap = cv2.VideoCapture(idx)
            if cap is not None and cap.isOpened():
                available.append(idx)
                cap.release()
        return available

    def set_camera_resolution(self, width: int = 3840, height: int = 2160):
        if self.cap is not None and self.cap.isOpened():
            self.cap.set(cv2.CAP_PROP_FRAME_WIDTH, width)
            self.cap.set(cv2.CAP_PROP_FRAME_HEIGHT, height)
            actual_width = int(self.cap.get(cv2.CAP_PROP_FRAME_WIDTH))
            actual_height = int(self.cap.get(cv2.CAP_PROP_FRAME_HEIGHT))
            self.status_var.set(f"Status: Camera resolution set to {actual_width}x{actual_height}")

    def refresh_camera_options(self):
        camera_indices = self.get_available_cameras(8)
        if not camera_indices:
            camera_indices = [0]
        self.available_camera_indices = camera_indices
        if hasattr(self, 'camera_selector'):
            camera_values = [str(i) for i in camera_indices]
            self.camera_selector.configure(values=camera_values)
            if self.camera_index in camera_indices:
                self.camera_selector.set(str(self.camera_index))
            else:
                self.camera_index = camera_indices[0]
                self.camera_selector.set(str(self.camera_index))
        self.status_var.set(f"Status: Available cameras: {', '.join(str(i) for i in camera_indices)}")

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
                images = convert_from_path(file_path, dpi=300, fmt='jpeg')
                if not images: raise Exception("Empty PDF Pages")
                preview_path = "upload_preview.jpg"
                images[0].save(preview_path, 'JPEG', quality=95, subsampling=0)
                is_temp = True
                temp_preview_created = True
            except Exception as e:
                self.show_error_dialog("Error", f"PDF converter failure: {e}")
                return

        saved_raw = None
        try:
            if is_temp and temp_preview_created:
                saved_raw = save_raw_capture(source_path=preview_path)
            else:
                saved_raw = save_raw_capture(source_path=file_path)
        except Exception:
            saved_raw = None
        
        model_input_path = saved_raw if saved_raw else preview_path

        preview_win = ctk.CTkToplevel(self.root)
        preview_win.title("Preview Window Confirmation")
        preview_win.attributes('-topmost', True)
        self.center_window(preview_win, 800, 760)
        preview_win.focus_set()
        preview_win.bind('<FocusOut>', lambda e: preview_win.attributes('-topmost', False))

        file_name = os.path.basename(file_path)
        file_format = os.path.splitext(file_path)[1].lstrip('.').upper() or 'Unknown'

        info_frame = ctk.CTkFrame(preview_win, fg_color="#1e1e1e")
        info_frame.pack(fill="x", padx=10, pady=(10, 0))
        ctk.CTkLabel(info_frame, text=f"File: {file_name}", font=("Arial", 13, "bold")).pack(fill="x")
        ctk.CTkLabel(info_frame, text=f"Format: {file_format}", font=("Arial", 12)).pack(fill="x")

        try:
            img_obj = Image.open(model_input_path)
            img_w, img_h = img_obj.size

            # Place canvas and scrollbars inside a dedicated frame to avoid mixing pack/grid
            canvas_container = tk.Frame(preview_win)
            canvas_container.pack(fill="both", expand=True, padx=10, pady=10)

            canvas = tk.Canvas(canvas_container, background="#222")
            hbar = ctk.CTkScrollbar(canvas_container, orientation="horizontal", command=canvas.xview,
                width=12, corner_radius=8, fg_color="#272727", button_color="#ffffff")
            vbar = ctk.CTkScrollbar(canvas_container, orientation="vertical", command=canvas.yview,
                width=12, corner_radius=8, fg_color="#272727", button_color="#ffffff")
            canvas.configure(xscrollcommand=hbar.set, yscrollcommand=vbar.set)
            canvas.grid(row=0, column=0, sticky="nsew")
            vbar.grid(row=0, column=1, sticky="ns")
            hbar.grid(row=1, column=0, sticky="ew")
            canvas_container.grid_rowconfigure(0, weight=1)
            canvas_container.grid_columnconfigure(0, weight=1)

            canvas_w_estimate = 760
            canvas_h_estimate = 560
            fit_zoom = min(canvas_w_estimate / img_w, canvas_h_estimate / img_h, 1.0)

            state = {"zoom": fit_zoom, "image": img_obj, "photo": None, "img_id": None}

            def render_zoom():
                img = state["image"]
                w, h = img.size
                new_size = (max(1, int(w * state["zoom"])), max(1, int(h * state["zoom"])))
                resized = img.resize(new_size, Image.LANCZOS)
                state["photo"] = ImageTk.PhotoImage(resized)
                canvas.delete("all")
                canvas_w = max(canvas_w_estimate, new_size[0])
                canvas_h = max(canvas_h_estimate, new_size[1])
                center_x = canvas_w / 2
                center_y = canvas_h / 2
                state["img_id"] = canvas.create_image(center_x, center_y, anchor="center", image=state["photo"])
                canvas.config(scrollregion=(0, 0, canvas_w, canvas_h))

            def on_zoom(event):
                delta = getattr(event, 'delta', 0)
                if delta > 0:
                    state["zoom"] *= 1.1
                elif delta < 0:
                    state["zoom"] *= 0.9
                state["zoom"] = max(0.2, min(state["zoom"], 5.0))
                render_zoom()

            canvas.bind("<MouseWheel>", on_zoom)
            canvas.bind("<ButtonPress-1>", lambda event: canvas.scan_mark(event.x, event.y))
            canvas.bind("<B1-Motion>", lambda event: canvas.scan_dragto(event.x, event.y, gain=1))
            render_zoom()
            tk.Label(preview_win, text="Use mouse wheel to zoom and drag to pan", background="#222", foreground="#fff").pack(fill="x", pady=4)
        except Exception as e:
            tk.Label(preview_win, text=f"Preview unavailable: {e}").pack()

        def do_process():
            self.processing_queue.put(model_input_path)
            total_in_queue = self.processing_queue.qsize()
            self.status_var.set(f"Status: File queued ({total_in_queue} waiting)...")
            self.refresh_queue_list()
            preview_win.destroy()
            if temp_preview_created and os.path.exists(preview_path):
                try: os.remove(preview_path)
                except: pass

        def change_document():
            preview_win.destroy()
            if is_temp and os.path.exists(preview_path):
                os.remove(preview_path)
            self.upload_and_process()

        btns = ctk.CTkFrame(preview_win, fg_color="#1e1e1e")
        btns.pack(pady=10)
        ctk.CTkButton(btns, text="Scan & Classify Document", fg_color="#4CAF50", text_color="white", command=do_process).pack(side="left", padx=10)
        ctk.CTkButton(btns, text="Change Document", command=change_document).pack(side="left", padx=10)
        ctk.CTkButton(btns, text="Cancel", command=preview_win.destroy).pack(side="left", padx=10)

    def load_history(self):
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
        if hasattr(row, 'to_dict'): row = row.to_dict()
        if isinstance(row, dict):
            for v in row.values():
                if isinstance(v, str): values.append(v)
                elif v is not None: values.append(str(v))
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

        for i in self.tree.get_children(): self.tree.delete(i)
        self.history_map.clear()

        sorted_entries = list(self.history_entries)
        if self.history_sort_column:
            index_map = {"Type": 0, "No": 1, "Date": 2, "Name": 3, "Amount": 4}
            sort_index = index_map.get(self.history_sort_column, 0)
            def sort_key(entry):
                value = entry['values'][sort_index] if sort_index < len(entry['values']) else ""
                if value is None: return ""
                if self.history_sort_column == "Date":
                    try: return time.strptime(str(value), "%Y-%m-%d")
                    except: return str(value).lower()
                if self.history_sort_column == "Amount":
                    v = str(value).strip()
                    if v == "-" or v == "": return float('inf') if not self.history_sort_reverse else float('-inf')
                    try: return float(v)
                    except: return str(value).lower()
                return str(value).lower()
            sorted_entries = sorted(sorted_entries, key=sort_key, reverse=self.history_sort_reverse)

        for entry in sorted_entries:
            values = entry['values']
            meta = entry['meta']
            combined = " ".join(
                [str(v) for v in values if v is not None] +
                [meta.get('extra','') or '', meta.get('source','') or '', meta.get('json','') or '']
            ).lower()

            if tokens and not all(token in combined for token in tokens): continue

            tag = "highlight" if tokens and any(token in combined for token in tokens) else "normal"
            item_id = self.tree.insert("", "end", values=values, tags=(tag,))
            self.history_map[item_id] = meta

        self.refresh_queue_list()

    def sort_history_column(self, column: str):
        if self.history_sort_column == column:
            self.history_sort_reverse = not self.history_sort_reverse
        else:
            self.history_sort_column = column
            self.history_sort_reverse = False
        self.refresh_history_tree()

    def set_current_processing_path(self, path: Optional[str]):
        self.current_processing_path = path
        self.refresh_queue_list()

    def clear_current_processing(self):
        self.current_processing_path = None
        self.refresh_queue_list()

    def _create_queue_thumbnail_card(self, parent, path, title_text):
        item_frame = tk.Frame(parent, bd=1, relief="solid", padx=4, pady=4, background="#2a2a2a")
        item_frame.pack(side="left", padx=4, pady=4)

        title_label = tk.Label(item_frame, text=title_text, font=("Arial", 12, "bold"), background="#2a2a2a", foreground="#ffffff")
        title_label.pack(pady=(0, 4))

        display_name = os.path.basename(str(path)) if path else "<unknown>"
        thumb = None
        try:
            if path and os.path.exists(str(path)):
                img = Image.open(str(path))
                img.thumbnail((130, 120), Image.LANCZOS)
                thumb = ImageTk.PhotoImage(img)
        except Exception:
            thumb = None

        if thumb:
            label_img = tk.Label(item_frame, image=thumb, background="#2a2a2a")
            label_img.image = thumb
            label_img.pack()
            label_img.bind("<Button-1>", lambda e, p=path: self.open_queue_image_preview(p, title_text))
            label_img.config(cursor="hand2")
            self.queue_image_refs.append(thumb)
        else:
            label_no = tk.Label(item_frame, text="No preview", background="#2a2a2a", foreground="#ffffff", width=18, height=5, anchor="center")
            label_no.pack(padx=2, pady=2)
            label_no.bind("<Button-1>", lambda e, p=path: self.open_queue_image_preview(p, title_text))
            label_no.config(cursor="hand2")

        tk.Label(item_frame, text=display_name, wraplength=140, justify="center", background="#2a2a2a", foreground="#ffffff", font=("Arial", 11)).pack(padx=2, pady=(4, 0))

    def open_queue_image_preview(self, image_path, title_text="Queue Preview"):
        if not image_path or not os.path.exists(str(image_path)):
            self.show_warning_dialog("Preview unavailable", "Image file is not available for preview.")
            return

        zoom_win = ctk.CTkToplevel(self.root)
        zoom_win.title(title_text)
        zoom_win.attributes('-topmost', True)
        self.center_window(zoom_win, 900, 700)
        zoom_win.bind('<FocusOut>', lambda e: zoom_win.attributes('-topmost', False))

        canvas = tk.Canvas(zoom_win, background="#222", highlightthickness=0)
        hbar = ctk.CTkScrollbar(zoom_win, orientation="horizontal", command=canvas.xview, width=12, corner_radius=8, fg_color="#272727")
        vbar = ctk.CTkScrollbar(zoom_win, orientation="vertical", command=canvas.yview, width=12, corner_radius=8, fg_color="#272727")
        canvas.configure(xscrollcommand=hbar.set, yscrollcommand=vbar.set)
        canvas.grid(row=0, column=0, sticky="nsew")
        vbar.grid(row=0, column=1, sticky="ns")
        hbar.grid(row=1, column=0, sticky="ew")
        zoom_win.grid_rowconfigure(0, weight=1)
        zoom_win.grid_columnconfigure(0, weight=1)

        img_obj = Image.open(str(image_path))
        img_w, img_h = img_obj.size
        
        canvas_w_estimate = 880
        canvas_h_estimate = 670
        fit_zoom = min(canvas_w_estimate / img_w, canvas_h_estimate / img_h, 1.0)

        state = {
            "zoom": fit_zoom,
            "image": img_obj,
            "photo": None,
            "img_id": None
        }

        def render_zoom():
            img = state["image"]
            w, h = img.size
            new_size = (max(1, int(w * state["zoom"])), max(1, int(h * state["zoom"])))
            resized = img.resize(new_size, Image.LANCZOS)
            state["photo"] = ImageTk.PhotoImage(resized)
            canvas.delete("all")
            canvas_w = max(canvas_w_estimate, new_size[0])
            canvas_h = max(canvas_h_estimate, new_size[1])
            center_x = canvas_w / 2
            center_y = canvas_h / 2
            state["img_id"] = canvas.create_image(center_x, center_y, anchor="center", image=state["photo"])
            canvas.config(scrollregion=(0, 0, canvas_w, canvas_h))

        def on_zoom(event):
            delta = getattr(event, 'delta', 0)
            if delta > 0:
                state["zoom"] *= 1.1
            elif delta < 0:
                state["zoom"] *= 0.9
            state["zoom"] = max(0.2, min(state["zoom"], 5.0))
            render_zoom()

        canvas.bind("<MouseWheel>", on_zoom)
        canvas.bind("<ButtonPress-1>", lambda event: canvas.scan_mark(event.x, event.y))
        canvas.bind("<B1-Motion>", lambda event: canvas.scan_dragto(event.x, event.y, gain=1))
        render_zoom()

        ctk.CTkLabel(zoom_win, text="Use mouse wheel to zoom and drag to pan", text_color="#ffffff").grid(row=2, column=0, columnspan=2, sticky="ew", pady=4)

    def refresh_queue_list(self):
        try:
            queue_items = list(self.processing_queue.queue)
        except Exception:
            queue_items = []

        for widget in self.queue_thumbnail_container.winfo_children():
            widget.destroy()
        self.queue_image_refs = []

        if self.current_processing_path:
            self._create_queue_thumbnail_card(self.queue_thumbnail_container, self.current_processing_path, "Current Processing")

        if queue_items:
            for idx, queued_path in enumerate(queue_items, start=1):
                self._create_queue_thumbnail_card(self.queue_thumbnail_container, queued_path, f"Queued {idx}")

        processing_count = 1 if self.current_processing_path else 0
        queued_count = len(queue_items)
        self.queue_status_var.set(f"Processing: {processing_count}, Queued: {queued_count}")

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
        if not sel: return
        item_id = sel[0]
        meta = self.history_map.get(item_id, {})
        if not meta:
            self.show_warning_dialog("Missing metadata", "This history record has no saved metadata.")
            return
        self.open_record_for_edit(meta)

    def open_record_for_edit(self, meta: Dict[str, Any]):
        json_path = meta.get('json')
        source = meta.get('source')
        if json_path and os.path.exists(json_path):
            try:
                with open(json_path, 'r', encoding='utf-8') as f: data = json.load(f)
                doc = DocumentData.model_validate(data)
                self.open_edit_dialog(doc, source or json_path, data.get('barcode_data'), history_meta=meta)
                return
            except Exception as e:
                self.show_error_dialog("Open Failed", f"Cannot open saved JSON: {e}")
                return
        if source and os.path.exists(source):
            try:
                document_data, barcode_data = extract_document_data(source)
                self.open_edit_dialog(document_data, source, barcode_data, history_meta=meta)
                return
            except Exception as e:
                self.show_error_dialog("Re-extraction Failed", f"Cannot re-extract document: {e}")
                return

    def delete_history_record(self, meta: Dict[str, Any], item_id: Optional[str] = None, edit_window: Optional[tk.Toplevel] = None):
        if not meta: return
        if not self.ask_confirm_dialog("Confirm Delete", "Delete saved image, JSON, and Excel record? This cannot be undone.", self.root): return
        source = meta.get('source')
        jsonp = meta.get('json')
        doc_type = meta.get('doc_type')
        number = meta.get('number')
        try:
            if source and os.path.exists(source): os.remove(source)
        except Exception: pass
        try:
            if jsonp and os.path.exists(jsonp): os.remove(jsonp)
        except Exception: pass
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
                    except: pass
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
                    except: pass
        except Exception as e: print(f"Error pruning excel: {e}")
        if item_id and item_id in self.history_map: del self.history_map[item_id]
        if edit_window: edit_window.destroy()
        self.load_history()

    def open_edit_dialog(self, document_data: DocumentData, image_path: str, barcode_data=None, history_meta: Optional[Dict[str, Any]] = None):
        history_meta = history_meta or {}
        is_invoice = (document_data.document_type == "Invoice")
        
        edit_win = ctk.CTkToplevel(self.root)
        edit_win.title(f"Validate Extracted Structured {document_data.document_type}")
        edit_win.geometry("1400x700")
        edit_win.attributes('-topmost', True)
        self.center_window(edit_win, 1400, 700)
        edit_win.focus_set()
        edit_win.bind('<FocusOut>', lambda e: edit_win.attributes('-topmost', False))

        left_frame = ctk.CTkFrame(edit_win, fg_color="#1e1e1e")
        left_frame.pack(side="left", fill="both", expand=True, padx=10, pady=10)

        right_frame = ctk.CTkFrame(edit_win, fg_color="#1e1e1e")
        right_frame.pack(side="left", fill="both", expand=True, padx=10, pady=10)

        state = {'zoom': None, 'image_tk': None, 'fit_done': False}

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
                if state['zoom'] is None: state['zoom'] = 0.35
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

            tk.Label(left_frame, text="Scroll and drag to zoom and pan the preview.", fg="white", bg="#1e1e1e").pack(pady=4)
            img_canvas.bind("<Configure>", redraw_image)
            img_canvas.bind("<MouseWheel>", lambda e: run_zoom_in() if getattr(e, 'delta', 0) > 0 else run_zoom_out())
            img_canvas.bind("<ButtonPress-1>", lambda e: img_canvas.scan_mark(e.x, e.y))
            img_canvas.bind("<B1-Motion>", lambda e: img_canvas.scan_dragto(e.x, e.y, gain=1))
            redraw_image()
        except Exception as err:
            tk.Label(left_frame, text=f"Preview unavailable: {err}").pack()

        notebook = tk.ttk.Notebook(right_frame)
        doc_tab = tk.Frame(notebook, bg="#1e1e1e")
        items_tab = tk.Frame(notebook, bg="#1e1e1e")
        style = tk.ttk.Style()
        style.theme_use('clam')
        style.configure('TNotebook',
            background='#1e1e1e',
            borderwidth=0,
            relief='flat'
        )
        style.configure('TNotebook.Tab',
            background='#292929',
            foreground='#ffffff',
            lightcolor='#383838',
            darkcolor='#1d1d1d',
            bordercolor='#3a3a3a',
            borderwidth=1,
            padding=(16, 8),
            font=("Arial", 14, "bold"),
            relief='raised'
        )
        style.map('TNotebook.Tab',
            background=[('selected', '#1f6feb'), ('active', '#353535'), ('!selected', '#292929')],
            foreground=[('selected', '#ffffff'), ('!selected', '#d0d0d0')]
        )
        notebook.add(doc_tab, text="Document Data")
        notebook.add(items_tab, text="Item Line Grid Data")
        notebook.pack(fill="both", expand=True)

        canvas = tk.Canvas(doc_tab, bg="#1e1e1e", highlightthickness=0)
        scrollable_frame = tk.Frame(canvas, bg="#1e1e1e")
        sb = ctk.CTkScrollbar(doc_tab, orientation="vertical", command=canvas.yview)
        scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=sb.set)
        canvas.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        fields = {}
        row_idx = 0

        def append_field(lbl, value):
            nonlocal row_idx
            tk.Label(scrollable_frame, text=lbl, bg="#1e1e1e", fg="#ffffff", font=("Arial", 16, "bold")).grid(row=row_idx, column=0, sticky="w", padx=10, pady=4)
            if lbl == "Description":
                ent = scrolledtext.ScrolledText(scrollable_frame, width=50, height=4, wrap=tk.WORD, bg="#2a2a2a", fg="#ffffff", insertbackground="#ffffff", font=("Arial", 15))
                ent.insert("1.0", str(value or ""))
            else:
                ent = tk.Entry(scrollable_frame, width=50, bg="#2a2a2a", fg="#ffffff", insertbackground="#ffffff", font=("Arial", 15))
                ent.insert(0, str(value or ""))
            ent.grid(row=row_idx, column=1, padx=10, pady=4, sticky="ew")
            fields[lbl] = ent
            row_idx += 1

        tk.Label(scrollable_frame, text="Document Type detected:", bg="#1e1e1e", fg="#ffffff", font=("Arial", 16, "bold")).grid(row=row_idx, column=0, sticky="w", padx=10, pady=4)
        tk.Label(scrollable_frame, text=document_data.document_type, font=("Arial", 16, "bold"), bg="#1e1e1e", fg="#4CAF50").grid(row=row_idx, column=1, sticky="w", padx=10, pady=4)
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
        style = tk.ttk.Style()
        style.configure('DocumentData.Treeview',
            background='#161616',
            foreground='#f5f5f5',
            fieldbackground='#161616',
            font=("Arial", 16),
            rowheight=40,
            bordercolor='#161616',
            borderwidth=0,
            relief='flat',
            highlightthickness=0
        )
        style.configure('DocumentData.Treeview.Heading',
            background='#1f1f1f',
            foreground='#f5f5f5',
            relief='raised',
            borderwidth=1,
            bordercolor="#FFFFFF",
            padding=(14, 12),
            font=("Arial", 16, "bold")
        )
        style.map('DocumentData.Treeview',
            background=[('selected', '#1f6feb')],
            foreground=[('selected', '#ffffff')],
            fieldbackground=[('selected', '#1f6feb')]
        )
        style.layout('DocumentData.Treeview', [
            ('Treeview.field', {'sticky': 'nswe', 'children': [
                ('Treeview.padding', {'sticky': 'nswe', 'children': [
                    ('Treeview.treearea', {'sticky': 'nswe'})
                ]})
            ]})
        ])
        
        tree_container = tk.Frame(items_tab, bg="#1e1e1e")
        tree_container.pack(fill="both", expand=True, padx=10, pady=5)
        
        tree = tk.ttk.Treeview(tree_container, columns=columns, show="headings", height=12, style="DocumentData.Treeview")
        for col in columns:
            tree.heading(col, text=col)
            if col == "Item Number":
                tree.column(col, width=100, anchor="w")
            elif col == "Description":
                tree.column(col, width=350, anchor="center")
            elif col == "Quantity":
                tree.column(col, width=80, anchor="center")
            elif col == "Unit Price":
                tree.column(col, width=100, anchor="center")
            elif col == "Total Price":
                tree.column(col, width=100, anchor="center")
            elif col == "UOM":
                tree.column(col, width=80, anchor="center")
            else:
                tree.column(col, width=120, anchor="w")
        for r_vals in initial_rows: tree.insert("", "end", values=r_vals)
        tree.pack(side="left", fill="both", expand=True)
        items_scrollbar = ctk.CTkScrollbar(tree_container, orientation="vertical", command=tree.yview)
        tree.configure(yscrollcommand=items_scrollbar.set)
        items_scrollbar.pack(side="right", fill="y")

        def edit_row_popup(item_id):
            vals = tree.item(item_id, "values")
            pop = ctk.CTkToplevel(edit_win)
            pop.title("Modify Item Entry")
            pop.geometry("750x380")
            pop.transient(edit_win)
            pop.attributes('-topmost', True)
            self.center_window(pop, 750, 380)
            pop.grid_columnconfigure(1, weight=1)
            pop.focus_set()
            pop.lift()
            pop.after(100, lambda: pop.attributes('-topmost', False))
            item_entries = {}
            for idx, label in enumerate(item_labels):
                ctk.CTkLabel(pop, text=label, font=("Arial", 14, "bold")).grid(row=idx, column=0, padx=10, pady=5, sticky="w")
                if label == "Description":
                    if hasattr(ctk, 'CTkTextbox'):
                        e = ctk.CTkTextbox(pop, width=520, height=120, corner_radius=8)
                        e.insert("0.0", vals[idx])
                    else:
                        e = scrolledtext.ScrolledText(pop, width=80, height=5, wrap=tk.WORD, bg="#2a2a2a", fg="#ffffff", insertbackground="#ffffff", font=("Arial", 13))
                        e.insert("1.0", vals[idx])
                else:
                    e = ctk.CTkEntry(pop, width=520, font=("Arial", 13))
                    e.insert(0, vals[idx])
                e.grid(row=idx, column=1, padx=10, pady=5, sticky="ew")
                item_entries[label] = e

            def save_row():
                new_vals = []
                for l in item_labels:
                    if l == "Description": new_vals.append(item_entries[l].get("1.0", tk.END).strip())
                    else: new_vals.append(item_entries[l].get())
                tree.item(item_id, values=new_vals)
                pop.destroy()

            def delete_row():
                if self.ask_confirm_dialog("Delete row", "Delete this row?", pop):
                    tree.delete(item_id)
                    pop.destroy()

            btn_frame = ctk.CTkFrame(pop, fg_color="#1e1e1e")
            btn_frame.grid(row=len(item_labels), column=0, columnspan=2, pady=10)
            ctk.CTkButton(btn_frame, text="Apply changes", command=save_row, width=160, font=("Arial", 13, "bold")).pack(side="left", padx=10)
            ctk.CTkButton(btn_frame, text="Delete row", command=delete_row, fg_color="#F44336", width=160, font=("Arial", 13, "bold")).pack(side="left", padx=10)

        def append_row_popup():
            pop = ctk.CTkToplevel(edit_win)
            pop.title("Add New Row Item")
            pop.geometry("750x380")
            pop.transient(edit_win)
            pop.attributes('-topmost', True)
            self.center_window(pop, 750, 380)
            pop.grid_columnconfigure(1, weight=1)
            pop.focus_set()
            pop.lift()
            pop.after(100, lambda: pop.attributes('-topmost', False))
            item_entries = {}
            for idx, label in enumerate(item_labels):
                ctk.CTkLabel(pop, text=label, font=("Arial", 14, "bold")).grid(row=idx, column=0, padx=10, pady=5, sticky="w")
                if label == "Description":
                    if hasattr(ctk, 'CTkTextbox'):
                        e = ctk.CTkTextbox(pop, width=520, height=120, corner_radius=8)
                    else:
                        e = scrolledtext.ScrolledText(pop, width=80, height=5, wrap=tk.WORD, bg="#2a2a2a", fg="#ffffff", insertbackground="#ffffff", font=("Arial", 13))
                else:
                    e = ctk.CTkEntry(pop, width=520, font=("Arial", 13))
                e.grid(row=idx, column=1, padx=10, pady=5, sticky="ew")
                item_entries[label] = e

            def add_row():
                new_vals = []
                for l in item_labels:
                    if l == "Description": new_vals.append(item_entries[l].get("1.0", tk.END).strip())
                    else: new_vals.append(item_entries[l].get())
                tree.insert("", "end", values=new_vals)
                pop.destroy()
            ctk.CTkButton(pop, text="Append Row", command=add_row, font=("Arial", 13, "bold")).grid(row=len(item_labels), column=0, columnspan=2, pady=10)

        tree.bind("<Double-1>", lambda e: edit_row_popup(tree.identify_row(e.y)))

        ctk.CTkButton(items_tab, text="+ Add Item Line Row", command=append_row_popup, font=("Arial", 15, "bold")).pack(pady=5)

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

                json_file = save_document_json(image_path, out_data, barcode_data)
                replace_source = history_meta.get('source') if history_meta else None
                old_number = history_meta.get('number') if history_meta else None
                save_to_excel(out_data, source_path=image_path, json_path=json_file, replace_source=replace_source, old_number=old_number)
                append_to_master_json(image_path, out_data, barcode_data, json_path=json_file)
                self.load_history()
                edit_win.destroy()
                self.show_info_dialog("Success", f"{out_data.document_type} successfully saved!")
            except Exception as ex:
                self.show_error_dialog("Validation Field Error", f"Review field structural formats: {ex}")

        btn_frame = ctk.CTkFrame(scrollable_frame, fg_color="#1e1e1e")
        btn_frame.grid(row=row_idx, column=0, columnspan=2, pady=15)
        ctk.CTkButton(btn_frame, text="Save Document", font=("Arial", 13, "bold"), fg_color="#4CAF50", text_color="white", command=save_compiled_dataset).pack(side="left", padx=10)
        if history_meta:
            ctk.CTkButton(btn_frame, text="Delete Record", font=("Arial", 13, "bold"), fg_color="#F44336", text_color="white", command=lambda: self.delete_history_record(history_meta, None, edit_win)).pack(side="left", padx=10)
        ctk.CTkButton(btn_frame, text="Cancel", command=edit_win.destroy).pack(side="left", padx=10)


    def on_closing(self):
        self.is_running = False
        self.arduino_running = False
        try:
            if self.arduino_ser and self.arduino_ser.is_open:
                self.arduino_ser.close()
        except:
            pass
        try:
            self.cap.release()
        except:
            pass
        self.root.destroy()


if __name__ == "__main__":
    root = ctk.CTk()
    app = DocumentProcessingApp(root)
    root.protocol("WM_DELETE_WINDOW", app.on_closing)
    root.mainloop()
