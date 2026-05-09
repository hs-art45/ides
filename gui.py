import json
import os
import time
import base64
import cv2
import pandas as pd
from pathlib import Path
from typing import List, Optional, Dict, Any
import tkinter as tk
from tkinter import messagebox, ttk, scrolledtext, filedialog
from PIL import Image, ImageTk

from pydantic import BaseModel
import ollama
from ollama import chat
from ollama._types import ResponseError

try:
    from pyzbar import pyzbar
    PYZBAR_AVAILABLE = True
except ImportError:
    PYZBAR_AVAILABLE = False

# =========================
# 1. JSON Schema (Original)
# =========================

class InvoiceItem(BaseModel):
    name: Optional[str] = None
    quantity: Optional[float] = None
    unit_price: Optional[float] = None
    total_price: Optional[float] = None

class InvoiceData(BaseModel):
    invoice_number: Optional[str] = None
    phone: Optional[str] = None
    email: Optional[str] = None
    address: Optional[str] = None
    date: Optional[str] = None
    items: List[InvoiceItem] = []
    subtotal: Optional[float] = None
    tax: Optional[float] = None
    total_price: Optional[float] = None

# =========================
# 2. Logic Functions (Original)
# =========================

def encode_image_to_base64(image_path: str) -> str:
    """Encode image file to base64 string for sending to Ollama."""
    with open(image_path, "rb") as image_file:
        return base64.b64encode(image_file.read()).decode("utf-8")

def save_to_excel(invoice_data: InvoiceData, filename: str = "invoice_output.xlsx"):
    """Save invoice data to Excel file, preserving existing rows."""
    data = invoice_data.model_dump()
    items_df = pd.DataFrame(data['items'])
    main_data = pd.DataFrame([{k: v for k, v in data.items() if k != 'items'}])

    if Path(filename).exists():
        try:
            existing_main = pd.read_excel(filename, sheet_name='Invoice', engine='openpyxl')
        except Exception: existing_main = pd.DataFrame()
        try:
            existing_items = pd.read_excel(filename, sheet_name='Items', engine='openpyxl')
        except Exception: existing_items = pd.DataFrame()

        combined_main = pd.concat([existing_main, main_data], ignore_index=True, sort=False) if not existing_main.empty else main_data
        combined_items = pd.concat([existing_items, items_df], ignore_index=True, sort=False) if not existing_items.empty else items_df

        with pd.ExcelWriter(filename, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            combined_main.to_excel(writer, sheet_name='Invoice', index=False)
            combined_items.to_excel(writer, sheet_name='Items', index=False)
    else:
        with pd.ExcelWriter(filename, engine='openpyxl') as writer:
            main_data.to_excel(writer, sheet_name='Invoice', index=False)
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

def extract_invoice_data(file_path: str) -> tuple[InvoiceData, Optional[Dict[str, Any]]]:
    image_base64 = encode_image_to_base64(file_path)
    prompt = """Analyze this invoice image and extract the following data in JSON format:
Rules:
1. The invoice may be in English, Chinese, or other languages. Understand the content in its original language, but extract the data as-is without translation.
2. Do not guess missing values. If a value is missing, use null.
3. Extract: invoice_number, phone, email, address, date, items (name, quantity, unit_price, total_price), subtotal, tax, total_price.
Return ONLY a JSON object."""

    model_name = get_ollama_model()
    
    # Measure Ollama model runtime
    import time
    start_time = time.time()
    
    response = chat(
        model=model_name,
        messages=[{"role": "user", "content": prompt, "images": [image_base64]}],
        format=InvoiceData.model_json_schema(),
        options={"temperature": 0}
    )
    
    end_time = time.time()
    runtime = end_time - start_time
    print(f"Ollama model ({model_name}) runtime: {runtime:.2f} seconds")
    
    image_data = InvoiceData.model_validate_json(response.message.content)
    barcode_data = extract_barcode_qr_data(file_path)
    
    return image_data, barcode_data

def compare_invoice_data(image_data: InvoiceData, barcode_data: Dict[str, Any]) -> bool:
    """Compare image-extracted data with barcode/QR extracted data."""
    if not barcode_data or "barcodes" not in barcode_data:
        return False
    
    for barcode in barcode_data["barcodes"]:
        data = barcode.get("data")
        if isinstance(data, dict):
            # Check if items list matches
            barcode_items = data.get("items", [])
            image_items = [item.model_dump() for item in image_data.items]
            
            # Normalize items for comparison
            barcode_items_normalized = [{k: v for k, v in item.items() if k in ["name", "quantity", "unit_price", "total_price"]} for item in barcode_items]
            image_items_normalized = [{k: v for k, v in item.items() if k in ["name", "quantity", "unit_price", "total_price"]} for item in image_items]
            
            if barcode_items_normalized == image_items_normalized:
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
        self.root.title("AI Invoice Camera & History")
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
        hist_frame = tk.LabelFrame(top_pane, text="Recent Records (Invoice Sheet)")
        top_pane.add(hist_frame, width=600)

        self.tree = ttk.Treeview(hist_frame, columns=("No", "Date", "Address", "Total"), show="headings")
        for col in ("No", "Date", "Address", "Total"):
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
                image_data, barcode_data = extract_invoice_data(path)
                # Check if barcode/QR data matches image data
                if barcode_data and compare_invoice_data(image_data, barcode_data):
                    save_barcode_qr_json(path, barcode_data)
                # Call original editing GUI
                self.open_edit_dialog(image_data, path, barcode_data)
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

        try:
            if file_path.lower().endswith('.pdf'):
                # Handle PDF files
                self.process_pdf(file_path)
            else:
                # Handle image files
                image_data, barcode_data = extract_invoice_data(file_path)
                # Check if barcode/QR data matches image data
                if barcode_data and compare_invoice_data(image_data, barcode_data):
                    save_barcode_qr_json(file_path, barcode_data)
                self.open_edit_dialog(image_data, file_path, barcode_data)
        except Exception as e:
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

            image_data, barcode_data = extract_invoice_data(temp_path)
            # Check if barcode/QR data matches image data
            if barcode_data and compare_invoice_data(image_data, barcode_data):
                save_barcode_qr_json(temp_path, barcode_data)
            self.open_edit_dialog(image_data, temp_path, barcode_data)

            # Clean up temp file
            os.remove(temp_path)

        except Exception as e:
            messagebox.showerror("Error", f"PDF processing failed: {e}")

    def load_history(self):
        for i in self.tree.get_children(): self.tree.delete(i)
        file = "invoice_output.xlsx"
        if Path(file).exists():
            try:
                df = pd.read_excel(file, sheet_name='Invoice')
                for _, row in df.tail(15).iterrows(): # Show last 15
                    self.tree.insert("", 0, values=(row.get('invoice_number',''), row.get('date',''), row.get('address',''), row.get('total_price','')))
            except: pass

    # =========================
    # 4. Detailed Editing GUI (Advanced with per-row edit/delete)
    # =========================
    def open_edit_dialog(self, invoice_data, image_path, barcode_data=None):
        edit_win = tk.Toplevel(self.root)
        edit_win.title("Verify and Edit Invoice Data")
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

        # Image preview thumbnail; open zoomable popup on click
        try:
            img_original = Image.open(image_path)
            thumbnail_size = (220, 220)
            thumbnail = img_original.copy()
            thumbnail.thumbnail(thumbnail_size, Image.LANCZOS)
            thumb_tk = ImageTk.PhotoImage(thumbnail)

            img_label = tk.Label(left_frame, image=thumb_tk, cursor="hand2")
            img_label.image = thumb_tk
            img_label.pack(pady=5)

            def open_image_popup():
                popup = tk.Toplevel(edit_win)
                popup.title("Image Preview")
                popup.geometry("900x700")
                self.center_window(popup, 900, 700)
                popup.transient(edit_win)
                popup.grab_set()

                popup_frame = tk.Frame(popup)
                popup_frame.pack(fill="both", expand=True, padx=10, pady=10)

                popup_fullscreen = {'on': False}
                def toggle_popup_fullscreen():
                    popup_fullscreen['on'] = not popup_fullscreen['on']
                    popup.attributes('-fullscreen', popup_fullscreen['on'])
                    popup_fs_button.config(text='Exit Fullscreen' if popup_fullscreen['on'] else 'Fullscreen')

                popup_control = tk.Frame(popup_frame)
                popup_control.pack(fill='x', pady=(0, 5))
                popup_fs_button = tk.Button(popup_control, text='Fullscreen', command=toggle_popup_fullscreen)
                popup_fs_button.pack(side='right')

                canvas_frame = tk.Frame(popup_frame)
                canvas_frame.pack(fill="both", expand=True)

                hbar = tk.Scrollbar(canvas_frame, orient="horizontal")
                hbar.pack(side="bottom", fill="x")
                vbar = tk.Scrollbar(canvas_frame, orient="vertical")
                vbar.pack(side="right", fill="y")

                popup_canvas = tk.Canvas(canvas_frame, xscrollcommand=hbar.set, yscrollcommand=vbar.set, highlightthickness=0)
                popup_canvas.pack(fill="both", expand=True)
                hbar.config(command=popup_canvas.xview)
                vbar.config(command=popup_canvas.yview)

                popup_zoom = 1.0
                original_width, original_height = img_original.size

                def update_popup(center=None):
                    nonlocal popup_zoom
                    display_width = max(1, int(original_width * popup_zoom))
                    display_height = max(1, int(original_height * popup_zoom))
                    display_img = img_original.resize((display_width, display_height), Image.LANCZOS)
                    popup_tk = ImageTk.PhotoImage(display_img)

                    popup_canvas.delete("all")
                    popup_canvas.create_image(0, 0, anchor="nw", image=popup_tk)
                    popup_canvas.image = popup_tk
                    popup_canvas.config(scrollregion=(0, 0, display_width, display_height))

                    if center is not None:
                        old_x, old_y, screen_x, screen_y, old_zoom = center
                        scale = popup_zoom / old_zoom
                        new_x = old_x * scale
                        new_y = old_y * scale
                        if display_width > 0:
                            popup_canvas.xview_moveto(max(0.0, min(1.0, (new_x - screen_x) / display_width)))
                        if display_height > 0:
                            popup_canvas.yview_moveto(max(0.0, min(1.0, (new_y - screen_y) / display_height)))

                def zoom_in(event=None):
                    nonlocal popup_zoom
                    popup_zoom = min(10.0, popup_zoom * 1.25)
                    update_popup()

                def zoom_out(event=None):
                    nonlocal popup_zoom
                    popup_zoom = max(0.1, popup_zoom * 0.8)
                    update_popup()

                def on_scroll(event):
                    nonlocal popup_zoom
                    delta = 0
                    if hasattr(event, 'delta') and event.delta:
                        delta = event.delta
                    elif event.num == 4:
                        delta = 120
                    elif event.num == 5:
                        delta = -120

                    old_zoom = popup_zoom
                    if delta > 0:
                        popup_zoom = min(10.0, popup_zoom * 1.1)
                    elif delta < 0:
                        popup_zoom = max(0.1, popup_zoom * 0.9)

                    old_x = popup_canvas.canvasx(event.x)
                    old_y = popup_canvas.canvasy(event.y)
                    update_popup((old_x, old_y, event.x, event.y, old_zoom))

                control_frame = tk.Frame(popup_frame)
                control_frame.pack(pady=10)
                tk.Button(control_frame, text="Zoom In", width=10, command=zoom_in).pack(side="left", padx=5)
                tk.Button(control_frame, text="Zoom Out", width=10, command=zoom_out).pack(side="left", padx=5)

                def on_button_press(event):
                    popup_canvas.scan_mark(event.x, event.y)

                def on_drag(event):
                    popup_canvas.scan_dragto(event.x, event.y, gain=1)

                popup_canvas.bind("<MouseWheel>", on_scroll)
                popup_canvas.bind("<Button-4>", on_scroll)
                popup_canvas.bind("<Button-5>", on_scroll)
                popup_canvas.bind("<ButtonPress-1>", on_button_press)
                popup_canvas.bind("<B1-Motion>", on_drag)

                update_popup()

            img_label.bind("<Button-1>", lambda e: open_image_popup())
        except Exception as e:
            tk.Label(left_frame, text=f"Image preview failed:\n{str(e)}", wraplength=220).pack()

        # Barcode/QR data display
        if barcode_data:
            tk.Label(left_frame, text="Barcode/QR Data:", font=("Arial", 10, "bold")).pack(pady=(10, 5))
            barcode_text = scrolledtext.ScrolledText(left_frame, height=12, width=45, wrap=tk.WORD)
            barcode_text.pack(padx=5, pady=5, fill="both", expand=True)
            barcode_json_str = json.dumps(barcode_data, indent=2, ensure_ascii=False)
            barcode_text.insert(tk.END, barcode_json_str)
            barcode_text.config(state=tk.DISABLED)
        else:
            tk.Label(left_frame, text="No Barcode/QR Data Found", wraplength=380).pack()

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
            dialog = tk.Toplevel(edit_win)
            dialog.title("Edit Item")
            dialog.geometry("300x220")
            self.center_window(dialog, 300, 220)

            fields = {}
            labels = ["Name", "Quantity", "Unit Price", "Total Price"]
            for i, label in enumerate(labels):
                tk.Label(dialog, text=label).grid(row=i, column=0, padx=10, pady=5)
                entry = tk.Entry(dialog)
                entry.insert(0, values[i] or "")
                entry.grid(row=i, column=1, padx=10, pady=5)
                fields[label] = entry

            def calculate_total(*args):
                try:
                    qty = float(fields["Quantity"].get() or 0)
                    unit_price = float(fields["Unit Price"].get() or 0)
                    total = qty * unit_price
                    fields["Total Price"].delete(0, tk.END)
                    fields["Total Price"].insert(0, str(total))
                except ValueError:
                    pass

            # Bind events to calculate total when quantity or unit price changes
            fields["Quantity"].bind("<KeyRelease>", calculate_total)
            fields["Unit Price"].bind("<KeyRelease>", calculate_total)

            def save():
                new_values = [fields[label].get() for label in labels] + ["Edit", "Delete"]
                tree.item(item_id, values=new_values)
                dialog.destroy()

            tk.Button(dialog, text="Save", command=save).grid(row=len(labels), column=0, pady=10)
            tk.Button(dialog, text="Cancel", command=dialog.destroy).grid(row=len(labels), column=1, pady=10)

        # Add Item button below
        def add_item_dialog():
            dialog = tk.Toplevel(edit_win)
            dialog.title("Add Item")
            dialog.geometry("300x220")
            self.center_window(dialog, 300, 220)

            fields = {}
            labels = ["Name", "Quantity", "Unit Price", "Total Price"]
            for i, label in enumerate(labels):
                tk.Label(dialog, text=label).grid(row=i, column=0, padx=10, pady=5)
                entry = tk.Entry(dialog)
                entry.grid(row=i, column=1, padx=10, pady=5)
                fields[label] = entry

            def calculate_total(*args):
                try:
                    qty = float(fields["Quantity"].get() or 0)
                    unit_price = float(fields["Unit Price"].get() or 0)
                    total = qty * unit_price
                    fields["Total Price"].delete(0, tk.END)
                    fields["Total Price"].insert(0, str(total))
                except ValueError:
                    pass

            # Bind events to calculate total when quantity or unit price changes
            fields["Quantity"].bind("<KeyRelease>", calculate_total)
            fields["Unit Price"].bind("<KeyRelease>", calculate_total)

            def save():
                new_values = [fields[label].get() for label in labels] + ["Edit", "Delete"]
                tree.insert("", "end", values=new_values)
                dialog.destroy()

            tk.Button(dialog, text="Add", command=save).grid(row=len(labels), column=0, pady=10)
            tk.Button(dialog, text="Cancel", command=dialog.destroy).grid(row=len(labels), column=1, pady=10)

        tk.Button(scrollable_frame, text="Add Item", command=add_item_dialog).grid(row=row, column=0, columnspan=2, pady=10)
        row += 1

        def save_final():
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
                save_to_excel(validated_data)
                self.load_history()
                edit_win.destroy()
                messagebox.showinfo("Saved", "Data appended to Excel!")
            except Exception as e:
                messagebox.showerror("Error", f"Invalid data: {str(e)}")

        def cancel():
            edit_win.destroy()

        tk.Button(scrollable_frame, text="Save", command=save_final).grid(row=row, column=0, pady=10)
        tk.Button(scrollable_frame, text="Cancel", command=cancel).grid(row=row, column=1, pady=10)

    def on_closing(self):
        self.cap.release()
        self.root.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    app = InvoiceApp(root)
    root.protocol("WM_DELETE_WINDOW", app.on_closing)
    root.mainloop()