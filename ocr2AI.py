import json
import os
import time
from pathlib import Path
from typing import List, Optional

# Disable oneDNN optimizations that can trigger unsupported runtime attributes on some Windows Paddle installations.
os.environ.setdefault("PADDLE_DISABLE_ONE_DNN", "1")
os.environ.setdefault("FLAGS_use_gpu", "0")

from pydantic import BaseModel
import ollama
from ollama import chat
from ollama._types import ResponseError
from paddleocr import PaddleOCR
from pdf2image import convert_from_path


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
# 2. Convert PDF to Images
# =========================

def pdf_to_images(pdf_path: str):
    pages = convert_from_path(pdf_path)
    image_paths = []

    output_dir = Path("temp_pages")
    output_dir.mkdir(exist_ok=True)

    for i, page in enumerate(pages):
        img_path = output_dir / f"page_{i + 1}.jpg"
        page.save(img_path, "JPEG")
        image_paths.append(str(img_path))

    return image_paths


# =========================
# 3. OCR Image to Text
# =========================

def save_ocr_log(ocr_text: str, log_path: str = "ocr_output.log") -> None:
    with open(log_path, "w", encoding="utf-8") as log_file:
        log_file.write(ocr_text)
    print(f"OCR text logged to {log_path}")


def extract_text_from_ocr_result(result) -> List[str]:
    if not result:
        return []

    ocr_block = result[0]
    if isinstance(ocr_block, dict) and "rec_texts" in ocr_block:
        return ocr_block["rec_texts"]

    # fallback for older PaddleOCR output shape
    extracted = []
    try:
        for line in ocr_block:
            if isinstance(line, list) and len(line) > 1:
                extracted.append(line[1][0])
            elif isinstance(line, (tuple, list)) and len(line) == 2:
                extracted.append(str(line[1]))
    except Exception:
        pass

    return extracted


def ocr_images(image_paths: List[str]) -> str:
    ocr = PaddleOCR(use_textline_orientation=True, lang="en")

    full_text = ""

    for image_path in image_paths:
        result = ocr.predict(image_path)
        text_lines = extract_text_from_ocr_result(result)

        full_text += f"\n--- OCR FROM {image_path} ---\n"
        if text_lines:
            full_text += "\n".join(text_lines) + "\n"
        else:
            full_text += "[no recognized text]\n"

    save_ocr_log(full_text)
    return full_text


# =========================
# 4. Extract Invoice Data
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


def extract_invoice_data(ocr_text: str) -> InvoiceData:
    prompt = f"""
Extract invoice data from the image below.

Return only valid JSON.

Rules:
- Do not guess missing values.
- If missing, use null.
- Extract invoice number, phone, email, address, and date.
- Extract all purchased items.
- For each item, extract name, quantity, unit_price, total_price.
- Extract subtotal, tax, total_price, and currency if available.

OCR text:
{ocr_text}
"""
    model_name = get_ollama_model()
    print(f"Using Ollama model: {model_name}")

    try:
        response = chat(
            model=model_name,
            messages=[
                {"role": "user", "content": prompt}
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

    if file_path.suffix.lower() == ".pdf":
        image_paths = pdf_to_images(str(file_path))
    else:
        image_paths = [str(file_path)]

    print("Running OCR...")
    ocr_text = ocr_images(image_paths)

    print("Extracting invoice data using Ollama...")
    invoice_data = extract_invoice_data(ocr_text)

    output = invoice_data.model_dump()

    with open("invoice_output.json", "w", encoding="utf-8") as f:
        json.dump(output, f, indent=2, ensure_ascii=False)

    print("\nExtracted Data:")
    print(json.dumps(output, indent=2, ensure_ascii=False))

    print("\nSaved to invoice_output.json")


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