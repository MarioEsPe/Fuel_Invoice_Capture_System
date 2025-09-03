"""
Module: modulo_ocr.py
Handles OCR (Optical Character Recognition) and text cleaning.

Functions:
    - extract_field(image, coords, lang="spa"): Extracts text from a cropped image region. 
    - clean_text(field_name, text): Cleans and normalizes OCR output for a given field.
"""


import pytesseract
from PIL import Image
import re

# Path to the Tesseract executable (adjust if necessary)
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

def extract_field(image, coords, lang="spa"):
    """
    Extract text from a specific region of an image using OCR.

    Args:
        - image (PIL.Image): The source image.
        - coords (tuple): (x1, y1, x2, y2) defining the crop region.
        - lang (str): The language for OCR (default is "spa").

    Returns:
        - str: Extracted raw text.
    """
    crop_img = image.crop(coords)
    text = pytesseract.image_to_string(crop_img, lang=lang, config="--psm 6")
    return text.strip()


def clean_text(field_name, text):
    """
    Clean and normalize OCR output depending on the field type.
    
    Args:
        - field_name (str): The name of the field being cleaned.
        - text (str): Raw OCR text.

    Returns:
        - str: Cleaned and normalized value, or empty string if not found.
    """
    text = text.replace("\n", " ").strip()

    if field_name == "numero_embarque":
        match = re.search(r"\d{6}", text)
        return match.group(0) if match else ""

    elif field_name == "fecha_carga":
        match = re.search(r"\d{2}/\d{2}/\d{4}", text)
        return match.group(0) if match else ""

    elif field_name == "clave_equipo":
        # Strict pattern: FZN + 4 digits
        match = re.search(r"FZN\d{4}", text)
        if match:
            return match.group(0)

        # Fallback: any 3 letters + 4 digits
        match = re.search(r"[A-Z]{3}\d{4}", text)
        return match.group(0) if match else ""

    elif field_name in ["cantidad_natural", "cantidad_20c"]:
        # Match large numbers with decimals, e.g. 31,999.000
        match = re.search(r"\d{2,3},?\d{3}\.\d{3}", text)
        if match:
            return match.group(0).replace(",", "")  # Normalize by removing comma
        else:
            return ""

    return text
