"""
Project: Fuel Invoice Capture System
CS50 Final Project

This project automates the process of capturing data from photos of fuel
invoices during a work shift. It uses OCR to extract key fields, stores the
information into an Excel template, calculates totals, and finally
generates a report in both Excel and PDF formats.

Modules:
    - modulo_ocr.py: Handles OCR extraction and text cleaning.
    - modulo_excel.py: Manages Excel sheet creation, data saving, and totals calculation.
    - modulo_reporte.py: Finalizes report fields and exports to PDF.

Author: M.S. Mario Espinosa Peniche
"""

from modulo_ocr import extract_field, clean_text
from modulo_excel import get_or_create_sheet, save_to_excel, calculate_totals  
from modulo_reporte import generar_reporte
from PIL import Image     


def main():
    """
    Main entry point for the Fuel Invoice Capture System.
    
    Workflow:
        1. Prompt the user for general shift information 
           (date, shift, responsible person).
        2. Create or copy a new Excel worksheet from a template
           corresponding to the given shift.
        3. Process invoice images provided by the user:
            - Extract predefined fields using OCR. 
            - Save results in the appropriate cells of the worksheet.
        4. Calculates total liters received (both natural and normalized at 20 °C).
        5. Generates a final Excel report and export it to PDF.
    """
    print("==== Fuel Invoice Capture System ====")

    # General shift data
    fecha_descarga, turno, responsable = pedir_datos_turno()

    # Create or copy Excel worksheet
    output_excel = r"C:\Users\mario\Desktop\facturas.xlsx"
    sheet = get_or_create_sheet(output_excel, fecha_descarga, turno)

    # Process invoice images
    procesar_facturas(output_excel, sheet)

    # Calculating totals and generate report
    total_cantidad_natural, total_cantidad_20c = calculate_totals(output_excel, sheet)
    print(f"📊 Total litros al natural: {total_cantidad_natural}")
    print(f"📊 Total litros a 20 °C: {total_cantidad_20c}")

    # Generate final report
    generar_reporte(
        output_excel, 
        sheet,
        turno, 
        responsable, 
        fecha_descarga, 
        total_cantidad_natural, 
        total_cantidad_20c,
        output_dir=r"C:\Users\mario\Desktop"
    )


def pedir_datos_turno():
    """
    Ask the user for general shift information.
    
    Returns:
        tuple: A tuple containing:
            - fecha_descarga (str): Download date in dd/mm/yyyy format.
            - turno (str): Shift range.
            - responsable (str): Person responsible for the shift.
    """
    fecha_descarga = input("Ingrese la fecha de descarga (dd/mm/yyyy): ")
    turno = input("Ingrese el Turno: ")
    responsable = input("Ingrese el responsable: ")
    return fecha_descarga, turno, responsable
    
    
def procesar_facturas(output_excel, sheet):
    """
    Allows the user to input and process multiple invoice images.

    Args:
        output_excel (str): Path to the Excel file where data is stored.
        sheet (str): The worksheet name corresponding to the current shift.
    """

    # OCR search coordinates (x1, y1, x2, y2)
    campos = {
        "numero_embarque": (2000, 500, 2500, 900),
        "fecha_carga": (1650, 350, 2000, 600),
        "clave_equipo": (820, 1180, 1150, 1420),
        "cantidad_natural": (60, 1680, 500, 1900),
        "cantidad_20c": (420, 1700, 820, 1900),
    }

    while True:
        path = input("Enter image path (or press ENTER to finish): ")
        if not path.strip(): 
            break   
        datos = extraer_datos_factura(path, campos)
        print("✅ Invoice processed:", datos)
        save_to_excel(output_excel, datos, sheet)
        
def extraer_datos_factura(path, campos):
    """
    Extracts and cleans data from a single invoice image.

    Args:
        path (str): Path to the invoice image.
        campos (dict): Dictionary of field names and OCR crop coordinates.

    Returns:
        dict: Extracted and cleaned values for each requiered field.
    """
    image = Image.open(path).convert("RGB")
    datos = {}
    for campo, coords in campos.items():
        raw_value = extract_field(image, coords)
        datos[campo] = clean_text(campo, raw_value)

    return datos
    

if __name__ == "__main__":
    main()