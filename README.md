# Fuel Invoice Capture System
#### Video Demo: <https://youtu.be/6S17Y5Hzoqo>

#### Description
This project automates the process of capturing data from **fuel invoice images** during work shifts.  
It uses **OCR** (Optical Character Recognition) to extract predefined fields from invoices, stores the results in a predefined **Excel template**, calculates totals, and finally generates a **report in both Excel and PDF formats**.

The system was designed as a **CS50P Final Project** and demonstrates skills in:
- Image processing with Python
- Text extraction via OCR
- Excel automation
- Report generation
- Automated testing with `pytest`

---

## Features
- Extracts the following fields from invoice images:
  - Shipment number (`numero_embarque`)
  - Load date (`fecha_carga`)
  - Equipment code (`clave_equipo`)
  - Fuel liters (natural and adjusted to 20 °C)
- Saves data into an Excel template with structured rows.
- Automatically calculates totals at the end of the shift.
- Generates a summary report (Excel + PDF).
- Includes unit and integration tests.

---

## How to Run

1. Clone this repository and move into the project folder:
   ```bash
   git clone <https://github.com/MarioEsPe/Fuel_Invoice_Capture_System>
   cd ProyectoCS50
2. (Optional but recommended) Create and activate a virtual environment:
   python -m venv .venv
   source .venv/bin/activate    # Mac/Linux
   .venv\Scripts\activate       # Windows
3. Install dependencies:
   pip install -r requirements.txt
4. Run the program:
   python project.py

## Example Run
Input:
==== Fuel Invoice Capture System ====
Enter download date (dd/mm/yyyy): 23/08/2025
Enter shift: 07:00-15:00
Enter responsible: Ing. Mario Espinosa
Enter the path to the image (or ENTER to finish): Factura01-3815.jpg
Output:
Invoice processed: {'numero_embarque': '257288', 'fecha_carga': '01/06/2024',
 'clave_equipo': 'FZS3815', 'cantidad_natural': '31999.000',
 'cantidad_20c': '31688.000'}

Total liters (natural): 31999.0
Total liters (20 °C): 31688.0
Report successfully generated: facturas/23-08-2025_Turno0700-1500.pdf

## Tests

This project includes unit tests and integration tests using pytest.

Run all tests with:
pytest

Example result:

collected 4 items

test_project.py ....                                                                 [100%]

## Requirements

Python 3.10+

Dependencies listed in requirements.txt:

pillow

pytesseract

openpyxl

reportlab

pytest


## Author

M.S. Mario Espinosa Peniche
🌐 LinkedIn: www.linkedin.com/in/marioespinosachemicalengineer
📧 Email: mario_espinosap@yahoo.com
Final Project for CS50’s Introduction to Programming with Python

## License

This project is licensed under the MIT License – free to use, modify, and distribute.

## Acknowledgments

This project was created as part of Harvard’s CS50P – Introduction to Programming with Python.
Special thanks to the CS50 team for their guidance and excellent course material.
