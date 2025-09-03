"""
Unit tests for project.py (CS50 Final Project)
"""

from project import pedir_datos_turno, extraer_datos_factura, procesar_facturas, get_or_create_sheet
from unittest.mock import patch, MagicMock
from PIL import Image, ImageDraw
import builtins


def test_pedir_datos_turno(monkeypatch):
    """
    Test that pedir_datos_turno correctly returns user inputs as a tuple.
    """
    inputs = iter(["23/08/2025", "07:00-15:00", "Ing. Mario Espinosa"])
    monkeypatch.setattr("builtins.input", lambda _: next(inputs))

    fecha, turno, responsable = pedir_datos_turno()
    assert fecha == "23/08/2025"
    assert turno == "07:00-15:00"
    assert responsable == "Ing. Mario Espinosa"


def test_extraer_datos_factura_integration(tmp_path):
    """Integration test: simulate invoice image and verify OCR extracts dictionary with expected keys."""
    img_path = tmp_path / "fake_invoice.jpg"
    img = Image.new("RGB", (3000, 2000), "white")
    d = ImageDraw.Draw(img)
    d.text((2100, 600), "257288", fill="black")   # numero_embarque
    d.text((1700, 500), "01/06/2024", fill="black")  # fecha_carga
    d.text((850, 1300), "FZS3815", fill="black")  # clave_equipo
    d.text((100, 1750), "31999.000", fill="black")  # cantidad_natural
    d.text((500, 1750), "31688.000", fill="black")  # cantidad_20c
    img.save(img_path)

    campos = {
        "numero_embarque": (2000, 500, 2500, 900),
        "fecha_carga": (1650, 350, 2000, 600),
        "clave_equipo": (820, 1180, 1150, 1420),
        "cantidad_natural": (60, 1680, 500, 1900),
        "cantidad_20c": (420, 1700, 820, 1900),
    }

    datos = extraer_datos_factura(str(img_path), campos)

    
    assert set(datos.keys()) == set(campos.keys())
    assert datos["numero_embarque"] is not None
    assert datos["fecha_carga"] is not None


def test_extraer_datos_factura_mock():
    """Unit test: mock both OCR and Image.open to isolate logic."""
    campos = {
        "numero_embarque": (0, 0, 0, 0),
        "fecha_carga": (0, 0, 0, 0),
        "clave_equipo": (0, 0, 0, 0),
        "cantidad_natural": (0, 0, 0, 0),
        "cantidad_20c": (0, 0, 0, 0),
    }

    fake_img = MagicMock()

    with patch("project.extract_field", side_effect=[
        "257288", "01/06/2024", "FZS3815", "31999.000", "31688.000"
    ]), patch("project.Image.open", return_value=fake_img):
        datos = extraer_datos_factura("fake_path", campos)

    
    assert datos["numero_embarque"] == "257288"
    assert datos["fecha_carga"] == "01/06/2024"
    assert datos["clave_equipo"] == "FZS3815"
    assert datos["cantidad_natural"] == "31999.000"
    assert datos["cantidad_20c"] == "31688.000"


def test_procesar_facturas(monkeypatch, tmp_path):
    """Test procesar_facturas with mocked user input and save_to_excel."""
    
    inputs = iter([
        str(tmp_path / "fake_invoice.jpg"),  
        ""  
    ])
    monkeypatch.setattr(builtins, "input", lambda _: next(inputs))

    
    from PIL import Image, ImageDraw
    img = Image.new("RGB", (3000, 2000), "white")
    d = ImageDraw.Draw(img)
    d.text((2100, 600), "257288", fill="black")
    d.text((1700, 500), "01/06/2024", fill="black")
    d.text((850, 1300), "FZS3815", fill="black")
    d.text((100, 1750), "31999.000", fill="black")
    d.text((500, 1750), "31688.000", fill="black")
    img_path = tmp_path / "fake_invoice.jpg"
    img.save(img_path)

    
    with patch("project.save_to_excel") as mock_save:
        procesar_facturas("fake_excel.xlsx", "FakeSheet")
        mock_save.assert_called_once()  