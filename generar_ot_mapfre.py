# -*- coding: utf-8 -*-
"""
Llena la "Orden de Trabajo" de Mapfre (basada en OT_MAPFRE.xls) y la
convierte a PDF, que es el formato que el portal exige (confirmado por
Carlos — el PDF original NO tiene campos rellenables, es un documento
plano, así que reconstruimos desde el Excel y convertimos).

⚠️ A diferencia de Plan Seguro, esta Orden de Trabajo NO lleva el detalle
de cada persona — solo los datos de la póliza/empresa (confirmado:
"FAVOR DE DAR DE ALTA ASEGURADA" es genérico). El detalle de las personas
va completo en el Excel "Concentrado de Altas" (archivo separado).

Requiere LibreOffice instalado en el servidor (soffice --headless) para
la conversión xls -> pdf.
"""

import os
import shutil
import subprocess
import tempfile
from datetime import date

import xlrd
import xlwt
from xlutils.copy import copy as xl_copy

TEMPLATE_PATH = "OT_MAPFRE_template.xls"  # copiar aquí la plantilla real (Hoja1)

# Datos fijos por grupo de pólizas — confirmados por Carlos.
DATOS_FIJOS_POR_POLIZA = {
    # Las 3 pólizas de Arrenda Max comparten los mismos datos fijos,
    # solo cambia el número de póliza.
    "2612500003880": {
        "razon_social": "GRUPO ARRENDA MAX DE MORELOS S.A. de C.V.",
        "domicilio_col": "CHIPITLAN",
        "domicilio_estado": "MORELOS",
        "vigencia_desde": (1, 11, 2025),   # (DD, MM, AA)
        "vigencia_hasta": (1, 11, 2026),
        "instrucciones": "FAVOR DE DAR DE ALTA ASEGURADA",
    },
    "2612500003895": {
        "razon_social": "GRUPO ARRENDA MAX DE MORELOS S.A. de C.V.",
        "domicilio_col": "CHIPITLAN",
        "domicilio_estado": "MORELOS",
        "vigencia_desde": (1, 11, 2025),
        "vigencia_hasta": (1, 11, 2026),
        "instrucciones": "FAVOR DE DAR DE ALTA ASEGURADA",
    },
    "2612500003973": {
        "razon_social": "GRUPO ARRENDA MAX DE MORELOS S.A. de C.V.",
        "domicilio_col": "CHIPITLAN",
        "domicilio_estado": "MORELOS",
        "vigencia_desde": (1, 11, 2025),
        "vigencia_hasta": (1, 11, 2026),
        "instrucciones": "FAVOR DE DAR DE ALTA ASEGURADA",
    },
    "2612600000892": {  # CYSE/SNAC
        "razon_social": "GRUPO CONSULTOR Y ASESOR CYSE S.A.\nde C.V.",
        "domicilio_col": "CHIPITLAN",
        "domicilio_estado": "MORELOS",
        "vigencia_desde": (20, 2, 2026),
        "vigencia_hasta": (20, 2, 2027),
        "instrucciones": "FAVOR DE DAR DE ALTA  A LOS  ASEGURADOS.",
    },
}

AGENTE_CLAVE = "33910"
AGENTE_NOMBRE = "FRANCISCO JAVIER PORRAS VELAZQUEZ"


def generar_orden_trabajo_mapfre(
    no_poliza: str,
    fecha_solicitud: date = None,
    template_path: str = TEMPLATE_PATH,
    salida_pdf: str = "/tmp/orden_trabajo_mapfre.pdf",
) -> str:
    datos = DATOS_FIJOS_POR_POLIZA.get(no_poliza)
    if not datos:
        raise ValueError(
            f"No hay datos fijos confirmados para la póliza {no_poliza!r}. "
            f"Pólizas soportadas: {list(DATOS_FIJOS_POR_POLIZA)}"
        )

    fecha_solicitud = fecha_solicitud or date.today()

    rb = xlrd.open_workbook(template_path, formatting_info=True)
    wb = xl_copy(rb)
    ws = wb.get_sheet(0)  # Hoja1

    formato_fecha = xlwt.easyxf(num_format_str='DD/MM/YYYY')

    # Fecha de solicitud (celda superior derecha, confirmada en el original)
    ws.write(0, 10, fecha_solicitud, formato_fecha)

    # No. Póliza
    ws.write(1, 10, no_poliza)

    # Vigencia (fila 7 = índice 7, columnas 0-5: DD MM AA DD MM AA)
    d1, m1, a1 = datos["vigencia_desde"]
    d2, m2, a2 = datos["vigencia_hasta"]
    ws.write(7, 0, d1)
    ws.write(7, 1, m1)
    ws.write(7, 2, a1)
    ws.write(7, 3, d2)
    ws.write(7, 4, m2)
    ws.write(7, 5, a2)

    # Nombre (razón social)
    ws.write(11, 2, datos["razon_social"])

    # Domicilio de pago
    ws.write(14, 10, datos["domicilio_col"])
    ws.write(15, 6, datos["domicilio_estado"])

    # Agente
    ws.write(18, 2, AGENTE_CLAVE)
    ws.write(18, 6, AGENTE_NOMBRE)

    # Instrucciones
    ws.write(25, 0, datos["instrucciones"])

    with tempfile.TemporaryDirectory() as tmpdir:
        xls_temp = os.path.join(tmpdir, "orden_trabajo.xls")
        wb.save(xls_temp)

        # Convertir a PDF con LibreOffice headless — SinglePageSheets fuerza
        # que la hoja completa quepa en UNA sola página (sin esto, se parte
        # en varias páginas de forma rara porque la hoja es más ancha que
        # el área imprimible normal).
        resultado = subprocess.run(
            [
                "soffice", "--headless", "--convert-to",
                'pdf:calc_pdf_Export:{"SinglePageSheets":{"type":"boolean","value":"true"}}',
                "--outdir", tmpdir, xls_temp,
            ],
            capture_output=True, text=True, timeout=60,
        )
        if resultado.returncode != 0:
            raise RuntimeError(f"Error convirtiendo a PDF con LibreOffice: {resultado.stderr}")

        pdf_generado = os.path.join(tmpdir, "orden_trabajo.pdf")
        if not os.path.exists(pdf_generado):
            raise RuntimeError(
                f"LibreOffice no generó el PDF esperado. Salida: {resultado.stdout} {resultado.stderr}"
            )

        # La plantilla original trae una segunda hoja ("ENDOSO", un ejemplo
        # viejo sin relación con nuestros datos) que también se exporta —
        # nos quedamos solo con la página 1, que es la que sí llenamos.
        from pypdf import PdfReader, PdfWriter
        lector = PdfReader(pdf_generado)
        escritor = PdfWriter()
        escritor.add_page(lector.pages[0])
        with open(salida_pdf, "wb") as f:
            escritor.write(f)

    return salida_pdf


if __name__ == "__main__":
    ruta = generar_orden_trabajo_mapfre(no_poliza="2612600000892", salida_pdf="/tmp/OT_MAPFRE_TEST.pdf")
    print("Generado:", ruta)
