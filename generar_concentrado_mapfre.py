# -*- coding: utf-8 -*-
"""
Genera el Excel "Concentrado de Altas" de Mapfre, en el formato real
confirmado (CONCENTRADO_ALTAS.xls que Carlos compartió, con el ejemplo
real de PAULINA PATIÑO GONZALEZ ARAGON).

⚠️ Formato .xls VIEJO (Excel 97-2003), no .xlsx — se usa xlwt para
mantener compatibilidad exacta con lo que espera el portal de Mapfre.

Estructura confirmada — 39 columnas en total, pero SOLO 9 se llenan
realmente para un alta normal (el resto se deja vacío):
    Número de póliza
    Apellido Paterno del Asegurado
    Apellido Materno del Asegurado
    Nombre(s) del Asegurado
    Cód.del Parentesco   (numérico: 4=Titular, 1=Cónyuge, 2=Hijo(a))
    Fecha de Nacimiento
    Edad
    Sexo                 (una sola letra: M/F)
    F.de Inicio de vig. del asegurado  (= fecha del alta, confirmado por Carlos)

Las otras 30 columnas (Número de Familia, Estatura, Peso, Exclusiones,
Sueldo, país de origen, etc.) se dejan vacías — no se usan para un alta
estándar.
"""

import xlwt
from datetime import datetime, date

HEADERS = [
    'Número de póliza', 'Número de Familia', 'Tipo de Doc.', 'Código del Documento',
    'Número de empleado', 'Apellido Paterno del Asegurado', 'Apellido Materno del Asegurado',
    'Nombre(s) del Asegurado', 'Cód.del Parentesco', 'Fecha de Nacimiento', 'Edad', 'Sexo',
    'Estatura', 'Peso', 'Calcula Prima R/N', 'F.de Ant. Tepeyac', 'F. de Rec. de Antigüedad',
    'F.de Ant. para Maternidad', 'F.de Ant. Cob. Internacional', 'F.de Ant. para Enf. Cat.',
    'F.de Inicio de vig. del asegurado', 'Cód.de Ocupación', 'Cód. de Deporte', 'Exclusión 1',
    'Exclusión 2', 'Exclusión 3', 'Exclusión 4', 'Exclusión 5', 'Cob.2037 (Maternidad)',
    'Sueldo', 'País de origen', 'Motivo del viaje', 'Inicio del periodo de estancia',
    'Fin del periodo de estancia', 'Compañía aseguradora', 'Suma Asegurada rec.',
    'Deducible reconocido', 'Porcentaje de coaseguro reconocido', 'Limite de coaseguro rec.',
]

# Confirmado por Carlos, coincide exactamente con las 3 opciones de
# x_studio_parentesco en Odoo — nunca inventar un código para un valor
# que no esté aquí.
CODIGO_PARENTESCO = {
    'Titular': 4,
    'Conyuge': 1,
    'Hijo(a)': 2,
}


def _calcular_edad(fecha_nac: date, referencia: date = None) -> int:
    referencia = referencia or date.today()
    edad = referencia.year - fecha_nac.year
    if (referencia.month, referencia.day) < (fecha_nac.month, fecha_nac.day):
        edad -= 1
    return edad


def _parse_fecha(fecha):
    if isinstance(fecha, (datetime, date)):
        return fecha if isinstance(fecha, date) and not isinstance(fecha, datetime) else fecha.date()
    return datetime.strptime(fecha, "%Y-%m-%d").date()


def generar_concentrado_altas_mapfre(asegurados: list[dict], no_poliza: str, salida: str) -> str:
    """
    asegurados: lista de dicts con:
        apellido_paterno, apellido_materno, nombre, parentesco
        ("Titular"/"Conyuge"/"Hijo(a)" — debe venir de x_studio_parentesco,
        sin inventar un valor si viene vacío), fecha_nacimiento (YYYY-MM-DD),
        sexo ("M"/"F")
    no_poliza: ej. "2612600000892"
    salida: ruta del archivo .xls a generar
    """
    wb = xlwt.Workbook()
    ws = wb.add_sheet('G.M.M.')

    formato_fecha = xlwt.XFStyle()
    formato_fecha.num_format_str = 'DD/MM/YYYY'

    for col, encabezado in enumerate(HEADERS):
        ws.write(0, col, encabezado)

    hoy = date.today()
    for fila_idx, a in enumerate(asegurados, start=1):
        fecha_nac = _parse_fecha(a["fecha_nacimiento"])
        codigo_parentesco = CODIGO_PARENTESCO.get(a.get("parentesco", ""))
        if codigo_parentesco is None:
            raise ValueError(
                f"Parentesco {a.get('parentesco')!r} no reconocido — debe ser "
                f"exactamente uno de {list(CODIGO_PARENTESCO)} (viene de "
                f"x_studio_parentesco en Odoo)."
            )

        ws.write(fila_idx, 0, no_poliza)  # Número de póliza
        ws.write(fila_idx, 5, a["apellido_paterno"])
        ws.write(fila_idx, 6, a.get("apellido_materno", ""))
        ws.write(fila_idx, 7, a["nombre"])
        ws.write(fila_idx, 8, codigo_parentesco)
        ws.write(fila_idx, 9, fecha_nac, formato_fecha)
        ws.write(fila_idx, 10, _calcular_edad(fecha_nac))
        ws.write(fila_idx, 11, a["sexo"])  # 'M' o 'F'
        ws.write(fila_idx, 20, hoy, formato_fecha)  # F.de Inicio de vig. del asegurado

    wb.save(salida)
    return salida


if __name__ == "__main__":
    ejemplo = [{
        "apellido_paterno": "PATIÑO",
        "apellido_materno": "GONZALEZ ARAGON",
        "nombre": "PAULINA",
        "parentesco": "Titular",
        "fecha_nacimiento": "1997-04-02",
        "sexo": "F",
    }]
    ruta = generar_concentrado_altas_mapfre(ejemplo, no_poliza="2612600000892", salida="/tmp/CONCENTRADO_TEST.xls")
    print("Generado:", ruta)
