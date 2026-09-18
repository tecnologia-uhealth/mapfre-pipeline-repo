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

    formato_fecha = xlwt.easyxf(
        'font: bold off; borders: left thin, right thin, top thin, bottom thin;'
        'align: horiz center;',
        num_format_str='DD/MM/YYYY'
    )

    # Encabezado: fondo azul oscuro, letra blanca en negrita, bordes —
    # solo en las 9 columnas que de verdad se usan, para que resalten;
    # el resto de los encabezados (sin usar) quedan sin formato.
    COLUMNAS_USADAS = {0, 5, 6, 7, 8, 9, 10, 11, 20}
    estilo_encabezado_usado = xlwt.easyxf(
        'font: bold on, color white;'
        'pattern: pattern solid, fore_color dark_blue;'
        'borders: left thin, right thin, top thin, bottom thin;'
        'align: horiz center, wrap on;'
    )
    estilo_encabezado_sin_usar = xlwt.easyxf(
        'font: bold on, color gray50;'
        'pattern: pattern solid, fore_color gray25;'
        'borders: left thin, right thin, top thin, bottom thin;'
        'align: horiz center, wrap on;'
    )
    estilo_celda = xlwt.easyxf(
        'borders: left thin, right thin, top thin, bottom thin;'
        'align: horiz center;'
    )
    estilo_celda_texto = xlwt.easyxf(
        'borders: left thin, right thin, top thin, bottom thin;'
        'align: horiz left;'
    )

    ws.set_panes_frozen(True)
    ws.set_horz_split_pos(1)  # Congela la fila de encabezados al hacer scroll

    for col, encabezado in enumerate(HEADERS):
        estilo = estilo_encabezado_usado if col in COLUMNAS_USADAS else estilo_encabezado_sin_usar
        ws.write(0, col, encabezado, estilo)
        ws.col(col).width = 256 * max(14, min(len(encabezado) + 2, 28))
    ws.row(0).height = 700  # más alta para que quepan encabezados en 2 líneas

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

        ws.write(fila_idx, 0, no_poliza, estilo_celda)  # Número de póliza
        ws.write(fila_idx, 5, a["apellido_paterno"], estilo_celda_texto)
        ws.write(fila_idx, 6, a.get("apellido_materno", ""), estilo_celda_texto)
        ws.write(fila_idx, 7, a["nombre"], estilo_celda_texto)
        ws.write(fila_idx, 8, codigo_parentesco, estilo_celda)
        ws.write(fila_idx, 9, fecha_nac, formato_fecha)
        ws.write(fila_idx, 10, _calcular_edad(fecha_nac), estilo_celda)
        ws.write(fila_idx, 11, a["sexo"], estilo_celda)  # 'M' o 'F'
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
