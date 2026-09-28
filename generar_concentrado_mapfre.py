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
        num_format_str='DD/MM/YYYY'
    )

    # Estilos confirmados contra la plantilla real (CONCENTRADO_ALTAS__1_.xls)
    # — NO son inventados: se sacaron inspeccionando el formato exacto de
    # cada fila con xlrd(formatting_info=True). El portal de Mapfre es
    # estricto con que el layout coincida, así que aquí no se adorna nada
    # que la plantilla original no tuviera.
    #   Fila 1 (encabezados):        sin relleno, sin negrita, sin bordes.
    #   Fila 2 (spec. de anchos):    relleno amarillo pálido RGB(255,255,153),
    #                                bordes delgados en las 4 direcciones.
    #   Filas de datos:              sin relleno, sin bordes.
    xlwt.add_palette_colour("amarillo_spec", 0x21)
    estilo_encabezado = xlwt.XFStyle()  # sin ningún formato, igual que el original
    estilo_fila_spec = xlwt.easyxf(
        'pattern: pattern solid, fore_color amarillo_spec;'
        'borders: left thin, right thin, top thin, bottom thin;'
    )
    estilo_celda = xlwt.XFStyle()  # sin formato, igual que el original
    estilo_celda_texto = xlwt.XFStyle()  # sin formato, igual que el original

    ws.set_panes_frozen(True)
    ws.set_horz_split_pos(1)  # Congela la fila de encabezados al hacer scroll

    for col, encabezado in enumerate(HEADERS):
        ws.write(0, col, encabezado, estilo_encabezado)
        ws.col(col).width = 256 * max(14, min(len(encabezado) + 2, 28))

    # Fila 2 (índice 1) — fila FIJA de especificación de anchos de columna
    # que el portal de Mapfre exige tal cual, con estos valores exactos
    # confirmados por Carlos. Si se omite o se sobreescribe con datos, el
    # portal rechaza el archivo por no parecerse al layout esperado.
    FILA_ESPECIFICACION_ANCHOS = [
        13, 5, 3, 20, 13, 30, 30, 30, 2, 8, 2, 1, 4, 3, 1, 8, 8, 8, 8, 8,
        8, 5, 5, 7, 7, 7, 7, 7, 1, 7, 30, 30, 8, 8, 30, 30, 30, 30, 30,
    ]
    wb.set_colour_RGB(0x21, 255, 255, 153)
    for col, valor in enumerate(FILA_ESPECIFICACION_ANCHOS):
        ws.write(1, col, valor, estilo_fila_spec)

    hoy = date.today()
    # Los datos reales empiezan en la fila 3 (índice 2) — la fila 2 ya la
    # ocupa la especificación de anchos, no debe pisarse con datos.
    for fila_idx, a in enumerate(asegurados, start=2):
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
