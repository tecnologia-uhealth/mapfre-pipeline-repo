# -*- coding: utf-8 -*-
"""
RPA de Mapfre (portal ZonAliados) — reconstruido a partir de una
grabación REAL con `playwright codegen` (rpa_mapfre_grabado.py, sept 2026).
Selectores confirmados contra el portal real.

⚠️ Alta y Baja ya están confirmadas con grabaciones reales. Diferencia
clave: Baja solo sube el Concentrado de Altas (1 archivo) — no lleva
Orden de Trabajo, a diferencia de Alta que sube ambos.

FLUJO CONFIRMADO:
    1. Login (usuario/contraseña) + aceptar cookies
    2. Cerrar banner/aviso post-login (si aparece)
    3. GAMA → Crear Solicitud
    4. Sector = AyE, Solicitud = Endoso → Continuar
    5. Modal con folio (temprano) → Aceptar
    6. Tipo de endoso = ALTA/BAJA DE ASEGURADOS, Núm. de póliza = <poliza>
       (el resto de los datos de la póliza se auto-rellenan solos)
    7. Subir Concentrado de Altas → diálogo nativo se descarta → Adjuntar
    8. [SOLO ALTA] Subir Orden de Trabajo → diálogo nativo se descarta → Adjuntar
    9. Escribir comentario → Agregar comentario
    10. Continuar → modal "Número de folio generado: XXXXX" → Aceptar

El folio SÍ es visible en pantalla (a diferencia de Plan Seguro) — se
captura del modal final, con el mismo valor mostrado desde el paso 5.
"""

import os
import re
import logging
from playwright.sync_api import sync_playwright

log = logging.getLogger("rpa_mapfre")

MAPFRE_USER = os.environ.get("MAPFRE_USER")
MAPFRE_PASS = os.environ.get("MAPFRE_PASS")

PORTAL_LOGIN_URL = "https://zonaliados.mapfre.com.mx/zonaliados/loginr.aspx"

# Confirmado con 2 grabaciones reales — alta y baja.
VALORES_TIPO_ENDOSO = {
    "alta": "ALTA DE ASEGURADOS",
    "baja": "BAJA DE ASEGURADOS",
}


class TipoMovimientoNoConfirmadoError(Exception):
    """Se lanza si se pide un tipo_movimiento distinto de 'alta'/'baja'
    (los únicos 2 confirmados con grabaciones reales)."""
    pass


def emitir_movimiento_mapfre(
    no_poliza: str,
    tipo_movimiento: str,  # "alta" o "baja" — ambos confirmados
    concentrado_altas_path: str,
    orden_trabajo_path: str | None = None,  # solo obligatorio para ALTA
    comentario: str = None,
    headless: bool = True,
) -> str:
    """Regresa el folio generado por el portal (ej. '226040190892964')."""
    if tipo_movimiento not in VALORES_TIPO_ENDOSO:
        raise TipoMovimientoNoConfirmadoError(
            f"El tipo de movimiento '{tipo_movimiento}' no está confirmado. "
            f"Solo 'alta' está confirmado con una grabación real."
        )

    if not MAPFRE_USER or not MAPFRE_PASS:
        raise RuntimeError("Faltan las variables de entorno MAPFRE_USER / MAPFRE_PASS")

    if tipo_movimiento == "alta" and not orden_trabajo_path:
        raise ValueError("orden_trabajo_path es obligatorio para ALTA.")

    if comentario is None:
        comentario = (
            "Hola, favor de dar de alta, gracias." if tipo_movimiento == "alta"
            else "Hola favor de dar de baja, gracias."
        )

    folio = None

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=headless)
        context = browser.new_context()
        page = context.new_page()

        try:
            # --- 1. Login ---
            page.goto(PORTAL_LOGIN_URL)
            page.get_by_role("button", name="Aceptar").click()  # aviso de cookies

            page.get_by_role("textbox", name="Usuario").click()
            page.get_by_role("textbox", name="Usuario").fill(MAPFRE_USER)
            page.get_by_role("textbox", name="Contraseña").click()
            page.get_by_role("textbox", name="Contraseña").fill(MAPFRE_PASS)
            page.get_by_role("button", name="Ingresar").click()

            # Banner/aviso post-login — no siempre aparece
            try:
                page.get_by_role("button", name="Cerrar").click(timeout=8000)
            except Exception:
                log.info("No apareció el banner post-login (o no había que cerrarlo).")

            # --- 2. GAMA -> Crear Solicitud ---
            page.get_by_role("link", name="• GAMA").click()
            page.get_by_role("link", name="Crear Solicitud").click()

            # --- 3. Sector + Solicitud ---
            page.locator("#ddlSector-button").click()
            page.get_by_role("option", name="AyE").click()

            page.locator("#ddlSolicitud-button").click()
            page.get_by_role("option", name="Endoso").click()

            page.get_by_role("button", name="Continuar").click()

            # --- 4. Modal temprano con folio — solo cerrar ---
            page.get_by_role("button", name="Aceptar").click(timeout=15000)

            # --- 5. Tipo de endoso + Núm. de póliza ---
            page.locator("#ddlTipoEndoso-button").click()
            page.get_by_role("option", name=VALORES_TIPO_ENDOSO[tipo_movimiento]).click()

            page.locator("#txtPoliza").click()
            page.locator("#txtPoliza").fill(no_poliza)
            # Alejar el foco del campo para que dispare el auto-relleno del
            # resto de los datos de la póliza (confirmado en la grabación
            # y en el manual: "Una vez poniendo el num de poliza se rellena
            # en automático los demás datos")
            page.locator("div").filter(has_text="* Núm. de póliza:").first.click()

            # --- 6. Subir Concentrado de Altas (siempre, alta y baja) ---
            page.get_by_role("button", name="Choose File").set_input_files(concentrado_altas_path)
            page.once("dialog", lambda dialog: dialog.dismiss())
            page.get_by_role("button", name="Adjuntar").click()

            # --- 7. Subir Orden de Trabajo — SOLO para alta, baja no lo necesita ---
            if tipo_movimiento == "alta":
                page.get_by_role("button", name="Choose File").set_input_files(orden_trabajo_path)
                page.once("dialog", lambda dialog: dialog.dismiss())
                page.get_by_role("button", name="Adjuntar").click()

            # --- 8. Comentario ---
            page.get_by_role("textbox", name="Agregar comentario...").click()
            page.get_by_role("textbox", name="Agregar comentario...").fill(comentario)
            page.locator("#BotonAgregar").click()

            # --- 9. Continuar y capturar folio final ---
            page.get_by_role("button", name="Continuar").click()

            page.wait_for_selector("text=Número de folio generado", timeout=20000)
            texto_modal = page.locator("body").inner_text()
            match = re.search(r"Número de folio generado:\s*(\d+)", texto_modal)
            folio = match.group(1) if match else "FOLIO_NO_DETECTADO"

            page.get_by_role("button", name="Aceptar").click()

        finally:
            context.close()
            browser.close()

    return folio


if __name__ == "__main__":
    # Prueba manual local — usa headless=False para ver el navegador.
    # ⚠️ Esto SÍ va a crear una solicitud real en Mapfre si se corre
    # contra credenciales reales. Usar con cuidado.
    logging.basicConfig(level=logging.INFO)

    # --- Ejemplo ALTA ---
    folio = emitir_movimiento_mapfre(
        no_poliza="2612600000892",
        tipo_movimiento="alta",
        concentrado_altas_path="/tmp/CONCENTRADO_TEST.xls",
        orden_trabajo_path="/tmp/OT_MAPFRE_TEST.pdf",
        headless=False,
    )
    print("Folio (alta):", folio)

    # --- Ejemplo BAJA (más simple: solo el Concentrado) ---
    # folio = emitir_movimiento_mapfre(
    #     no_poliza="2612600000892",
    #     tipo_movimiento="baja",
    #     concentrado_altas_path="/tmp/CONCENTRADO_TEST.xls",
    #     headless=False,
    # )
    # print("Folio (baja):", folio)
