# -*- coding: utf-8 -*-
"""
Servidor Flask que recibe la llamada del cron de Odoo, orquesta:
generar Excel "Concentrado de Altas" (con TODOS los asegurados del grupo)
-> generar PDF "Orden de Trabajo" (solo para alta, a nivel póliza, sin
detalle de personas) -> correr el RPA contra el portal ZonAliados de
Mapfre.

Mismo patrón que webhook_plan_seguro.py, corre en el mismo VPS/EasyPanel
(o en uno nuevo — ver INSTALACION_MAPFRE.txt).
"""

import os
import base64
import logging
from flask import Flask, request, jsonify

from generar_concentrado_mapfre import generar_concentrado_altas_mapfre
from generar_ot_mapfre import generar_orden_trabajo_mapfre
from rpa_mapfre import emitir_movimiento_mapfre, TipoMovimientoNoConfirmadoError

logging.basicConfig(level=logging.INFO)
log = logging.getLogger("webhook_mapfre")

app = Flask(__name__)

TRABAJO_DIR = os.environ.get("TRABAJO_DIR", "/tmp/mapfre_pipeline")
os.makedirs(TRABAJO_DIR, exist_ok=True)

WEBHOOK_SECRET = os.environ.get("MAPFRE_WEBHOOK_SECRET", "")


@app.route("/webhook/mapfre/alta", methods=["POST"])
def webhook_mapfre():
    if WEBHOOK_SECRET:
        recibido = request.headers.get("X-Webhook-Secret", "")
        if recibido != WEBHOOK_SECRET:
            return jsonify({"ok": False, "error": "Secreto inválido"}), 403

    payload = request.get_json(force=True)

    try:
        orden_id = payload["order_id"]
        tipo_movimiento = payload.get("tipo_movimiento", "alta")
        no_poliza = payload["no_poliza"]
        asegurados = payload["asegurados"]  # lista — 1 o varios

        if not asegurados:
            return jsonify({"ok": False, "error": "No se recibió ningún asegurado."}), 400

        base = f"{TRABAJO_DIR}/orden_{orden_id}_{tipo_movimiento}"

        # 1. Generar Excel "Concentrado de Altas" — lleva a TODOS, sin límite
        concentrado_asegurados = [
            {
                "apellido_paterno": a["apellido_paterno"],
                "apellido_materno": a.get("apellido_materno", ""),
                "nombre": a["nombre"],
                "parentesco": a.get("parentesco", ""),
                "fecha_nacimiento": a["fecha_nacimiento"],
                "sexo": a["sexo"],  # 'M' o 'F'
            }
            for a in asegurados
        ]
        concentrado_path = generar_concentrado_altas_mapfre(
            concentrado_asegurados,
            no_poliza=no_poliza,
            salida=f"{base}_concentrado.xls",
        )

        # 2. Generar PDF "Orden de Trabajo" — SOLO para alta, no lleva
        # detalle de personas (es a nivel póliza/empresa)
        ot_path = None
        if tipo_movimiento == "alta":
            ot_path = generar_orden_trabajo_mapfre(
                no_poliza=no_poliza,
                salida_pdf=f"{base}_orden_trabajo.pdf",
            )

        # 3. Correr el RPA
        folio = emitir_movimiento_mapfre(
            no_poliza=no_poliza,
            tipo_movimiento=tipo_movimiento,
            concentrado_altas_path=concentrado_path,
            orden_trabajo_path=ot_path,  # None para baja, la función lo ignora
        )

        log.info(
            f"Orden {orden_id}: {tipo_movimiento} registrada en Mapfre para "
            f"{len(asegurados)} asegurado(s) — folio {folio}"
        )

        # Regresar los archivos generados en base64 para que Odoo los adjunte
        with open(concentrado_path, "rb") as f:
            concentrado_base64 = base64.b64encode(f.read()).decode()

        ot_base64 = None
        if ot_path:
            with open(ot_path, "rb") as f:
                ot_base64 = base64.b64encode(f.read()).decode()

        return jsonify({
            "ok": True,
            "folio": folio,
            "concentrado_base64": concentrado_base64,
            "concentrado_filename": os.path.basename(concentrado_path),
            "ot_base64": ot_base64,
            "ot_filename": os.path.basename(ot_path) if ot_path else None,
        }), 200

    except TipoMovimientoNoConfirmadoError as e:
        log.error(f"Tipo de movimiento no soportado: {e}")
        return jsonify({"ok": False, "error": str(e)}), 500

    except Exception as e:
        log.exception("Error procesando movimiento Mapfre")
        return jsonify({"ok": False, "error": str(e)}), 500


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5003))
    app.run(host="0.0.0.0", port=port)
