# -*- coding: utf-8 -*-
"""
Lógica de negocio de la app de Anexos V (Fajas UE).

Reutiliza sin cambios las reglas de generador_anexos.py (day_tramos / build_day)
y agrega: lectura de varias bajadas, detección de archivos duplicados, alertas de
control y empaquetado.
"""

import hashlib
import importlib.util
import json
import os
import re
import zipfile

import openpyxl
import pandas as pd

from generador_anexos import MESES, build_day, day_tramos, parse_faja

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PLANTILLA = os.environ.get(
    "PLANTILLA", os.path.join(BASE_DIR, "plantilla_Fajas_UE_fixed.xlsx")
)

# Umbrales de control (configurables por variable de entorno)
VOLUMEN_ALTO = int(os.environ.get("VOLUMEN_ALTO", "4000"))   # fajas/día
VOLUMEN_BAJO = int(os.environ.get("VOLUMEN_BAJO", "600"))    # fajas/día
TRAMO_CHICO = int(os.environ.get("TRAMO_CHICO", "5"))        # fajas en un tramo "suelto"
DISTANCIA_LEJANA = int(os.environ.get("DISTANCIA_LEJANA", "10000"))
MAX_DIAS = int(os.environ.get("MAX_DIAS", "62"))

# Sin Pillow, openpyxl descarta el logo de SENASA al abrir la plantilla
PILLOW_OK = importlib.util.find_spec("PIL") is not None


class BajadaError(Exception):
    """Error de validación con mensaje apto para mostrar al usuario."""


# --------------------------------------------------------------------------- #
# Lectura
# --------------------------------------------------------------------------- #
def _sha256(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for bloque in iter(lambda: f.read(1 << 20), b""):
            h.update(bloque)
    return h.hexdigest()


def leer_bajadas(archivos):
    """archivos: lista de (nombre_original, ruta). Devuelve (DataFrame, avisos)."""
    avisos, vistos, frames = [], {}, []
    for nombre, ruta in archivos:
        huella = _sha256(ruta)
        if huella in vistos:
            avisos.append(f"«{nombre}» es idéntico a «{vistos[huella]}»: se usó una sola vez.")
            continue
        vistos[huella] = nombre
        try:
            libro = openpyxl.load_workbook(ruta, read_only=True)
            hojas = libro.sheetnames
            libro.close()
        except Exception:
            raise BajadaError(f"No se pudo leer «{nombre}». ¿Es un archivo .xlsx válido?")
        if "ag-grid" not in hojas:
            raise BajadaError(
                f"«{nombre}» no tiene la hoja «ag-grid». Verifique que sea una bajada de scans."
            )
        df = pd.read_excel(ruta, sheet_name="ag-grid", dtype=str)
        faltan = {"Fecha", "Faja1", "Faja2"} - set(df.columns)
        if faltan:
            raise BajadaError(
                f"«{nombre}» no tiene las columnas requeridas: {', '.join(sorted(faltan))}."
            )
        frames.append(df)
    if not frames:
        raise BajadaError("No se recibió ninguna bajada.")
    return pd.concat(frames, ignore_index=True), avisos


# --------------------------------------------------------------------------- #
# Análisis por día y alertas
# --------------------------------------------------------------------------- #
def _numeros_por_serie(d):
    nums = {}
    for col in ("Faja1", "Faja2"):
        for v in d[col].dropna():
            p = parse_faja(v)
            if p:
                nums.setdefault(p[0], set()).add(p[1])
    return nums


def _alertas(d, tramos, util):
    alertas = []
    nums = _numeros_por_serie(d)
    for t in tramos:
        t["cantidad"] = sum(1 for n in nums.get(t["serie"], ()) if t["du"] <= n <= t["hu"])

    if len(tramos) > 1:
        principal = max(tramos, key=lambda t: t["cantidad"])
        for t in tramos:
            if t is principal or t["serie"] != principal["serie"]:
                continue
            distancia = min(abs(t["du"] - principal["hu"]), abs(principal["du"] - t["hu"]))
            rango = f"{t['du']}" if t["du"] == t["hu"] else f"{t['du']}–{t['hu']}"
            if t["cantidad"] <= TRAMO_CHICO:
                alertas.append(
                    f"Tramo {rango} con solo {t['cantidad']} faja(s), separado del tramo principal "
                    f"({principal['du']}–{principal['hu']}). Posible escaneo aislado."
                )
            elif distancia > DISTANCIA_LEJANA:
                alertas.append(
                    f"Tramo {rango} muy alejado del tramo principal "
                    f"({principal['du']}–{principal['hu']}). Posible escaneo de fajas antiguas."
                )
    if util > VOLUMEN_ALTO:
        alertas.append(
            f"Volumen alto ({util} fajas). Confirmar que la bajada no mezcle más de un día."
        )
    elif util < VOLUMEN_BAJO:
        alertas.append(
            f"Volumen bajo ({util} fajas). Confirmar que la bajada esté completa o que sea una jornada reducida."
        )
    return alertas


def analizar(df, solo_ultimo=False):
    fechas = sorted(
        {d for d in df["Fecha"].dropna().unique() if isinstance(d, str)
         and re.fullmatch(r"\d{4}-\d{2}-\d{2}", d)}
    )
    dias = []
    for fecha in fechas:
        d = df[df["Fecha"] == fecha]
        tramos, util = day_tramos(d)
        if not tramos:
            continue
        dias.append({"fecha": fecha, "tramos": tramos, "util": util,
                     "alertas": _alertas(d, tramos, util)})
    if not dias:
        raise BajadaError(
            "No se encontraron fajas UE (series FE / FF) en la bajada. "
            "Las fajas FC (USA/Kosher) no se incluyen."
        )
    if solo_ultimo:
        dias = dias[-1:]
    if len(dias) > MAX_DIAS:
        raise BajadaError(
            f"La bajada tiene {len(dias)} días con producción UE; el máximo por vez es {MAX_DIAS}. "
            "Divida la carga en partes."
        )
    return dias


# --------------------------------------------------------------------------- #
# Generación
# --------------------------------------------------------------------------- #
def nombre_xlsx(fecha):
    yy, mm, dd = fecha.split("-")
    return f"Anexo_V_Fajas_UE_{dd}-{mm}-{yy}.xlsx"


def generar(df, dias, job_dir):
    """Genera un Excel por día (y un ZIP si hay más de uno) dentro de job_dir."""
    xlsx_dir = os.path.join(job_dir, "xlsx")
    os.makedirs(xlsx_dir, exist_ok=True)

    for dia in dias:
        ruta = os.path.join(xlsx_dir, nombre_xlsx(dia["fecha"]))
        build_day(PLANTILLA, dia["fecha"], dia["tramos"], dia["util"], ruta)
        dia["xlsx"] = os.path.basename(ruta)

    zip_nombre = None
    if len(dias) > 1:
        zip_nombre = "Anexos_V_Fajas_UE.zip"
        with zipfile.ZipFile(os.path.join(job_dir, zip_nombre), "w", zipfile.ZIP_DEFLATED) as z:
            for dia in dias:
                yy, mm, dd = dia["fecha"].split("-")
                z.write(os.path.join(xlsx_dir, dia["xlsx"]),
                        f"{MESES.get(mm, mm)}/Anexo_V_Fajas_UE_{yy}-{mm}-{dd}.xlsx")

    meta = {
        "dias": _serializable(dias),
        "avisos": [] if PILLOW_OK else [
            "Atención: falta la librería Pillow en el servidor y los anexos salen SIN el logo de SENASA. "
            "Avise al administrador antes de imprimir."
        ],
        "zip": zip_nombre,
        "total_util": sum(d["util"] for d in dias),
    }
    with open(os.path.join(job_dir, "meta.json"), "w", encoding="utf-8") as f:
        json.dump(meta, f, ensure_ascii=False)
    return meta


def _serializable(dias):
    return [
        {
            "fecha": d["fecha"],
            "fecha_fmt": "/".join(reversed(d["fecha"].split("-"))),
            "util": d["util"],
            "tramos": [
                {k: t[k] for k in ("serie", "du", "hu", "sh", "cantidad")} for t in d["tramos"]
            ],
            "alertas": d["alertas"],
            "xlsx": d["xlsx"],
        }
        for d in dias
    ]
