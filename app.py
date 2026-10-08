# -*- coding: utf-8 -*-
"""App web para generar los Anexos V (Fajas UE - SENASA) a partir de las bajadas."""

import hmac
import json
import os
import re
import shutil
import tempfile
import time
import uuid

from flask import (Flask, Response, abort, redirect, render_template, request,
                   send_from_directory, url_for)

import service

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = int(os.environ.get("MAX_UPLOAD_MB", "60")) * 1024 * 1024

JOBS_DIR = os.environ.get("JOBS_DIR", os.path.join(tempfile.gettempdir(), "anexos_jobs"))
JOB_TTL_SEGUNDOS = int(os.environ.get("JOB_TTL_HORAS", "6")) * 3600
MAX_ARCHIVOS = int(os.environ.get("MAX_ARCHIVOS", "20"))
APP_USER = os.environ.get("APP_USER", "fajas")
APP_PASSWORD = os.environ.get("APP_PASSWORD", "")

JOB_ID = re.compile(r"^[0-9a-f]{32}$")
NOMBRE_SEGURO = re.compile(r"^[A-Za-z0-9_.\-]+$")

os.makedirs(JOBS_DIR, exist_ok=True)


# --------------------------------------------------------------------------- #
# Acceso (opcional): si APP_PASSWORD está definida, se pide usuario y contraseña
# --------------------------------------------------------------------------- #
@app.before_request
def _requiere_clave():
    if not APP_PASSWORD or request.path == "/health":
        return None
    auth = request.authorization
    if auth and hmac.compare_digest(auth.username or "", APP_USER) \
            and hmac.compare_digest(auth.password or "", APP_PASSWORD):
        return None
    return Response("Acceso restringido.", 401,
                    {"WWW-Authenticate": 'Basic realm="Anexos V Fajas UE"'})


def _limpiar_viejos():
    ahora = time.time()
    for nombre in os.listdir(JOBS_DIR):
        ruta = os.path.join(JOBS_DIR, nombre)
        try:
            if ahora - os.path.getmtime(ruta) > JOB_TTL_SEGUNDOS:
                shutil.rmtree(ruta, ignore_errors=True)
        except OSError:
            pass


def _job_dir(job_id):
    if not JOB_ID.match(job_id):
        abort(404)
    ruta = os.path.join(JOBS_DIR, job_id)
    if not os.path.isdir(ruta):
        abort(404)
    return ruta


# --------------------------------------------------------------------------- #
# Rutas
# --------------------------------------------------------------------------- #
@app.get("/health")
def health():
    return {"estado": "ok"}


@app.get("/")
def inicio():
    return render_template("index.html", error=None)


@app.post("/generar")
def generar():
    _limpiar_viejos()
    archivos = [f for f in request.files.getlist("bajadas") if f and f.filename]
    if not archivos:
        return render_template("index.html", error="Seleccione al menos un archivo de bajada."), 400
    if len(archivos) > MAX_ARCHIVOS:
        return render_template(
            "index.html", error=f"Se admiten hasta {MAX_ARCHIVOS} archivos por vez."), 400
    for f in archivos:
        if not f.filename.lower().endswith(".xlsx"):
            return render_template(
                "index.html", error=f"«{f.filename}» no es un archivo .xlsx."), 400

    job_id = uuid.uuid4().hex
    job_dir = os.path.join(JOBS_DIR, job_id)
    entrada = os.path.join(job_dir, "entrada")
    os.makedirs(entrada)
    try:
        rutas = []
        for i, f in enumerate(archivos):
            ruta = os.path.join(entrada, f"{i}.xlsx")
            f.save(ruta)
            rutas.append((f.filename, ruta))

        df, avisos = service.leer_bajadas(rutas)
        solo_ultimo = request.form.get("modo") == "ultimo"
        dias = service.analizar(df, solo_ultimo=solo_ultimo)
        meta = service.generar(df, dias, job_dir)
        meta["avisos"] = avisos + meta["avisos"]
        with open(os.path.join(job_dir, "meta.json"), "w", encoding="utf-8") as fh:
            json.dump(meta, fh, ensure_ascii=False)
    except service.BajadaError as e:
        shutil.rmtree(job_dir, ignore_errors=True)
        return render_template("index.html", error=str(e)), 400
    except Exception:
        app.logger.exception("Error inesperado al generar anexos")
        shutil.rmtree(job_dir, ignore_errors=True)
        return render_template(
            "index.html",
            error="Ocurrió un error inesperado al generar los anexos. Intente nuevamente o revise el archivo.",
        ), 500
    finally:
        shutil.rmtree(entrada, ignore_errors=True)  # no se conservan las bajadas subidas
    return redirect(url_for("resultado", job_id=job_id))


@app.get("/resultado/<job_id>")
def resultado(job_id):
    job_dir = _job_dir(job_id)
    with open(os.path.join(job_dir, "meta.json"), encoding="utf-8") as fh:
        meta = json.load(fh)
    con_alertas = sum(1 for d in meta["dias"] if d["alertas"])
    return render_template("resultado.html", job_id=job_id, meta=meta, con_alertas=con_alertas)


@app.get("/descargar/<job_id>/<tipo>/<nombre>")
def descargar(job_id, tipo, nombre):
    job_dir = _job_dir(job_id)
    if not NOMBRE_SEGURO.match(nombre):
        abort(404)
    carpetas = {"xlsx": "xlsx", "pdf": "pdf", "raiz": ""}
    if tipo not in carpetas:
        abort(404)
    return send_from_directory(os.path.join(job_dir, carpetas[tipo]), nombre, as_attachment=True)


@app.errorhandler(413)
def demasiado_grande(_e):
    return render_template(
        "index.html", error="Los archivos superan el tamaño máximo permitido."), 413


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", "8000")), debug=False)
