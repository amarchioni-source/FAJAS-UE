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

from werkzeug.exceptions import HTTPException
from flask import (Flask, Response, abort, redirect, render_template, request,
                   send_from_directory, url_for)

import service

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
app = Flask(__name__, template_folder=BASE_DIR, static_folder=None)
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
def _iguales(a, b):
    """Comparación en tiempo constante; admite caracteres no ASCII (ñ, tildes)."""
    return hmac.compare_digest((a or "").encode("utf-8"), (b or "").encode("utf-8"))


@app.before_request
def _requiere_clave():
    if not APP_PASSWORD or request.path == "/health":
        return None
    auth = request.authorization
    if auth and _iguales(auth.username, APP_USER) and _iguales(auth.password, APP_PASSWORD):
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
    carpetas = {"xlsx": "xlsx", "raiz": ""}
    if tipo not in carpetas:
        abort(404)
    return send_from_directory(os.path.join(job_dir, carpetas[tipo]), nombre, as_attachment=True)


@app.errorhandler(Exception)
def error_no_controlado(e):
    if isinstance(e, HTTPException):
        return e
    codigo = uuid.uuid4().hex[:8]
    app.logger.exception("Error no controlado [%s] en %s %s", codigo, request.method, request.path)
    cuerpo = (
        "<!doctype html><meta charset='utf-8'><title>Error</title>"
        "<body style='font-family:system-ui,sans-serif;max-width:640px;margin:48px auto;padding:0 16px'>"
        "<h1>Ocurrió un error</h1>"
        f"<p>Código de error: <b>{codigo}</b>. Está registrado en los logs del servidor con el detalle.</p>"
        "<p>Para revisar la instalación, abra <a href='/diagnostico'>/diagnostico</a>.</p>"
        "<p><a href='/'>Volver al inicio</a></p></body>"
    )
    return Response(cuerpo, 500, mimetype="text/html")


@app.get("/diagnostico")
def diagnostico():
    """Revisa que la instalación esté completa."""
    import platform

    import flask
    import openpyxl
    import pandas

    base = os.path.dirname(os.path.abspath(__file__))
    archivos = ["base.html", "index.html", "resultado.html",
                "generador_anexos.py", "service.py"]
    info = {
        "archivos": {a: os.path.exists(os.path.join(base, a)) for a in archivos},
        "plantilla_xlsx": os.path.exists(service.PLANTILLA),
        "pillow_instalado": service.PILLOW_OK,
        "logo_en_plantilla": len(openpyxl.load_workbook(service.PLANTILLA).active._images) == 1,
        "acceso_con_clave": bool(APP_PASSWORD),
        "versiones": {"python": platform.python_version(), "flask": flask.__version__,
                      "pandas": pandas.__version__, "openpyxl": openpyxl.__version__},
    }
    try:
        prueba = os.path.join(JOBS_DIR, ".prueba")
        with open(prueba, "w") as fh:
            fh.write("ok")
        os.remove(prueba)
        info["carpeta_temporal_escribible"] = True
    except OSError as exc:
        info["carpeta_temporal_escribible"] = f"no: {exc}"

    return info


@app.errorhandler(413)
def demasiado_grande(_e):
    return render_template(
        "index.html", error="Los archivos superan el tamaño máximo permitido."), 413


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", "8000")), debug=False)
