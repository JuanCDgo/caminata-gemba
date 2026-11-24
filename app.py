"""Local Flask server for Caminata Gemba presentation generator."""

from __future__ import annotations

import base64
import json
import os
import subprocess
from pathlib import Path

from flask import Flask, Response, jsonify, render_template, request, send_file

import pptx_engine

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
DATA_DIR.mkdir(exist_ok=True)
DEFAULT_PPTX = DATA_DIR / "Caminata Gemba.pptx"

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 200 * 1024 * 1024  # 200 MB max upload


def _active_pptx() -> Path:
    if DEFAULT_PPTX.exists():
        return DEFAULT_PPTX
    candidates = sorted(DATA_DIR.glob("*.pptx"))
    if candidates:
        return candidates[0]
    return DEFAULT_PPTX


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/status", methods=["GET"])
def api_status():
    pptx_path = _active_pptx()
    if not pptx_path.exists():
        return jsonify({"loaded": False, "message": "No hay ningún archivo .pptx cargado todavía."})
    info = pptx_engine.inspect_presentation(pptx_path)
    info["loaded"] = True
    info["path"] = str(pptx_path)
    return jsonify(info)


@app.route("/api/upload-pptx", methods=["POST"])
def api_upload_pptx():
    file = request.files.get("pptx")
    if not file or not file.filename.lower().endswith(".pptx"):
        return jsonify({"error": "Seleccioná un archivo .pptx válido."}), 400
    if DEFAULT_PPTX.exists():
        pptx_engine.backup_pptx(DEFAULT_PPTX)
    file.save(str(DEFAULT_PPTX))
    info = pptx_engine.inspect_presentation(DEFAULT_PPTX)
    info["loaded"] = True
    info["path"] = str(DEFAULT_PPTX)
    return jsonify(info)


@app.route("/api/slide-image/<int:slide_index>/<int:pic_index>", methods=["GET"])
def api_slide_image(slide_index: int, pic_index: int):
    pptx_path = _active_pptx()
    if not pptx_path.exists():
        return Response(status=404)
    try:
        jpg = pptx_engine.extract_slide_picture(pptx_path, slide_index, pic_index)
        return Response(jpg, mimetype="image/jpeg", headers={"Cache-Control": "no-store"})
    except Exception:
        return Response(status=404)


@app.route("/api/add-walk", methods=["POST"])
def api_add_walk():
    pptx_path = _active_pptx()
    if not pptx_path.exists():
        return jsonify({"error": "Primero cargá el archivo Caminata Gemba.pptx."}), 400

    cover_b64 = request.form.get("cover_png", "")
    if not cover_b64:
        return jsonify({"error": "Falta la imagen de portada."}), 400
    if "," in cover_b64:
        cover_b64 = cover_b64.split(",", 1)[1]
    cover_bytes = base64.b64decode(cover_b64)

    files = request.files.getlist("photos")
    if not files:
        return jsonify({"error": "Agregá al menos una foto del ANTES."}), 400

    photo_bytes_list = [f.read() for f in files if f and f.filename]
    if not photo_bytes_list:
        return jsonify({"error": "No se recibieron fotos válidas."}), 400

    result = pptx_engine.append_walk(
        pptx_path=pptx_path,
        cover_png_bytes=cover_bytes,
        antes_photos=photo_bytes_list,
        clean_trailing_empty=True,
    )
    return jsonify(result)


@app.route("/api/update-despues", methods=["POST"])
def api_update_despues():
    pptx_path = _active_pptx()
    if not pptx_path.exists():
        return jsonify({"error": "Primero cargá el archivo Caminata Gemba.pptx."}), 400

    meta_raw = request.form.get("updates_meta", "[]")
    try:
        meta = json.loads(meta_raw)
    except json.JSONDecodeError:
        return jsonify({"error": "Formato inválido en updates_meta."}), 400

    updates = []
    for item in meta:
        slide_index = int(item["slide_index"])
        note = item.get("note", "")
        file_key = f"photo_{slide_index}"
        file_obj = request.files.get(file_key)
        photo_bytes = file_obj.read() if file_obj and file_obj.filename else None
        if photo_bytes or note.strip():
            updates.append(
                {
                    "slide_index": slide_index,
                    "photo_bytes": photo_bytes,
                    "note": note,
                }
            )

    if not updates:
        return jsonify({"error": "No hay cambios para guardar."}), 400

    result = pptx_engine.update_despues_slides(pptx_path, updates)
    return jsonify(result)


@app.route("/api/download-pptx", methods=["GET"])
def api_download_pptx():
    pptx_path = _active_pptx()
    if not pptx_path.exists():
        return jsonify({"error": "No existe el archivo."}), 404
    return send_file(
        str(pptx_path),
        as_attachment=True,
        download_name="Caminata Gemba.pptx",
        mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
    )


@app.route("/api/open-folder", methods=["POST"])
def api_open_folder():
    pptx_path = _active_pptx()
    if os.name == "nt":
        if pptx_path.exists():
            subprocess.Popen(["explorer", "/select,", str(pptx_path)])
        else:
            subprocess.Popen(["explorer", str(DATA_DIR)])
    return jsonify({"ok": True})


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5150, debug=False)
