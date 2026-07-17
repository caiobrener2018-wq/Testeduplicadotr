"""
Aplicação web para verificação de fotos duplicadas de atendimentos Sebrae.

Uso:
    pip install flask openpyxl pillow requests
    python app.py
    Abra http://localhost:5000
"""

import os
import uuid
import threading
import tempfile
from dataclasses import asdict

from flask import (Flask, request, render_template, jsonify,
                   send_file, abort)
from werkzeug.utils import secure_filename

import detector

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 25 * 1024 * 1024  # 25 MB

BASE = os.path.dirname(os.path.abspath(__file__))
UPLOAD_DIR = os.path.join(BASE, "uploads")
OUTPUT_DIR = os.path.join(BASE, "outputs")
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

# Estado dos jobs em memória (para produção, usar Redis/DB).
JOBS = {}
LOCK = threading.Lock()


def _set(job_id, **kw):
    with LOCK:
        JOBS[job_id].update(kw)


def _run_job(job_id, caminho_entrada, caminho_saida, fotos_cols):
    def progress(feito, total):
        _set(job_id, feito=feito, total=total,
             pct=round(feito / total * 100) if total else 0)
    try:
        _set(job_id, status="processando", etapa="Lendo planilha e baixando imagens…")
        res = detector.processar(caminho_entrada, caminho_saida,
                                 progress_cb=progress, fotos_cols=fotos_cols)
        dups = [{
            "linha_dup": d.foto.row,
            "coluna_dup": d.foto.col_letra,
            "pos": d.foto.idx_na_celula + 1,
            "linha_orig": d.original.row,
            "coluna_orig": d.original.col_letra,
            "url": d.foto.url,
        } for d in sorted(res.duplicatas, key=lambda x: (x.foto.row, x.foto.col))]
        falhas = [{"linha": f.row, "coluna": f.col_letra, "pos": f.idx_na_celula + 1,
                   "erro": f.erro, "url": f.url} for f in res.falhas]
        _set(job_id, status="concluido", etapa="Concluído",
             resumo={
                 "total_links": res.total_links,
                 "total_baixadas": res.total_baixadas,
                 "total_duplicatas": len(res.duplicatas),
                 "total_falhas": len(res.falhas),
             },
             duplicatas=dups, falhas=falhas,
             download=os.path.basename(caminho_saida))
    except Exception as e:
        _set(job_id, status="erro", etapa=f"Erro: {e}")


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/processar", methods=["POST"])
def processar():
    if "arquivo" not in request.files:
        return jsonify({"erro": "Nenhum arquivo enviado"}), 400
    f = request.files["arquivo"]
    if not f.filename.lower().endswith((".xlsx", ".xlsm")):
        return jsonify({"erro": "Envie um arquivo .xlsx"}), 400

    job_id = uuid.uuid4().hex
    nome = secure_filename(f.filename)
    entrada = os.path.join(UPLOAD_DIR, f"{job_id}_{nome}")
    saida = os.path.join(OUTPUT_DIR, f"{job_id}_verificado_{nome}")
    f.save(entrada)

    with LOCK:
        JOBS[job_id] = {"status": "iniciando", "pct": 0, "etapa": "Preparando…"}

    t = threading.Thread(target=_run_job,
                         args=(job_id, entrada, saida, detector.FOTOS_COLS), daemon=True)
    t.start()
    return jsonify({"job_id": job_id})


@app.route("/status/<job_id>")
def status(job_id):
    with LOCK:
        job = JOBS.get(job_id)
    if not job:
        return jsonify({"erro": "job não encontrado"}), 404
    return jsonify(job)


@app.route("/download/<job_id>")
def download(job_id):
    with LOCK:
        job = JOBS.get(job_id)
    if not job or job.get("status") != "concluido":
        abort(404)
    caminho = os.path.join(OUTPUT_DIR, job["download"])
    if not os.path.exists(caminho):
        abort(404)
    return send_file(caminho, as_attachment=True,
                     download_name=job["download"].split("_", 1)[1])


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
