#!/usr/bin/env python3
"""
VO Generator — Web UI
Run with: python3 app.py  (or via the `vogen` alias)
"""
import json
import os
import queue
import shutil
import socket
import tempfile
import threading
import uuid
import webbrowser
import zipfile

from dotenv import load_dotenv, set_key
from flask import Flask, Response, jsonify, render_template, request, send_file, stream_with_context
from werkzeug.utils import secure_filename

from vo_generator import GenerationAborted, generate_from_table, inspect_document

load_dotenv()

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 32 * 1024 * 1024  # 32 MB upload limit

ENV_PATH = os.path.join(os.path.dirname(__file__), ".env")
UPLOAD_DIR = os.path.join(tempfile.gettempdir(), "vo-generator-uploads")
os.makedirs(UPLOAD_DIR, exist_ok=True)

_jobs = {}
_jobs_lock = threading.Lock()

PORT = 5000


def in_docker():
    return os.path.exists("/.dockerenv")


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/status")
def status():
    return jsonify({"has_key": bool(os.environ.get("ELEVENLABS_API_KEY"))})


@app.route("/setup", methods=["POST"])
def setup():
    data = request.get_json(force=True)
    api_key = (data.get("api_key") or "").strip()
    if not api_key:
        return jsonify({"error": "No API key provided"}), 400

    import requests as req
    try:
        resp = req.get(
            "https://api.elevenlabs.io/v1/voices",
            headers={"xi-api-key": api_key},
            timeout=8,
        )
        if resp.status_code == 401:
            return jsonify({"error": "Invalid API key — ElevenLabs rejected it."}), 400
        resp.raise_for_status()
    except Exception as e:
        return jsonify({"error": f"Could not reach ElevenLabs to validate key: {e}"}), 400

    set_key(ENV_PATH, "ELEVENLABS_API_KEY", api_key)
    os.environ["ELEVENLABS_API_KEY"] = api_key
    return jsonify({"status": "ok"})


SUPPORTED_EXTENSIONS = {".docx", ".pptx", ".xlsx"}


@app.route("/upload", methods=["POST"])
def upload():
    if "file" not in request.files:
        return jsonify({"error": "No file provided"}), 400
    f = request.files["file"]
    ext = os.path.splitext(f.filename)[1].lower()
    if ext not in SUPPORTED_EXTENSIONS:
        return jsonify({"error": f"Unsupported file type '{ext}'. Accepted: .docx, .pptx, .xlsx"}), 422

    session_id = str(uuid.uuid4())
    session_dir = os.path.join(UPLOAD_DIR, session_id)
    os.makedirs(session_dir, exist_ok=True)

    filename = secure_filename(f.filename)
    doc_path = os.path.join(session_dir, filename)
    f.save(doc_path)

    try:
        result = inspect_document(doc_path)
    except Exception as e:
        shutil.rmtree(session_dir, ignore_errors=True)
        return jsonify({"error": str(e)}), 500

    if not result["tables"]:
        shutil.rmtree(session_dir, ignore_errors=True)
        return jsonify({"error": "No tables found in this document"}), 422

    result["session_id"] = session_id
    result["doc_path"] = doc_path

    threading.Timer(3600, lambda: shutil.rmtree(session_dir, ignore_errors=True)).start()

    return jsonify(result)


@app.route("/generate", methods=["POST"])
def generate():
    data = request.get_json(force=True)

    doc_path = (data.get("doc_path") or "").strip()
    if not doc_path or not os.path.isfile(doc_path):
        return jsonify({"error": "Uploaded file not found — please re-upload"}), 400

    try:
        table_idx = int(data.get("table_idx", 0))
        vo_col = int(data["vo_col"])
        file_col = int(data["file_col"])
        header_rows = int(data.get("header_rows", 2))
    except (KeyError, ValueError) as e:
        return jsonify({"error": f"Invalid parameter: {e}"}), 400

    voice_name = (data.get("voice_name") or "Mikko").strip() or "Mikko"
    file_type = (data.get("file_type") or "docx").strip()
    api_key = (data.get("api_key") or "").strip() or os.environ.get("ELEVENLABS_API_KEY")

    if not api_key:
        return jsonify({"error": "No API key — complete setup first"}), 400

    audio_dir = os.path.join(os.path.dirname(doc_path), "audio")
    os.makedirs(audio_dir, exist_ok=True)

    job_id = str(uuid.uuid4())
    q = queue.Queue(maxsize=500)
    job = {"queue": q, "aborted": False, "audio_dir": audio_dir}

    with _jobs_lock:
        _jobs[job_id] = job

    def progress_cb(event):
        if job["aborted"] and event.get("status") == "generating":
            raise GenerationAborted()
        try:
            q.put_nowait(event)
        except queue.Full:
            pass

    def run():
        try:
            generate_from_table(
                doc_path, table_idx, vo_col, file_col,
                header_rows, api_key, voice_name, progress_cb,
                output_dir=audio_dir,
                file_type=file_type,
            )
        except Exception as e:
            try:
                q.put_nowait({"type": "error", "message": str(e)})
            except queue.Full:
                pass
        finally:
            try:
                q.put_nowait(None)
            except queue.Full:
                pass
            threading.Timer(300, lambda: _jobs.pop(job_id, None)).start()

    threading.Thread(target=run, daemon=True).start()
    return jsonify({"job_id": job_id}), 202


@app.route("/stream/<job_id>")
def stream(job_id):
    with _jobs_lock:
        job = _jobs.get(job_id)
    if not job:
        return jsonify({"error": "Unknown job"}), 404

    q = job["queue"]

    def event_stream():
        while True:
            try:
                event = q.get(timeout=25)
            except queue.Empty:
                yield ": keepalive\n\n"
                continue
            if event is None:
                break
            yield f"data: {json.dumps(event)}\n\n"

    return Response(
        stream_with_context(event_stream()),
        mimetype="text/event-stream",
        headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"},
    )


@app.route("/download/<job_id>")
def download(job_id):
    with _jobs_lock:
        job = _jobs.get(job_id)
    if not job:
        return jsonify({"error": "Unknown job or already expired"}), 404

    audio_dir = job.get("audio_dir")
    if not audio_dir or not os.path.isdir(audio_dir):
        return jsonify({"error": "Audio files not found"}), 404

    wav_files = [f for f in os.listdir(audio_dir) if f.endswith(".wav")]
    if not wav_files:
        return jsonify({"error": "No .wav files were generated"}), 404

    zip_path = audio_dir + ".zip"
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for fname in sorted(wav_files):
            zf.write(os.path.join(audio_dir, fname), fname)

    return send_file(zip_path, as_attachment=True, download_name="vo_audio.zip")


@app.route("/abort/<job_id>", methods=["POST"])
def abort(job_id):
    with _jobs_lock:
        job = _jobs.get(job_id)
    if not job:
        return jsonify({"error": "Unknown job"}), 404
    job["aborted"] = True
    return jsonify({"status": "aborted"})


if __name__ == "__main__":
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        if s.connect_ex(("127.0.0.1", PORT)) == 0:
            print(f"\nERROR: Port {PORT} is already in use.")
            print(f"  Kill the conflicting process with: lsof -ti:{PORT} | xargs kill\n")
            raise SystemExit(1)

    if not in_docker():
        threading.Timer(1.2, lambda: webbrowser.open(f"http://localhost:{PORT}")).start()

    print(f"\n  VO Generator running at http://localhost:{PORT}\n")
    app.run(host="0.0.0.0", port=PORT, debug=False, threaded=True)