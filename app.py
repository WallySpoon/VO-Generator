#!/usr/bin/env python3
"""
VO Generator — Web UI
Run with: python3 app.py  (or via the `vogen` alias)
"""
import json
import os
import queue
import socket
import threading
import uuid
import webbrowser

from dotenv import load_dotenv, set_key
from flask import Flask, Response, jsonify, render_template, request, stream_with_context

from vo_generator import GenerationAborted, generate_from_table, inspect_docx

load_dotenv()

app = Flask(__name__)

_jobs = {}
_jobs_lock = threading.Lock()

PORT = 5000


ENV_PATH = os.path.join(os.path.dirname(__file__), ".env")


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

    # Validate the key against ElevenLabs before saving
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

    # Save to .env and update the running process
    set_key(ENV_PATH, "ELEVENLABS_API_KEY", api_key)
    os.environ["ELEVENLABS_API_KEY"] = api_key
    return jsonify({"status": "ok"})


@app.route("/inspect", methods=["POST"])
def inspect():
    data = request.get_json(force=True)
    path = (data.get("path") or "").strip()
    if not path:
        return jsonify({"error": "No path provided"}), 400

    resolved = os.path.abspath(os.path.expanduser(path))
    if not os.path.isfile(resolved):
        return jsonify({"error": f"File not found: {resolved}"}), 400

    ext = os.path.splitext(resolved)[1].lower()
    if ext != ".docx":
        return jsonify({"error": f"Unsupported file type '{ext}' — only .docx is supported currently."}), 422

    try:
        result = inspect_docx(resolved)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

    if not result["tables"]:
        return jsonify({"error": "No tables found in this document."}), 422

    return jsonify(result)


@app.route("/generate", methods=["POST"])
def generate():
    data = request.get_json(force=True)

    path = (data.get("path") or "").strip()
    if not path or not os.path.isfile(path):
        return jsonify({"error": "File not found"}), 400

    try:
        table_idx = int(data.get("table_idx", 0))
        vo_col = int(data["vo_col"])
        file_col = int(data["file_col"])
        header_rows = int(data.get("header_rows", 2))
    except (KeyError, ValueError) as e:
        return jsonify({"error": f"Invalid parameter: {e}"}), 400

    voice_name = (data.get("voice_name") or "Mikko").strip() or "Mikko"
    api_key = (data.get("api_key") or "").strip() or os.environ.get("ELEVENLABS_API_KEY")

    if not api_key:
        return jsonify({"error": "No API key — add ELEVENLABS_API_KEY to .env or enter it in the form."}), 400

    job_id = str(uuid.uuid4())
    q = queue.Queue(maxsize=500)
    job = {"queue": q, "aborted": False}

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
                path, table_idx, vo_col, file_col,
                header_rows, api_key, voice_name, progress_cb,
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


@app.route("/abort/<job_id>", methods=["POST"])
def abort(job_id):
    with _jobs_lock:
        job = _jobs.get(job_id)
    if not job:
        return jsonify({"error": "Unknown job"}), 404
    job["aborted"] = True
    return jsonify({"status": "aborted"})


if __name__ == "__main__":
    # macOS AirPlay Receiver also uses port 5000 — catch the conflict early
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        if s.connect_ex(("127.0.0.1", PORT)) == 0:
            print(f"\nERROR: Port {PORT} is already in use.")
            print(f"  Kill the conflicting process with: lsof -ti:{PORT} | xargs kill\n")
            raise SystemExit(1)

    threading.Timer(1.2, lambda: webbrowser.open(f"http://localhost:{PORT}")).start()
    print(f"\n  VO Generator running at http://localhost:{PORT}\n")
    app.run(host="127.0.0.1", port=PORT, debug=False, threaded=True)
