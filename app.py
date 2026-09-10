"""Forge (Gen 2) — an Ollama Cloud-powered, downloadable file workspace."""
import base64
import io
import json
import os
import re
import uuid
import zipfile
from datetime import datetime, timezone
from pathlib import Path, PurePosixPath

import requests
from flask import Flask, abort, jsonify, render_template, request, send_file
from docx import Document
from openpyxl import Workbook
from pptx import Presentation
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 12 * 1024 * 1024
WORKSPACES = Path(os.environ.get("WORKSPACE_DIR", "data/workspaces"))
WORKSPACES.mkdir(parents=True, exist_ok=True)

# Single server-side token for Ollama Cloud (https://ollama.com). Falls back
# to the key provided at setup time so this runs out of the box; override by
# setting OLLAMA_API_KEY in the environment (preferred for anything but a
# quick local test, since env vars don't end up committed to source control).
OLLAMA_API_KEY = os.environ.get("OLLAMA_API_KEY", "38a2805ec9ba40abb2cfbece6d81b664.fcj4jrAZ0Vz8FVPPU3OF3joq").strip()
# Ollama Cloud uses its native /api/chat shape, not the OpenAI-style /v1 route.
OLLAMA_CHAT_URL = "https://ollama.com/api/chat"

# Ollama Cloud's hosted catalogue. gpt-oss:20b is the default: it's a strong,
# fast open-weight instruction/coding model sized to run well on the cloud
# tier without the latency of the much larger 120b/671b models below.
MODELS = [
    {"id": "gpt-oss:20b", "name": "GPT-OSS 20B", "family": "OpenAI OSS", "tag": "Recommended · fast & capable"},
    {"id": "gpt-oss:120b", "name": "GPT-OSS 120B", "family": "OpenAI OSS", "tag": "Larger · slower · stronger reasoning"},
    {"id": "qwen3:32b", "name": "Qwen3 32B", "family": "Qwen", "tag": "Strong general & code"},
    {"id": "deepseek-v3.1:671b", "name": "DeepSeek V3.1 671B", "family": "DeepSeek", "tag": "Largest · slowest · frontier-scale"},
]
DEFAULT_MODEL = MODELS[0]["id"]

SYSTEM = '''You are Forge, an expert software and artifact builder. Turn the request into a concise response plus files. Respond with ONLY valid JSON, no prose before or after it, no markdown code fences, using this schema:
{"reply":"short helpful Markdown response","files":[{"path":"safe relative filename.ext","kind":"text|docx|xlsx|pptx|pdf|stl|base64","content":"content for artifact"}]}
Create useful, complete files. Use text for code, HTML, CSS, JSON, CSV, SVG, Markdown and arbitrary plain text. For docx/pdf use paragraphs separated by blank lines. For xlsx use JSON rows like [["Header"],["value"]]. For pptx use JSON slides like [{"title":"...","body":"..."}]. For stl use a JSON shape: {"shape":"cube|pyramid","size":20}; use base64 only for true binary payloads. If the request only needs a text answer, return an empty files list. Never use absolute paths, traversal, or more than 12 files.'''



def decode_model_result(content):
    """Accept strict JSON, fenced JSON, and imperfect free-model output."""
    text = str(content or "").strip()
    if not text:
        raise ValueError("The selected model returned an empty response. Try another free model or retry.")
    candidates = [text]
    fenced = re.search(r"```(?:json)?\s*([\s\S]*?)\s*```", text, re.IGNORECASE)
    if fenced: candidates.append(fenced.group(1))
    start, end = text.find("{"), text.rfind("}")
    if start >= 0 and end > start: candidates.append(text[start:end + 1])
    for candidate in candidates:
        try:
            parsed = json.loads(candidate)
            if isinstance(parsed, dict): return parsed
        except json.JSONDecodeError:
            continue
    # Free models sometimes ignore structured-output instructions. Preserve their work.
    return {"reply": "The model returned unstructured output, saved below.", "files": [{"path": "generation.md", "kind": "text", "content": text}]}


def safe_path(value):
    path = PurePosixPath(value.replace("\\", "/"))
    if path.is_absolute() or ".." in path.parts or not path.parts or path.name in ("", "."):
        raise ValueError("Unsafe output filename")
    return path


def write_artifact(root, item):
    path = root / safe_path(item["path"])
    path.parent.mkdir(parents=True, exist_ok=True)
    kind, content = item.get("kind", "text"), item.get("content", "")
    if kind == "text": path.write_text(str(content), encoding="utf-8")
    elif kind == "base64": path.write_bytes(base64.b64decode(content))
    elif kind == "docx":
        doc = Document()
        for block in str(content).split("\n\n"): doc.add_paragraph(block)
        doc.save(path)
    elif kind == "xlsx":
        wb = Workbook(); sheet = wb.active; sheet.title = "Sheet1"
        for row in json.loads(content): sheet.append(row if isinstance(row, list) else [row])
        wb.save(path)
    elif kind == "pptx":
        pres = Presentation()
        for slide_data in json.loads(content):
            slide = pres.slides.add_slide(pres.slide_layouts[1])
            slide.shapes.title.text = slide_data.get("title", "Untitled")
            slide.placeholders[1].text = slide_data.get("body", "")
        pres.save(path)
    elif kind == "pdf":
        pdf = canvas.Canvas(str(path), pagesize=letter); y = 750
        for line in str(content).splitlines() or [""]:
            if y < 50: pdf.showPage(); y = 750
            pdf.drawString(54, y, line[:110]); y -= 18
        pdf.save()
    elif kind == "stl":
        spec = json.loads(content); size = float(spec.get("size", 20)); h = size / 2
        if spec.get("shape") == "pyramid": vertices = [(-h,-h,0),(h,-h,0),(h,h,0),(-h,h,0),(0,0,size)] ; faces = [(0,1,2),(0,2,3),(0,1,4),(1,2,4),(2,3,4),(3,0,4)]
        else: vertices = [(-h,-h,-h),(h,-h,-h),(h,h,-h),(-h,h,-h),(-h,-h,h),(h,-h,h),(h,h,h),(-h,h,h)]; faces = [(0,2,1),(0,3,2),(4,5,6),(4,6,7),(0,1,5),(0,5,4),(1,2,6),(1,6,5),(2,3,7),(2,7,6),(3,0,4),(3,4,7)]
        lines = ["solid forge"]
        for a,b,c in faces:
            lines += [" facet normal 0 0 0", "  outer loop"] + [f"   vertex {vertices[i][0]} {vertices[i][1]} {vertices[i][2]}" for i in (a,b,c)] + ["  endloop", " endfacet"]
        path.write_text("\n".join(lines + ["endsolid forge"]), encoding="ascii")
    else: raise ValueError(f"Unsupported artifact kind: {kind}")


@app.get("/")
def index(): return render_template("index.html")

@app.get("/api/models")
def models():
    return jsonify(MODELS)


@app.post("/api/chat")
def chat():
    if not OLLAMA_API_KEY:
        return jsonify(error="Forge isn't configured yet: set the OLLAMA_API_KEY environment variable on the server to an Ollama Cloud API key (ollama.com/settings/keys), then restart."), 500
    data = request.get_json(force=True); prompt = str(data.get("prompt", "")).strip()
    if not prompt: return jsonify(error="Enter a request."), 400
    messages = [{"role":"system","content":SYSTEM}] + data.get("history", [])[-10:] + [{"role":"user","content":prompt}]
    model_id = data.get("model") or DEFAULT_MODEL
    # Ollama's native /api/chat shape differs from OpenAI-style APIs: no
    # response_format, generation options nest under "options", and a
    # non-streaming call needs "stream": false or it returns line-delimited
    # JSON chunks instead of one object.
    payload = {"model": model_id, "messages": messages, "stream": False, "options": {"temperature": 0.35, "num_predict": 4096}}
    try:
        response = requests.post(OLLAMA_CHAT_URL, headers={"Authorization": f"Bearer {OLLAMA_API_KEY}"}, json=payload, timeout=150)
        if response.status_code == 401:
            return jsonify(error="Ollama Cloud rejected the API key. Check OLLAMA_API_KEY on the server."), 502
        if response.status_code == 429:
            return jsonify(error=f"{model_id} is rate-limited on Ollama Cloud right now. Wait a bit or switch models."), 502
        response.raise_for_status(); body = response.json()
        if body.get("done") and body.get("done_reason") == "length":
            # The model hit num_predict and cut off mid-generation. Surfacing
            # this explicitly is clearer than showing a silently truncated reply.
            return jsonify(error="The model ran out of room before finishing its response. Try a shorter request, break it into steps, or switch to a different model."), 502
        result = decode_model_result(body["message"]["content"])
        files = result.get("files", [])[:12]; workspace_id = uuid.uuid4().hex; root = WORKSPACES / workspace_id; root.mkdir()
        for item in files: write_artifact(root, item)
        manifest = [{"path":str(p.relative_to(root)).replace("\\", "/"), "bytes":p.stat().st_size} for p in root.rglob("*") if p.is_file()]
        return jsonify(reply=result.get("reply", "Done."), workspace=workspace_id, files=manifest)
    except requests.HTTPError as error:
        # The provider message is useful to the owner but must never include request headers/tokens.
        detail = error.response.text[:500] if error.response is not None else str(error)
        return jsonify(error=f"Ollama Cloud rejected this request ({error.response.status_code}). {detail}"), 502
    except (requests.RequestException, KeyError, IndexError, json.JSONDecodeError, ValueError) as error:
        return jsonify(error=f"Generation failed: {error}"), 502


@app.get("/api/download/<workspace_id>")
def download(workspace_id):
    if not re.fullmatch(r"[a-f0-9]{32}", workspace_id): abort(404)
    root = WORKSPACES / workspace_id
    if not root.is_dir(): abort(404)
    payload = io.BytesIO()
    with zipfile.ZipFile(payload, "w", zipfile.ZIP_DEFLATED) as archive:
        for file in root.rglob("*"):
            if file.is_file(): archive.write(file, file.relative_to(root))
    payload.seek(0)
    return send_file(payload, as_attachment=True, download_name=f"forge-{workspace_id[:8]}.zip", mimetype="application/zip")


@app.get("/api/download/<workspace_id>/<path:filename>")
def download_single(workspace_id, filename):
    if not re.fullmatch(r"[a-f0-9]{32}", workspace_id): abort(404)
    root = WORKSPACES / workspace_id
    if not root.is_dir(): abort(404)
    try:
        target = (root / safe_path(filename)).resolve()
    except ValueError:
        abort(404)
    root_resolved = root.resolve()
    if root_resolved not in target.parents or not target.is_file(): abort(404)
    return send_file(target, as_attachment=True, download_name=target.name)


if __name__ == "__main__": app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)), debug=True)
