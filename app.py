"""Forge — an OpenRouter-powered, downloadable file workspace."""
import base64
import io
import json
import os
import re
import uuid
import zipfile
from datetime import datetime, timezone
from pathlib import Path, PurePosixPath
from urllib.parse import urlparse

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

FALLBACK_MODELS = [
    {"id": "openrouter/free", "name": "Free Models Router", "family": "OpenRouter", "tag": "Free · automatic routing", "source": "openrouter"},
    {"id": "poolside/laguna-s-2.1:free", "name": "Poolside Laguna S 2.1", "family": "Poolside", "tag": "Free · coding", "source": "openrouter"},
]

SYSTEM = '''You are Forge, an expert software and artifact builder. Turn the request into a concise response plus files. Return ONLY valid JSON using this schema:
{"reply":"short helpful Markdown response","files":[{"path":"safe relative filename.ext","kind":"text|docx|xlsx|pptx|pdf|stl|base64","content":"content for artifact"}]}
Create useful, complete files. Use text for code, HTML, CSS, JSON, CSV, SVG, Markdown and arbitrary plain text. For docx/pdf use paragraphs separated by blank lines. For xlsx use JSON rows like [["Header"],["value"]]. For pptx use JSON slides like [{"title":"...","body":"..."}]. For stl use a JSON shape: {"shape":"cube|pyramid","size":20}; use base64 only for true binary payloads. Never use absolute paths, traversal, or more than 12 files.'''


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
    output = [{"id": "openrouter/free", "name": "Free Models Router", "family": "OpenRouter", "tag": "Free · automatic routing", "source": "openrouter"}]
    try:
        response = requests.get("https://openrouter.ai/api/v1/models", timeout=8)
        response.raise_for_status()
        preferred = [m for m in response.json()["data"] if "text" in m.get("architecture",{}).get("output_modalities",["text"]) and (is_free(m) or m["id"].startswith("poolside/"))]
        def price_tag(model):
            pricing = model.get("pricing", {})
            values = [pricing.get(field, "0") for field in ("prompt", "completion", "request")]
            return "Free" if all(str(value) in ("0", "0.0", "0.00") for value in values) else "Paid"
        output.extend([{ "id":m["id"], "name":m["name"], "family":m["id"].split("/")[0].title(), "tag":f"{price_tag(m)} · {m.get('context_length',0)//1000}K context", "source":"openrouter"} for m in preferred[:120]])
    except requests.RequestException: output.extend(FALLBACK_MODELS[1:])
    ollama_url = request.args.get("ollama_url", "").rstrip("/")
    if valid_ollama_url(ollama_url):
        try:
            ollama = requests.get(f"{ollama_url}/api/tags", timeout=3).json().get("models", [])
            output.extend({"id": model["name"], "name": model["name"], "family": "Ollama", "tag": "Local · free", "source": "ollama"} for model in ollama)
        except requests.RequestException: pass
    return jsonify(output)


def is_free(model):
    pricing = model.get("pricing", {})
    return all(str(pricing.get(field, "0")) in ("0", "0.0", "0.00") for field in ("prompt", "completion", "request"))


def valid_ollama_url(value):
    parsed = urlparse(value)
    return parsed.scheme in ("http", "https") and bool(parsed.netloc) and len(value) < 200


@app.post("/api/chat")
def chat():
    data = request.get_json(force=True); prompt = str(data.get("prompt", "")).strip()
    if not prompt: return jsonify(error="Enter a request."), 400
    source = data.get("source", "openrouter")
    key = str(data.get("apiKey", "")).strip()
    if source == "openrouter" and not key: return jsonify(error="Add your OpenRouter API key with the API key button."), 400
    messages = [{"role":"system","content":SYSTEM}] + data.get("history", [])[-10:] + [{"role":"user","content":prompt}]
    payload = {"model": data.get("model") or FALLBACK_MODELS[0]["id"], "messages": messages, "temperature": 0.35, "response_format":{"type":"json_object"}}
    try:
        if source == "ollama":
            ollama_url = str(data.get("ollamaUrl", "")).rstrip("/")
            if not valid_ollama_url(ollama_url): return jsonify(error="Enter a valid Ollama server URL."), 400
            response = requests.post(f"{ollama_url}/api/chat", json={"model": payload["model"], "messages": messages, "stream": False, "format": "json"}, timeout=120)
            response.raise_for_status(); result = json.loads(response.json()["message"]["content"])
        else:
            response = requests.post("https://openrouter.ai/api/v1/chat/completions", headers={"Authorization":f"Bearer {key}","HTTP-Referer":request.host_url,"X-Title":"Forge"}, json=payload, timeout=120)
            response.raise_for_status(); result = json.loads(response.json()["choices"][0]["message"]["content"])
        files = result.get("files", [])[:12]; workspace_id = uuid.uuid4().hex; root = WORKSPACES / workspace_id; root.mkdir()
        for item in files: write_artifact(root, item)
        manifest = [{"path":str(p.relative_to(root)).replace("\\", "/"), "bytes":p.stat().st_size} for p in root.rglob("*") if p.is_file()]
        return jsonify(reply=result.get("reply", "Done."), workspace=workspace_id, files=manifest)
    except requests.HTTPError as error:
        # The provider message is useful to the owner but must never include request headers/API keys.
        detail = error.response.text[:500] if error.response is not None else str(error)
        return jsonify(error=f"OpenRouter rejected this request ({error.response.status_code}). {detail}"), 502
    except (requests.RequestException, KeyError, json.JSONDecodeError, ValueError) as error:
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


if __name__ == "__main__": app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)), debug=True)
