"""Forge (Gen 2) — an Ollama Cloud-powered, downloadable file workspace."""
import base64
import io
import json
import math
import mimetypes
import os
import re
import time
import uuid
import zipfile
from datetime import datetime, timezone
from pathlib import Path, PurePosixPath
from urllib.parse import quote

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
WORKSPACE_MAX_AGE_SECONDS = 2 * 60 * 60  # ephemeral disk: prune old workspaces so it never fills up

# Single server-side token for Ollama Cloud (https://ollama.com). Falls back
# to the key provided at setup time so this runs out of the box; override by
# setting OLLAMA_API_KEY in the environment (preferred for anything but a
# quick local test, since env vars don't end up committed to source control).
OLLAMA_API_KEY = os.environ.get("OLLAMA_API_KEY", "38a2805ec9ba40abb2cfbece6d81b664.fcj4jrAZ0Vz8FVPPU3OF3joq").strip()
# Ollama Cloud uses its native /api/chat shape, not the OpenAI-style /v1 route.
OLLAMA_CHAT_URL = "https://ollama.com/api/chat"

# Pollinations.ai — free, keyless text-to-image API. Used for the "image" file kind.
POLLINATIONS_URL = "https://image.pollinations.ai/prompt/{prompt}"

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
IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg", ".webp", ".gif", ".svg"}

SYSTEM = '''You are Forge, an expert software and artifact builder. Turn the request into a concise response plus files. Respond with ONLY valid JSON, no prose before or after it, no markdown code fences, using this schema:
{"reply":"short helpful Markdown response","files":[{"path":"safe relative filename.ext","kind":"text|docx|xlsx|pptx|pdf|stl|image|base64","content":"content for artifact"}]}
Create useful, complete files. Use text for code, HTML, CSS, JSON, CSV, SVG (vector images), Markdown and arbitrary plain text.
For docx/pdf use paragraphs separated by blank lines. For xlsx use JSON rows like [["Header"],["value"]]. For pptx use JSON slides like [{"title":"...","body":"..."}].
For a raster/photographic image, use kind "image" with a .png/.jpg path; content must be ONLY a vivid, detailed English image-generation prompt describing the picture (no JSON, no extra commentary) — it is rendered by an external image model.

For a 3D model, use kind "stl" with a .stl path. content is JSON describing a short BUILD PROGRAM that Forge parses and executes step by step — you are writing code, not drawing:
{"plan":"1-2 sentences: what real-world parts does this object have, and how do they connect?","ops":[
  {"op":"add","shape":"box|sphere|cylinder|cone|torus|pyramid","size":20,"radius":10,"height":20,"tube":4,"segments":16,"position":[x,y,z],"rotation":[rx,ry,rz],"scale":[sx,sy,sz]},
  {"op":"repeat","count":6,"rotate":[0,0,60],"around":[0,0,0]}
]}
Always write "plan" first and actually design the object as distinct parts before listing ops — do not default to one bare primitive.
Shape params: "box" size is [w,d,h] (or one number for a cube); "sphere"/"cylinder"/"cone" use "radius" (+"height" for cylinder/cone); "torus" uses "radius" (ring radius) and "tube" (tube thickness); "pyramid" uses "size" (+optional "height"). "segments" (8-48) controls roundness, default 16.
Every shape is built centered on its own origin, then: scaled by "scale" [sx,sy,sz] (stretches it, e.g. a sphere into an egg or a cylinder into a plank), then rotated by "rotation" [rx,ry,rz] in degrees (X then Y then Z), then moved to "position" [x,y,z]. All optional, default no scale/rotation and position [0,0,0].
"repeat" duplicates the shape from the immediately preceding "add" op "count"-1 more times: "rotate":[rx,ry,rz] rotates each successive copy by that many more degrees around the "around" pivot point (default world origin) — use for radial patterns like gear teeth, wheel spokes, or flower petals; "translate":[dx,dy,dz] offsets each successive copy further along that vector — use for linear patterns like fence posts, stairs, or table legs. Combine both for spirals.
Compose real objects from several add/repeat ops (roughly 4-14 total): e.g. a table = one flat box top + 4 cylinder legs via one add plus one repeat with translate; a gear = a short cylinder body + one tooth box positioned at its edge + a 12x repeat rotating around the center; a rocket = a tall cylinder body + a cone nose on top + 3-4 fin boxes near the base via one add plus a 4x repeat rotating around the body's own axis. Keep every coordinate within roughly -200..200.

Use base64 only for true binary payloads that don't fit the kinds above. If the request only needs a text answer, return an empty files list. Never use absolute paths, traversal, or more than 12 files.'''



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
    if len(path.parts) > 1 and re.match(r"^[a-zA-Z]:$", path.parts[0]):
        raise ValueError("Unsafe output filename")  # reject Windows-style drive prefixes too
    return path


# ---- Parametric solid-build engine for the "stl" kind ---------------------
# Rather than trust free models to emit raw, hand-rolled vertex/face lists
# (which are easy to get non-manifold or malformed), Forge exposes a small
# instruction set — add a primitive, repeat it with a rotation/translation —
# and executes that program itself. The model writes the build steps; Forge
# turns them into real, valid geometry.
MAX_TRIANGLES = 150_000  # safety cap so a runaway program can't hang the worker or produce a useless file


def _clamp_segments(value, lo=6, hi=48):
    try: return max(lo, min(int(value), hi))
    except (TypeError, ValueError): return 16


def _box_triangles(size):
    w, d, h = ((size, size, size) if not isinstance(size, (list, tuple)) else (list(size) + [size[0] if size else 20]*3)[:3])
    w, d, h = float(w), float(d), float(h)
    hw, hd, hh = w/2, d/2, h/2
    v = [(-hw,-hd,-hh),(hw,-hd,-hh),(hw,hd,-hh),(-hw,hd,-hh),(-hw,-hd,hh),(hw,-hd,hh),(hw,hd,hh),(-hw,hd,hh)]
    faces = [(0,2,1),(0,3,2),(4,5,6),(4,6,7),(0,1,5),(0,5,4),(1,2,6),(1,6,5),(2,3,7),(2,7,6),(3,0,4),(3,4,7)]
    return [(v[a], v[b], v[c]) for a, b, c in faces]


def _pyramid_triangles(size, height=None):
    h = float(size) / 2; height = float(height) if height is not None else float(size); z0, z1 = -height/2, height/2
    v = [(-h,-h,z0),(h,-h,z0),(h,h,z0),(-h,h,z0),(0,0,z1)]
    faces = [(0,2,1),(0,3,2),(0,1,4),(1,2,4),(2,3,4),(3,0,4)]
    return [(v[a], v[b], v[c]) for a, b, c in faces]


def _sphere_triangles(radius, segments=16):
    segments = _clamp_segments(segments); stacks = max(4, segments // 2)
    tris = []
    for i in range(stacks):
        lat0 = math.pi * (-0.5 + i / stacks); lat1 = math.pi * (-0.5 + (i + 1) / stacks)
        for j in range(segments):
            lon0 = 2 * math.pi * j / segments; lon1 = 2 * math.pi * (j + 1) / segments
            def pt(lat, lon): return (radius*math.cos(lat)*math.cos(lon), radius*math.cos(lat)*math.sin(lon), radius*math.sin(lat))
            p00, p01, p10, p11 = pt(lat0,lon0), pt(lat0,lon1), pt(lat1,lon0), pt(lat1,lon1)
            if i != 0: tris.append((p00, p01, p11))
            if i != stacks - 1: tris.append((p00, p11, p10))
    return tris


def _cylinder_triangles(radius, height, segments=16):
    segments = _clamp_segments(segments); h = height / 2
    tris = []
    for j in range(segments):
        a0, a1 = 2*math.pi*j/segments, 2*math.pi*(j+1)/segments
        x0,y0,x1,y1 = radius*math.cos(a0), radius*math.sin(a0), radius*math.cos(a1), radius*math.sin(a1)
        top0,top1,bot0,bot1 = (x0,y0,h),(x1,y1,h),(x0,y0,-h),(x1,y1,-h)
        tris += [(bot0,bot1,top1),(bot0,top1,top0),(top0,top1,(0,0,h)),(bot1,bot0,(0,0,-h))]
    return tris


def _cone_triangles(radius, height, segments=16):
    segments = _clamp_segments(segments); apex = (0,0,height/2); h = height/2
    tris = []
    for j in range(segments):
        a0, a1 = 2*math.pi*j/segments, 2*math.pi*(j+1)/segments
        x0,y0,x1,y1 = radius*math.cos(a0), radius*math.sin(a0), radius*math.cos(a1), radius*math.sin(a1)
        base0, base1 = (x0,y0,-h), (x1,y1,-h)
        tris += [(base0, base1, apex), (base1, base0, (0,0,-h))]
    return tris


def _torus_triangles(major_radius, tube_radius, segments=24, tube_segments=12):
    segments = _clamp_segments(segments, 8, 48); tube_segments = _clamp_segments(tube_segments, 6, 32)
    tris = []
    def pt(u, v): return ((major_radius+tube_radius*math.cos(v))*math.cos(u), (major_radius+tube_radius*math.cos(v))*math.sin(u), tube_radius*math.sin(v))
    for i in range(segments):
        u0, u1 = 2*math.pi*i/segments, 2*math.pi*(i+1)/segments
        for j in range(tube_segments):
            v0, v1 = 2*math.pi*j/tube_segments, 2*math.pi*(j+1)/tube_segments
            p00, p01, p10, p11 = pt(u0,v0), pt(u0,v1), pt(u1,v0), pt(u1,v1)
            tris += [(p00, p10, p11), (p00, p11, p01)]
    return tris


def _shape_radius(s, default=10):
    if "radius" in s: return float(s["radius"])
    if "size" in s and not isinstance(s["size"], (list, tuple)): return float(s["size"]) / 2
    return float(default)


def build_local_shape(spec):
    """Build a shape centered on its own local origin, unrotated/unscaled/unplaced."""
    shape = spec.get("shape", "box")
    if shape in ("box", "cube"): return _box_triangles(spec.get("size", 20))
    if shape == "pyramid": return _pyramid_triangles(spec.get("size", 20), spec.get("height"))
    if shape == "sphere": return _sphere_triangles(_shape_radius(spec), spec.get("segments", 16))
    if shape == "cylinder": return _cylinder_triangles(_shape_radius(spec), float(spec.get("height", spec.get("size", 20))), spec.get("segments", 16))
    if shape == "cone": return _cone_triangles(_shape_radius(spec), float(spec.get("height", spec.get("size", 20))), spec.get("segments", 16))
    if shape == "torus": return _torus_triangles(float(spec.get("radius", 20)), float(spec.get("tube", spec.get("minor_radius", 5))), spec.get("segments", 24), spec.get("tube_segments", 12))
    raise ValueError(f"Unknown shape '{shape}'")


def _rotate_point(p, rotation_deg):
    x, y, z = p
    rx, ry, rz = (math.radians(v) for v in rotation_deg)
    y, z = y*math.cos(rx)-z*math.sin(rx), y*math.sin(rx)+z*math.cos(rx)
    x, z = x*math.cos(ry)+z*math.sin(ry), -x*math.sin(ry)+z*math.cos(ry)
    x, y = x*math.cos(rz)-y*math.sin(rz), x*math.sin(rz)+y*math.cos(rz)
    return (x, y, z)


def _place_triangles(tris, scale=(1,1,1), rotation=(0,0,0), position=(0,0,0)):
    sx, sy, sz = scale
    out = []
    for tri in tris:
        placed = []
        for (x, y, z) in tri:
            x, y, z = x*sx, y*sy, z*sz
            x, y, z = _rotate_point((x, y, z), rotation)
            placed.append((x+position[0], y+position[1], z+position[2]))
        out.append(tuple(placed))
    return out


def _rotate_triangles_around(tris, rotation_deg, pivot):
    px, py, pz = pivot
    out = []
    for tri in tris:
        rotated = []
        for (x, y, z) in tri:
            rx, ry, rz = _rotate_point((x-px, y-py, z-pz), rotation_deg)
            rotated.append((rx+px, ry+py, rz+pz))
        out.append(tuple(rotated))
    return out


def _vec3(value, default=(0.0, 0.0, 0.0)):
    if not value: return default
    values = list(value) + list(default)
    return tuple(float(v) for v in values[:3])


def run_stl_program(spec):
    """Interpret the model's ordered build steps ("ops") into world-space
    triangles. Supports "add" (place a primitive, optionally scaled/rotated)
    and "repeat" (duplicate the previous add with a cumulative rotation
    and/or translation per copy — radial or linear patterns)."""
    ops = spec.get("ops") if isinstance(spec, dict) else None
    if not ops:
        # Back-compat with the earlier, simpler schemas.
        if isinstance(spec, dict) and spec.get("shapes"): ops = [{"op": "add", **item} for item in spec["shapes"]]
        elif isinstance(spec, dict) and spec.get("shape"): ops = [{"op": "add", **spec}]
        else: raise ValueError("STL spec has no ops/shapes/shape to build from")

    triangles, last_placed = [], None
    for op in ops[:80]:
        kind = op.get("op", "add")
        if kind == "add":
            local = build_local_shape(op)
            placed = _place_triangles(local, _vec3(op.get("scale"), (1, 1, 1)), _vec3(op.get("rotation")), _vec3(op.get("position")))
            triangles += placed
            last_placed = placed
        elif kind == "repeat" and last_placed:
            count = max(1, min(int(op.get("count", 1)), 60))
            translate_step, rotate_step, pivot = _vec3(op.get("translate")), _vec3(op.get("rotate")), _vec3(op.get("around"))
            for i in range(1, count):
                step = last_placed
                if any(rotate_step): step = _rotate_triangles_around(step, tuple(a*i for a in rotate_step), pivot)
                if any(translate_step):
                    dx, dy, dz = (a*i for a in translate_step)
                    step = [tuple((x+dx, y+dy, z+dz) for x, y, z in tri) for tri in step]
                triangles += step
        if len(triangles) > MAX_TRIANGLES:
            raise ValueError("That design is too complex to build (too many triangles) — simplify it")
    if not triangles:
        raise ValueError("STL program produced no geometry")
    return triangles


def render_ascii_stl(triangles):
    lines = ["solid forge"]
    for a, b, c in triangles:
        ax, ay, az = a; bx, by, bz = b; cx, cy, cz = c
        ux, uy, uz = bx-ax, by-ay, bz-az
        vx, vy, vz = cx-ax, cy-ay, cz-az
        nx, ny, nz = uy*vz-uz*vy, uz*vx-ux*vz, ux*vy-uy*vx
        length = math.sqrt(nx*nx+ny*ny+nz*nz) or 1.0
        lines += [f" facet normal {nx/length:.6f} {ny/length:.6f} {nz/length:.6f}", "  outer loop"]
        lines += [f"   vertex {p[0]:.4f} {p[1]:.4f} {p[2]:.4f}" for p in (a, b, c)]
        lines += ["  endloop", " endfacet"]
    lines.append("endsolid forge")
    return "\n".join(lines)


def write_artifact(root, item):
    path = root / safe_path(item["path"])
    path.parent.mkdir(parents=True, exist_ok=True)
    kind, content = item.get("kind", "text"), item.get("content", "")
    if kind == "text": path.write_text(str(content), encoding="utf-8")
    elif kind == "base64": path.write_bytes(base64.b64decode(content))
    elif kind == "image":
        prompt = str(content).strip()[:800] or "abstract art"
        url = POLLINATIONS_URL.format(prompt=quote(prompt)) + "?width=1024&height=1024&nologo=true&model=flux"
        response = requests.get(url, timeout=90)
        response.raise_for_status()
        if not response.headers.get("content-type", "").startswith("image/") and len(response.content) < 500:
            raise ValueError("Image generation did not return an image")
        path.write_bytes(response.content)
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
        spec = json.loads(content)
        path.write_text(render_ascii_stl(run_stl_program(spec)), encoding="ascii")
    else: raise ValueError(f"Unsupported artifact kind: {kind}")


def prune_old_workspaces():
    """Ephemeral disk hygiene: delete workspace folders older than the cutoff
    so a long-running process never silently fills its disk with old ZIPs."""
    cutoff = time.time() - WORKSPACE_MAX_AGE_SECONDS
    try:
        for entry in WORKSPACES.iterdir():
            if entry.is_dir() and entry.stat().st_mtime < cutoff:
                for f in sorted(entry.rglob("*"), reverse=True):
                    (f.rmdir() if f.is_dir() else f.unlink())
                entry.rmdir()
    except OSError:
        pass  # best-effort cleanup; never let this break a request


@app.get("/")
def index(): return render_template("index.html")

@app.get("/api/models")
def models():
    return jsonify(MODELS)


@app.post("/api/chat")
def chat():
    if not OLLAMA_API_KEY:
        return jsonify(error="Forge isn't configured yet: set the OLLAMA_API_KEY environment variable on the server to an Ollama Cloud API key (ollama.com/settings/keys), then restart."), 500
    # get_json(force=True) raises Flask's own HTML 400 page on a malformed
    # body, which broke the frontend's JSON parsing. silent=True + a manual
    # check keeps every response on this route JSON, even for bad input.
    data = request.get_json(silent=True)
    if not isinstance(data, dict):
        return jsonify(error="Malformed request body."), 400
    prompt = str(data.get("prompt", "")).strip()
    if not prompt: return jsonify(error="Enter a request."), 400
    prune_old_workspaces()
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
        manifest = sorted(
            [{"path": str(p.relative_to(root)).replace("\\", "/"), "bytes": p.stat().st_size,
              "isImage": p.suffix.lower() in IMAGE_EXTENSIONS}
             for p in root.rglob("*") if p.is_file()],
            key=lambda f: f["path"],
        )
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


def resolve_workspace_file(workspace_id, filename):
    if not re.fullmatch(r"[a-f0-9]{32}", workspace_id): abort(404)
    root = WORKSPACES / workspace_id
    if not root.is_dir(): abort(404)
    try:
        target = (root / safe_path(filename)).resolve()
    except ValueError:
        abort(404)
    if root.resolve() not in target.parents or not target.is_file(): abort(404)
    return target


@app.get("/api/download/<workspace_id>/<path:filename>")
def download_single(workspace_id, filename):
    target = resolve_workspace_file(workspace_id, filename)
    return send_file(target, as_attachment=True, download_name=target.name)


@app.get("/api/preview/<workspace_id>/<path:filename>")
def preview_single(workspace_id, filename):
    # Same safety checks as the download route, but served inline (not as an
    # attachment) with a guessed mimetype, so <img> tags can render it directly.
    target = resolve_workspace_file(workspace_id, filename)
    mimetype = mimetypes.guess_type(target.name)[0] or "application/octet-stream"
    return send_file(target, as_attachment=False, mimetype=mimetype)


if __name__ == "__main__": app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)), debug=True)
