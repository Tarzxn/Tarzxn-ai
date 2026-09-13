"""Forge (Gen 2) — an Ollama Cloud-powered, downloadable file workspace."""
import base64
import io
import json
import math
import mimetypes
import os
import re
import textwrap
import time
import uuid
import zipfile
from datetime import datetime, timezone
from pathlib import Path, PurePosixPath
from urllib.parse import quote

import requests
from flask import Flask, abort, jsonify, render_template, request, send_file
from docx import Document
from docx.shared import Pt
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
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

# Tavily — web search, used to ground answers in current information before
# the model responds. No key is baked in (unlike Ollama) because none was
# provided; set TAVILY_API_KEY in the environment to enable the "Web search"
# toggle in the composer. Get a free key at https://app.tavily.com.
TAVILY_API_KEY = os.environ.get("TAVILY_API_KEY", "").strip()
TAVILY_SEARCH_URL = "https://api.tavily.com/search"

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

For docx/pdf, content is Markdown-lite: lines starting "# "/"## "/"### " become headings, lines starting "- " become bullets, and **bold** spans are rendered bold; separate paragraphs with blank lines.
For xlsx, content is JSON rows like [["Header1","Header2"],["value",1]] — the first row is treated as a header and gets bold styling and auto-sized columns automatically.
For pptx, content is JSON slides like [{"title":"...","body":"one bullet per line, separated by \\n"}] — each line in "body" becomes its own bullet point.
For a raster/photographic image, use kind "image" with a .png/.jpg path. content is either a plain English image-generation prompt, or JSON {"prompt":"...","aspect":"square|portrait|landscape"} for more control over framing — use vivid, specific, detailed prompts.

For a 3D model, use kind "stl" with a .stl path. Take real time to think this through — you are the CAD engineer: mentally model the object as an assembly of real, distinct parts and their spatial relationships before writing anything. content is JSON describing a BUILD PROGRAM that Forge parses and executes step by step:
{"plan":"a few sentences: what real-world parts does this object have, roughly what size is each, and how do they connect/align?","ops":[
  {"op":"add","shape":"box|sphere|cylinder|cone|torus|tube|capsule|wedge|pyramid","size":20,"radius":10,"height":20,"tube":4,"segments":16,"position":[x,y,z],"rotation":[rx,ry,rz],"scale":[sx,sy,sz]},
  {"op":"repeat","count":6,"rotate":[0,0,60],"around":[0,0,0]}
]}
Shape params — "box": size [w,d,h] (or one number for a cube). "sphere"/"cylinder"/"cone": "radius" (+"height" for cylinder/cone). "torus": "radius" (ring) + "tube" (thickness). "tube": a hollow pipe/ring — "radius" (outer) + "inner_radius" + "height". "capsule": a pill shape — "radius" + "height" (straight section length; total length is height + 2*radius). "wedge": a ramp/doorstop/roof — size [w,d,h], sloped down along x. "pyramid": "size" (+optional "height"). "cylinder" with a low "segments" (e.g. 5, 6, 8) becomes a pentagonal/hexagonal/octagonal prism — use this for nuts, bolts, multi-sided posts, etc. instead of a separate prism shape. Leave "segments" unset to let Forge auto-pick a smooth value from the part's size; only set it explicitly for a deliberately low-poly/faceted look.
Every shape is centered on its own local origin, then: scaled by "scale" [sx,sy,sz] (stretch into an ellipsoid, plank, etc.), rotated by "rotation" [rx,ry,rz] degrees (X then Y then Z, e.g. tilt a fin or lay a cylinder on its side), then moved to "position" [x,y,z]. All optional, default no scale/rotation, position [0,0,0].
"repeat" duplicates the shape from the immediately preceding "add" "count"-1 more times: "rotate":[rx,ry,rz] rotates each successive copy further around the "around" pivot (default world origin) — radial patterns (gear teeth, wheel spokes, flower petals, fins around a body). "translate":[dx,dy,dz] offsets each successive copy further along that vector — linear patterns (fence posts, stair treads, table legs, shelf slats). Combine both for a spiral/helix.
Build real objects from several parts (roughly 6-20 ops is normal for something detailed) — e.g. a mug = a "tube" body + a "torus" or bent-"capsule" handle positioned at the side; a table = one flat box top + 4 cylinder legs via one add + one repeat with translate; a gear = a short cylinder body + one tooth box at its edge + a repeat rotating around the center; a rocket = a cylinder body + a cone nose + a capsule or sphere tip + fin boxes via one add + a radial repeat. Prefer the shape that is actually hollow/rounded when the real object is (a cup or pipe should be a "tube", not a solid cylinder; a pill or rounded handle should be a "capsule", not a box). Keep coordinates within roughly -200..200. If one of your ops is invalid Forge will skip just that piece and keep the rest, so don't let one uncertain part stop you from building the others.

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
MAX_TRIANGLES = 260_000  # generous cap (user explicitly OK with slower/bigger builds) so a runaway program still can't hang the worker indefinitely
MAX_OPS = 160
MAX_REPEAT_COUNT = 120


def _clamp_segments(value, lo=6, hi=64):
    try: return max(lo, min(int(round(float(value))), hi))
    except (TypeError, ValueError): return 16


def _auto_segments(size_metric, explicit):
    """When the model doesn't specify a segment count, scale it with the
    part's own size instead of using one fixed default — bigger round parts
    get smoother curves automatically, which reads as far more realistic
    without requiring the model to reason about facet counts itself."""
    if explicit is not None: return _clamp_segments(explicit)
    return _clamp_segments(round(abs(size_metric) * 1.3) + 12, 14, 64)


def _box_triangles(size):
    w, d, h = ((size, size, size) if not isinstance(size, (list, tuple)) else (list(size) + [size[0] if size else 20]*3)[:3])
    w, d, h = float(w), float(d), float(h)
    hw, hd, hh = w/2, d/2, h/2
    v = [(-hw,-hd,-hh),(hw,-hd,-hh),(hw,hd,-hh),(-hw,hd,-hh),(-hw,-hd,hh),(hw,-hd,hh),(hw,hd,hh),(-hw,hd,hh)]
    faces = [(0,2,1),(0,3,2),(4,5,6),(4,6,7),(0,1,5),(0,5,4),(1,2,6),(1,6,5),(2,3,7),(2,7,6),(3,0,4),(3,4,7)]
    return [(v[a], v[b], v[c]) for a, b, c in faces]


def _wedge_triangles(size):
    """A ramp/doorstop/roof shape: a rectangular base tapering up to a ridge
    along one edge, sloped down along x. Useful for ramps, roofs, chocks."""
    w, d, h = ((size, size, size) if not isinstance(size, (list, tuple)) else (list(size) + [size[0] if size else 20]*3)[:3])
    w, d, h = float(w), float(d), float(h)
    hw, hd, hh = w/2, d/2, h/2
    b0,b1,b2,b3 = (-hw,-hd,-hh),(hw,-hd,-hh),(hw,hd,-hh),(-hw,hd,-hh)
    t0,t1 = (-hw,0,hh),(hw,0,hh)
    return [
        (b0,b2,b1),(b0,b3,b2),          # bottom
        (b0,b1,t1),(b0,t1,t0),          # front slope (y=-hd side)
        (b3,t0,t1),(b3,t1,b2),          # back slope (y=+hd side)
        (b0,t0,b3),                     # left end cap
        (b1,b2,t1),                     # right end cap
    ]


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


def _hemisphere_triangles(radius, segments=16, upper=True):
    """Half a sphere, flat/open side on the z=0 plane — used to cap capsules
    so they seal flush against the cylinder body. Unlike a full sphere, only
    ONE end (the pole, away from z=0) collapses to a point; the z=0 ring is
    a full-radius rim and must keep both triangles of every quad so its
    boundary edges exist to seal against the adjoining cylinder wall."""
    segments = _clamp_segments(segments); stacks = max(3, segments // 4)
    sign = 1 if upper else -1
    tris = []
    for i in range(stacks):
        lat0 = sign * (math.pi/2) * (i / stacks); lat1 = sign * (math.pi/2) * ((i + 1) / stacks)
        pole_row = (i == stacks - 1)  # only the far row degenerates to a point
        for j in range(segments):
            lon0 = 2 * math.pi * j / segments; lon1 = 2 * math.pi * (j + 1) / segments
            def pt(lat, lon): return (radius*math.cos(lat)*math.cos(lon), radius*math.cos(lat)*math.sin(lon), radius*math.sin(lat))
            p00, p01, p10, p11 = pt(lat0,lon0), pt(lat0,lon1), pt(lat1,lon0), pt(lat1,lon1)
            if upper:
                tris.append((p00, p01, p11))
                if not pole_row: tris.append((p00, p11, p10))
            else:
                tris.append((p00, p11, p01))
                if not pole_row: tris.append((p00, p10, p11))
    return tris


def _cylinder_side_triangles(radius, height, segments=16):
    segments = _clamp_segments(segments); h = height / 2
    tris = []
    for j in range(segments):
        a0, a1 = 2*math.pi*j/segments, 2*math.pi*(j+1)/segments
        x0,y0,x1,y1 = radius*math.cos(a0), radius*math.sin(a0), radius*math.cos(a1), radius*math.sin(a1)
        top0,top1,bot0,bot1 = (x0,y0,h),(x1,y1,h),(x0,y0,-h),(x1,y1,-h)
        tris += [(bot0,bot1,top1),(bot0,top1,top0)]
    return tris


def _cylinder_triangles(radius, height, segments=16):
    segments = _clamp_segments(segments); h = height / 2
    tris = _cylinder_side_triangles(radius, height, segments)
    for j in range(segments):
        a0, a1 = 2*math.pi*j/segments, 2*math.pi*(j+1)/segments
        x0,y0,x1,y1 = radius*math.cos(a0), radius*math.sin(a0), radius*math.cos(a1), radius*math.sin(a1)
        top0,top1,bot0,bot1 = (x0,y0,h),(x1,y1,h),(x0,y0,-h),(x1,y1,-h)
        tris += [(top0,top1,(0,0,h)), (bot1,bot0,(0,0,-h))]
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
    segments = _clamp_segments(segments, 8, 64); tube_segments = _clamp_segments(tube_segments, 6, 32)
    tris = []
    def pt(u, v): return ((major_radius+tube_radius*math.cos(v))*math.cos(u), (major_radius+tube_radius*math.cos(v))*math.sin(u), tube_radius*math.sin(v))
    for i in range(segments):
        u0, u1 = 2*math.pi*i/segments, 2*math.pi*(i+1)/segments
        for j in range(tube_segments):
            v0, v1 = 2*math.pi*j/tube_segments, 2*math.pi*(j+1)/tube_segments
            p00, p01, p10, p11 = pt(u0,v0), pt(u0,v1), pt(u1,v0), pt(u1,v1)
            tris += [(p00, p10, p11), (p00, p11, p01)]
    return tris


def _tube_triangles(outer_radius, inner_radius, height, segments=16):
    """A hollow pipe/ring/washer: two concentric cylindrical walls joined by
    flat annular caps top and bottom — genuinely hollow, not an approximation."""
    segments = _clamp_segments(segments); inner_radius = max(0.001, min(inner_radius, outer_radius - 0.001)); h = height / 2
    tris = []
    for j in range(segments):
        a0, a1 = 2*math.pi*j/segments, 2*math.pi*(j+1)/segments
        ox0,oy0,ox1,oy1 = outer_radius*math.cos(a0), outer_radius*math.sin(a0), outer_radius*math.cos(a1), outer_radius*math.sin(a1)
        ix0,iy0,ix1,iy1 = inner_radius*math.cos(a0), inner_radius*math.sin(a0), inner_radius*math.cos(a1), inner_radius*math.sin(a1)
        o_top0,o_top1,o_bot0,o_bot1 = (ox0,oy0,h),(ox1,oy1,h),(ox0,oy0,-h),(ox1,oy1,-h)
        i_top0,i_top1,i_bot0,i_bot1 = (ix0,iy0,h),(ix1,iy1,h),(ix0,iy0,-h),(ix1,iy1,-h)
        tris += [(o_bot0,o_bot1,o_top1),(o_bot0,o_top1,o_top0)]           # outer wall
        tris += [(i_bot1,i_bot0,i_top0),(i_bot1,i_top0,i_top1)]           # inner wall (reversed so it faces inward)
        tris += [(o_top0,o_top1,i_top1),(o_top0,i_top1,i_top0)]           # top annulus
        tris += [(o_bot1,o_bot0,i_bot0),(o_bot1,i_bot0,i_bot1)]           # bottom annulus
    return tris


def _capsule_triangles(radius, height=0.0, segments=16):
    """A pill/stadium shape: a straight cylindrical section capped with two
    hemispheres — for handles, pills, rounded rods, fingers, rounded ends."""
    segments = _clamp_segments(segments); half = max(float(height), 0.0) / 2
    tris = _cylinder_side_triangles(radius, height, segments) if height > 0 else []
    tris += [tuple((x, y, z + half) for x, y, z in tri) for tri in _hemisphere_triangles(radius, segments, upper=True)]
    tris += [tuple((x, y, z - half) for x, y, z in tri) for tri in _hemisphere_triangles(radius, segments, upper=False)]
    return tris


def _shape_radius(s, default=10):
    if "radius" in s: return float(s["radius"])
    if "size" in s and not isinstance(s["size"], (list, tuple)): return float(s["size"]) / 2
    return float(default)


def build_local_shape(spec):
    """Build a shape centered on its own local origin, unrotated/unscaled/unplaced."""
    shape = spec.get("shape", "box")
    if shape in ("box", "cube"): return _box_triangles(spec.get("size", 20))
    if shape == "wedge": return _wedge_triangles(spec.get("size", 20))
    if shape == "pyramid": return _pyramid_triangles(spec.get("size", 20), spec.get("height"))
    if shape == "sphere":
        r = _shape_radius(spec); return _sphere_triangles(r, _auto_segments(r, spec.get("segments")))
    if shape == "cylinder":
        r = _shape_radius(spec); return _cylinder_triangles(r, float(spec.get("height", spec.get("size", 20))), _auto_segments(r, spec.get("segments")))
    if shape == "cone":
        r = _shape_radius(spec); return _cone_triangles(r, float(spec.get("height", spec.get("size", 20))), _auto_segments(r, spec.get("segments")))
    if shape == "torus":
        r = float(spec.get("radius", 20)); return _torus_triangles(r, float(spec.get("tube", spec.get("minor_radius", 5))), _auto_segments(r, spec.get("segments")), spec.get("tube_segments", 12))
    if shape == "tube":
        outer = _shape_radius(spec, 15); inner = float(spec.get("inner_radius", spec.get("inner", outer * 0.6)))
        return _tube_triangles(outer, inner, float(spec.get("height", 20)), _auto_segments(outer, spec.get("segments")))
    if shape == "capsule":
        r = _shape_radius(spec, 8); return _capsule_triangles(r, float(spec.get("height", 0)), _auto_segments(r, spec.get("segments")))
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
    and/or translation per copy — radial or linear patterns). Each op is
    executed independently: if one is malformed, it's skipped with a
    recorded warning instead of failing the whole model, so a single bad
    part never throws away an otherwise-good design.
    Returns (triangles, warnings)."""
    ops = spec.get("ops") if isinstance(spec, dict) else None
    if not ops:
        # Back-compat with the earlier, simpler schemas.
        if isinstance(spec, dict) and spec.get("shapes"): ops = [{"op": "add", **item} for item in spec["shapes"]]
        elif isinstance(spec, dict) and spec.get("shape"): ops = [{"op": "add", **spec}]
        else: raise ValueError("STL spec has no ops/shapes/shape to build from")

    triangles, last_placed, warnings = [], None, []
    for index, op in enumerate(ops[:MAX_OPS]):
        kind = op.get("op", "add")
        try:
            if kind == "add":
                local = build_local_shape(op)
                placed = _place_triangles(local, _vec3(op.get("scale"), (1, 1, 1)), _vec3(op.get("rotation")), _vec3(op.get("position")))
                triangles += placed
                last_placed = placed
            elif kind == "repeat":
                if not last_placed: raise ValueError("repeat with nothing preceding it to repeat")
                count = max(1, min(int(op.get("count", 1)), MAX_REPEAT_COUNT))
                translate_step, rotate_step, pivot = _vec3(op.get("translate")), _vec3(op.get("rotate")), _vec3(op.get("around"))
                for i in range(1, count):
                    step = last_placed
                    if any(rotate_step): step = _rotate_triangles_around(step, tuple(a*i for a in rotate_step), pivot)
                    if any(translate_step):
                        dx, dy, dz = (a*i for a in translate_step)
                        step = [tuple((x+dx, y+dy, z+dz) for x, y, z in tri) for tri in step]
                    triangles += step
            else:
                warnings.append(f"Step {index+1}: unknown op '{kind}' — skipped.")
        except (ValueError, TypeError, KeyError, ZeroDivisionError, ArithmeticError) as error:
            warnings.append(f"Step {index+1} ({kind}): {error} — skipped, rest of the model was still built.")
        if len(triangles) > MAX_TRIANGLES:
            raise ValueError("That design is too complex to build (too many triangles) — simplify it")
    if not triangles:
        raise ValueError("STL program produced no geometry")
    return triangles, warnings


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


def _apply_bold_runs(paragraph, text):
    """Split "**bold**" spans out of a line of text and add them as bold runs."""
    for i, chunk in enumerate(re.split(r"\*\*(.+?)\*\*", text)):
        if not chunk: continue
        run = paragraph.add_run(chunk)
        if i % 2 == 1: run.bold = True


def write_artifact(root, item):
    """Writes one artifact to disk. Returns a list of non-fatal warning
    strings (only ever populated for "stl", where a bad build step is
    skipped rather than failing the whole file)."""
    path = root / safe_path(item["path"])
    path.parent.mkdir(parents=True, exist_ok=True)
    kind, content = item.get("kind", "text"), item.get("content", "")
    if kind == "text":
        path.write_text(str(content), encoding="utf-8")
    elif kind == "base64":
        path.write_bytes(base64.b64decode(content))
    elif kind == "image":
        # content is either a plain prompt string, or JSON {"prompt":...,"aspect":...}
        # for finer control over framing.
        prompt, aspect = str(content), "square"
        try:
            parsed = json.loads(content)
            if isinstance(parsed, dict) and "prompt" in parsed:
                prompt, aspect = str(parsed["prompt"]), str(parsed.get("aspect", "square"))
        except (json.JSONDecodeError, ValueError):
            pass
        width, height = {"portrait": (832, 1216), "landscape": (1216, 832)}.get(aspect, (1024, 1024))
        prompt = prompt.strip()[:800] or "abstract art"
        url = POLLINATIONS_URL.format(prompt=quote(prompt)) + f"?width={width}&height={height}&nologo=true&model=flux"
        response = requests.get(url, timeout=90)
        response.raise_for_status()
        if not response.headers.get("content-type", "").startswith("image/") and len(response.content) < 500:
            raise ValueError("Image generation did not return an image")
        path.write_bytes(response.content)
    elif kind == "docx":
        # Lightweight Markdown: "#"-headings, "- " bullets, **bold** spans —
        # instead of dumping everything as identical plain paragraphs.
        doc = Document()
        for block in str(content).split("\n\n"):
            for line in block.split("\n") or [""]:
                stripped = line.strip()
                if not stripped: continue
                heading_match = re.match(r"^(#{1,3})\s+(.*)", stripped)
                if heading_match:
                    doc.add_heading(heading_match.group(2), level=len(heading_match.group(1)))
                elif stripped.startswith("- "):
                    _apply_bold_runs(doc.add_paragraph(style="List Bullet"), stripped[2:])
                else:
                    _apply_bold_runs(doc.add_paragraph(), stripped)
        doc.save(path)
    elif kind == "xlsx":
        wb = Workbook(); sheet = wb.active; sheet.title = "Sheet1"
        rows = json.loads(content)
        for row_index, row in enumerate(rows):
            sheet.append(row if isinstance(row, list) else [row])
            if row_index == 0:
                for cell in sheet[1]: cell.font = Font(bold=True)
        sheet.freeze_panes = "A2"
        widths = {}
        for row in rows:
            for col_index, value in enumerate(row if isinstance(row, list) else [row]):
                widths[col_index] = max(widths.get(col_index, 8), min(len(str(value)) + 2, 40))
        for col_index, width in widths.items():
            sheet.column_dimensions[get_column_letter(col_index + 1)].width = width
        wb.save(path)
    elif kind == "pptx":
        pres = Presentation()
        for slide_data in json.loads(content):
            slide = pres.slides.add_slide(pres.slide_layouts[1])
            slide.shapes.title.text = slide_data.get("title", "Untitled")
            body = slide.placeholders[1].text_frame
            lines = str(slide_data.get("body", "")).split("\n") or [""]
            body.text = lines[0]
            for line in lines[1:]:
                body.add_paragraph().text = line
        pres.save(path)
    elif kind == "pdf":
        pdf = canvas.Canvas(str(path), pagesize=letter); y = 750
        for raw_line in str(content).splitlines() or [""]:
            heading_match = re.match(r"^(#{1,3})\s+(.*)", raw_line.strip())
            font, size, text = ("Helvetica-Bold", 15, heading_match.group(2)) if heading_match else ("Helvetica", 11, raw_line)
            pdf.setFont(font, size)
            wrapped = textwrap.wrap(text, width=95) or [""]
            for line in wrapped:
                if y < 50: pdf.showPage(); pdf.setFont(font, size); y = 750
                pdf.drawString(54, y, line)
                y -= (size + 6)
            if heading_match: y -= 4
        pdf.save()
    elif kind == "stl":
        spec = json.loads(content)
        triangles, warnings = run_stl_program(spec)
        path.write_text(render_ascii_stl(triangles), encoding="ascii")
        return warnings
    else:
        raise ValueError(f"Unsupported artifact kind: {kind}")
    return []


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


def tavily_search(query):
    """Search the live web via Tavily and return a compact text block the
    model can read as extra context. Raises on failure — the caller decides
    whether that should abort the request or just proceed without results."""
    response = requests.post(
        TAVILY_SEARCH_URL,
        headers={"Authorization": f"Bearer {TAVILY_API_KEY}"},
        json={"query": query, "search_depth": "basic", "max_results": 6, "include_answer": True},
        timeout=25,
    )
    response.raise_for_status()
    body = response.json()
    lines = []
    if body.get("answer"): lines.append(f"Summary: {body['answer']}")
    for result in body.get("results", [])[:6]:
        title = str(result.get("title", "")).strip()
        url = str(result.get("url", "")).strip()
        snippet = str(result.get("content", "")).strip()[:500]
        lines.append(f"- {title} ({url}): {snippet}")
    if not lines:
        raise ValueError("Tavily returned no results")
    return "\n".join(lines)


@app.get("/")
def index(): return render_template("index.html")

@app.get("/api/models")
def models():
    return jsonify(MODELS)


@app.get("/api/config")
def config():
    return jsonify(webSearchEnabled=bool(TAVILY_API_KEY))


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

    messages = [{"role": "system", "content": SYSTEM}] + data.get("history", [])[-10:]
    if data.get("web_search"):
        if not TAVILY_API_KEY:
            return jsonify(error="Web search isn't configured yet: set the TAVILY_API_KEY environment variable on the server (get a free key at app.tavily.com), then restart."), 500
        try:
            search_context = tavily_search(prompt)
            messages.append({"role": "system", "content": f"Live web search results for the user's request — use them to inform your answer, and mention where information came from where it's helpful, but don't fabricate beyond what's here:\n{search_context}"})
        except (requests.RequestException, ValueError, KeyError) as error:
            return jsonify(error=f"Web search failed: {error}"), 502
    messages.append({"role": "user", "content": prompt})

    model_id = data.get("model") or DEFAULT_MODEL
    # Ollama's native /api/chat shape differs from OpenAI-style APIs: no
    # response_format, generation options nest under "options", and a
    # non-streaming call needs "stream": false or it returns line-delimited
    # JSON chunks instead of one object.
    payload = {"model": model_id, "messages": messages, "stream": False, "options": {"temperature": 0.35, "num_predict": 4096}}
    try:
        # The person explicitly wants better, more detailed builds over speed,
        # so this timeout runs long — kept a little under gunicorn's own
        # worker timeout so Forge's own JSON error wins the race if it fires.
        response = requests.post(OLLAMA_CHAT_URL, headers={"Authorization": f"Bearer {OLLAMA_API_KEY}"}, json=payload, timeout=260)
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
        warnings = []
        for item in files: warnings += write_artifact(root, item)
        manifest = sorted(
            [{"path": str(p.relative_to(root)).replace("\\", "/"), "bytes": p.stat().st_size,
              "isImage": p.suffix.lower() in IMAGE_EXTENSIONS}
             for p in root.rglob("*") if p.is_file()],
            key=lambda f: f["path"],
        )
        reply = result.get("reply", "Done.")
        if warnings:
            reply += "\n\n" + "\n".join(f"⚠️ {w}" for w in warnings)
        return jsonify(reply=reply, workspace=workspace_id, files=manifest)
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
