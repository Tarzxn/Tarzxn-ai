# Forge (Gen 2)

Forge is a Flask-based AI build workspace with a ChatGPT-style chat interface: sign in, describe what to create, watch the reply stream in live, then download the generated files as a ZIP. It runs on **Ollama Cloud** (https://ollama.com) for text/code, **Pollinations.ai** for images, and optionally **Tavily** for live web search — all server-authenticated, so nothing but your own Forge login is entered in the browser.

## Look and feel

The UI uses a "liquid glass" style: translucent, blurred panels (sidebar, header, composer, message bubbles, code blocks, file cards, the login card) floating over an animated, colorful blurred backdrop, each with a soft specular highlight along its top edge.

## Chat architecture

- **Real token streaming.** Replies appear live, word by word, as the model generates them — Forge incrementally decodes the `"reply"` field straight out of the model's still-in-progress JSON output (a small custom streaming-JSON-string parser), so you see text immediately instead of waiting for the whole generation (which, for a file-heavy request, can take a while). Verified correct against Python's own JSON decoder across chunk boundaries that split mid-escape-sequence, mid-unicode-escape, and mid-key.
- **Stop generating** mid-stream — the send button becomes a Stop button; any text that had already streamed in is kept rather than thrown away.
- **Regenerate** any assistant reply (not just the latest) — drops it and everything after it, then re-asks the same prompt.
- **Copy buttons** on individual code blocks and on whole replies.
- **Rich Markdown rendering** — headings, bulleted/numbered lists, blockquotes, links, bold/italic, and syntax-styled code blocks.
- **Automatic retry** for transient upstream failures (a 502/503/504 or dropped connection from Ollama Cloud gets one quick retry before surfacing an error), and a single bad file in a response never discards the rest of a good reply — each file, and each step of a 3D build program, is attempted independently.

## Authentication — no cookies, nothing remembered after a restart

Forge sits behind a login screen. This is deliberately **not** cookie-based:

- The server never sets a cookie. On login it hands back an opaque bearer token, which the browser holds in `sessionStorage` (not `localStorage`) and sends explicitly (`Authorization: Bearer ...`) on every request. `sessionStorage` is wiped the instant the tab or browser closes — there is no mechanism for a returning visitor to be silently auto-logged-in.
- Tokens live only in an in-memory dict on the server, so **restarting the server invalidates every session** — nothing about who was logged in, or what they said, survives a shutdown.
- Conversation history lives in the same `sessionStorage`, for the same reason: it lets you switch between a few chats in one active tab session, but nothing is retained once that session ends. (Generated *files* are separate — they live on the server's ephemeral workspace disk and are pruned after 2 hours regardless of the browser session.)
- Set `FORGE_PASSWORD` (required) and optionally `FORGE_USERNAME` (defaults to `admin`) to enable login; without `FORGE_PASSWORD` set, every request is rejected with a clear message telling you to configure it — there is no default password.

## What it can build

- Code and text files, `.docx`, `.xlsx`, `.pptx`, `.pdf` — each with real formatting: `.docx`/`.pdf` understand a Markdown-lite subset (`#`/`##`/`###` headings, `- ` bullets, `**bold**`, and `| a | b |` tables with a `|---|---|` separator row render as real formatted tables — grid tables with a bold header row in Word, aligned columns in PDF), `.xlsx` gets a bold auto-width header row with the top row frozen, and `.pptx` splits multi-line bodies into proper bullet points.
- **Images** — raster PNG/JPG via a free, keyless call to Pollinations.ai, with optional aspect-ratio control (`square`/`portrait`/`landscape`), or vector `.svg` written directly as text. Generated images/SVGs get an inline thumbnail gallery in chat.
- **Data charts** — real bar/line/pie/scatter charts rendered from actual numbers via matplotlib (not an AI-generated approximation of a chart). Distinct from "image": use this when the user wants their data plotted accurately.
- **3D models (.stl)** — a real parametric CAD-lite engine. The model writes a design "plan" plus an ordered build program ("ops") using primitives — box, sphere, cylinder, cone, torus, a genuinely hollow **tube** (pipe/ring/washer), a **capsule** (pill/rounded-rod shape with true hemispherical caps), a **wedge** (ramp/roof), and pyramid — each with position/rotation/scale. `repeat` creates radial or linear patterns (gear teeth, table legs, fence posts, fins, stair treads); `mirror` reflects a part across an axis-aligned plane for symmetric designs (wings, hull halves, paired brackets) without describing both sides by hand. Segment counts auto-scale with part size for smooth curves. Every op runs independently — a malformed step is skipped with a warning instead of failing the whole model. Every shape ships tested watertight (manifold — every edge shared by exactly two triangles) with outward-facing normals, and the reply includes the model's overall size (bounding box) and triangle count.
- **Live web research** — an optional "🔎 Web search" toggle in the composer runs the request through Tavily first and feeds the results to the model as context. Off by default; only enabled if `TAVILY_API_KEY` is set.
- Arbitrary base64 binary payloads for anything else.
- Plain-text answers with no files render as ordinary chat replies.

All generated paths are restricted to a per-request workspace and delivered as a ZIP, with individual files also downloadable (or, for images/SVGs, previewable) on their own.

## Models

The picker offers Ollama Cloud's hosted catalogue:

- **GPT-OSS 20B** (default) — fast, capable, and the best balance of speed vs. quality for interactive use.
- GPT-OSS 120B — larger, slower, stronger reasoning.
- Qwen3 32B — strong general-purpose and code alternative.
- DeepSeek V3.1 671B — frontier-scale, noticeably slower; use for the hardest requests.

## Run locally

1. Install Python 3.12+.
2. Create a virtual environment and install dependencies: `pip install -r requirements.txt`.
3. Set a login password (required — Forge refuses all requests without it):
   ```
   export FORGE_PASSWORD="choose-a-real-password"
   export FORGE_USERNAME="admin"   # optional, this is the default
   ```
4. Forge ships with an Ollama Cloud API key already set as the default in `app.py`, so text/code/file generation works immediately once you've logged in. To use a different key, set it in the environment instead — it overrides the built-in default:
   ```
   export OLLAMA_API_KEY="<your-ollama-cloud-api-key>"
   ```
   Optionally enable live web search too — no default is baked in for this one, since none was provided:
   ```
   export TAVILY_API_KEY="<your-tavily-api-key>"
   ```
5. Run `flask --app app run` and open the displayed address. Log in with the username/password you set above. Image generation needs no key at all.

## Deploy to Render

Push this directory to a Git repository and create a Render Blueprint from it (or a Python Web Service with the commands in `render.yaml`). Set `FORGE_PASSWORD` (and optionally `FORGE_USERNAME`) in the Render dashboard — both are declared as non-synced secrets in `render.yaml`, so you'll be prompted to supply them; login is refused until `FORGE_PASSWORD` is set. Set `OLLAMA_API_KEY` there too if you want to use a key other than the one built into `app.py`. `render.yaml` and `gunicorn.conf.py` both set a longer worker timeout (300s) since Ollama Cloud generations — especially file-heavy ones — routinely exceed gunicorn's 30s default, and use threaded (`gthread`) workers so one process can serve several concurrent streaming chats instead of a single request occupying a whole worker. Render's local disk is ephemeral, which is fine here since ZIP workspaces are meant for immediate download; Forge also prunes workspace folders older than 2 hours on every request so disk usage never creeps up on a long-running process.

## Security and privacy

**The Ollama Cloud API key is currently hardcoded as the default value in `app.py`, at the person's explicit request, so the app runs without extra setup.** This is fine for personal/local use but means anyone with access to this source (e.g. if pushed to a public repo) can read and use the key. Before deploying anywhere shared or public, either remove the hardcoded default and require `OLLAMA_API_KEY` to be set, or rotate the key at https://ollama.com/settings/keys if it's ever exposed. The key is never sent to or stored in the browser — only the server holds it. Login credentials (`FORGE_USERNAME`/`FORGE_PASSWORD`) gate every API route; there is no hardcoded default password. Prompts and generated content are sent to Ollama Cloud (text), Pollinations.ai (image prompts), and — only when you enable the toggle — Tavily (your search query), so do not enter credentials, private keys, regulated data, or sensitive files unless you've reviewed each provider's retention/privacy terms. Workspaces are stored only on the server's ephemeral local disk, pruned automatically after 2 hours, and reachable only by someone holding both a valid login session and the unguessable 128-bit workspace ID — but they are not encrypted at rest.
