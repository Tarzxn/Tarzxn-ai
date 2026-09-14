# Forge (Gen 2)

Forge is a Flask-based AI build workspace: describe what to create, then download the generated workspace as a ZIP. It presents a ChatGPT-style chat surface with a Codex-like model picker. Gen 2 runs on **Ollama Cloud** (https://ollama.com) for text/code, plus **Pollinations.ai** for images — both server-authenticated, so nothing is entered in the browser.

## Look and feel

The UI uses a "liquid glass" style: translucent, blurred panels (sidebar, header, composer, message bubbles, code blocks, file cards) floating over an animated, colorful blurred backdrop, each with a soft specular highlight along its top edge.

## What it can build

- Code and text files, `.docx`, `.xlsx`, `.pptx`, `.pdf` — each with real formatting, not just plain text dumped in: `.docx`/`.pdf` understand a Markdown-lite subset (`#`/`##`/`###` headings, `- ` bullets, `**bold**`), `.xlsx` gets a bold auto-width header row with the top row frozen, and `.pptx` splits multi-line bodies into proper bullet points.
- **Images** — raster PNG/JPG via a free, keyless call to Pollinations.ai, with optional aspect-ratio control (`square`/`portrait`/`landscape`), or vector `.svg` written directly as text. Generated images/SVGs get an inline thumbnail gallery in chat.
- **Data charts** — real bar/line/pie/scatter charts rendered from actual numbers via matplotlib (not an AI-generated approximation of a chart). Distinct from "image": use this when the user wants their data plotted accurately.
- **3D models (.stl)** — a real parametric CAD-lite engine. The model writes a design "plan" plus an ordered build program ("ops") using primitives — box, sphere, cylinder, cone, torus, a genuinely hollow **tube** (pipe/ring/washer), a **capsule** (pill/rounded-rod shape with true hemispherical caps), a **wedge** (ramp/roof), and pyramid — each with position/rotation/scale, plus a `repeat` step for radial or linear patterns (gear teeth, table legs, fence posts, fins, stair treads). Segment counts auto-scale with part size for smooth curves unless the model deliberately wants a low-poly look. Every op runs independently — a malformed step is skipped with a warning instead of failing the whole model. Every shape ships tested watertight (manifold — every edge shared by exactly two triangles) with outward-facing normals.
- **Live web research** — an optional "🔎 Web search" toggle in the composer runs the request through Tavily first and feeds the results to the model as context, so answers about current events/prices/recent releases/etc. are grounded in real, fresh sources instead of the model's training data. Off by default; only appears active if `TAVILY_API_KEY` is set on the server.
- Arbitrary base64 binary payloads for anything else.
- Plain-text answers with no files render as ordinary chat replies.

All generated paths are restricted to a per-request workspace and delivered as a ZIP, with individual files also downloadable (or, for images/SVGs, previewable) on their own.

## Chatbot features

- **Multiple persistent conversations.** Every chat is saved to the browser's `localStorage` (not the server) — the sidebar's "Recent" list lets you switch between past conversations or delete one, and your most recent conversation is restored automatically on reload. Note: this only persists the conversation transcript, not generated files — those still live on the server's ephemeral workspace disk and get pruned after 2 hours, so very old conversations' download links may stop working.
- **Rich replies.** Assistant messages render real Markdown — headings, bulleted/numbered lists, blockquotes, links, bold/italic, and syntax-styled code blocks — instead of one flat wall of text.
- **Stop generating.** The send button turns into a Stop button mid-request; clicking it cancels the in-flight request immediately.
- **Regenerate.** Hover any assistant reply (not just the latest) for a "↻ Regenerate" button that drops it and everything after it, then re-asks the same prompt.
- **Copy buttons** on both individual code blocks and whole assistant replies.

## Models

The picker offers Ollama Cloud's hosted catalogue:

- **GPT-OSS 20B** (default) — fast, capable, and the best balance of speed vs. quality for interactive use.
- GPT-OSS 120B — larger, slower, stronger reasoning.
- Qwen3 32B — strong general-purpose and code alternative.
- DeepSeek V3.1 671B — frontier-scale, noticeably slower; use for the hardest requests.

## Run locally

1. Install Python 3.12+.
2. Create a virtual environment and install dependencies: `pip install -r requirements.txt`.
3. Forge ships with an Ollama Cloud API key already set as the default in `app.py`, so it works immediately. To use a different key, set it in the environment instead — it overrides the built-in default:
   ```
   export OLLAMA_API_KEY="<your-ollama-cloud-api-key>"
   ```
   Optionally enable live web search too — no default is baked in for this one, since none was provided:
   ```
   export TAVILY_API_KEY="<your-tavily-api-key>"
   ```
4. Run `flask --app app run` and open the displayed address. No key needs to be entered in the browser — Forge authenticates every visitor with the server-side key. Image generation needs no key at all.

## Deploy to Render

Push this directory to a Git repository and create a Render Blueprint from it (or a Python Web Service with the commands in `render.yaml`). Set `OLLAMA_API_KEY` in the Render dashboard if you want to use a key other than the one built into `app.py` — `render.yaml` declares it as a non-synced secret so you're prompted to supply it. `render.yaml` and `gunicorn.conf.py` both set a longer worker timeout (180s) and 2 workers, since Ollama Cloud generations — especially file-heavy ones — routinely exceed gunicorn's 30s default; without this, a slow generation gets killed mid-request and the platform's proxy serves an HTML error page instead of Forge's own JSON error. Render's local disk is ephemeral, which is fine here since ZIP workspaces are meant for immediate download; Forge also prunes workspace folders older than 2 hours on every request so disk usage never creeps up on a long-running process.

## Security and privacy

**The Ollama Cloud API key is currently hardcoded as the default value in `app.py`, at the person's explicit request, so the app runs without extra setup.** This is fine for personal/local use but means anyone with access to this source (e.g. if pushed to a public repo) can read and use the key. Before deploying anywhere shared or public, either remove the hardcoded default and require `OLLAMA_API_KEY` to be set, or rotate the key at https://ollama.com/settings/keys if it's ever exposed. The key is never sent to or stored in the browser — only the server holds it, and every visitor shares its quota. Prompts and generated content are sent to Ollama Cloud (text) and Pollinations.ai (image prompts), so do not enter credentials, private keys, regulated data, or sensitive files unless you've reviewed each provider's retention/privacy terms. Workspaces are stored only on the server's ephemeral local disk, pruned automatically after 2 hours, and reachable only through an unguessable 128-bit workspace ID — but they are not encrypted at rest.
