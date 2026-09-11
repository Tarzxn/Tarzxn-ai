# Forge (Gen 2)

Forge is a Flask-based AI build workspace: describe what to create, then download the generated workspace as a ZIP. It presents a ChatGPT-style chat surface with a Codex-like model picker. Gen 2 runs on **Ollama Cloud** (https://ollama.com) for text/code, plus **Pollinations.ai** for images — both server-authenticated, so nothing is entered in the browser.

## What it can build

- Code and text files, `.docx`, `.xlsx`, `.pptx`, `.pdf`.
- **Images** — raster PNG/JPG via a free, keyless call to Pollinations.ai (the model writes an image prompt, Forge renders and saves it), or vector `.svg` written directly as text. Generated images and SVGs get an inline thumbnail gallery in chat, not just a download link.
- **3D models (.stl)** — built by Forge itself from a small, safe set of parametric primitives (cube, pyramid, sphere, cylinder, cone) with position offsets, so the model can compose several primitives into one compound shape (e.g. a cylinder + cone + sphere for a rocket) instead of hand-rolling raw, error-prone vertex lists.
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
3. Forge ships with an Ollama Cloud API key already set as the default in `app.py`, so it works immediately. To use a different key, set it in the environment instead — it overrides the built-in default:
   ```
   export OLLAMA_API_KEY="<your-ollama-cloud-api-key>"
   ```
4. Run `flask --app app run` and open the displayed address. No key needs to be entered in the browser — Forge authenticates every visitor with the server-side key. Image generation needs no key at all.

## Deploy to Render

Push this directory to a Git repository and create a Render Blueprint from it (or a Python Web Service with the commands in `render.yaml`). Set `OLLAMA_API_KEY` in the Render dashboard if you want to use a key other than the one built into `app.py` — `render.yaml` declares it as a non-synced secret so you're prompted to supply it. `render.yaml` and `gunicorn.conf.py` both set a longer worker timeout (180s) and 2 workers, since Ollama Cloud generations — especially file-heavy ones — routinely exceed gunicorn's 30s default; without this, a slow generation gets killed mid-request and the platform's proxy serves an HTML error page instead of Forge's own JSON error. Render's local disk is ephemeral, which is fine here since ZIP workspaces are meant for immediate download; Forge also prunes workspace folders older than 2 hours on every request so disk usage never creeps up on a long-running process.

## Security and privacy

**The Ollama Cloud API key is currently hardcoded as the default value in `app.py`, at the person's explicit request, so the app runs without extra setup.** This is fine for personal/local use but means anyone with access to this source (e.g. if pushed to a public repo) can read and use the key. Before deploying anywhere shared or public, either remove the hardcoded default and require `OLLAMA_API_KEY` to be set, or rotate the key at https://ollama.com/settings/keys if it's ever exposed. The key is never sent to or stored in the browser — only the server holds it, and every visitor shares its quota. Prompts and generated content are sent to Ollama Cloud (text) and Pollinations.ai (image prompts), so do not enter credentials, private keys, regulated data, or sensitive files unless you've reviewed each provider's retention/privacy terms. Workspaces are stored only on the server's ephemeral local disk, pruned automatically after 2 hours, and reachable only through an unguessable 128-bit workspace ID — but they are not encrypted at rest.
