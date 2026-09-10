# Forge (Gen 2)

Forge is a Flask-based AI build workspace: describe what to create, then download the generated workspace as a ZIP. It presents a ChatGPT-style chat surface with a Codex-like model picker. Gen 2 runs on **Ollama Cloud** (https://ollama.com), using a single server-side API key instead of a per-user key.

## Artifact support

The server natively produces Python/code and text files, `.docx`, `.xlsx`, `.pptx`, `.pdf`, and ASCII `.stl` (cube and pyramid starter geometry). The model may also return safe base64 payloads for any other file extension. Plain-text answers with no files are shown as ordinary chat replies. All generated paths are restricted to a per-request workspace and delivered as a ZIP, with individual files also downloadable on their own.

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
4. Run `flask --app app run` and open the displayed address. No key needs to be entered in the browser — Forge authenticates every visitor with the server-side key.

## Deploy to Render

Push this directory to a Git repository and create a Render Blueprint from it (or a Python Web Service with the commands in `render.yaml`). Set `OLLAMA_API_KEY` in the Render dashboard if you want to use a key other than the one built into `app.py` — `render.yaml` declares it as a non-synced secret so you're prompted to supply it. Render's local disk is ephemeral, which is appropriate here because ZIP workspaces are intended for immediate download.

## Security and privacy

**The API key is currently hardcoded as the default value in `app.py`, at the person's explicit request, so the app runs without extra setup.** This is fine for personal/local use but means anyone with access to this source (e.g. if pushed to a public repo) can read and use the key. Before deploying anywhere shared or public, either remove the hardcoded default and require `OLLAMA_API_KEY` to be set, or rotate the key at https://ollama.com/settings/keys if it's ever exposed. The key is never sent to or stored in the browser — only the server holds it, and every visitor shares its quota. Prompts and generated content are sent to Ollama Cloud, so do not enter credentials, private keys, regulated data, or sensitive files unless you've reviewed Ollama's retention/privacy terms. Workspaces are stored only on the server's ephemeral local disk and are reachable through an unguessable 128-bit workspace ID, but they are not encrypted at rest or automatically deleted during the process lifetime.
