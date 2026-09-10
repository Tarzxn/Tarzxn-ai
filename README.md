# Forge (Gen 2)

Forge is a Flask-based AI build workspace: describe what to create, then download the generated workspace as a ZIP. It presents a ChatGPT-style chat surface with a Codex-like model picker. Gen 2 runs entirely on Hugging Face's free serverless Inference Providers, using a single server-side token instead of a per-user API key.

## Artifact support

The server natively produces Python/code and text files, `.docx`, `.xlsx`, `.pptx`, `.pdf`, and ASCII `.stl` (cube and pyramid starter geometry). The model may also return safe base64 payloads for any other file extension. Plain-text answers with no files are shown as ordinary chat replies. All generated paths are restricted to a per-request workspace and delivered as a ZIP, with individual files also downloadable on their own.

## Models

The picker offers a curated set of ungated models that work on Hugging Face's free tier:

- **Qwen2.5 7B Instruct** (default) — best all-round pick for this app: strong instruction-following and code generation, ungated, and reliably available on the free tier.
- Mistral 7B Instruct v0.3 — solid general-purpose alternative.
- Phi-3.5 Mini Instruct — smaller and faster, useful if Qwen is rate-limited.
- Zephyr 7B Beta — another chat-tuned fallback.

## Run locally

1. Install Python 3.12+.
2. Create a virtual environment and install dependencies: `pip install -r requirements.txt`.
3. Get a free access token at https://huggingface.co/settings/tokens (read access is enough) and set it: `export HF_TOKEN=hf_...`.
4. Run `flask --app app run` and open the displayed address. No key needs to be entered in the browser — Forge is pre-authenticated for every visitor using `HF_TOKEN`.

## Deploy to Render

Push this directory to a Git repository and create a Render Blueprint from it (or a Python Web Service with the commands in `render.yaml`). Set the `HF_TOKEN` environment variable in the Render dashboard — `render.yaml` declares it as a required, non-synced secret. Render's local disk is ephemeral, which is appropriate here because ZIP workspaces are intended for immediate download.

## Security and privacy

The Hugging Face token lives only in the server's environment (`HF_TOKEN`) — it is never sent to or stored in the browser, and every visitor shares the same server-side quota. Prompts and generated content are sent to Hugging Face and whichever inference provider serves the selected model, so do not enter credentials, private keys, regulated data, or sensitive files unless you've reviewed Hugging Face's and the relevant provider's retention/privacy terms. Workspaces are stored only on the server's ephemeral local disk and are reachable through an unguessable 128-bit workspace ID, but they are not encrypted at rest or automatically deleted during the process lifetime. Because this is a shared free tier, expect occasional 429 rate limits under heavy use — switch models or retry shortly if that happens.
