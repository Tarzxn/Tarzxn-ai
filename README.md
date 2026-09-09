# Forge

Forge is a Flask-based AI build workspace: choose an OpenRouter model, describe what to create, then download the generated workspace as a ZIP. It presents a familiar ChatGPT-style chat surface with a Codex-like model picker.

## Artifact support

The server natively produces Python/code and text files, `.docx`, `.xlsx`, `.pptx`, `.pdf`, and ASCII `.stl` (cube and pyramid starter geometry). The model may also return safe base64 payloads for any other file extension. All generated paths are restricted to a per-request workspace and delivered in a ZIP.

## Run locally

1. Install Python 3.12+.
2. Create a virtual environment and install dependencies: `pip install -r requirements.txt`.
3. Copy `.env.example` to `.env` and set `OPENROUTER_API_KEY`.
4. Run `flask --app app run` and open the displayed address.

## Deploy to Render

Push this directory to a Git repository and create a Render Blueprint from it (or create a Python Web Service with the commands in `render.yaml`). Set `OPENROUTER_API_KEY` as a secret environment variable. Render’s local disk is ephemeral, which is appropriate here because ZIP workspaces are intended for immediate download.

The model picker loads OpenRouter's live model catalogue and provides a curated fallback if it is unavailable.
