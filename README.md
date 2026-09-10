# Forge

Forge is a Flask-based AI build workspace: choose an OpenRouter model, describe what to create, then download the generated workspace as a ZIP. It presents a familiar ChatGPT-style chat surface with a Codex-like model picker.

## Artifact support

The server natively produces Python/code and text files, `.docx`, `.xlsx`, `.pptx`, `.pdf`, and ASCII `.stl` (cube and pyramid starter geometry). The model may also return safe base64 payloads for any other file extension. All generated paths are restricted to a per-request workspace and delivered in a ZIP.

## Run locally

1. Install Python 3.12+.
2. Create a virtual environment and install dependencies: `pip install -r requirements.txt`.
3. Run `flask --app app run` and open the displayed address.
4. Select **API key** in the interface and paste your OpenRouter key for the current browser session.

## Deploy to Render

Push this directory to a Git repository and create a Render Blueprint from it (or create a Python Web Service with the commands in `render.yaml`). No API-key environment variable is required: each user adds their own key through the interface for their current browser session. Render’s local disk is ephemeral, which is appropriate here because ZIP workspaces are intended for immediate download.

The model picker loads OpenRouter's live catalogue with free models first. The default **Free Models Router** (`openrouter/free`) automatically chooses a currently available free model. Paid models are omitted except for Poolside, which is included at the user's request and labelled accurately. The live list also includes Poolside's free variants when available.

## Ollama

Add an Ollama server URL in **Connections** to list and run its installed models, which are labelled `Local · free`. The URL must be reachable from the Flask server—not merely from your browser. `http://localhost:11434` works when Forge and Ollama run on the same computer. A Render deployment cannot reach Ollama running only on your home computer; use a securely hosted Ollama server reachable by the Render service instead.

## Security and privacy

The API key is held in the browser's session storage and sent over HTTPS only when a generation request is made. It is not placed in source control, the deployment ZIP, the server's environment, or the generated workspaces. Prompts and generated content are nevertheless sent to OpenRouter and the provider selected for the request, so do not enter credentials, private keys, regulated data, or sensitive files unless you have reviewed OpenRouter and the chosen provider's retention/privacy terms. Workspaces are stored only on Render's ephemeral local disk and are reachable through an unguessable 128-bit workspace ID, but they are not encrypted at rest or automatically deleted during the process lifetime.
