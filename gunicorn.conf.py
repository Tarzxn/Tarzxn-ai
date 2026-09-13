# Gunicorn auto-loads this file from the working directory, so the longer
# timeout applies even on hosts that just run `gunicorn app:app` without extra
# flags. Ollama Cloud generations (especially file-heavy STL/document
# requests, or ones with a web-search pass first) can take well over
# gunicorn's 30s default, which otherwise kills the worker mid-request and
# makes the platform's proxy return an HTML error page instead of Forge's
# own JSON error response. Kept above app.py's own 260s request timeout so
# gunicorn never wins that race.
timeout = 300
workers = 2
