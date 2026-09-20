# Gunicorn auto-loads this file from the working directory, so these settings
# apply even on hosts that just run `gunicorn app:app` without extra flags.
#
# timeout: Ollama Cloud generations (especially file-heavy STL/document
# requests, a web-search pass first, or High/Max reasoning effort on the
# largest model) can take well over gunicorn's 30s default, which otherwise
# kills the worker mid-request and makes the platform's proxy return an HTML
# error page instead of Forge's own JSON error response. Kept above app.py's
# own request timeout (up to 280s at High/Max power) so gunicorn never wins
# that race.
#
# worker_class/threads: /api/chat now streams its response instead of
# buffering it. A plain sync worker still blocks for the whole duration of
# one request either way, so streaming alone doesn't need this — but gthread
# workers let each worker interleave several concurrent streaming requests
# (mostly idle, waiting on network I/O) instead of one request fully
# monopolizing a worker, which meaningfully raises how many people can use
# Forge at once on the same process count.
timeout = 340
workers = 2
worker_class = "gthread"
threads = 4
