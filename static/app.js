const state = {
  model: 'openrouter/free',
  history: [],
  apiKey: sessionStorage.getItem('forgeOpenRouterKey') || ''
};
const $ = s => document.querySelector(s);
const convo = $('#conversation');
const promptEl = $('#prompt');
const sendBtn = $('#send');

function escapeHtml(t) {
  return String(t).replace(/[&<>'"]/g, c => ({
    '&': '&amp;', '<': '&lt;', '>': '&gt;', "'": '&#39;', '"': '&quot;'
  }[c]));
}

// Minimal, safe markdown-ish renderer: fenced code blocks, inline code, line breaks.
// Everything is escaped first, so this never introduces raw HTML from model output.
function renderReply(text) {
  const escaped = escapeHtml(text || '');
  const withBlocks = escaped.replace(/```(\w*)\n?([\s\S]*?)```/g, (m, lang, code) => {
    return `<pre><code>${code.replace(/\n$/, '')}</code></pre>`;
  });
  const withInline = withBlocks.replace(/`([^`\n]+)`/g, '<code class="inline">$1</code>');
  return `<div class="reply-text">${withInline.replace(/\n/g, '<br>')}</div>`;
}

function formatBytes(n) {
  if (n < 1024) return `${n} B`;
  if (n < 1024 * 1024) return `${(n / 1024).toFixed(1)} KB`;
  return `${(n / (1024 * 1024)).toFixed(1)} MB`;
}

function add(role, html) {
  const el = document.createElement('div');
  el.className = `message ${role}`;
  el.innerHTML = html;
  convo.append(el);
  el.scrollIntoView({ behavior: 'smooth', block: 'end' });
  return el;
}

function buildArtifactHtml(data) {
  if (!data.files || data.files.length === 0) return '';
  const rows = data.files.map(f => `
    <li>
      <span class="fname">${escapeHtml(f.path)}</span>
      <span class="fbytes">${formatBytes(f.bytes)}</span>
      <a href="/api/download/${data.workspace}/${encodeURIComponent(f.path)}" download>Download</a>
    </li>`).join('');
  return `
    <div class="artifact">
      <div class="artifact-head"><strong>${data.files.length} file${data.files.length === 1 ? '' : 's'} created</strong></div>
      <ul>${rows}</ul>
      <a class="download-all" href="/api/download/${data.workspace}">Download all (.zip) ↓</a>
    </div>`;
}

async function loadModels() {
  try {
    const res = await fetch('/api/models');
    const models = await res.json();
    $('#models').innerHTML = models.map(m => `
      <button class="model-choice" data-id="${escapeHtml(m.id)}" data-name="${escapeHtml(m.name)}">
        <strong>${escapeHtml(m.name)}</strong>
        <small>${escapeHtml(m.family)} · ${escapeHtml(m.tag)}</small>
      </button>`).join('');
    document.querySelectorAll('.model-choice').forEach(b => b.onclick = () => {
      state.model = b.dataset.id;
      $('#modelName').textContent = b.dataset.name;
      $('#modelMenu').classList.remove('open');
    });
  } catch (e) {
    $('#models').textContent = 'Model list unavailable.';
  }
}

function closeMenus() {
  $('#modelMenu').classList.remove('open');
  $('#keyModal').classList.remove('open');
}

$('#modelButton').onclick = (e) => { e.stopPropagation(); $('#modelMenu').classList.toggle('open'); };
$('#modelMenu').onclick = (e) => e.stopPropagation();

$('#newChat').onclick = () => {
  state.history = [];
  convo.innerHTML = $('#hero') ? $('#hero').outerHTML : '';
  if (!$('#hero')) location.reload();
  closeMenus();
  promptEl.value = '';
  autoResize();
  promptEl.focus();
  rebindSuggestions();
};

function rebindSuggestions() {
  document.querySelectorAll('.suggestions button').forEach(b => b.onclick = () => {
    promptEl.value = b.textContent;
    autoResize();
    promptEl.focus();
  });
}
rebindSuggestions();

$('#keyButton').onclick = (e) => {
  e.stopPropagation();
  $('#apiKey').value = state.apiKey;
  $('#keyModal').classList.add('open');
  $('#apiKey').focus();
};
$('#closeKey').onclick = () => $('#keyModal').classList.remove('open');
$('#keyModal').onclick = (e) => { if (e.target.id === 'keyModal') $('#keyModal').classList.remove('open'); };
$('#saveKey').onclick = () => {
  state.apiKey = $('#apiKey').value.trim();
  if (state.apiKey) sessionStorage.setItem('forgeOpenRouterKey', state.apiKey);
  else sessionStorage.removeItem('forgeOpenRouterKey');
  $('#keyModal').classList.remove('open');
  promptEl.focus();
};

document.addEventListener('click', closeMenus);
document.addEventListener('keydown', (e) => {
  if (e.key === 'Escape') closeMenus();
});

function autoResize() {
  promptEl.style.height = 'auto';
  promptEl.style.height = Math.min(promptEl.scrollHeight, 220) + 'px';
}
promptEl.addEventListener('input', autoResize);

// Enter sends the message; Shift+Enter inserts a newline.
promptEl.addEventListener('keydown', (e) => {
  if (e.key === 'Enter' && !e.shiftKey) {
    e.preventDefault();
    $('#composer').requestSubmit();
  }
});

async function submitPrompt() {
  const prompt = promptEl.value.trim();
  if (!prompt) return;
  if (!state.apiKey) { $('#keyButton').click(); return; }

  $('#hero')?.remove();
  add('user', escapeHtml(prompt));
  promptEl.value = '';
  autoResize();

  const pending = add('assistant typing', '<div class="reply-text">Forge is working…</div>');
  sendBtn.disabled = true;

  try {
    const r = await fetch('/api/chat', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ prompt, model: state.model, history: state.history, apiKey: state.apiKey })
    });
    const data = await r.json();
    if (!r.ok) throw new Error(data.error || 'Request failed');

    pending.classList.remove('typing');
    // Text-only replies (no files) render as plain chat text; only show the
    // file panel when the model actually produced files.
    pending.innerHTML = renderReply(data.reply) + buildArtifactHtml(data);
    pending.scrollIntoView({ behavior: 'smooth', block: 'end' });

    state.history.push({ role: 'user', content: prompt }, { role: 'assistant', content: data.reply });
  } catch (err) {
    pending.classList.remove('typing');
    pending.innerHTML = `<span class="error-text">${escapeHtml(err.message)}</span>`;
  } finally {
    sendBtn.disabled = false;
    promptEl.focus();
  }
}

$('#composer').addEventListener('submit', (e) => {
  e.preventDefault();
  submitPrompt();
});

loadModels();
autoResize();
