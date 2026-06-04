#!/usr/bin/env python3
"""
VERIFICADOR DE CITAÇÕES — Interface Web (v8.6)
Compatível com Android (Termux), Windows, Linux, Mac.

Instalação rápida (Termux no Android):
  pkg update && pkg install python
  pip install flask anthropic pymupdf
  termux-setup-storage          ← dá acesso ao armazenamento externo (pen drive)

Uso:
  python app_web.py
  Abrir no Chrome: http://localhost:5000
"""

import json
import os
import queue
import tempfile
import threading
import uuid
from pathlib import Path

from flask import Flask, Response, jsonify, request

app = Flask(__name__)

# Estado dos jobs em andamento
_JOBS: dict = {}   # job_id → {log_q, stop, resultado, done, diss_tmp}


# ════════════════════════════════════════════════════════════════════════════
# PÁGINA PRINCIPAL
# ════════════════════════════════════════════════════════════════════════════

_HTML = """<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1,maximum-scale=1">
<title>Verificador de Citações</title>
<style>
*{box-sizing:border-box;-webkit-tap-highlight-color:transparent}
body{font-family:Segoe UI,Arial,sans-serif;margin:0;padding:0;background:#f0f2f5;color:#222}
header{background:#1a2744;color:#fff;padding:1em 1.2em;position:sticky;top:0;z-index:99}
header h1{margin:0;font-size:1.15em;font-weight:700}
header p{margin:.2em 0 0;font-size:.8em;color:#aac;opacity:.9}
main{padding:1em;max-width:720px;margin:0 auto}
.card{background:#fff;border-radius:10px;padding:1.2em;margin-bottom:1em;box-shadow:0 1px 4px rgba(0,0,0,.1)}
label{display:block;font-weight:600;font-size:.9em;margin-bottom:.3em;color:#333}
input[type=text],input[type=password]{width:100%;padding:.75em 1em;border:1.5px solid #ccd;border-radius:8px;font-size:.95em;background:#f9f9ff;outline:none}
input[type=text]:focus,input[type=password]:focus{border-color:#1a2744}
.file-btn{display:flex;align-items:center;gap:.6em;padding:.75em 1em;border:2px dashed #aac;border-radius:8px;cursor:pointer;background:#f9f9ff;font-size:.9em;color:#555;transition:border-color .2s}
.file-btn:active{border-color:#1a2744}
.file-btn input{display:none}
.file-name{font-size:.8em;color:#1a2744;margin-top:.3em;word-break:break-all}
.check-row{display:flex;align-items:center;gap:.6em;margin-top:.5em}
.check-row input{width:1.2em;height:1.2em}
.btn{display:block;width:100%;padding:.85em;border:none;border-radius:8px;font-size:1em;font-weight:700;cursor:pointer;letter-spacing:.02em;transition:opacity .15s}
.btn:active{opacity:.75}
.btn-green{background:#27ae60;color:#fff;margin-top:.5em}
.btn-red{background:#e74c3c;color:#fff}
.btn-blue{background:#2980b9;color:#fff}
.btn-sm{font-size:.85em;padding:.5em .9em;width:auto;display:inline-block;border-radius:6px}
.storage-list{display:none;margin-top:.5em;border:1px solid #dde;border-radius:8px;max-height:180px;overflow-y:auto}
.storage-item{padding:.6em 1em;border-bottom:1px solid #eee;font-size:.85em;cursor:pointer;color:#1a2744;word-break:break-all}
.storage-item:active{background:#eef}
.storage-item:last-child{border-bottom:none}
#log-box{background:#111;color:#7fff7f;font-family:monospace;font-size:.78em;padding:.8em;border-radius:8px;height:280px;overflow-y:auto;white-space:pre-wrap;word-break:break-word;line-height:1.4}
.status-bar{display:flex;align-items:center;gap:.5em;margin-bottom:.6em;font-size:.88em;font-weight:600}
.dot{width:10px;height:10px;border-radius:50%;background:#888;flex-shrink:0}
.dot.running{background:#f39c12;animation:pulse 1s infinite}
.dot.done{background:#27ae60}
.dot.error{background:#e74c3c}
@keyframes pulse{0%,100%{opacity:1}50%{opacity:.3}}
.result-btns{display:flex;gap:.6em;flex-wrap:wrap;margin-bottom:.8em}
.hidden{display:none}
.row{display:flex;gap:.6em;align-items:center;margin-top:.4em}
.row input{flex:1}
.toggle-eye{background:none;border:1.5px solid #ccd;border-radius:8px;padding:.6em .8em;cursor:pointer;font-size:.9em}
</style>
</head>
<body>

<header>
  <h1>📚 Verificador de Citações Acadêmicas</h1>
  <p>Lê a dissertação · Cruza referências · Verifica com IA</p>
</header>

<main>

<!-- ── FORMULÁRIO ── -->
<div class="card" id="sec-form">
  <label>Dissertação (.docx)</label>
  <label class="file-btn" id="btn-docx">
    <span>📄 Selecionar arquivo…</span>
    <input type="file" id="inp-docx" accept=".docx,.doc">
  </label>
  <div class="file-name" id="nome-docx"></div>

  <label style="margin-top:1em">Pasta de referências (caminho local)</label>
  <input type="text" id="inp-refs"
         placeholder="Ex: /storage/XXXX-XXXX/fusions centers">
  <button class="btn btn-blue btn-sm" style="margin-top:.4em"
          onclick="listarStorage()">📁 Ver armazenamentos disponíveis</button>
  <div class="storage-list" id="storage-list"></div>

  <label style="margin-top:1em">Chave API Claude <small style="font-weight:400">(opcional — necessária para verificação semântica)</small></label>
  <div class="row">
    <input type="password" id="inp-api" placeholder="sk-ant-…">
    <button class="toggle-eye" onclick="toggleApi()">👁</button>
  </div>

  <div class="check-row" style="margin-top:.8em">
    <input type="checkbox" id="chk-sem" value="true">
    <label for="chk-sem" style="margin:0;font-weight:400">Apenas cruzamento (sem verificação IA — mais rápido)</label>
  </div>

  <button class="btn btn-green" onclick="iniciar()">▶ Iniciar análise</button>
  <p id="msg-erro" style="color:#e74c3c;font-size:.88em;margin:.4em 0 0"></p>
</div>

<!-- ── PROGRESSO ── -->
<div class="card hidden" id="sec-prog">
  <div class="status-bar">
    <div class="dot running" id="dot"></div>
    <span id="status-txt">Analisando…</span>
  </div>
  <div id="log-box"></div>
  <button class="btn btn-red" style="margin-top:.6em" onclick="parar()">⏹ Parar análise</button>
</div>

<!-- ── RESULTADO ── -->
<div class="card hidden" id="sec-result">
  <p style="font-weight:700;color:#27ae60;font-size:1.05em">✅ Análise concluída!</p>
  <div class="result-btns">
    <button class="btn btn-blue btn-sm" onclick="verRelatorio()">📊 Ver relatório</button>
    <button class="btn btn-sm" style="background:#8e44ad;color:#fff" onclick="novaAnalise()">🔄 Nova análise</button>
  </div>
</div>

</main>

<script>
let _jobId = null;
let _sse   = null;

// ── Selecionar arquivo ───────────────────────────────────────────────────
document.getElementById('inp-docx').addEventListener('change', function() {
  const n = this.files[0] ? this.files[0].name : '';
  document.getElementById('nome-docx').textContent = n ? '✓ ' + n : '';
});

// ── Mostrar/ocultar API key ──────────────────────────────────────────────
function toggleApi() {
  const i = document.getElementById('inp-api');
  i.type = i.type === 'password' ? 'text' : 'password';
}

// ── Listar armazenamentos ────────────────────────────────────────────────
function listarStorage() {
  fetch('/storage_list')
    .then(r => r.json())
    .then(paths => {
      const el = document.getElementById('storage-list');
      if (!paths.length) {
        el.innerHTML = '<div class="storage-item" style="color:#888">Nenhum armazenamento encontrado</div>';
      } else {
        el.innerHTML = paths.map(p =>
          `<div class="storage-item" onclick="selecionarPath('${p.replace(/'/g,"\\'")}')">📁 ${p}</div>`
        ).join('');
      }
      el.style.display = el.style.display === 'block' ? 'none' : 'block';
    })
    .catch(() => alert('Erro ao listar armazenamentos'));
}

function selecionarPath(p) {
  document.getElementById('inp-refs').value = p;
  document.getElementById('storage-list').style.display = 'none';
}

// ── Iniciar análise ──────────────────────────────────────────────────────
function iniciar() {
  const docxFile = document.getElementById('inp-docx').files[0];
  const refsDir  = document.getElementById('inp-refs').value.trim();
  const apiKey   = document.getElementById('inp-api').value.trim();
  const semV     = document.getElementById('chk-sem').checked;
  const erroEl   = document.getElementById('msg-erro');

  erroEl.textContent = '';
  if (!docxFile) { erroEl.textContent = 'Selecione o arquivo da dissertação.'; return; }
  if (!refsDir)  { erroEl.textContent = 'Informe o caminho da pasta de referências.'; return; }

  const fd = new FormData();
  fd.append('docx', docxFile);
  fd.append('refs_dir', refsDir);
  fd.append('api_key', apiKey);
  fd.append('sem_verificacao', semV ? 'true' : 'false');

  fetch('/iniciar', { method: 'POST', body: fd })
    .then(r => r.json())
    .then(data => {
      if (data.erro) { erroEl.textContent = data.erro; return; }
      _jobId = data.job_id;
      mostrar('sec-prog');
      document.getElementById('log-box').textContent = '';
      conectarSSE();
    })
    .catch(e => { erroEl.textContent = 'Erro ao iniciar: ' + e; });
}

// ── SSE — logs em tempo real ─────────────────────────────────────────────
function conectarSSE() {
  if (_sse) _sse.close();
  _sse = new EventSource('/logs/' + _jobId);
  const box = document.getElementById('log-box');
  _sse.onmessage = function(ev) {
    const msg = JSON.parse(ev.data);
    if (msg === '__PING__') return;
    if (msg === '__DONE__') {
      _sse.close();
      setStatus('done', 'Concluído');
      mostrar('sec-result');
      return;
    }
    if (msg === '__ERRO__') {
      _sse.close();
      setStatus('error', 'Erro — verifique o log');
      return;
    }
    box.textContent += msg + '\\n';
    box.scrollTop = box.scrollHeight;
  };
  _sse.onerror = function() {
    if (_sse.readyState === EventSource.CLOSED) return;
    setStatus('error', 'Conexão perdida');
  };
}

// ── Parar ────────────────────────────────────────────────────────────────
function parar() {
  if (_jobId) fetch('/parar/' + _jobId, { method: 'POST' });
  if (_sse)   _sse.close();
  setStatus('error', 'Interrompido pelo usuário');
}

// ── Ver relatório ────────────────────────────────────────────────────────
function verRelatorio() {
  if (_jobId) window.open('/resultado/' + _jobId, '_blank');
}

// ── Nova análise ─────────────────────────────────────────────────────────
function novaAnalise() {
  _jobId = null;
  document.getElementById('inp-docx').value = '';
  document.getElementById('nome-docx').textContent = '';
  document.getElementById('msg-erro').textContent = '';
  mostrar('sec-form');
}

// ── Helpers ──────────────────────────────────────────────────────────────
function mostrar(id) {
  ['sec-form','sec-prog','sec-result'].forEach(s => {
    document.getElementById(s).classList.toggle('hidden', s !== id);
  });
}

function setStatus(tipo, txt) {
  const dot = document.getElementById('dot');
  dot.className = 'dot ' + tipo;
  document.getElementById('status-txt').textContent = txt;
}
</script>
</body>
</html>"""


# ════════════════════════════════════════════════════════════════════════════
# ROTAS
# ════════════════════════════════════════════════════════════════════════════

@app.route("/")
def index():
    return _HTML


@app.route("/storage_list")
def storage_list():
    """Lista caminhos de armazenamento acessíveis (útil no Android/Termux)."""
    candidatos = [
        "/storage",
        "/sdcard",
        "/mnt/sdcard",
        str(Path.home()),
        str(Path.home() / "storage"),
        str(Path.home() / "storage" / "external-1"),
        str(Path.home() / "storage" / "shared"),
    ]
    resultado = []
    for base in candidatos:
        p = Path(base)
        if not p.exists():
            continue
        try:
            subs = sorted(
                str(s) for s in p.iterdir()
                if s.is_dir() and not s.name.startswith(".")
            )
            resultado.extend(subs[:30])
        except PermissionError:
            resultado.append(base)
    return jsonify(list(dict.fromkeys(resultado)))  # deduplica mantendo ordem


@app.route("/iniciar", methods=["POST"])
def iniciar():
    api_key  = request.form.get("api_key", "").strip()
    refs_dir = request.form.get("refs_dir", "").strip()
    sem_v    = request.form.get("sem_verificacao", "false") == "true"

    if "docx" not in request.files or not request.files["docx"].filename:
        return jsonify({"erro": "Selecione o arquivo da dissertação (.docx)."}), 400

    refs_path = Path(refs_dir) if refs_dir else None
    if not refs_path or not refs_path.is_dir():
        return jsonify({"erro": f"Pasta não encontrada: {refs_dir}"}), 400

    # Salva DOCX em arquivo temporário
    f = request.files["docx"]
    tmp = tempfile.NamedTemporaryFile(suffix=".docx", delete=False)
    f.save(tmp.name)
    tmp.close()
    diss_tmp = tmp.name

    job_id = uuid.uuid4().hex[:12]
    log_q  = queue.Queue()
    stop_e = threading.Event()

    _JOBS[job_id] = {
        "log_q":    log_q,
        "stop":     stop_e,
        "resultado": None,
        "done":     False,
        "diss_tmp": diss_tmp,
    }

    def _run():
        try:
            import analisar
            res = analisar.analisar(
                diss_path=diss_tmp,
                refs_dir=str(refs_path),
                api_key=api_key,
                log_fn=log_q.put,
                stop_fn=stop_e.is_set,
                sem_verificacao=sem_v,
            )
            if res:
                _JOBS[job_id]["resultado"] = analisar.gerar_html(res)
                log_q.put("__DONE__")
            else:
                log_q.put("__ERRO__")
        except Exception as exc:
            log_q.put(f"ERRO FATAL: {exc}")
            log_q.put("__ERRO__")
        finally:
            _JOBS[job_id]["done"] = True
            try:
                os.unlink(diss_tmp)
            except OSError:
                pass

    threading.Thread(target=_run, daemon=True).start()
    return jsonify({"job_id": job_id})


@app.route("/logs/<job_id>")
def logs(job_id):
    if job_id not in _JOBS:
        return "job não encontrado", 404

    def _gen():
        q = _JOBS[job_id]["log_q"]
        while True:
            try:
                msg = q.get(timeout=25)
                yield f"data: {json.dumps(msg)}\n\n"
                if msg in ("__DONE__", "__ERRO__"):
                    break
            except queue.Empty:
                yield 'data: "__PING__"\n\n'

    return Response(
        _gen(),
        mimetype="text/event-stream",
        headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"},
    )


@app.route("/resultado/<job_id>")
def resultado(job_id):
    if job_id not in _JOBS:
        return "Job não encontrado", 404
    job = _JOBS[job_id]
    html = job.get("resultado")
    if not html:
        if job.get("done"):
            return "<p style='font-family:sans-serif;color:#c00'>Análise não produziu resultado. Verifique o log.</p>", 200
        return "Análise ainda em andamento…", 202
    return html


@app.route("/parar/<job_id>", methods=["POST"])
def parar(job_id):
    if job_id in _JOBS:
        _JOBS[job_id]["stop"].set()
    return jsonify({"ok": True})


# ════════════════════════════════════════════════════════════════════════════
# PONTO DE ENTRADA
# ════════════════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    import sys

    port = 5000
    print()
    print("=" * 55)
    print("  VERIFICADOR DE CITAÇÕES — Interface Web v8.6")
    print("=" * 55)
    print(f"  Servidor iniciado em: http://localhost:{port}")
    print("  Abra esse endereço no seu navegador (Chrome).")
    print("  Para parar: pressione Ctrl+C")
    print("=" * 55)
    print()

    # No Android/Termux, webbrowser pode não funcionar — só imprime o URL
    try:
        import webbrowser
        webbrowser.open(f"http://localhost:{port}")
    except Exception:
        pass

    app.run(host="0.0.0.0", port=port, debug=False, threaded=True)
