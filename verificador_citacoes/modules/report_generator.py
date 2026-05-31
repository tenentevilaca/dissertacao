"""
Gera o relatório final em formato texto e HTML.
"""

import json
from datetime import datetime
from pathlib import Path

from modules.dissertation_parser import CitacaoInText, EntradaReferencia
from modules.citation_verifier import ResultadoVerificacao, Veredicto


# ──────────────────────────────────────────────
# Gerador de relatório
# ──────────────────────────────────────────────

class ReportGenerator:
    def __init__(
        self,
        citadas_sem_ref: list[CitacaoInText],
        refs_sem_citacao: list[EntradaReferencia],
        verificacoes: list[ResultadoVerificacao],
        caminho_dissertacao: str,
    ):
        self.citadas_sem_ref    = citadas_sem_ref
        self.refs_sem_citacao   = refs_sem_citacao
        self.verificacoes       = verificacoes
        self.caminho_dissertacao = caminho_dissertacao
        self.timestamp = datetime.now().strftime("%Y-%m-%d %H:%M")

    # ── Contadores ───────────────────────────

    @property
    def incorretas(self) -> list[ResultadoVerificacao]:
        return [v for v in self.verificacoes if v.veredicto == Veredicto.INCORRETO]

    @property
    def parciais(self) -> list[ResultadoVerificacao]:
        return [v for v in self.verificacoes if v.veredicto == Veredicto.PARCIAL]

    @property
    def corretas(self) -> list[ResultadoVerificacao]:
        return [v for v in self.verificacoes if v.veredicto == Veredicto.CORRETO]

    @property
    def sem_fonte(self) -> list[ResultadoVerificacao]:
        return [v for v in self.verificacoes if v.veredicto == Veredicto.SEM_FONTE]

    # ── Relatório TXT ─────────────────────────

    def gerar_txt(self, destino: Path) -> Path:
        linhas = self._montar_txt()
        destino.parent.mkdir(parents=True, exist_ok=True)
        destino.write_text("\n".join(linhas), encoding="utf-8")
        return destino

    def _montar_txt(self) -> list[str]:
        L = []
        sep = "=" * 70

        L += [
            sep,
            "RELATÓRIO DE VERIFICAÇÃO DE CITAÇÕES E REFERÊNCIAS",
            f"Dissertação: {self.caminho_dissertacao}",
            f"Gerado em: {self.timestamp}",
            sep,
            "",
        ]

        # ── Resumo executivo ─────────────────
        total_cit = len(self.verificacoes) + len(self.citadas_sem_ref)
        total_ref = len(self.verificacoes) + len(self.refs_sem_citacao) + len(self.sem_fonte)

        L += [
            "RESUMO EXECUTIVO",
            "-" * 40,
            f"  Total de citações identificadas      : {total_cit}",
            f"  Total de referências identificadas   : {total_ref}",
            f"  Citações verificadas (com fonte)     : {len(self.verificacoes)}",
            f"    ✓ Corretas                         : {len(self.corretas)}",
            f"    ⚠ Parcialmente suportadas          : {len(self.parciais)}",
            f"    ✗ Incorretas / distorcidas         : {len(self.incorretas)}",
            f"    ? Fonte não encontrada na pasta    : {len(self.sem_fonte)}",
            f"  Citações SEM entrada de referência   : {len(self.citadas_sem_ref)}",
            f"  Referências SEM citação no texto     : {len(self.refs_sem_citacao)}",
            "",
        ]

        # ── Seção 1 ──────────────────────────
        L += [
            sep,
            "SEÇÃO 1 — CITAÇÕES SEM ENTRADA NA LISTA DE REFERÊNCIAS",
            "(Obras citadas no texto que não constam na lista de referências)",
            sep,
            "",
        ]
        if self.citadas_sem_ref:
            for i, cit in enumerate(self.citadas_sem_ref, 1):
                autores_str = "; ".join(cit.autores)
                L += [
                    f"[{i}] Citação: {cit.texto_original}",
                    f"     Autores identificados: {autores_str}",
                    f"     Ano: {cit.ano}",
                    f"     Contexto: …{cit.contexto[200:400]}…",
                    f"     → AÇÃO: Adicionar referência completa no formato ABNT.",
                    "",
                ]
        else:
            L += ["  Nenhuma ocorrência encontrada.", ""]

        # ── Seção 2 ──────────────────────────
        L += [
            sep,
            "SEÇÃO 2 — REFERÊNCIAS SEM CITAÇÃO NO TEXTO",
            "(Entradas que constam na lista de referências mas não foram citadas)",
            sep,
            "",
        ]
        if self.refs_sem_citacao:
            for i, ref in enumerate(self.refs_sem_citacao, 1):
                L += [
                    f"[{i}] {ref.texto_completo}",
                    f"     → AÇÃO: Verificar se deve ser removida ou se falta a citação no texto.",
                    "",
                ]
        else:
            L += ["  Nenhuma ocorrência encontrada.", ""]

        # ── Seção 3 ──────────────────────────
        L += [
            sep,
            "SEÇÃO 3 — CITAÇÕES QUE NÃO REFLETEM O CONTEÚDO DA FONTE",
            "(Verificação semântica por IA — requer revisão humana)",
            sep,
            "",
        ]

        problemas = self.incorretas + self.parciais
        if problemas:
            for i, v in enumerate(problemas, 1):
                emoji = "✗" if v.veredicto == Veredicto.INCORRETO else "⚠"
                L += [
                    f"[{i}] {emoji} {v.veredicto.value}",
                    f"     Citação    : {v.citacao.texto_original}",
                    f"     Referência : {v.referencia.texto_completo[:120] if v.referencia else 'N/A'}",
                    f"     Arquivo    : {v.documento_fonte.nome_arquivo if v.documento_fonte else 'N/A'}",
                    f"     Diagnóstico: {v.justificativa}",
                    f"     Contexto dissertação:",
                    f"       …{v.citacao.contexto[100:350]}…",
                    "",
                ]
        else:
            L += ["  Nenhuma citação problemática encontrada.", ""]

        # ── Seção 4 — Fontes não localizadas ──
        if self.sem_fonte:
            L += [
                sep,
                "SEÇÃO 4 — CITAÇÕES CUJAS FONTES NÃO FORAM LOCALIZADAS NA PASTA",
                sep,
                "",
            ]
            for i, v in enumerate(self.sem_fonte, 1):
                L += [
                    f"[{i}] Citação    : {v.citacao.texto_original}",
                    f"     Referência : {v.referencia.texto_completo[:120] if v.referencia else 'N/A'}",
                    f"     → AÇÃO: Garantir que o arquivo da obra esteja na pasta 'fusion centers'.",
                    "",
                ]

        # ── Seção 5 — Citações corretas ───────
        L += [
            sep,
            "SEÇÃO 5 — CITAÇÕES VERIFICADAS E CORRETAS",
            sep,
            "",
        ]
        if self.corretas:
            for i, v in enumerate(self.corretas, 1):
                L += [
                    f"[{i}] ✓ {v.citacao.texto_original}",
                    f"     Arquivo: {v.documento_fonte.nome_arquivo if v.documento_fonte else 'N/A'}",
                    f"     Diagnóstico: {v.justificativa}",
                    "",
                ]
        else:
            L += ["  Nenhuma.", ""]

        L += ["", sep, "FIM DO RELATÓRIO", sep]
        return L

    # ── Relatório HTML ────────────────────────

    def gerar_html(self, destino: Path) -> Path:
        destino.parent.mkdir(parents=True, exist_ok=True)
        destino.write_text(self._montar_html(), encoding="utf-8")
        return destino

    def _montar_html(self) -> str:
        def esc(s: str) -> str:
            return (s or "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

        problemas = self.incorretas + self.parciais
        total_cit = len(self.verificacoes) + len(self.citadas_sem_ref)

        html = f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<title>Relatório de Citações</title>
<style>
  body {{ font-family: Arial, sans-serif; max-width: 900px; margin: 40px auto; color: #222; }}
  h1   {{ color: #1a3a5c; }}
  h2   {{ color: #1a3a5c; border-bottom: 2px solid #1a3a5c; padding-bottom: 4px; margin-top: 40px; }}
  .card {{ background: #f9f9f9; border-left: 5px solid #aaa; padding: 12px 16px; margin: 12px 0; border-radius: 4px; }}
  .card.incorreto {{ border-color: #d9534f; background: #fff5f5; }}
  .card.parcial   {{ border-color: #f0ad4e; background: #fffdf0; }}
  .card.correto   {{ border-color: #5cb85c; background: #f5fff5; }}
  .card.semref    {{ border-color: #5bc0de; background: #f0faff; }}
  .badge {{ display:inline-block; padding: 2px 8px; border-radius: 3px; font-size:.85em; font-weight:bold; color:#fff; }}
  .badge.incorreto {{ background:#d9534f; }}
  .badge.parcial   {{ background:#f0ad4e; color:#333; }}
  .badge.correto   {{ background:#5cb85c; }}
  .badge.semfonte  {{ background:#777; }}
  .sumario {{ display:flex; gap:16px; flex-wrap:wrap; margin:16px 0; }}
  .stat {{ background:#1a3a5c; color:#fff; padding:12px 20px; border-radius:6px; text-align:center; }}
  .stat strong {{ display:block; font-size:2em; }}
  pre {{ background:#eee; padding:8px; font-size:.85em; overflow-x:auto; }}
  footer {{ margin-top:60px; font-size:.8em; color:#888; }}
</style>
</head>
<body>
<h1>Relatório de Verificação de Citações e Referências</h1>
<p><strong>Dissertação:</strong> {esc(self.caminho_dissertacao)}</p>
<p><strong>Gerado em:</strong> {self.timestamp}</p>

<div class="sumario">
  <div class="stat"><strong>{total_cit}</strong>Citações totais</div>
  <div class="stat"><strong>{len(self.corretas)}</strong>Corretas ✓</div>
  <div class="stat" style="background:#f0ad4e;color:#333"><strong>{len(self.parciais)}</strong>Parciais ⚠</div>
  <div class="stat" style="background:#d9534f"><strong>{len(self.incorretas)}</strong>Incorretas ✗</div>
  <div class="stat" style="background:#777"><strong>{len(self.sem_fonte)}</strong>Sem fonte</div>
  <div class="stat" style="background:#5bc0de"><strong>{len(self.citadas_sem_ref)}</strong>Sem referência</div>
  <div class="stat" style="background:#444"><strong>{len(self.refs_sem_citacao)}</strong>Ref. sem citação</div>
</div>

<h2>1 — Citações sem entrada na lista de referências</h2>
"""
        if self.citadas_sem_ref:
            for cit in self.citadas_sem_ref:
                html += f"""<div class="card semref">
  <span class="badge semfonte">SEM REFERÊNCIA</span>
  <p><strong>Citação:</strong> {esc(cit.texto_original)}</p>
  <p><strong>Autores:</strong> {esc("; ".join(cit.autores))} &nbsp; <strong>Ano:</strong> {esc(cit.ano)}</p>
  <p><em>Contexto:</em> …{esc(cit.contexto[200:400])}…</p>
  <p style="color:#d9534f">→ Adicionar referência completa no formato ABNT.</p>
</div>
"""
        else:
            html += "<p>Nenhuma ocorrência.</p>"

        html += "<h2>2 — Referências sem citação no texto</h2>\n"
        if self.refs_sem_citacao:
            for ref in self.refs_sem_citacao:
                html += f"""<div class="card">
  <p>{esc(ref.texto_completo)}</p>
  <p style="color:#d9534f">→ Verificar se deve ser removida ou se falta citação no texto.</p>
</div>
"""
        else:
            html += "<p>Nenhuma ocorrência.</p>"

        html += "<h2>3 — Citações que não refletem o conteúdo da fonte</h2>\n"
        if problemas:
            for v in problemas:
                cls = "incorreto" if v.veredicto == Veredicto.INCORRETO else "parcial"
                html += f"""<div class="card {cls}">
  <span class="badge {cls}">{esc(v.veredicto.value)}</span>
  <p><strong>Citação:</strong> {esc(v.citacao.texto_original)}</p>
  <p><strong>Referência:</strong> {esc(v.referencia.texto_completo[:150] if v.referencia else "N/A")}</p>
  <p><strong>Arquivo-fonte:</strong> {esc(v.documento_fonte.nome_arquivo if v.documento_fonte else "N/A")}</p>
  <p><strong>Diagnóstico:</strong> {esc(v.justificativa)}</p>
  <pre>{esc(v.citacao.contexto[100:400])}</pre>
</div>
"""
        else:
            html += "<p>Nenhuma citação problemática encontrada.</p>"

        if self.sem_fonte:
            html += "<h2>4 — Citações cujas fontes não foram localizadas na pasta</h2>\n"
            for v in self.sem_fonte:
                html += f"""<div class="card">
  <span class="badge semfonte">SEM ARQUIVO</span>
  <p><strong>Citação:</strong> {esc(v.citacao.texto_original)}</p>
  <p><strong>Referência:</strong> {esc(v.referencia.texto_completo[:150] if v.referencia else "N/A")}</p>
  <p style="color:#d9534f">→ Garantir que o arquivo da obra esteja na pasta 'fusion centers'.</p>
</div>
"""

        html += "<h2>5 — Citações verificadas e corretas</h2>\n"
        if self.corretas:
            for v in self.corretas:
                html += f"""<div class="card correto">
  <span class="badge correto">CORRETO</span>
  <p><strong>Citação:</strong> {esc(v.citacao.texto_original)}</p>
  <p><strong>Arquivo:</strong> {esc(v.documento_fonte.nome_arquivo if v.documento_fonte else "N/A")}</p>
  <p>{esc(v.justificativa)}</p>
</div>
"""
        else:
            html += "<p>Nenhuma.</p>"

        html += f"""
<footer>Relatório gerado automaticamente pelo Verificador de Citações — {self.timestamp}</footer>
</body></html>"""
        return html

    # ── JSON para auditoria ───────────────────

    def gerar_json(self, destino: Path) -> Path:
        destino.parent.mkdir(parents=True, exist_ok=True)
        dados = {
            "metadados": {
                "dissertacao": self.caminho_dissertacao,
                "gerado_em": self.timestamp,
            },
            "resumo": {
                "citacoes_sem_referencia": len(self.citadas_sem_ref),
                "referencias_sem_citacao": len(self.refs_sem_citacao),
                "corretas": len(self.corretas),
                "parciais": len(self.parciais),
                "incorretas": len(self.incorretas),
                "sem_fonte": len(self.sem_fonte),
            },
            "citadas_sem_referencia": [
                {"citacao": c.texto_original, "autores": c.autores, "ano": c.ano}
                for c in self.citadas_sem_ref
            ],
            "referenciadas_sem_citacao": [
                {"referencia": r.texto_completo, "chave": r.chave}
                for r in self.refs_sem_citacao
            ],
            "verificacoes": [
                {
                    "citacao": v.citacao.texto_original,
                    "referencia": v.referencia.texto_completo if v.referencia else None,
                    "arquivo": v.documento_fonte.nome_arquivo if v.documento_fonte else None,
                    "veredicto": v.veredicto.value,
                    "justificativa": v.justificativa,
                }
                for v in self.verificacoes
            ],
        }
        destino.write_text(json.dumps(dados, ensure_ascii=False, indent=2), encoding="utf-8")
        return destino
