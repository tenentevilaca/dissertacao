"""
Analisador de Dissertação — análise crítica estrutural baseada no template genérico.
Reusa ler_docx, ler_arquivo e _extrair_trecho_relevante do verificador.py.
"""

import json
import re
import zipfile
from pathlib import Path
from typing import Optional

import anthropic

# ── Importações do módulo irmão ──────────────────────────────────────────
from verificador import ler_docx, ler_arquivo, _extrair_trecho_relevante, _STOPWORDS_PT

_MAX_CHARS_CAP   = 14_000
_MAX_CHARS_META  = 7_000
_MAX_CHARS_APOIO = 5_000


# ════════════════════════════════════════════════════════════════════════
# DOWNLOAD DO GOOGLE DRIVE
# ════════════════════════════════════════════════════════════════════════

def _extrair_id_drive(url: str) -> str:
    """Extrai o file ID de qualquer formato de URL do Google Drive."""
    # /file/d/ID/view  ou  /d/ID/
    m = re.search(r'/d/([a-zA-Z0-9_-]{10,})', url)
    if m:
        return m.group(1)
    # ?id=ID  ou  &id=ID
    m = re.search(r'[?&]id=([a-zA-Z0-9_-]{10,})', url)
    if m:
        return m.group(1)
    raise ValueError(
        "Não foi possível identificar o ID do arquivo na URL. "
        "Certifique-se de usar o link de compartilhamento do Google Drive."
    )


def baixar_google_drive(url: str, destino: Path, log_fn=None) -> Path:
    """
    Baixa um arquivo público do Google Drive usando httpx (já incluído via anthropic).
    Retorna o caminho do arquivo salvo.
    """
    import httpx

    file_id = _extrair_id_drive(url)
    if log_fn:
        log_fn(f"   ID do arquivo: {file_id}")

    headers = {"User-Agent": "Mozilla/5.0"}
    base_url = "https://drive.google.com/uc"
    params = {"export": "download", "id": file_id}

    if log_fn:
        log_fn("   Conectando ao Google Drive…")

    with httpx.Client(headers=headers, follow_redirects=True, timeout=300) as client:
        resp = client.get(base_url, params=params)

        content_type = resp.headers.get("Content-Type", "")
        if "text/html" in content_type:
            if log_fn:
                log_fn("   Arquivo grande — obtendo token de confirmação…")
            m = re.search(r'name="uuid"\s+value="([^"]+)"', resp.text)
            uuid = m.group(1) if m else None
            m2 = re.search(r'action="([^"]+confirm[^"]+)"', resp.text)
            if m2 and uuid:
                confirm_url = m2.group(1).replace("&amp;", "&")
                resp = client.post(confirm_url, data={"uuid": uuid})
            else:
                params["confirm"] = "t"
                resp = client.get(base_url, params=params)

        resp.raise_for_status()

        ext = ".zip"
        cd = resp.headers.get("Content-Disposition", "")
        m_cd = re.search(r'filename[^;=\n]*=\s*["\']?([^"\'\n;]+)', cd)
        if m_cd:
            ext = Path(m_cd.group(1).strip()).suffix or ext
        elif "pdf" in content_type:
            ext = ".pdf"
        elif "msword" in content_type or "wordprocessing" in content_type:
            ext = ".docx"

        caminho_final = destino.with_suffix(ext)
        if log_fn:
            log_fn(f"   Salvando como {caminho_final.name}…")

        with open(caminho_final, "wb") as f:
            f.write(resp.content)

        total = len(resp.content)
        if log_fn:
            log_fn(f"   Download concluido: {total/(1024*1024):.1f} MB -> {caminho_final.name}")

        return caminho_final


# ════════════════════════════════════════════════════════════════════════
# PROMPTS
# ════════════════════════════════════════════════════════════════════════

_PROMPT_META = """Você é um assistente especializado em análise de textos acadêmicos brasileiros.

Leia o início desta dissertação e extraia as informações bibliográficas e estruturais.

Retorne SOMENTE um JSON válido (sem markdown, sem texto antes ou depois):
{
  "autor": "Nome completo do autor ou null",
  "titulo": "Título completo da dissertação ou null",
  "programa": "Nome do programa de pós-graduação ou null",
  "instituicao": "Nome da instituição ou null",
  "orientador": "Nome do orientador ou null",
  "area": "Área de concentração ou null",
  "ano": "Ano de conclusão/defesa ou null",
  "problema_pesquisa": "O problema central da pesquisa em 1-2 frases ou null",
  "hipotese": "A hipótese ou tese central em 1-2 frases ou null",
  "objetivos": ["objetivo específico 1", "objetivo específico 2"],
  "capitulos_detectados": ["Título Cap 1", "Título Cap 2"]
}

INÍCIO DA DISSERTAÇÃO:
"""

_PROMPT_ANALISE = """Você é um orientador acadêmico experiente analisando criticamente uma dissertação brasileira.

METADADOS DA DISSERTAÇÃO:
{metadados}

CAPÍTULO EM ANÁLISE: «{titulo_cap}»

TEXTO DO CAPÍTULO (até {n_chars} caracteres):
\"\"\"
{texto_cap}
\"\"\"

Realize a análise crítica completa conforme o framework abaixo.
Retorne SOMENTE um JSON válido (sem markdown, sem texto antes ou depois):
{{
  "titulo_capitulo": "{titulo_cap}",
  "problemas_estruturais": [
    {{
      "descricao": "descrição objetiva do problema",
      "gravidade": "ALTA",
      "localizacao": "onde ocorre (seção, início, fim, ao longo do texto)"
    }}
  ],
  "adequacao_teorica": {{
    "avaliacao": "ADEQUADA",
    "pontos_positivos": ["ponto 1", "ponto 2"],
    "lacunas": ["lacuna 1"]
  }},
  "integracao_empirica": {{
    "avaliacao": "INTEGRADA",
    "observacoes": "descrição de como teoria e dados se articulam (ou não)"
  }},
  "conformidade_abnt": {{
    "problemas_encontrados": ["descrição de problema ABNT"],
    "status": "OK"
  }},
  "estilo_escrita": {{
    "paragrafos_que_iniciam_com_autor": 0,
    "uso_de_travessao": false,
    "uso_excessivo_dois_pontos": false,
    "transicoes_entre_secoes": "PRESENTES",
    "listas_desnecessarias": false,
    "observacoes": "observações gerais sobre estilo"
  }},
  "recomendacoes_priorizadas": [
    "1. [ALTA] recomendação prioritária",
    "2. [MEDIA] recomendação secundária"
  ],
  "pontuacao_geral": "APROVADO_COM_RESSALVAS",
  "resumo_executivo": "Em 3-4 frases: o que está bem e o que precisa de atenção."
}}

Onde "avaliacao"/"status"/"pontuacao_geral" usam apenas: APROVADO | APROVADO_COM_RESSALVAS | REQUER_REVISAO | ADEQUADA | INSUFICIENTE | AUSENTE | INTEGRADA | PARCIAL | OK | COM_PROBLEMAS
"""

_PROMPT_APOIO = """Você é um assistente de pesquisa acadêmica ajudando a validar argumentos de uma dissertação.

CAPÍTULO: «{titulo_cap}»

TRECHO DO CAPÍTULO QUE PRECISA DE VALIDAÇÃO:
\"\"\"
{trecho_cap}
\"\"\"

TRECHOS DO MATERIAL DE APOIO DISPONÍVEL:
\"\"\"
{trechos_apoio}
\"\"\"

Analise se os trechos do material de apoio sustentam, contradizem ou complementam os argumentos do capítulo.

Retorne SOMENTE um JSON válido (sem markdown, sem texto antes ou depois):
{{
  "pares_encontrados": [
    {{
      "argumento_dissertacao": "argumento ou afirmação do capítulo (cite diretamente)",
      "trecho_apoio": "trecho exato do material que se relaciona (até 300 chars)",
      "arquivo_fonte": "nome do arquivo de onde veio o trecho",
      "relacao": "VALIDA",
      "explicacao": "por que valida/contradiz/complementa"
    }}
  ],
  "cobertura": "ALTA",
  "sintese": "Em 2-3 frases: avaliação geral do suporte do material de apoio para este capítulo."
}}

Onde "relacao" usa: VALIDA | CONTRADIZ | COMPLEMENTA
Onde "cobertura" usa: ALTA | MEDIA | BAIXA | NENHUMA
"""


# ════════════════════════════════════════════════════════════════════════
# DETECÇÃO DE CAPÍTULOS
# ════════════════════════════════════════════════════════════════════════

# Padrões de título de capítulo
_RE_CAP = re.compile(
    r"^(?:"
    r"CAP[IÍ]TULO\s+\d+"                    # CAPÍTULO 1
    r"|(?:CHAPTER|PARTE)\s+\d+"              # CHAPTER 1
    r"|\d{1,2}(?:\.\d{1,2})?\s+[A-ZÁÉÍ]"   # 1. Introdução / 2.1 Contexto
    r"|[IVX]{1,5}\s+[A-ZÁÉÍ]"              # IV Metodologia
    r")",
    re.IGNORECASE,
)

# Também captura linhas curtas em maiúsculas (provável título)
_RE_TITULO_CURTO = re.compile(
    r"^[A-ZÁÉÍÓÚÂÊÎÔÛÃÕÀÇ][A-ZÁÉÍÓÚÂÊÎÔÛÃÕÀÇ\s\-\:]{3,60}$"
)

_SECOES_IGNORAR = {
    "REFERÊNCIAS", "REFERENCIAS", "REFERÊNCIAS BIBLIOGRÁFICAS",
    "REFERÊNCIAS BIBLIOGRAFICAS", "SUMÁRIO", "SUMARIO", "ABSTRACT",
    "RESUMO", "LISTA DE FIGURAS", "LISTA DE TABELAS", "LISTA DE ABREVIATURAS",
    "AGRADECIMENTOS", "DEDICATÓRIA", "EPÍGRAFE", "ÍNDICE",
}


def detectar_capitulos(texto: str) -> list[tuple[str, str]]:
    """
    Detecta capítulos no texto e retorna lista de (titulo, conteudo).
    Estratégia em dois níveis:
    - Padrões numéricos/CAPÍTULO → sempre são cortes
    - Títulos curtos em caps → apenas se distantes 15+ linhas do anterior
    """
    paragrafos = texto.split("\n")
    cortes: list[tuple[int, str]] = []  # (idx_paragrafo, titulo)

    for i, p in enumerate(paragrafos):
        stripped = p.strip()
        if not stripped or len(stripped) > 90:
            continue
        upper = stripped.upper()
        if upper in _SECOES_IGNORAR:
            continue

        e_cap_explicito = bool(_RE_CAP.match(stripped))
        e_titulo_curto  = (not e_cap_explicito
                           and bool(_RE_TITULO_CURTO.match(stripped))
                           and len(stripped.split()) >= 2)

        if e_cap_explicito:
            cortes.append((i, stripped))
        elif e_titulo_curto:
            # Só inclui se distante 15+ linhas do último corte
            if not cortes or (i - cortes[-1][0]) >= 15:
                cortes.append((i, stripped))

    if not cortes:
        return [("Dissertação Completa", texto[:_MAX_CHARS_CAP])]

    # Mescla cortes seguidos sem conteúdo entre eles (títulos multi-linha)
    cortes_filtrados: list[tuple[int, str]] = []
    for j, (idx, titulo) in enumerate(cortes):
        if not cortes_filtrados:
            cortes_filtrados.append((idx, titulo))
            continue
        anterior_idx, anterior_titulo = cortes_filtrados[-1]
        linhas_entre = [
            paragrafos[k].strip()
            for k in range(anterior_idx + 1, idx)
            if paragrafos[k].strip()
        ]
        if len(linhas_entre) < 3:
            # Provavelmente sub-título ou continuação — agrega ao título anterior
            cortes_filtrados[-1] = (anterior_idx, f"{anterior_titulo} — {titulo}")
        else:
            cortes_filtrados.append((idx, titulo))

    capitulos: list[tuple[str, str]] = []
    for k, (idx, titulo) in enumerate(cortes_filtrados):
        inicio = idx + 1
        fim = cortes_filtrados[k + 1][0] if k + 1 < len(cortes_filtrados) else len(paragrafos)
        conteudo = "\n".join(paragrafos[inicio:fim]).strip()
        if len(conteudo) > 150:
            capitulos.append((titulo, conteudo))

    return capitulos if capitulos else [("Dissertação Completa", texto[:_MAX_CHARS_CAP])]


# ════════════════════════════════════════════════════════════════════════
# EXTRAÇÃO DE MATERIAL DE APOIO (ZIP)
# ════════════════════════════════════════════════════════════════════════

EXTENSOES_APOIO = {".pdf", ".docx", ".doc", ".txt"}


def extrair_material_de_zip(zip_path: Path, destino: Path) -> dict[str, str]:
    """
    Extrai um ZIP e lê todos os arquivos suportados.
    Retorna dict nome_arquivo → texto.
    """
    destino.mkdir(parents=True, exist_ok=True)

    # Ignora arquivos temporários do Office
    with zipfile.ZipFile(str(zip_path)) as z:
        for info in z.infolist():
            nome = Path(info.filename).name
            if nome.startswith("~$") or info.is_dir():
                continue
            ext = Path(info.filename).suffix.lower()
            if ext in EXTENSOES_APOIO:
                z.extract(info, destino)

    material: dict[str, str] = {}
    for arq in destino.rglob("*"):
        if arq.is_file() and arq.suffix.lower() in EXTENSOES_APOIO:
            if arq.name.startswith("~$"):
                continue
            texto = ler_arquivo(arq)
            if texto and len(texto) > 100 and not texto.startswith("[ERRO"):
                material[arq.name] = texto

    return material


def carregar_material_de_pasta(pasta: Path) -> dict[str, str]:
    """
    Lê arquivos de uma pasta já extraída (quando vieram como uploads individuais).
    """
    material: dict[str, str] = {}
    for arq in pasta.rglob("*"):
        if arq.is_file() and arq.suffix.lower() in EXTENSOES_APOIO:
            if arq.name.startswith("~$"):
                continue
            texto = ler_arquivo(arq)
            if texto and len(texto) > 100 and not texto.startswith("[ERRO"):
                material[arq.name] = texto
    return material


# ════════════════════════════════════════════════════════════════════════
# CHAMADAS AO CLAUDE
# ════════════════════════════════════════════════════════════════════════

def _chamar_claude(client: anthropic.Anthropic, prompt: str, max_tokens: int = 4096) -> dict:
    """Chama o Claude e parseia o JSON retornado."""
    resp = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=max_tokens,
        messages=[{"role": "user", "content": prompt}],
    )
    raw = resp.content[0].text.strip()
    # Remove markdown se presente
    raw = re.sub(r"^```(?:json)?\s*", "", raw, flags=re.MULTILINE)
    raw = re.sub(r"\s*```\s*$", "", raw, flags=re.MULTILINE)
    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        m = re.search(r"\{[\s\S]*\}", raw)
        if m:
            try:
                return json.loads(m.group())
            except json.JSONDecodeError:
                pass
        return {"_erro_parse": raw[:500]}


def extrair_metadados(texto: str, api_key: str) -> dict:
    client = anthropic.Anthropic(api_key=api_key)
    prompt = _PROMPT_META + texto[:_MAX_CHARS_META]
    return _chamar_claude(client, prompt, max_tokens=2048)


def analisar_capitulo(
    titulo_cap: str,
    texto_cap: str,
    metadados: dict,
    api_key: str,
) -> dict:
    client = anthropic.Anthropic(api_key=api_key)
    meta_str = json.dumps(metadados, ensure_ascii=False, indent=2)
    trecho = texto_cap[:_MAX_CHARS_CAP]
    prompt = _PROMPT_ANALISE.format(
        metadados=meta_str[:1000],
        titulo_cap=titulo_cap,
        texto_cap=trecho,
        n_chars=len(trecho),
    )
    return _chamar_claude(client, prompt, max_tokens=3000)


def buscar_argumentos_apoio(
    titulo_cap: str,
    texto_cap: str,
    material: dict[str, str],
    api_key: str,
) -> dict:
    if not material:
        return {"pares_encontrados": [], "cobertura": "NENHUMA",
                "sintese": "Nenhum material de apoio disponível."}

    # Usa a função já testada do verificador para encontrar trecho relevante
    # Agrega os top-3 documentos mais relevantes
    trechos_selecionados: list[str] = []
    chars_usados = 0

    for nome_arq, texto_arq in material.items():
        trecho = _extrair_trecho_relevante(texto_cap, texto_arq, janela=2000)
        if trecho.strip() and len(trecho.strip()) > 100:
            bloco = f"[ARQUIVO: {nome_arq}]\n{trecho}"
            if chars_usados + len(bloco) <= _MAX_CHARS_APOIO * 3:
                trechos_selecionados.append(bloco)
                chars_usados += len(bloco)
        if chars_usados >= _MAX_CHARS_APOIO * 3:
            break

    if not trechos_selecionados:
        return {"pares_encontrados": [], "cobertura": "NENHUMA",
                "sintese": "Nenhum trecho relevante encontrado no material de apoio."}

    trechos_apoio = "\n\n---\n\n".join(trechos_selecionados)

    client = anthropic.Anthropic(api_key=api_key)
    prompt = _PROMPT_APOIO.format(
        titulo_cap=titulo_cap,
        trecho_cap=texto_cap[:2000],
        trechos_apoio=trechos_apoio[:_MAX_CHARS_APOIO * 3],
    )
    return _chamar_claude(client, prompt, max_tokens=3000)


# ════════════════════════════════════════════════════════════════════════
# ORQUESTRADOR PRINCIPAL
# ════════════════════════════════════════════════════════════════════════

def analisar_dissertacao(
    diss_path: str,
    material: dict[str, str],
    api_key: str,
    log_fn,
) -> dict:
    """
    Pipeline completo:
      1. Lê dissertação
      2. Extrai metadados via IA
      3. Detecta capítulos
      4. Analisa cada capítulo via IA
      5. Busca argumentos de apoio no material
    """

    # ── Etapa 1: Leitura ────────────────────────────────────────────────
    log_fn("📄 ETAPA 1: Lendo a dissertação…", "etapa")
    texto = ler_docx(diss_path)
    n = len(texto)
    log_fn(f"   {n:,} caracteres extraídos")
    if n < 300:
        log_fn("   ⚠ Texto muito curto — o arquivo pode estar vazio ou corrompido", "warn")

    # ── Etapa 2: Metadados ──────────────────────────────────────────────
    log_fn("🔍 ETAPA 2: Extraindo metadados via IA…", "etapa")
    try:
        metadados = extrair_metadados(texto, api_key)
        autor = metadados.get("autor") or "Não identificado"
        titulo = metadados.get("titulo") or "Não identificado"
        log_fn(f"   Autor : {autor}")
        log_fn(f"   Título: {titulo}")
        log_fn(f"   Instituição: {metadados.get('instituicao') or '—'}")
        log_fn(f"   Capítulos detectados pela IA: {len(metadados.get('capitulos_detectados') or [])}")
    except Exception as e:
        log_fn(f"   ⚠ Erro ao extrair metadados: {e}", "warn")
        metadados = {}

    # ── Etapa 3: Detecção de capítulos ──────────────────────────────────
    log_fn("📑 ETAPA 3: Detectando estrutura de capítulos…", "etapa")
    capitulos = detectar_capitulos(texto)
    log_fn(f"   {len(capitulos)} capítulo(s) / seção(ões) detectado(s)")
    for titulo_cap, conteudo in capitulos[:5]:
        log_fn(f"   • {titulo_cap[:60]} ({len(conteudo):,} chars)")

    # ── Etapas 4+5: Análise por capítulo ────────────────────────────────
    resultados_caps: list[dict] = []
    total = len(capitulos)
    n_mat = len(material)

    for i, (titulo_cap, conteudo) in enumerate(capitulos, 1):
        log_fn(f"🤖 [{i}/{total}] Analisando: «{titulo_cap[:55]}»…", "etapa")

        # Análise crítica
        try:
            analise = analisar_capitulo(titulo_cap, conteudo, metadados, api_key)
            pontuacao = analise.get("pontuacao_geral", "—")
            resumo = analise.get("resumo_executivo", "")
            log_fn(f"   Pontuação: {pontuacao}")
            if resumo:
                log_fn(f"   {resumo[:120]}")
        except Exception as e:
            log_fn(f"   ✗ Erro na análise: {e}", "warn")
            analise = {"titulo_capitulo": titulo_cap, "_erro": str(e)}

        # Busca de argumentos de apoio
        apoio: dict = {}
        if n_mat > 0:
            log_fn(f"   🔎 Buscando argumentos de apoio ({n_mat} arquivo(s))…")
            try:
                apoio = buscar_argumentos_apoio(titulo_cap, conteudo, material, api_key)
                n_pares = len(apoio.get("pares_encontrados") or [])
                cobertura = apoio.get("cobertura", "—")
                log_fn(f"   {n_pares} par(es) encontrado(s) | cobertura: {cobertura}")
            except Exception as e:
                log_fn(f"   ⚠ Erro na busca de apoio: {e}", "warn")
                apoio = {"_erro": str(e)}
        else:
            log_fn("   ℹ Sem material de apoio — etapa pulada")

        resultados_caps.append({
            "titulo": titulo_cap,
            "n_chars": len(conteudo),
            "analise": analise,
            "apoio": apoio,
        })

    log_fn(f"✅ Análise concluída — {len(resultados_caps)} capítulo(s) processado(s)", "ok")

    return {
        "metadados": metadados,
        "n_capitulos": len(capitulos),
        "n_arquivos_apoio": n_mat,
        "capitulos": resultados_caps,
    }


# ════════════════════════════════════════════════════════════════════════
# GERAÇÃO DE RELATÓRIO
# ════════════════════════════════════════════════════════════════════════

_COR_PONT = {
    "APROVADO": "#16a34a",
    "APROVADO_COM_RESSALVAS": "#d97706",
    "REQUER_REVISAO": "#dc2626",
}
_ICON_PONT = {
    "APROVADO": "✅",
    "APROVADO_COM_RESSALVAS": "⚠️",
    "REQUER_REVISAO": "❌",
}
_COR_REL = {
    "VALIDA": "#16a34a",
    "CONTRADIZ": "#dc2626",
    "COMPLEMENTA": "#2563eb",
}


def _esc(s) -> str:
    return str(s or "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def gerar_relatorio_analise(resultado: dict, rel_dir: Path) -> None:
    from datetime import datetime

    rel_dir = Path(rel_dir)
    rel_dir.mkdir(parents=True, exist_ok=True)

    meta = resultado.get("metadados") or {}
    caps = resultado.get("capitulos") or []
    ts = datetime.now().strftime("%Y-%m-%d %H:%M")

    # ── JSON ─────────────────────────────────────────────────────────────
    (rel_dir / "relatorio_analise.json").write_text(
        json.dumps(resultado, ensure_ascii=False, indent=2), encoding="utf-8"
    )

    # ── Estatísticas gerais ──────────────────────────────────────────────
    pontuacoes = [c["analise"].get("pontuacao_geral", "") for c in caps]
    n_aprovado    = sum(1 for p in pontuacoes if p == "APROVADO")
    n_ressalvas   = sum(1 for p in pontuacoes if p == "APROVADO_COM_RESSALVAS")
    n_revisao     = sum(1 for p in pontuacoes if p == "REQUER_REVISAO")
    total_pares   = sum(len(c["apoio"].get("pares_encontrados") or []) for c in caps)

    # ── HTML ─────────────────────────────────────────────────────────────
    autor     = _esc(meta.get("autor") or "Não identificado")
    titulo_d  = _esc(meta.get("titulo") or "Dissertação")
    inst      = _esc(meta.get("instituicao") or "—")
    orientador= _esc(meta.get("orientador") or "—")
    programa  = _esc(meta.get("programa") or "—")
    area      = _esc(meta.get("area") or "—")
    ano       = _esc(meta.get("ano") or "—")

    problema  = _esc(meta.get("problema_pesquisa") or "—")
    hipotese  = _esc(meta.get("hipotese") or "—")
    objetivos = meta.get("objetivos") or []

    html = f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Análise — {titulo_d}</title>
<style>
*, *::before, *::after {{ box-sizing: border-box; margin: 0; padding: 0; }}
:root {{
  --azul: #1a3a5c; --azul2: #2563eb; --verde: #16a34a;
  --amarelo: #d97706; --vermelho: #dc2626; --cinza: #64748b;
  --fundo: #f1f5f9; --branco: #ffffff; --borda: #e2e8f0;
}}
body {{ font-family: 'Segoe UI', system-ui, sans-serif; background: var(--fundo);
        color: #1e293b; line-height: 1.6; }}
header {{ background: var(--azul); color: #fff; padding: 24px 32px; }}
header h1 {{ font-size: 1.3rem; margin-bottom: 4px; }}
header p  {{ font-size: .85rem; opacity: .75; }}
main {{ max-width: 960px; margin: 0 auto; padding: 32px 16px 80px; }}
h2 {{ font-size: 1.05rem; color: var(--azul); margin: 28px 0 14px; }}
h3 {{ font-size: .95rem; color: #334155; margin: 16px 0 8px; }}
.card {{ background: var(--branco); border-radius: 12px; border: 1px solid var(--borda);
         padding: 24px; margin-bottom: 20px; box-shadow: 0 2px 8px rgba(0,0,0,.06); }}
.meta-grid {{ display: grid; grid-template-columns: repeat(auto-fill, minmax(200px, 1fr));
              gap: 12px; margin-bottom: 20px; }}
.meta-item {{ background: #f8fafc; border-radius: 8px; padding: 12px;
              border: 1px solid var(--borda); }}
.meta-item label {{ font-size: .72rem; font-weight: 700; color: var(--cinza);
                    text-transform: uppercase; letter-spacing: .05em; display: block; margin-bottom: 4px; }}
.meta-item span  {{ font-size: .9rem; color: #1e293b; font-weight: 600; }}
.stats-row {{ display: flex; gap: 12px; flex-wrap: wrap; margin-bottom: 24px; }}
.stat {{ border-radius: 10px; padding: 14px 20px; text-align: center; min-width: 110px;
         border: 1px solid var(--borda); }}
.stat strong {{ display: block; font-size: 2rem; font-weight: 800; }}
.stat span   {{ font-size: .75rem; color: var(--cinza); }}
.s-verde  {{ background: #f0fdf4; color: var(--verde); }}
.s-amarelo{{ background: #fffbeb; color: var(--amarelo); }}
.s-vermelho{{ background: #fef2f2; color: var(--vermelho); }}
.s-azul   {{ background: #eff6ff; color: var(--azul2); }}
.cap-card {{ background: var(--branco); border-radius: 12px; border: 1px solid var(--borda);
             margin-bottom: 20px; overflow: hidden; }}
.cap-header {{ padding: 16px 20px; cursor: pointer; display: flex; justify-content: space-between;
               align-items: center; user-select: none; }}
.cap-header:hover {{ background: #f8fafc; }}
.cap-title  {{ font-weight: 700; font-size: .95rem; color: var(--azul); }}
.badge {{ display: inline-block; padding: 3px 10px; border-radius: 20px;
          font-size: .72rem; font-weight: 700; letter-spacing: .04em; }}
.badge-verde    {{ background: #dcfce7; color: var(--verde); }}
.badge-amarelo  {{ background: #fef9c3; color: #92400e; }}
.badge-vermelho {{ background: #fee2e2; color: var(--vermelho); }}
.badge-azul     {{ background: #dbeafe; color: var(--azul2); }}
.badge-cinza    {{ background: #f1f5f9; color: var(--cinza); }}
.cap-body {{ padding: 0 20px 20px; border-top: 1px solid var(--borda); display: none; }}
.cap-body.aberto {{ display: block; }}
.bloco {{ margin-top: 16px; }}
.bloco-titulo {{ font-size: .78rem; font-weight: 700; color: var(--cinza);
                 text-transform: uppercase; letter-spacing: .05em; margin-bottom: 8px; }}
ul.lista {{ padding-left: 18px; }}
ul.lista li {{ margin-bottom: 4px; font-size: .88rem; }}
.par-item {{ background: #f8fafc; border-radius: 8px; padding: 12px; margin-bottom: 8px;
             border-left: 4px solid var(--borda); }}
.par-item.valida     {{ border-left-color: var(--verde); }}
.par-item.contradiz  {{ border-left-color: var(--vermelho); }}
.par-item.complementa{{ border-left-color: var(--azul2); }}
.par-item .rel-tag {{ font-size: .7rem; font-weight: 700; padding: 2px 8px;
                      border-radius: 10px; display: inline-block; margin-bottom: 6px; }}
.par-item .argumento {{ font-size: .85rem; color: #334155; margin-bottom: 6px; }}
.par-item .trecho    {{ font-size: .82rem; color: #64748b; font-style: italic;
                        border-left: 2px solid #cbd5e1; padding-left: 8px; }}
.par-item .fonte     {{ font-size: .72rem; color: var(--cinza); margin-top: 6px; }}
.sintese {{ background: #f0f9ff; border-radius: 8px; padding: 14px;
            border: 1px solid #bae6fd; font-size: .88rem; color: #0369a1; margin-top: 12px; }}
.resumo-exec {{ background: #f0fdf4; border-radius: 8px; padding: 14px;
                border: 1px solid #bbf7d0; font-size: .88rem; color: #166534; margin-bottom: 12px; }}
.prob-item {{ background: #fef2f2; border-radius: 6px; padding: 10px 12px;
              margin-bottom: 6px; font-size: .85rem; border-left: 3px solid #fca5a5; }}
.prob-item.media  {{ background: #fffbeb; border-left-color: #fcd34d; }}
.prob-item.baixa  {{ background: #f8fafc; border-left-color: #cbd5e1; }}
.prob-loc {{ font-size: .72rem; color: var(--cinza); margin-top: 4px; }}
table.abnt {{ width: 100%; border-collapse: collapse; font-size: .84rem; margin-top: 8px; }}
table.abnt td {{ padding: 6px 10px; border-bottom: 1px solid var(--borda); }}
table.abnt tr:last-child td {{ border-bottom: none; }}
.check {{ color: var(--verde); font-weight: 700; }}
.cross {{ color: var(--vermelho); font-weight: 700; }}
.problema-pesquisa {{ background: #eff6ff; border-radius: 8px; padding: 14px;
                      border: 1px solid #bfdbfe; font-size: .9rem; margin-bottom: 12px; }}
.chip {{ display: inline-block; background: #dbeafe; color: var(--azul2);
         border-radius: 20px; padding: 3px 12px; font-size: .78rem;
         font-weight: 600; margin: 3px 3px 0 0; }}
footer {{ text-align: center; font-size: .78rem; color: var(--cinza); margin-top: 40px; padding: 20px; }}
@media (max-width: 600px) {{
  .meta-grid {{ grid-template-columns: 1fr 1fr; }}
  .stats-row {{ gap: 8px; }}
  .stat {{ min-width: 80px; padding: 10px; }}
}}
</style>
</head>
<body>
<header>
  <h1>📋 Análise Crítica de Dissertação</h1>
  <p>Gerado em {_esc(ts)} · Framework de Análise Genérico v1.0</p>
</header>
<main>

<!-- ── Metadados ──────────────────────────────────────────────────────── -->
<div class="card">
  <h2>📌 Identificação da Obra</h2>
  <div class="meta-grid">
    <div class="meta-item"><label>Autor</label><span>{autor}</span></div>
    <div class="meta-item"><label>Instituição</label><span>{inst}</span></div>
    <div class="meta-item"><label>Programa</label><span>{programa}</span></div>
    <div class="meta-item"><label>Orientador</label><span>{orientador}</span></div>
    <div class="meta-item"><label>Área</label><span>{area}</span></div>
    <div class="meta-item"><label>Ano</label><span>{ano}</span></div>
  </div>
  <div style="background:#f8fafc;border-radius:8px;padding:14px;border:1px solid var(--borda)">
    <b style="font-size:.92rem;">{titulo_d}</b>
  </div>

  {"<div class='problema-pesquisa' style='margin-top:12px'><b>Problema de pesquisa:</b><br>" + problema + "</div>" if meta.get("problema_pesquisa") else ""}
  {"<div class='problema-pesquisa' style='background:#f0fdf4;border-color:#bbf7d0;margin-top:8px'><b>Hipótese central:</b><br>" + hipotese + "</div>" if meta.get("hipotese") else ""}

  {('<div style="margin-top:12px"><b style="font-size:.8rem;color:var(--cinza)">OBJETIVOS ESPECÍFICOS</b><br>'
    + "".join(f'<span class="chip">{_esc(o)}</span>' for o in objetivos)
    + '</div>') if objetivos else ""}
</div>

<!-- ── Painel de resultados ────────────────────────────────────────────── -->
<div class="card">
  <h2>📊 Resumo da Análise</h2>
  <div class="stats-row">
    <div class="stat s-verde"><strong>{n_aprovado}</strong><span>Aprovados</span></div>
    <div class="stat s-amarelo"><strong>{n_ressalvas}</strong><span>Com ressalvas</span></div>
    <div class="stat s-vermelho"><strong>{n_revisao}</strong><span>Requer revisão</span></div>
    <div class="stat s-azul"><strong>{total_pares}</strong><span>Args. apoio</span></div>
    <div class="stat" style="background:#f8fafc;color:var(--cinza)"><strong>{len(caps)}</strong><span>Capítulos</span></div>
  </div>
</div>

<!-- ── Capítulos ──────────────────────────────────────────────────────── -->
<h2>📖 Análise por Capítulo</h2>
<p style="font-size:.84rem;color:var(--cinza);margin-bottom:16px;">Clique em cada capítulo para expandir os detalhes.</p>
"""

    # ── Cards por capítulo ───────────────────────────────────────────────
    for i, cap in enumerate(caps):
        titulo_cap = _esc(cap.get("titulo", f"Capítulo {i+1}"))
        analise = cap.get("analise") or {}
        apoio   = cap.get("apoio") or {}

        pont = analise.get("pontuacao_geral", "")
        cor_pont = _COR_PONT.get(pont, "#64748b")
        badge_cls = {"APROVADO": "badge-verde", "APROVADO_COM_RESSALVAS": "badge-amarelo",
                     "REQUER_REVISAO": "badge-vermelho"}.get(pont, "badge-cinza")
        icon_pont = _ICON_PONT.get(pont, "•")

        resumo_exec = _esc(analise.get("resumo_executivo") or "")

        # Problemas estruturais
        probs = analise.get("problemas_estruturais") or []
        probs_html = ""
        for prob in probs:
            grav = str(prob.get("gravidade", "")).upper()
            cls_grav = {"MEDIA": "media", "BAIXA": "baixa"}.get(grav, "")
            loc = _esc(prob.get("localizacao") or "")
            probs_html += f"""<div class="prob-item {cls_grav}">
              <b>[{_esc(grav)}]</b> {_esc(prob.get("descricao", ""))}
              {f'<div class="prob-loc">📍 {loc}</div>' if loc else ""}
            </div>"""

        # Adequação teórica
        teo = analise.get("adequacao_teorica") or {}
        teo_av = _esc(teo.get("avaliacao", "—"))
        teo_pos = teo.get("pontos_positivos") or []
        teo_lac = teo.get("lacunas") or []

        # Integração empírica
        emp = analise.get("integracao_empirica") or {}
        emp_av = _esc(emp.get("avaliacao", "—"))
        emp_obs = _esc(emp.get("observacoes") or "")

        # ABNT
        abnt = analise.get("conformidade_abnt") or {}
        abnt_probs = abnt.get("problemas_encontrados") or []
        abnt_status = _esc(abnt.get("status") or "")

        # Estilo
        estilo = analise.get("estilo_escrita") or {}
        est_autor = estilo.get("paragrafos_que_iniciam_com_autor", 0)
        est_trav  = estilo.get("uso_de_travessao", False)
        est_2pts  = estilo.get("uso_excessivo_dois_pontos", False)
        est_trans = _esc(estilo.get("transicoes_entre_secoes", "—"))
        est_listas= estilo.get("listas_desnecessarias", False)
        est_obs   = _esc(estilo.get("observacoes") or "")

        def _check(v) -> str:
            if isinstance(v, bool):
                return '<span class="cross">✗</span>' if v else '<span class="check">✓</span>'
            return _esc(str(v))

        # Recomendações
        recs = analise.get("recomendacoes_priorizadas") or []

        # Argumentos de apoio
        pares = apoio.get("pares_encontrados") or []
        sintese_apoio = _esc(apoio.get("sintese") or "")
        cobertura = _esc(apoio.get("cobertura") or "—")

        pares_html = ""
        for par in pares:
            rel = str(par.get("relacao", "")).upper()
            rel_cls = rel.lower()
            rel_cor = _COR_REL.get(rel, "#64748b")
            pares_html += f"""<div class="par-item {rel_cls}">
              <span class="rel-tag" style="background:{rel_cor}22;color:{rel_cor}">{_esc(rel)}</span>
              <div class="argumento"><b>Dissertação:</b> {_esc(par.get("argumento_dissertacao",""))}</div>
              <div class="trecho">{_esc(par.get("trecho_apoio",""))}</div>
              <div class="fonte">📄 {_esc(par.get("arquivo_fonte",""))} — {_esc(par.get("explicacao",""))}</div>
            </div>"""

        n_chars_fmt = f"{cap.get('n_chars', 0):,}"

        html += f"""
<div class="cap-card">
  <div class="cap-header" onclick="toggleCap(this)">
    <div>
      <div class="cap-title">{icon_pont} {titulo_cap}</div>
      <div style="font-size:.75rem;color:var(--cinza);margin-top:3px">{n_chars_fmt} caracteres</div>
    </div>
    <div style="display:flex;gap:8px;align-items:center">
      <span class="badge {badge_cls}">{_esc(pont or "—")}</span>
      <span style="color:var(--cinza);font-size:1.1rem">▸</span>
    </div>
  </div>
  <div class="cap-body">

    {"<div class='resumo-exec'>" + resumo_exec + "</div>" if resumo_exec else ""}

    {f'''<div class="bloco">
      <div class="bloco-titulo">⚠ Problemas Estruturais</div>
      {probs_html}
    </div>''' if probs_html else '<div class="bloco"><div class="bloco-titulo">✅ Sem problemas estruturais detectados</div></div>'}

    <div class="bloco" style="display:grid;grid-template-columns:1fr 1fr;gap:12px">
      <div>
        <div class="bloco-titulo">📚 Adequação Teórica: {teo_av}</div>
        {"<ul class='lista'>" + "".join(f"<li>✓ {_esc(p)}</li>" for p in teo_pos) + "</ul>" if teo_pos else ""}
        {"<div style='margin-top:6px'><b style='font-size:.78rem;color:var(--vermelho)'>Lacunas:</b><ul class='lista'>" + "".join(f"<li>{_esc(l)}</li>" for l in teo_lac) + "</ul></div>" if teo_lac else ""}
      </div>
      <div>
        <div class="bloco-titulo">🔗 Integração Empírica: {emp_av}</div>
        <p style="font-size:.85rem">{emp_obs}</p>
      </div>
    </div>

    <div class="bloco">
      <div class="bloco-titulo">✍ Estilo & Escrita</div>
      <table class="abnt">
        <tr><td>Parágrafos iniciando com autor</td><td>{_check(est_autor > 0)} {est_autor}</td></tr>
        <tr><td>Uso de travessão em prosa</td><td>{_check(est_trav)}</td></tr>
        <tr><td>Dois-pontos excessivos</td><td>{_check(est_2pts)}</td></tr>
        <tr><td>Transições entre seções</td><td>{est_trans}</td></tr>
        <tr><td>Listas desnecessárias</td><td>{_check(est_listas)}</td></tr>
      </table>
      {"<p style='font-size:.82rem;color:var(--cinza);margin-top:8px'>" + est_obs + "</p>" if est_obs else ""}
    </div>

    <div class="bloco">
      <div class="bloco-titulo">📏 Conformidade ABNT: {abnt_status}</div>
      {"<ul class='lista'>" + "".join(f"<li class='cross'>✗ {_esc(p)}</li>" for p in abnt_probs) + "</ul>" if abnt_probs else "<p style='font-size:.85rem;color:var(--verde)'>✓ Nenhum problema ABNT identificado</p>"}
    </div>

    {f'''<div class="bloco">
      <div class="bloco-titulo">📋 Recomendações Priorizadas</div>
      <ul class="lista">{"".join(f"<li>{_esc(r)}</li>" for r in recs)}</ul>
    </div>''' if recs else ""}

    {f'''<div class="bloco">
      <div class="bloco-titulo">🔍 Argumentos do Material de Apoio (cobertura: {cobertura})</div>
      {pares_html}
      {"<div class='sintese'>" + sintese_apoio + "</div>" if sintese_apoio else ""}
    </div>''' if (pares_html or sintese_apoio) else ""}

  </div>
</div>
"""

    html += f"""
</main>
<footer>Análise gerada automaticamente em {_esc(ts)} · Framework genérico de dissertações</footer>
<script>
function toggleCap(header) {{
  const body = header.nextElementSibling;
  const arrow = header.querySelector('span:last-child');
  body.classList.toggle('aberto');
  arrow.textContent = body.classList.contains('aberto') ? '▾' : '▸';
}}
</script>
</body>
</html>"""

    (rel_dir / "relatorio_analise.html").write_text(html, encoding="utf-8")

    # ── TXT resumido ─────────────────────────────────────────────────────
    sep = "=" * 70
    L = [sep, "  ANÁLISE CRÍTICA DE DISSERTAÇÃO", f"  Gerado em: {ts}", sep, ""]
    L += [
        f"Autor       : {meta.get('autor') or '—'}",
        f"Título      : {meta.get('titulo') or '—'}",
        f"Instituição : {meta.get('instituicao') or '—'}",
        f"Orientador  : {meta.get('orientador') or '—'}",
        f"Programa    : {meta.get('programa') or '—'}",
        f"Ano         : {meta.get('ano') or '—'}",
        "",
        f"RESUMO: {n_aprovado} aprovados | {n_ressalvas} com ressalvas | {n_revisao} requerem revisão",
        f"Material de apoio: {resultado.get('n_arquivos_apoio', 0)} arquivo(s) | {total_pares} par(es) encontrado(s)",
        "",
    ]
    for cap in caps:
        an = cap.get("analise") or {}
        L += [
            f"{'─' * 60}",
            f"CAPÍTULO: {cap.get('titulo', '')}",
            f"Pontuação: {an.get('pontuacao_geral', '—')}",
        ]
        resumo_exec = an.get("resumo_executivo")
        if resumo_exec:
            L.append(f"Resumo: {resumo_exec}")
        prbs = an.get("problemas_estruturais") or []
        if prbs:
            L.append("Problemas:")
            for p in prbs:
                L.append(f"  [{p.get('gravidade', '?')}] {p.get('descricao', '')}")
        recs = an.get("recomendacoes_priorizadas") or []
        if recs:
            L.append("Recomendações:")
            for r in recs:
                L.append(f"  {r}")
        pares = (cap.get("apoio") or {}).get("pares_encontrados") or []
        if pares:
            L.append(f"Argumentos de apoio encontrados: {len(pares)}")
        L.append("")

    (rel_dir / "relatorio_analise.txt").write_text("\n".join(L), encoding="utf-8")
