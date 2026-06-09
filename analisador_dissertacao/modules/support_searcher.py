"""
Varre o material de apoio (pasta zipada) em busca de trechos que validam,
aprofundam ou contradizem os argumentos de cada capítulo.
Usa IA para avaliar relevância semântica.
"""

import json
import re
import zipfile
from pathlib import Path

import anthropic

from config import ANTHROPIC_API_KEY, CLAUDE_MODEL, MAX_CHARS_APOIO, TOP_K_APOIO
from modules.docx_utils import extrair_texto_arquivo

_EXTENSOES = {".pdf", ".docx", ".txt", ".md"}


def extrair_materiais_apoio(caminho_zip: Path, pasta_tmp: Path) -> list[dict]:
    """
    Descompacta o zip e extrai texto de cada arquivo suportado.
    Retorna lista de {arquivo, texto}.
    """
    pasta_tmp.mkdir(parents=True, exist_ok=True)

    with zipfile.ZipFile(str(caminho_zip), "r") as z:
        z.extractall(str(pasta_tmp))

    materiais = []
    for arquivo in sorted(pasta_tmp.rglob("*")):
        # ignora arquivos temporários do Office
        if arquivo.name.startswith("~$"):
            continue
        if arquivo.suffix.lower() not in _EXTENSOES:
            continue
        texto = extrair_texto_arquivo(arquivo)
        if texto.strip():
            materiais.append({"arquivo": arquivo.name, "texto": texto})

    return materiais


def _chunkar(texto: str, tamanho: int = MAX_CHARS_APOIO) -> list[str]:
    """Divide texto em chunks de até `tamanho` caracteres, nos limites de parágrafo."""
    paragrafos = texto.split("\n")
    chunks = []
    atual = []
    tamanho_atual = 0

    for p in paragrafos:
        if tamanho_atual + len(p) > tamanho and atual:
            chunks.append("\n".join(atual))
            atual = []
            tamanho_atual = 0
        atual.append(p)
        tamanho_atual += len(p) + 1

    if atual:
        chunks.append("\n".join(atual))

    return chunks


def buscar_apoio_para_capitulo(
    capitulo: dict,
    materiais: list[dict],
    *,
    on_progress: callable = None,
) -> list[dict]:
    """
    Para um capítulo, busca nos materiais de apoio os trechos mais relevantes.
    Retorna lista de {arquivo, trecho, relevancia, tipo}.
    """
    if not materiais:
        return []

    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

    # Resumo do capítulo para guiar a busca
    texto_cap = capitulo["texto"][:3000]

    # Coleta candidatos: todos os chunks de todos os materiais
    candidatos = []
    for mat in materiais:
        for chunk in _chunkar(mat["texto"]):
            if len(chunk.strip()) > 100:
                candidatos.append({"arquivo": mat["arquivo"], "trecho": chunk})

    if not candidatos:
        return []

    if on_progress:
        on_progress(
            f"  → Buscando material de apoio para '{capitulo['titulo']}' "
            f"({len(candidatos)} trechos candidatos)..."
        )

    # Envia candidatos em lotes para a IA avaliar relevância
    resultado: list[dict] = []
    lote_size = 8

    for i in range(0, len(candidatos), lote_size):
        lote = candidatos[i : i + lote_size]
        resultado += _avaliar_lote(client, texto_cap, lote, capitulo["titulo"])
        if len(resultado) >= TOP_K_APOIO * 3:
            break

    # Ordena por pontuação e retorna top K
    resultado.sort(key=lambda x: x.get("pontuacao", 0), reverse=True)
    return resultado[:TOP_K_APOIO]


def _avaliar_lote(
    client: anthropic.Anthropic,
    resumo_capitulo: str,
    lote: list[dict],
    titulo_cap: str,
) -> list[dict]:
    trechos_str = "\n\n".join(
        f"[TRECHO {i+1}] Arquivo: {c['arquivo']}\n{c['trecho'][:MAX_CHARS_APOIO]}"
        for i, c in enumerate(lote)
    )

    prompt = f"""Você avalia relevância de trechos de material bibliográfico para um capítulo de dissertação.

CAPÍTULO: "{titulo_cap}"
RESUMO/INÍCIO DO CAPÍTULO:
{resumo_capitulo}

TRECHOS DO MATERIAL DE APOIO:
{trechos_str}

Para cada trecho, avalie:
- pontuacao: 0-10 (10 = altamente relevante e útil para enriquecer o capítulo)
- tipo: "valida" | "aprofunda" | "contradiz" | "irrelevante"
- relevancia: uma frase explicando por que o trecho é útil (ou não)

Retorne SOMENTE JSON:
[
  {{
    "indice": 1,
    "arquivo": "nome do arquivo",
    "pontuacao": 8,
    "tipo": "valida",
    "relevancia": "Confirma o argumento sobre X com dados empíricos de Y"
  }},
  ...
]
"""

    try:
        resp = client.messages.create(
            model=CLAUDE_MODEL,
            max_tokens=1024,
            messages=[{"role": "user", "content": prompt}],
        )
        raw = resp.content[0].text.strip()
        raw = re.sub(r"^```(?:json)?\s*", "", raw)
        raw = re.sub(r"\s*```$", "", raw)
        avaliacoes = json.loads(raw)
    except Exception:
        return []

    resultado = []
    for av in avaliacoes:
        idx = av.get("indice", 1) - 1
        if 0 <= idx < len(lote) and av.get("pontuacao", 0) >= 5:
            entrada = dict(lote[idx])
            entrada.update(
                {
                    "pontuacao": av.get("pontuacao", 0),
                    "tipo": av.get("tipo", ""),
                    "relevancia": av.get("relevancia", ""),
                }
            )
            resultado.append(entrada)

    return resultado
