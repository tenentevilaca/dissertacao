"""
Analisa criticamente cada capítulo conforme o template genérico,
e produz uma versão reescrita do capítulo com as melhorias aplicadas.
"""

import re
from typing import Generator

import anthropic

from config import ANTHROPIC_API_KEY, CLAUDE_MODEL, MAX_CHARS_CAPITULO

_PROMPT_ANALISE = """Você é um orientador acadêmico sênior especializado em revisão de dissertações de mestrado brasileiras.

## CONTEXTO DA DISSERTAÇÃO
- Autor: {autor}
- Título: {titulo}
- Programa: {programa} — {instituicao}
- Hipótese central: {hipotese}
- Metodologia: {metodologia}

## TAREFA: ANÁLISE CRÍTICA DO CAPÍTULO

Analise rigorosamente o capítulo abaixo, verificando:

### 1. Problemas Estruturais
- Repetições de conceitos, autores ou exemplos (cite trechos específicos)
- Falhas na ordem lógica (deve ser: geral → específico; teoria → dados; causa → efeito)
- Lacunas temáticas em relação ao que o capítulo deveria cobrir

### 2. Adequação Teórica
- Achados empíricos sem diálogo com referencial teórico
- Teoria citada mas não mobilizada (apenas mencionada, não integrada)
- Convergências ou tensões entre autores não exploradas

### 3. Conformidade ABNT NBR 10520:2023
- Citações que iniciam parágrafos (devem ser argumentos, não autores)
- Travessões em prosa corrida (remover, substituir por vírgulas/parênteses)
- Excesso de dois-pontos
- Falta de transições explícitas entre seções

### 4. Integração Teórico-Empírica
- Dados apresentados sem interpretação teórica
- Hipótese central não sendo validada, refutada ou matizada
- Cruzamentos e padrões não explorados

### 5. Coesão e Coerência
- Falta de transição para o próximo capítulo
- Inconsistências internas
- Parágrafos sem tópico frasal claro

## FORMATO DE SAÍDA

Retorne SOMENTE o JSON abaixo, sem texto adicional:

{{
  "capitulo": "{titulo_cap}",
  "numero": {numero_cap},
  "diagnostico": {{
    "pontuacao_geral": <0-10>,
    "problemas_estruturais": ["problema 1 com localização", "problema 2..."],
    "problemas_teoricos": ["..."],
    "problemas_abnt": ["..."],
    "problemas_integracao": ["..."],
    "problemas_coesao": ["..."],
    "pontos_positivos": ["..."],
    "lacunas": ["conceito/norma/dado ausente 1", "..."]
  }},
  "resumo_diagnostico": "parágrafo de 3-5 linhas resumindo os principais achados da análise",
  "recomendacoes": ["ação específica 1", "ação específica 2", "..."]
}}

## CAPÍTULO A ANALISAR:

{texto_capitulo}
"""

_PROMPT_REESCRITA = """Você é um redator acadêmico especializado em dissertações de mestrado brasileiras.

## CONTEXTO DA DISSERTAÇÃO
- Autor: {autor}
- Título: {titulo}
- Programa: {programa} — {instituicao}
- Hipótese central: {hipotese}
- Metodologia: {metodologia}

## DIAGNÓSTICO DO CAPÍTULO (já realizado)
{diagnostico_json}

## TRECHOS DE APOIO DO MATERIAL DE REFERÊNCIA
Os seguintes trechos foram encontrados nos materiais de referência e PODEM ser usados para enriquecer argumentos onde apropriado:
{trechos_apoio}

## TAREFA: REESCREVA O CAPÍTULO COMPLETO

Reescreva o capítulo integralmente, aplicando TODAS as correções identificadas no diagnóstico:

### Regras obrigatórias:
1. **Argumentos abrem parágrafos** — nunca comece com nome de autor
   ❌ "Silva (2020) argumenta que..."
   ✅ "A integração informacional é condição de confiança. Silva (2020) demonstra que..."

2. **Sem travessões em prosa** — substitua por vírgulas, parênteses ou ponto

3. **Dois-pontos reduzidos** — máximo 1-2 por página; prefira ponto ou vírgula

4. **Estrutura geral → específico** em cada seção

5. **Transições explícitas** entre seções e ao final do capítulo para o próximo

6. **Integração teórico-empírica real**: para cada dado apresentado, há uma frase explicando o que ele significa teoricamente

7. **Sem listas desnecessárias** — prosa em parágrafos

8. **ABNT NBR 10520:2023**: sobrenomes com primeira letra maiúscula apenas, et al. sem espaço antes do ponto

9. **Preserve todos os dados, citações e referências originais** — apenas reorganize e melhore a argumentação

10. **Se trechos de apoio forem relevantes**, integre-os naturalmente no texto com citação adequada

### IMPORTANTE:
- Mantenha o mesmo número e estrutura de seções do original
- Não invente dados, estatísticas ou referências
- Preserve fidelidade ao argumento original — melhore a forma, não mude a substância
- O texto reescrito deve ser académicamente robusto e fluente

## CAPÍTULO ORIGINAL:

{texto_capitulo}

---

Retorne APENAS o texto reescrito do capítulo completo, em prosa acadêmica, começando pelo título do capítulo.
Não inclua JSON, comentários, nem explicações — apenas o texto do capítulo reescrito.
"""


def analisar_capitulo(
    capitulo: dict,
    metadados: dict,
    trechos_apoio: list[dict],
    *,
    on_progress: callable = None,
) -> dict:
    """
    Executa análise crítica + reescrita de um capítulo.
    Retorna dict com {diagnostico, texto_reescrito}.
    """
    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

    texto = capitulo["texto"][:MAX_CHARS_CAPITULO]
    autor = metadados.get("autor") or "Não identificado"
    titulo = metadados.get("titulo") or "Dissertação"
    programa = metadados.get("programa") or ""
    instituicao = metadados.get("instituicao") or ""
    hipotese = metadados.get("hipotese") or "Não identificada"
    metodologia = metadados.get("metodologia") or "Não especificada"

    # ── FASE 1: Análise crítica ──────────────────────────────────────────
    if on_progress:
        on_progress(f"  → Analisando criticamente '{capitulo['titulo']}'...")

    prompt_analise = _PROMPT_ANALISE.format(
        autor=autor,
        titulo=titulo,
        programa=programa,
        instituicao=instituicao,
        hipotese=hipotese,
        metodologia=metodologia,
        titulo_cap=capitulo["titulo"],
        numero_cap=capitulo["numero"],
        texto_capitulo=texto,
    )

    resp_analise = client.messages.create(
        model=CLAUDE_MODEL,
        max_tokens=2048,
        messages=[{"role": "user", "content": prompt_analise}],
    )

    raw_diag = resp_analise.content[0].text.strip()
    raw_diag = re.sub(r"^```(?:json)?\s*", "", raw_diag)
    raw_diag = re.sub(r"\s*```$", "", raw_diag)

    import json
    try:
        diagnostico = json.loads(raw_diag)
    except json.JSONDecodeError:
        diagnostico = {"raw": raw_diag, "erro": "JSON inválido na análise"}

    # ── FASE 2: Reescrita ────────────────────────────────────────────────
    if on_progress:
        on_progress(f"  → Reescrevendo '{capitulo['titulo']}'...")

    apoio_texto = _formatar_trechos_apoio(trechos_apoio)

    prompt_reescrita = _PROMPT_REESCRITA.format(
        autor=autor,
        titulo=titulo,
        programa=programa,
        instituicao=instituicao,
        hipotese=hipotese,
        metodologia=metodologia,
        diagnostico_json=json.dumps(diagnostico, ensure_ascii=False, indent=2),
        trechos_apoio=apoio_texto,
        texto_capitulo=texto,
    )

    resp_reescrita = client.messages.create(
        model=CLAUDE_MODEL,
        max_tokens=8192,
        messages=[{"role": "user", "content": prompt_reescrita}],
    )

    texto_reescrito = resp_reescrita.content[0].text.strip()

    return {
        "numero": capitulo["numero"],
        "titulo": capitulo["titulo"],
        "texto_original": capitulo["texto"],
        "diagnostico": diagnostico,
        "texto_reescrito": texto_reescrito,
    }


def _formatar_trechos_apoio(trechos: list[dict]) -> str:
    if not trechos:
        return "(nenhum trecho de apoio encontrado para este capítulo)"
    linhas = []
    for i, t in enumerate(trechos, 1):
        arquivo = t.get("arquivo", "desconhecido")
        trecho = t.get("trecho", "")[:500]
        relevancia = t.get("relevancia", "")
        linhas.append(
            f"[APOIO {i}] Arquivo: {arquivo}\n"
            f"Relevância: {relevancia}\n"
            f"Trecho: {trecho}\n"
        )
    return "\n".join(linhas)
