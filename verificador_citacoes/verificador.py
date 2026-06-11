"""
Verificador de citações — módulo principal.

Fluxo em 3 etapas:
  1. Lê a dissertação e extrai texto
  2. Lê todos os documentos de apoio
  3. Identifica citações, cruza com referências e verifica com Claude
"""

import json
import re
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Optional
from xml.etree import ElementTree as ET


# ═══════════════════════════════════════════════════════════════════════
# ETAPA 1 — LEITURA DOS ARQUIVOS
# ═══════════════════════════════════════════════════════════════════════

_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _xml_para_paragrafos(xml_bytes: bytes) -> list[tuple[str, str]]:
    """Extrai (texto, estilo_do_paragrafo) de cada <w:p> de um XML do Word."""
    try:
        root = ET.fromstring(xml_bytes)
        paragrafos = []
        for para in root.iter(f"{{{_NS}}}p"):
            tokens = [t.text for t in para.iter(f"{{{_NS}}}t") if t.text]
            texto = "".join(tokens)
            estilo = ""
            ppr = para.find(f"{{{_NS}}}pPr")
            if ppr is not None:
                pstyle = ppr.find(f"{{{_NS}}}pStyle")
                if pstyle is not None:
                    estilo = pstyle.get(f"{{{_NS}}}val", "") or ""
            paragrafos.append((texto, estilo))
        return paragrafos
    except Exception:
        return []


def ler_docx_paragrafos(caminho: str | Path) -> list[tuple[str, str]]:
    """Lê o corpo de um .docx e retorna lista de (texto, estilo) por parágrafo."""
    try:
        with zipfile.ZipFile(str(caminho)) as z:
            if "word/document.xml" in z.namelist():
                return _xml_para_paragrafos(z.read("word/document.xml"))
    except Exception:
        pass
    return []


def _xml_para_texto(xml_bytes: bytes) -> str:
    """Extrai texto parágrafo a parágrafo de um XML do Word."""
    try:
        root = ET.fromstring(xml_bytes)
        linhas = []
        for para in root.iter(f"{{{_NS}}}p"):
            tokens = [t.text for t in para.iter(f"{{{_NS}}}t") if t.text]
            linhas.append("".join(tokens))
        return "\n".join(linhas)
    except Exception:
        return ""


def ler_docx(caminho: str | Path) -> str:
    """Lê um .docx e retorna o texto completo (corpo + notas de rodapé)."""
    partes: list[str] = []
    try:
        with zipfile.ZipFile(str(caminho)) as z:
            nomes = z.namelist()
            if "word/document.xml" in nomes:
                partes.append(_xml_para_texto(z.read("word/document.xml")))
            if "word/footnotes.xml" in nomes:
                partes.append("--- NOTAS DE RODAPÉ ---")
                partes.append(_xml_para_texto(z.read("word/footnotes.xml")))
            if "word/endnotes.xml" in nomes:
                partes.append(_xml_para_texto(z.read("word/endnotes.xml")))
    except zipfile.BadZipFile:
        try:
            import docx as _docx
            doc = _docx.Document(str(caminho))
            partes.append("\n".join(p.text for p in doc.paragraphs))
        except Exception as e:
            return f"[ERRO ao ler .docx: {e}]"
    except Exception as e:
        return f"[ERRO ao abrir arquivo: {e}]"

    texto = "\n".join(partes)
    texto = re.sub(r"\n{3,}", "\n\n", texto)
    return texto.strip()


def ler_pdf(caminho: str | Path) -> str:
    """Lê um PDF e retorna o texto de todas as páginas."""
    try:
        import fitz
        doc = fitz.open(str(caminho))
        return "\n".join(page.get_text() for page in doc)
    except Exception as e:
        return f"[ERRO ao ler PDF: {e}]"


def ler_arquivo(caminho: Path) -> str:
    """Lê qualquer arquivo suportado e retorna o texto."""
    ext = caminho.suffix.lower()
    if ext == ".pdf":
        return ler_pdf(caminho)
    elif ext in (".docx", ".doc"):
        return ler_docx(caminho)
    elif ext == ".txt":
        try:
            return caminho.read_text(encoding="utf-8", errors="replace")
        except Exception as e:
            return f"[ERRO ao ler .txt: {e}]"
    return ""


# ═══════════════════════════════════════════════════════════════════════
# ETAPA 2 — IDENTIFICAR CITAÇÕES E REFERÊNCIAS
# ═══════════════════════════════════════════════════════════════════════

@dataclass
class Citacao:
    texto: str
    autores: list[str]
    ano: str
    pagina: Optional[str]
    contexto: str
    paragrafo_idx: int


@dataclass
class Referencia:
    texto: str
    sobrenome: str          # primeiro autor
    sobrenomes: list[str]   # TODOS os autores (para cruzamento amplo)
    ano: str
    titulo: str


def _normalizar_autores(raw: str) -> list[str]:
    raw = re.sub(r",?\s*et\s+al\.?", "", raw, flags=re.IGNORECASE)
    partes = re.split(r"(?:\s*[;]\s*|\s+e\s+)", raw, flags=re.IGNORECASE)
    result = []
    for p in partes:
        tokens = p.strip().split()
        if tokens:
            result.append(tokens[0].upper())
    return result or ["DESCONHECIDO"]


def _contexto(paragrafos: list[str], idx: int, janela: int = 300) -> str:
    antes = " ".join(paragrafos[max(0, idx - 2):idx])
    depois = " ".join(paragrafos[idx + 1:idx + 3])
    return (antes[-janela:] + " " + paragrafos[idx] + " " + depois[:janela]).strip()


# Padrões ABNT
# Sobrenome: começa com MAIÚSCULA obrigatória (sem IGNORECASE no grupo inicial)
_SOBR = r"[A-ZÁÉÍÓÚÂÊÎÔÛÃÕÀÇ][A-ZÁÉÍÓÚÂÊÎÔÛÃÕÀÇa-záéíóúâêîôûãõàç\-\.']{1,}"
_ANO  = r"\d{4}[a-z]?"
_PAG  = r"(?:[,;:\s]+(?:p\.?p?|pág\.?|f\.?)\s*([\d\-–]+))?"

# Lista de autores separados por ponto-e-vírgula (sem espaço como separador)
# Captura o bloco de autores inteiro no grupo 1
_AUTORES_CAP = (
    r"("
    + _SOBR                              # primeiro sobrenome
    + r"(?:\s+" + _SOBR + r")?"         # segundo token opcional (ex: FONT DO)
    + r"(?:,?\s*et\s+al\.?)?"           # et al.
    + r"(?:\s*;\s*"                      # demais autores separados por ;
    + _SOBR
    + r"(?:\s+" + _SOBR + r")?"
    + r"(?:,?\s*et\s+al\.?)?)*"
    + r")"
)

# (AUTOR, ANO) ou (AUTOR, ANO, p. N)
_RE_PAREN = re.compile(
    r"\(" + _AUTORES_CAP + r"[,;]\s*(" + _ANO + r")" + _PAG + r"\)"
)

# AUTOR (ANO) — narrativo — NÃO usa IGNORECASE para evitar capturar palavras minúsculas
_RE_NARR = re.compile(
    _AUTORES_CAP + r"\s+\((" + _ANO + r")" + _PAG + r"\)"
)

# Início da seção de referências
_RE_INICIO_REFS = re.compile(
    r"^\s*REFER[EÊ]NCIAS\s*(BIBLIOGR[ÁA]FICAS?)?\s*$",
    re.IGNORECASE,
)

# Entrada de referência ABNT
_RE_ENTRADA_REF = re.compile(
    r"^(" + _SOBR + r")(?:,\s*.+?)?[.;]\s*(.+?)[.;]\s*.*?(\d{4})",
    re.DOTALL,
)


# Palavras que NÃO são sobrenomes em citações narrativas
_NAO_SOBRENOME = {
    "PARA", "SEGUNDO", "CONFORME", "SOBRE", "COMO", "POR", "COM",
    "PELO", "PELA", "PELOS", "PELAS", "EM", "DE", "DA", "DO",
    "APUD", "APUD", "VER", "CF", "VIDE", "ASSIM",
}


def _autores_paren(raw: str) -> list[str]:
    """Para (AUTOR, ANO): usa o PRIMEIRO token de cada grupo."""
    raw = re.sub(r",?\s*et\s+al\.?", "", raw, flags=re.IGNORECASE)
    partes = re.split(r"\s*[;]\s*", raw)
    result = []
    for p in partes:
        tokens = p.strip().split()
        if tokens:
            result.append(tokens[0].upper())
    return result or ["?"]


def _autores_narr(raw: str) -> list[str]:
    """Para AUTOR (ANO): usa o ÚLTIMO token não-preposicional (= sobrenome)."""
    raw = re.sub(r",?\s*et\s+al\.?", "", raw, flags=re.IGNORECASE)
    raw = raw.strip().rstrip(".")
    partes = re.split(r"\s*[;,]\s*", raw)
    result = []
    for p in partes:
        tokens = [t.strip(".,") for t in p.strip().split()
                  if t.strip(".,").upper() not in _NAO_SOBRENOME]
        if tokens:
            result.append(tokens[-1].upper())
        elif p.strip().split():
            result.append(p.strip().split()[-1].strip(".,").upper())
    return result or ["?"]


def _grupos_citacao(m, narrativo: bool = False) -> tuple[list[str], str, Optional[str]]:
    """Extrai (autores, ano, pagina) de um match de _RE_PAREN ou _RE_NARR."""
    raw = m.group(1) or ""
    autores = _autores_narr(raw) if narrativo else _autores_paren(raw)
    ano = m.group(2) or "?"
    pag = m.group(3) if m.lastindex and m.lastindex >= 3 else None
    return autores, ano, pag


def encontrar_citacoes(texto: str) -> tuple[list[Citacao], int]:
    """
    Retorna (lista_de_citacoes, índice_início_referências).
    """
    paragrafos = texto.split("\n")

    # Localiza seção REFERÊNCIAS
    idx_refs = -1
    for i, p in enumerate(paragrafos):
        if _RE_INICIO_REFS.match(p.strip()):
            idx_refs = i
            break

    limite = idx_refs if idx_refs > 0 else len(paragrafos)
    citacoes: list[Citacao] = []
    vistos: set[tuple] = set()

    for idx, para in enumerate(paragrafos[:limite]):
        ctx = _contexto(paragrafos[:limite], idx)

        for m in _RE_PAREN.finditer(para):
            autores, ano, pag = _grupos_citacao(m, narrativo=False)
            chave = (tuple(autores), ano)
            if chave not in vistos:
                vistos.add(chave)
                citacoes.append(Citacao(m.group(0), autores, ano, pag, ctx, idx))

        for m in _RE_NARR.finditer(para):
            autores, ano, pag = _grupos_citacao(m, narrativo=True)
            chave = (tuple(autores), ano)
            if chave not in vistos:
                vistos.add(chave)
                citacoes.append(Citacao(m.group(0), autores, ano, pag, ctx, idx))

    return citacoes, idx_refs


def extrair_referencias(texto: str, idx_refs: int) -> list[Referencia]:
    """Extrai entradas da lista de referências."""
    if idx_refs < 0:
        return []

    paragrafos = texto.split("\n")
    referencias: list[Referencia] = []
    buffer: list[str] = []

    _RE_NOVA = re.compile(
        r"^[A-ZÁÉÍÓÚÂÊÎÔÛÃÕÀÇ][A-ZÁÉÍÓÚÂÊÎÔÛÃÕÀÇa-záéíóúâêîôûãõàç\'\-]+[,;]"
    )

    # Extrai sobrenomes de todos os autores de uma entrada de referência
    _RE_AUTORES_REF = re.compile(
        r"([A-ZÁÉÍÓÚÂÊÎÔÛÃÕÀÇ][A-ZÁÉÍÓÚÂÊÎÔÛÃÕÀÇa-záéíóúâêîôûãõàç\-\']+),"
    )
    # Ano plausível (1900–2030) — evita capturar números de volume/página
    _RE_ANO_PLAUSIVEL = re.compile(r"\b((?:19|20)\d{2})\b")

    def _sobrenomes_ref(t: str) -> list[str]:
        """Extrai sobrenomes de todos os autores antes do título."""
        # Pega todos os tokens SOBRENOME, que aparecem antes do ponto final do último autor
        # Exemplo: "CEPIK, Marco; BORBA, Pedro. Título..." → ["CEPIK", "BORBA"]
        resultado = []
        # Procura até o primeiro ". " ou "." que não seja seguido de letra maiúscula de nome próprio
        cabecalho = re.split(r"\.\s+[A-ZÁÉÍÓÚÂÊÎÔÛÃÕÀÇ][a-záéíóúâêîôûãõàç]", t)[0]
        for m in _RE_AUTORES_REF.finditer(cabecalho):
            sob = m.group(1).upper()
            if len(sob) > 1:
                resultado.append(sob)
        if not resultado:
            m_sob = re.match(r"^([A-ZÁÉÍÓÚÂÊÎÔÛÃÕÀÇ][A-Za-záéíóúâêîôûãõàç\-\']+)", t)
            if m_sob:
                resultado = [m_sob.group(1).upper()]
        return resultado or ["DESCONHECIDO"]

    def processar(buf: list[str]) -> Optional[Referencia]:
        t = " ".join(buf).strip()
        if not t:
            return None
        sobrenomes = _sobrenomes_ref(t)
        sobrenome = sobrenomes[0]
        # Ano plausível primeiro
        m_ano = _RE_ANO_PLAUSIVEL.search(t)
        ano = m_ano.group(1) if m_ano else "s.d."
        # Título: texto após o bloco de autores
        m_titulo = re.search(r"\.\s+([A-Z][^\.\n]{5,})", t)
        titulo = m_titulo.group(1)[:80] if m_titulo else t[:80]
        return Referencia(t, sobrenome, sobrenomes, ano, titulo)

    for linha in paragrafos[idx_refs + 1:]:
        linha = linha.strip()
        if not linha:
            if buffer:
                r = processar(buffer)
                if r:
                    referencias.append(r)
                buffer = []
            continue
        if _RE_NOVA.match(linha) and buffer:
            r = processar(buffer)
            if r:
                referencias.append(r)
            buffer = [linha]
        else:
            buffer.append(linha)

    if buffer:
        r = processar(buffer)
        if r:
            referencias.append(r)

    return referencias


# ═══════════════════════════════════════════════════════════════════════
# ETAPA 3 — VERIFICAÇÃO SEMÂNTICA COM CLAUDE
# ═══════════════════════════════════════════════════════════════════════

_PROMPT_VERIF = """Você é um verificador de integridade acadêmica especializado em dissertações.

TRECHO DA DISSERTAÇÃO que usa a citação {citacao}:
\"\"\"{contexto}\"\"\"

REFERÊNCIA CITADA: {referencia}
ARQUIVO FONTE: {arquivo}

TRECHO RELEVANTE DA OBRA CITADA (extraído automaticamente):
\"\"\"{fonte}\"\"\"

Analise com atenção:
1. O argumento apresentado na dissertação é coerente com o que a obra realmente afirma?
2. Há distorção de sentido, generalização indevida ou uso fora de contexto?
3. A ideia atribuída ao autor realmente consta na obra?
4. Se há número de página na citação, o conteúdo bate com o trecho encontrado?

Responda SOMENTE com JSON (sem markdown, sem explicação extra):
{{"veredicto": "CORRETO" | "INCORRETO" | "PARCIAL" | "SEM_FONTE", "justificativa": "uma frase explicando o veredicto", "trecho_fonte": "trecho exato da obra que confirma ou contradiz (até 200 chars)"}}
"""

_STOPWORDS_PT = {
    "de", "da", "do", "das", "dos", "e", "a", "o", "as", "os", "um", "uma",
    "que", "em", "para", "com", "por", "se", "na", "no", "nas", "nos",
    "ao", "aos", "à", "às", "pelo", "pela", "pelos", "pelas", "como",
    "mais", "mas", "ou", "sua", "seu", "suas", "seus", "este", "esta",
    "isso", "esse", "essa", "esses", "essas", "ser", "foi", "são", "está",
    "não", "também", "quando", "sobre", "entre", "mesmo", "ainda", "pode",
}


def _extrair_trecho_relevante(contexto_cit: str, texto_fonte: str, janela: int = 4000) -> str:
    """
    Localiza no texto fonte o trecho mais pertinente para a citação.
    Usa matching de palavras-chave extraídas do contexto da dissertação.
    """
    palavras = re.findall(r'[a-záéíóúâêîôûãõàç]{4,}', contexto_cit.lower())
    palavras = [p for p in palavras if p not in _STOPWORDS_PT]
    if not palavras:
        return texto_fonte[:janela]

    paragrafos = [p.strip() for p in texto_fonte.split('\n') if len(p.strip()) > 40]
    if not paragrafos:
        return texto_fonte[:janela]

    def pontuar(para: str) -> int:
        pl = para.lower()
        return sum(1 for p in palavras if p in pl)

    scores = sorted(range(len(paragrafos)), key=lambda i: -pontuar(paragrafos[i]))
    melhor = scores[0]
    # Inclui parágrafos vizinhos para dar contexto
    indices = sorted({max(0, melhor - 2), max(0, melhor - 1), melhor,
                      min(len(paragrafos) - 1, melhor + 1), min(len(paragrafos) - 1, melhor + 2)})
    trecho = '\n'.join(paragrafos[i] for i in indices)
    if len(trecho) < janela // 2:
        # Adiciona mais parágrafos ao redor se o trecho for pequeno
        extras = [scores[j] for j in range(1, min(4, len(scores)))]
        for e in sorted(extras):
            trecho += '\n' + paragrafos[e]
    return trecho[:janela]


def verificar_com_claude(
    citacao: Citacao,
    referencia: "Referencia",
    arquivo_fonte: str,
    texto_fonte: str,
    api_key: str,
    model: str = "claude-sonnet-4-6",
) -> dict:
    import anthropic
    client = anthropic.Anthropic(api_key=api_key)
    trecho = _extrair_trecho_relevante(citacao.contexto, texto_fonte)
    prompt = _PROMPT_VERIF.format(
        citacao=citacao.texto,
        contexto=citacao.contexto[:600],
        referencia=referencia.texto[:150],
        arquivo=arquivo_fonte,
        fonte=trecho,
    )
    try:
        resp = client.messages.create(
            model=model,
            max_tokens=512,
            messages=[{"role": "user", "content": prompt}],
        )
        raw = resp.content[0].text.strip()
        raw = re.sub(r"^```(?:json)?", "", raw).strip().strip("`").strip()
        try:
            return json.loads(raw)
        except json.JSONDecodeError:
            m = re.search(r"\{[^}]+\}", raw, re.DOTALL)
            if m:
                try:
                    return json.loads(m.group())
                except Exception:
                    pass
        return {"veredicto": "ERRO", "justificativa": raw[:300], "trecho_fonte": ""}
    except Exception as e:
        return {"veredicto": "ERRO", "justificativa": str(e)[:200], "trecho_fonte": ""}


_PROMPT_VERIF_GRUPO = """Você é um verificador de integridade acadêmica especializado em dissertações.

REFERÊNCIA CITADA: {referencia}
ARQUIVO FONTE: {arquivo}

TRECHO RELEVANTE DA OBRA CITADA (extraído automaticamente):
\"\"\"{fonte}\"\"\"

A seguir estão {n} trecho(s) da dissertação que citam essa obra. Para CADA trecho, avalie:
1. O argumento apresentado é coerente com o que a obra realmente afirma?
2. Há distorção de sentido, generalização indevida ou uso fora de contexto?
3. A ideia atribuída ao autor realmente consta na obra?

{trechos}

Responda SOMENTE com um array JSON (sem markdown, sem explicação extra), um item por trecho, na MESMA ORDEM em que foram apresentados:
[{{"indice": 1, "veredicto": "CORRETO" | "INCORRETO" | "PARCIAL" | "SEM_FONTE", "justificativa": "uma frase explicando o veredicto", "trecho_fonte": "trecho exato da obra que confirma ou contradiz (até 200 chars)"}}]
"""


def verificar_grupo_com_claude(
    citacoes: list["Citacao"],
    referencia: "Referencia",
    arquivo_fonte: str,
    texto_fonte: str,
    api_key: str,
    model: str = "claude-sonnet-4-6",
) -> list[dict]:
    """Verifica TODAS as citações de uma mesma obra em UMA única chamada à API."""
    import anthropic
    client = anthropic.Anthropic(api_key=api_key)

    contexto_combinado = "\n".join(c.contexto for c in citacoes)
    trecho = _extrair_trecho_relevante(contexto_combinado, texto_fonte, janela=6000)

    trechos_fmt = "\n".join(
        f'TRECHO {i}:\n  Citação: {c.texto}\n  Contexto: """{c.contexto[:600]}"""'
        for i, c in enumerate(citacoes, 1)
    )

    prompt = _PROMPT_VERIF_GRUPO.format(
        referencia=referencia.texto[:150],
        arquivo=arquivo_fonte,
        fonte=trecho,
        n=len(citacoes),
        trechos=trechos_fmt,
    )

    padrao = [{"veredicto": "ERRO", "justificativa": "sem resposta", "trecho_fonte": ""} for _ in citacoes]
    try:
        resp = client.messages.create(
            model=model,
            max_tokens=400 * len(citacoes) + 256,
            messages=[{"role": "user", "content": prompt}],
        )
        raw = resp.content[0].text.strip()
        raw = re.sub(r"^```(?:json)?", "", raw).strip().strip("`").strip()
        try:
            itens = json.loads(raw)
        except json.JSONDecodeError:
            m = re.search(r"\[.*\]", raw, re.DOTALL)
            itens = json.loads(m.group()) if m else None

        if not isinstance(itens, list):
            for d in padrao:
                d["justificativa"] = raw[:300]
            return padrao

        resultados = list(padrao)
        for item in itens:
            idx = item.get("indice")
            if isinstance(idx, int) and 1 <= idx <= len(resultados):
                resultados[idx - 1] = {
                    "veredicto": item.get("veredicto", "ERRO"),
                    "justificativa": item.get("justificativa", ""),
                    "trecho_fonte": item.get("trecho_fonte", ""),
                }
        return resultados
    except Exception as e:
        for d in padrao:
            d["justificativa"] = str(e)[:200]
        return padrao


def _buscar_fonte(sobrenome: str, ano: str, titulo: str, material: dict[str, str]) -> tuple[str, str]:
    """
    Tenta encontrar o arquivo de apoio correspondente à referência.
    Estratégia: nome do arquivo → (sobrenome + ano no texto) → (sobrenome no texto).
    """
    sob_lower = sobrenome.lower()
    titulo_palavras = [p for p in re.findall(r'[a-záéíóúâêîôûãõàç]{4,}', titulo.lower())
                       if p not in _STOPWORDS_PT][:5]

    candidatos: list[tuple[int, str, str]] = []  # (score, nome, texto)

    for nome, texto in material.items():
        score = 0
        texto_lower = texto.lower()
        nome_lower = nome.lower()

        if sob_lower in nome_lower:
            score += 10
        if ano in nome:
            score += 5
        if sob_lower in texto_lower:
            score += 3
        if ano in texto:
            score += 2
        # Bônus por palavras do título na primeiras páginas do arquivo
        cabecalho = texto_lower[:2000]
        for p in titulo_palavras:
            if p in cabecalho:
                score += 1

        if score > 0:
            candidatos.append((score, nome, texto))

    if not candidatos:
        return "", ""

    candidatos.sort(key=lambda x: -x[0])
    _, nome_melhor, texto_melhor = candidatos[0]
    return nome_melhor, texto_melhor


# ═══════════════════════════════════════════════════════════════════════
# ORQUESTRADOR PRINCIPAL
# ═══════════════════════════════════════════════════════════════════════

def analisar(
    diss_path: str,
    refs_dir: Optional[str],
    api_key: str,
    sem_verificacao: bool,
    log_fn,
    material_preimportado: Optional[dict[str, str]] = None,
) -> dict:
    """
    Executa as 3 etapas e retorna um dicionário pronto para o relatório.
    """

    # ──────────────────────────────────────────────────────────────────
    # ETAPA 1 — Ler a dissertação
    # ──────────────────────────────────────────────────────────────────
    log_fn("📄 ETAPA 1: Lendo a dissertação…", "etapa")

    texto_diss = ler_docx(diss_path)
    n_chars = len(texto_diss)
    log_fn(f"   Texto extraído: {n_chars:,} caracteres")

    if n_chars < 200:
        log_fn("   ⚠ Texto muito curto — arquivo pode ser escaneado ou corrompido", "warn")
        log_fn(f"   Conteúdo: {repr(texto_diss[:300])}", "warn")
    else:
        log_fn("   Primeiras linhas do texto extraído:")
        for linha in texto_diss.split("\n")[:6]:
            if linha.strip():
                log_fn(f"   | {linha[:110]}")

    # ──────────────────────────────────────────────────────────────────
    # ETAPA 2 — Ler material de apoio
    # ──────────────────────────────────────────────────────────────────
    log_fn("📂 ETAPA 2: Lendo material de apoio…", "etapa")

    material: dict[str, str]
    if material_preimportado is not None:
        material = material_preimportado
        log_fn(f"   {len(material)} arquivo(s) importado(s) da Localização de Referências")
        for nome, texto in material.items():
            log_fn(f"   ✓ {nome} ({len(texto):,} chars)")
    else:
        extensoes = {".pdf", ".docx", ".doc", ".txt"}
        arquivos = [
            f for f in Path(refs_dir).rglob("*")
            if f.is_file() and f.suffix.lower() in extensoes
        ]
        log_fn(f"   {len(arquivos)} arquivo(s) encontrado(s)")

        material = {}
        for arq in arquivos:
            texto = ler_arquivo(arq)
            if texto and len(texto) > 50 and not texto.startswith("[ERRO"):
                material[arq.name] = texto
                log_fn(f"   ✓ {arq.name} ({len(texto):,} chars)")
            else:
                log_fn(f"   ✗ {arq.name} — não foi possível ler")

    log_fn(f"   {len(material)} documento(s) lido(s) com sucesso")

    # ──────────────────────────────────────────────────────────────────
    # ETAPA 3 — Identificar citações e referências
    # ──────────────────────────────────────────────────────────────────
    log_fn("🔍 ETAPA 3: Identificando citações…", "etapa")

    citacoes, idx_refs = encontrar_citacoes(texto_diss)
    log_fn(f"   {len(citacoes)} citação(ões) encontrada(s) (padrões ABNT)")

    if idx_refs >= 0:
        log_fn(f"   Seção REFERÊNCIAS encontrada (parágrafo {idx_refs})")
    else:
        log_fn("   ⚠ Seção REFERÊNCIAS não encontrada — verificar título no .docx", "warn")

    if citacoes:
        log_fn("   Exemplos:")
        for c in citacoes[:5]:
            log_fn(f"   ✓ {c.texto!r}")
    else:
        log_fn("   ⚠ Nenhuma citação encontrada com regex ABNT", "warn")
        log_fn("   Padrões buscados: (AUTOR, ANO) e AUTOR (ANO)", "warn")
        paragrafos = [p for p in texto_diss.split("\n") if len(p.strip()) > 20]
        log_fn("   Exemplos do texto (para verificar o formato das citações):")
        for p in paragrafos[:5]:
            log_fn(f"   | {p[:110]}")

    referencias = extrair_referencias(texto_diss, idx_refs)
    log_fn(f"   {len(referencias)} referência(s) na lista")
    for r in referencias[:3]:
        log_fn(f"   ✓ {r.sobrenome} ({r.ano}): {r.titulo[:60]}")

    # ──────────────────────────────────────────────────────────────────
    # Cruzamento
    # ──────────────────────────────────────────────────────────────────
    log_fn("🔗 Cruzando citações com referências…", "etapa")

    # Índice por TODOS os co-autores → mesma referência
    idx_refs_dict: dict[str, Referencia] = {}
    for r in referencias:
        for sob in r.sobrenomes:
            chave = f"{sob}_{r.ano}"
            if chave not in idx_refs_dict:  # primeiro autor tem prioridade
                idx_refs_dict[chave] = r

    chaves_primarias: set[str] = set()   # chave do primeiro autor de cada ref citada
    citadas_sem_ref: list[Citacao] = []
    pareamentos: list[tuple[Citacao, Referencia]] = []

    for cit in citacoes:
        encontrou = False
        for autor in cit.autores:
            chave = f"{autor}_{cit.ano}"
            if chave in idx_refs_dict:
                ref = idx_refs_dict[chave]
                chave_prim = f"{ref.sobrenome}_{ref.ano}"
                chaves_primarias.add(chave_prim)
                pareamentos.append((cit, ref))
                encontrou = True
                break
        if not encontrou:
            citadas_sem_ref.append(cit)

    refs_sem_citacao = [r for r in referencias if f"{r.sobrenome}_{r.ano}" not in chaves_primarias]

    log_fn(f"   {len(pareamentos)} par(es) | "
           f"{len(citadas_sem_ref)} cit. sem ref | "
           f"{len(refs_sem_citacao)} ref. sem citação")

    # ──────────────────────────────────────────────────────────────────
    # Verificação semântica
    # ──────────────────────────────────────────────────────────────────
    # Páginas estimadas (offset de caracteres / chars por página, padrão ABNT ~1800)
    paragrafos_diss = texto_diss.split("\n")
    CHARS_POR_PAGINA = 1800
    offsets = []
    acumulado = 0
    for p in paragrafos_diss:
        offsets.append(acumulado)
        acumulado += len(p) + 1

    def _pagina_estimada(idx_paragrafo: int) -> int:
        if 0 <= idx_paragrafo < len(offsets):
            return offsets[idx_paragrafo] // CHARS_POR_PAGINA + 1
        return 0

    verificacoes: list[dict] = []

    if api_key.strip() and not sem_verificacao and pareamentos:
        # Agrupa citações por obra (mesma referência) para evitar releitura/reanálise repetida
        grupos: dict[tuple[str, str], list[tuple[Citacao, Referencia]]] = {}
        ordem_grupos: list[tuple[str, str]] = []
        for cit, ref in pareamentos:
            chave = (ref.sobrenome, ref.ano)
            if chave not in grupos:
                grupos[chave] = []
                ordem_grupos.append(chave)
            grupos[chave].append((cit, ref))

        log_fn(f"🤖 ETAPA 4: Verificando {len(pareamentos)} citação(ões) "
               f"agrupadas em {len(ordem_grupos)} obra(s) com Claude…", "etapa")

        for gi, chave in enumerate(ordem_grupos, 1):
            grupo = grupos[chave]
            ref = grupo[0][1]
            citacoes_grupo = [cit for cit, _ in grupo]
            log_fn(f"   [{gi}/{len(ordem_grupos)}] {ref.sobrenome} ({ref.ano}) "
                   f"— {len(citacoes_grupo)} citação(ões)")

            arquivo_fonte, texto_fonte = _buscar_fonte(citacoes_grupo[0].autores[0], ref.ano, ref.titulo, material)

            if not texto_fonte:
                log_fn(f"      → fonte não encontrada: {ref.sobrenome} ({ref.ano})")
                for cit, _ in grupo:
                    verificacoes.append({
                        "citacao": cit.texto, "autores": cit.autores, "ano": cit.ano,
                        "contexto": cit.contexto, "referencia": ref.texto[:200],
                        "pagina_estimada": _pagina_estimada(cit.paragrafo_idx),
                        "arquivo_fonte": "", "veredicto": "SEM_FONTE",
                        "justificativa": f"Arquivo para {ref.sobrenome} ({ref.ano}) não encontrado na pasta de referências",
                        "trecho_fonte": "",
                    })
                continue

            resultados = verificar_grupo_com_claude(citacoes_grupo, ref, arquivo_fonte, texto_fonte, api_key)
            for cit, resultado in zip(citacoes_grupo, resultados):
                emoji = {"CORRETO": "✓", "INCORRETO": "✗", "PARCIAL": "⚠", "SEM_FONTE": "?", "ERRO": "!"}.get(
                    resultado.get("veredicto", ""), "?")
                pag = _pagina_estimada(cit.paragrafo_idx)
                log_fn(f"      {emoji} (p.~{pag}) {resultado.get('veredicto')} — {resultado.get('justificativa','')[:80]}")
                verificacoes.append({
                    "citacao": cit.texto, "autores": cit.autores, "ano": cit.ano,
                    "contexto": cit.contexto, "referencia": ref.texto[:200],
                    "pagina_estimada": pag,
                    "arquivo_fonte": arquivo_fonte, **resultado,
                })
    elif not api_key.strip() and pareamentos:
        log_fn("⏭ Verificação semântica ignorada (sem chave API).", "etapa")
        log_fn(f"   Para verificar as {len(pareamentos)} citação(ões), informe sua chave API Claude.")
    else:
        log_fn("⏭ Verificação semântica ignorada.", "etapa")

    return {
        "citacoes": citacoes,
        "referencias": referencias,
        "citadas_sem_ref": citadas_sem_ref,
        "refs_sem_citacao": refs_sem_citacao,
        "pareamentos": pareamentos,
        "verificacoes": verificacoes,
    }


# ═══════════════════════════════════════════════════════════════════════
# GERAÇÃO DE RELATÓRIO
# ═══════════════════════════════════════════════════════════════════════

def gerar_relatorio(resultado: dict, rel_dir: Path, diss_path: str) -> None:
    """Gera os relatórios TXT, HTML e JSON a partir do resultado de analisar()."""
    from datetime import datetime
    rel_dir = Path(rel_dir)
    rel_dir.mkdir(parents=True, exist_ok=True)

    citadas_sem_ref  = resultado["citadas_sem_ref"]
    refs_sem_citacao = resultado["refs_sem_citacao"]
    verificacoes     = resultado["verificacoes"]

    corretas   = [v for v in verificacoes if v.get("veredicto") == "CORRETO"]
    incorretas = [v for v in verificacoes if v.get("veredicto") == "INCORRETO"]
    parciais   = [v for v in verificacoes if v.get("veredicto") == "PARCIAL"]
    sem_fonte  = [v for v in verificacoes if v.get("veredicto") == "SEM_FONTE"]

    ts = datetime.now().strftime("%Y-%m-%d %H:%M")

    # ── JSON ──────────────────────────────────────────────────────────
    dados_json = {
        "metadados": {"dissertacao": diss_path, "gerado_em": ts},
        "resumo": {
            "citacoes_sem_referencia": len(citadas_sem_ref),
            "referencias_sem_citacao": len(refs_sem_citacao),
            "corretas":   len(corretas),
            "parciais":   len(parciais),
            "incorretas": len(incorretas),
            "sem_fonte":  len(sem_fonte),
        },
        "citadas_sem_referencia": [
            {"citacao": c.texto, "autores": c.autores, "ano": c.ano}
            for c in citadas_sem_ref
        ],
        "referenciadas_sem_citacao": [
            {"referencia": r.texto, "sobrenome": r.sobrenome, "sobrenomes": r.sobrenomes, "ano": r.ano}
            for r in refs_sem_citacao
        ],
        "verificacoes": verificacoes,
    }
    (rel_dir / "relatorio.json").write_text(
        json.dumps(dados_json, ensure_ascii=False, indent=2), encoding="utf-8"
    )

    # ── TXT ───────────────────────────────────────────────────────────
    sep = "=" * 70
    L = [sep, "  VERIFICADOR DE CITAÇÕES ACADÊMICAS", f"  Gerado em: {ts}", sep, ""]

    L += [f"RESUMO", "-" * 40,
          f"Citações sem referência:     {len(citadas_sem_ref)}",
          f"Referências sem citação:     {len(refs_sem_citacao)}",
          f"Verificações corretas:       {len(corretas)}",
          f"Verificações parciais:       {len(parciais)}",
          f"Verificações incorretas:     {len(incorretas)}",
          f"Fonte não encontrada:        {len(sem_fonte)}", ""]

    if citadas_sem_ref:
        L += ["CITAÇÕES SEM REFERÊNCIA", "-" * 40]
        for c in citadas_sem_ref:
            L.append(f"  {c.texto}")
        L.append("")

    if refs_sem_citacao:
        L += ["REFERÊNCIAS SEM CITAÇÃO NO TEXTO", "-" * 40]
        for r in refs_sem_citacao:
            L.append(f"  {r.texto[:100]}")
        L.append("")

    def _bloco_verif(prefixo: str, v: dict) -> list[str]:
        linhas = [
            f"  {prefixo} {v['citacao']}",
            f"    Veredicto  : {v.get('veredicto','')}",
            f"    Justificativa: {v.get('justificativa','')}",
        ]
        if v.get("pagina_estimada"):
            linhas.append(f"    Página estimada na dissertação: ~{v['pagina_estimada']}")
        if v.get("arquivo_fonte"):
            linhas.append(f"    Arquivo fonte: {v['arquivo_fonte']}")
        if v.get("trecho_fonte"):
            linhas.append(f"    Trecho na obra: \"{v['trecho_fonte'][:200]}\"")
        if v.get("contexto"):
            linhas.append(f"    Contexto (dissertação): {v['contexto'][:200]}")
        linhas.append("")
        return linhas

    if incorretas:
        L += ["CITAÇÕES COM PROBLEMAS (INCORRETAS)", "-" * 40]
        for v in incorretas:
            L.extend(_bloco_verif("✗", v))

    if parciais:
        L += ["CITAÇÕES PARCIAIS / COM RESSALVAS", "-" * 40]
        for v in parciais:
            L.extend(_bloco_verif("⚠", v))

    if corretas:
        L += [f"CITAÇÕES CORRETAS ({len(corretas)})", "-" * 40]
        for v in corretas:
            L.extend(_bloco_verif("✓", v))

    (rel_dir / "relatorio.txt").write_text("\n".join(L), encoding="utf-8")

    # ── HTML ──────────────────────────────────────────────────────────
    def esc(s: str) -> str:
        return str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

    def secao(titulo: str, items: list, cor: str, fn) -> str:
        if not items:
            return ""
        h = f'<h2 style="color:{cor}">{esc(titulo)} ({len(items)})</h2><ul>'
        return h + "".join(f"<li>{fn(i)}</li>" for i in items) + "</ul>"

    html = f"""<!DOCTYPE html>
<html lang="pt-BR"><head><meta charset="UTF-8">
<title>Relatório de Citações</title>
<style>
  body{{font-family:sans-serif;max-width:900px;margin:2em auto;padding:0 1em}}
  h1{{background:#2c3e50;color:#fff;padding:.8em;border-radius:4px}}
  h2{{margin-top:1.5em}} ul{{list-style:none;padding:0}}
  li{{border-left:4px solid #ccc;padding:.5em 1em;margin:.4em 0;background:#f9f9f9}}
  .ok{{border-color:#27ae60}} .err{{border-color:#e74c3c}} .warn{{border-color:#f39c12}}
  .info{{border-color:#3498db}} .resumo{{display:flex;gap:1em;flex-wrap:wrap}}
  .card{{background:#ecf0f1;border-radius:6px;padding:1em;min-width:130px;text-align:center}}
  .card span{{font-size:2em;font-weight:bold}}
</style></head><body>
<h1>Verificador de Citações Acadêmicas</h1>
<p>Gerado em: {esc(ts)} | Arquivo: {esc(diss_path)}</p>
<div class="resumo">
  <div class="card"><span>{len(citadas_sem_ref)}</span><br>Cit. sem ref.</div>
  <div class="card"><span>{len(refs_sem_citacao)}</span><br>Ref. sem cit.</div>
  <div class="card"><span>{len(corretas)}</span><br>Corretas</div>
  <div class="card"><span>{len(parciais)}</span><br>Parciais</div>
  <div class="card"><span>{len(incorretas)}</span><br>Incorretas</div>
  <div class="card"><span>{len(sem_fonte)}</span><br>Sem fonte</div>
</div>
"""
    def _card_verif(v: dict, cls: str) -> str:
        trecho = v.get("trecho_fonte", "")
        trecho_html = (f'<blockquote style="margin:.4em 0;font-style:italic;color:#555">'
                       f'"{esc(trecho)}"</blockquote>') if trecho else ""
        pag = v.get("pagina_estimada")
        pag_html = f'<br><small>Página estimada: ~{pag}</small>' if pag else ""
        return (
            f'<b class="{cls}">{esc(v["citacao"])}</b>'
            f'{pag_html}'
            f'<br><small>Fonte: <em>{esc(v.get("arquivo_fonte","?"))}</em></small>'
            f'<br><small>Ref.: {esc(v.get("referencia","")[:120])}</small>'
            f'<br>{esc(v.get("justificativa",""))}'
            f'{trecho_html}'
            f'<br><small style="color:#888">Contexto na dissertação: {esc(v.get("contexto","")[:200])}</small>'
        )

    html += secao("Citações sem referência na lista", citadas_sem_ref, "#e74c3c",
                  lambda c: f'<b class="err">{esc(c.texto)}</b> <small>({esc(", ".join(c.autores))}, {esc(c.ano)})</small>'
                            f'<br><small style="color:#888">{esc(c.contexto[:200])}</small>')
    html += secao("Referências sem citação no texto", refs_sem_citacao, "#e67e22",
                  lambda r: f'<span class="warn">{esc(r.texto[:200])}</span>')
    html += secao("Citações INCORRETAS", incorretas, "#c0392b",
                  lambda v: _card_verif(v, "err"))
    html += secao("Citações PARCIAIS / com ressalvas", parciais, "#d35400",
                  lambda v: _card_verif(v, "warn"))
    html += secao("Citações CORRETAS", corretas, "#27ae60",
                  lambda v: _card_verif(v, "ok"))
    html += "</body></html>"
    (rel_dir / "relatorio.html").write_text(html, encoding="utf-8")
