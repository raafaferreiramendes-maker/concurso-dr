"""
concursos_dr.py — Concursos públicos de obras (Diário da República, 2.ª série, Parte L)

Lê o RSS da Parte L, descarrega o PDF de cada anúncio de procedimento,
extrai os campos e acrescenta a uma base de
 dados Excel (concursos_base.xlsx)
as empreitadas que passam nos filtros. Anúncios já presentes na base não são
descarregados de novo; a versão anterior da base fica em concursos_base.bak.xlsx.

LIMITAÇÃO IMPORTANTE: o RSS só traz o ÚLTIMO Diário publicado.
Para não perder dias, o script tem de correr todos os dias úteis.

Instalar:  pip install requests pdfplumber pandas openpyxl
Correr:    python concursos_dr.py            (aplica os filtros abaixo)
           python concursos_dr.py --todas    (todas as obras do dia, sem filtros)
"""

import io
import re
import shutil
import sys
import time
import unicodedata
import xml.etree.ElementTree as ET
from datetime import datetime
from pathlib import Path

import pandas as pd
import pdfplumber
import requests
from openpyxl.utils import get_column_letter

# ============================ CONFIGURAÇÃO ============================
# Zona geográfica: deixar a lista vazia = sem filtro nesse nível.
# Basta um dos níveis coincidir (distrito OU concelho OU NUT III).
DISTRITOS = []                      # ex.: ["Porto", "Aveiro", "Braga"]
CONCELHOS = []                      # ex.: ["Penafiel", "Paredes", "Santa Maria da Feira"]
NUTS3 = []                          # ex.: ["PT11A"]  (Área Metropolitana do Porto)

PRECO_MIN = None                    # € sem IVA (preço base total do procedimento); None = sem mínimo
PRECO_MAX = None                    # None = sem teto
CPV_PREFIXOS = ["45",               # trabalhos de construção
                "507",              # reparação e manutenção de instalações de edifícios
                "5023",             # reparação e manutenção de estradas e equipamentos afins
                "5051",             # reparação e manutenção de bombas, válvulas e torneiras
                "50870",            # reparação e manutenção de equipamento de recreio público
                "7731",             # plantação e manutenção de espaços verdes
                "90650",            # remoção de amianto
                "51",               # serviços de instalação (exceto software)
                "44211",            # edifícios pré-fabricados
                "44221",            # janelas, portas e artigos afins
                "34928",            # mobiliário rodoviário (guardas, sinalização)
                "90722",            # recuperação ambiental
                "71",               # arquitetura, engenharia, fiscalização e inspeção
                "92522",            # preservação de locais e monumentos históricos
                "50225",            # manutenção de vias férreas
                "50246",            # manutenção de equipamento portuário e estruturas flutuantes
                "50413200",         # manutenção de equipamento de combate a incêndios
                "77211500",         # manutenção de árvores
                "77320"]            # manutenção de campos de desporto

PASTA_SAIDA = Path(".")
BASE = PASTA_SAIDA / "concursos_base.xlsx"
FORMATO_CONTABILIDADE = '_-* #,##0.00 "€"_-;-* #,##0.00 "€"_-;_-* "-"?? "€"_-;_-@_-'
PAUSA_ENTRE_PEDIDOS = 1.0           # segundos — não sobrecarregar o servidor do DR
# =====================================================================

RSS_URL = "https://files.diariodarepublica.pt/rss/serie2&parte=l.xml"
HEADERS = {"User-Agent": "Mozilla/5.0 (monitor de concursos; uso pessoal)"}

PLATAFORMAS = {             # valor no anúncio -> nome normalizado (lista IMPIC)
    "ACIN": "acinGov",
    "ACINGOV": "acinGov",
    "VORTAL": "VortalGOV",
    "VORTALGOV": "VortalGOV",
    "ANOGOV": "anoGov",
    "COMPRASPT": "ComprasPT",
}

ALVARA_MAX = 500           # caracteres da coluna "Alvará exigido" (o resto fica no PDF do DR)
ALVARA_EM_FALTA = "Não indicado no anúncio (ver Programa do Procedimento)"

COLUNAS = [
    "Data publicação", "Nº anúncio", "Entidade adjudicante", "Objeto",
    "Tipo de obra (CPV)", "Preço base (€)", "Prazo propostas", "Prazo execução", "Local",
    "Alvará exigido", "Plataforma", "Link peças", "PDF DR",
]


# ----------------------------- utilitários -----------------------------
def sem_acentos(s: str) -> str:
    s = unicodedata.normalize("NFKD", s or "")
    return "".join(c for c in s if not unicodedata.combining(c)).lower().strip()


def valor(texto: str, rotulo: str) -> str:
    """Primeiro valor de 'Rótulo: valor' (uma linha)."""
    m = re.search(rf"^{rotulo}\s*(.+)$", texto, re.MULTILINE)
    return m.group(1).strip() if m else ""


def todos(texto: str, rotulo: str) -> list:
    return [v.strip() for v in re.findall(rf"^{rotulo}\s*(.+)$", texto, re.MULTILINE)]


def euros(s: str):
    """'1.014.415,00 EUR' ou '3,265,000.00 EUR' -> float"""
    s = re.sub(r"[^\d,\.]", "", s or "")
    if not s:
        return None
    if re.search(r",\d{2}$", s):          # formato PT: 1.234,56
        s = s.replace(".", "").replace(",", ".")
    else:                                 # formato antigo: 1,234.56
        s = s.replace(",", "")
    try:
        return float(s)
    except ValueError:
        return None


def seccao(texto: str, inicio: str, fim: str) -> str:
    """Texto entre o cabeçalho 'inicio' e o cabeçalho 'fim' (ex.: '9 - LOCAL', '10 - ')."""
    m = re.search(rf"^{inicio}.*?(?=^{fim}|\Z)", texto, re.MULTILINE | re.DOTALL)
    return m.group(0) if m else ""


def limpar_rodapes(texto: str) -> str:
    """Remove o cabeçalho/rodapé repetido em cada página (série, nº, data, página)."""
    texto = re.sub(r"^2\.ª série$", "", texto, flags=re.MULTILINE)
    texto = re.sub(r"^N\.º \d+$", "", texto, flags=re.MULTILINE)
    texto = re.sub(r"^\d{2}-\d{2}-\d{4}$", "", texto, flags=re.MULTILINE)
    texto = re.sub(r"^\d{1,3}$", "", texto, flags=re.MULTILINE)
    texto = re.sub(r"^\d{9}$", "", texto, flags=re.MULTILINE)
    # texto vertical da margem ("Anúncio de procedimento n.º 26245/2026", lido ao contrário)
    texto = re.sub(r"^(\d{4}/\d+|º\.n|otnemidecorp|ed|oicnúnA)$", "", texto, flags=re.MULTILINE)
    return texto


def juntar_local(distrito: str, concelho: str) -> str:
    """'Santo Tirso (Porto)'; só um dos dois quando falta o outro."""
    if distrito and concelho:
        return f"{concelho} ({distrito})"
    return concelho or distrito


def texto_multilinha(texto: str, rotulo: str, ate: str) -> str:
    """Valor que pode ocupar várias linhas: do rótulo até à linha que começa por 'ate'."""
    m = re.search(rf"^{rotulo}\s*(.+?)\n(?={ate})", texto, re.MULTILINE | re.DOTALL)
    return " ".join(m.group(1).split()) if m else valor(texto, rotulo)


# ------------------------------ parsing --------------------------------
def parse_anuncio(texto: str) -> dict:
    t = limpar_rodapes(texto)
    local = seccao(t, r"9 - LOCAL DA EXECU", r"10 - ")
    if "Concelho:" not in local:            # anúncio sem local de execução: usa a morada da entidade
        local = seccao(t, r"1 - IDENTIFICA", r"2 - ")
    # Só o bloco do PROCEDIMENTO (os blocos por lote costumam repetir o mesmo local)
    distritos = list(dict.fromkeys(todos(local, "Distrito:")))
    concelhos = list(dict.fromkeys(todos(local, "Concelho:")))
    nuts = list(dict.fromkeys(todos(local, "NUT III:")))

    preco = valor(t, r"Preço base s/IVA:") or valor(t, r"Valor do preço base do procedimento:")
    # "Vocabulário Principal: 45233140 - Obras em estradas" (repete-se por lote; guardam-se os distintos).
    # Qualquer outro rótulo "Vocabulário ...:" (CPV adicionais) também entra na coluna.
    cpv = re.search(r"Vocabul[áa]rio [Pp]rincipal:\s*(\d{8})", t)
    cpvs = list(dict.fromkeys((c, d.strip()) for c, d in
                              re.findall(r"Vocabul[áa]rio \w+:\s*(\d{8})(?:-\d)?\s*-?\s*(.*)", t)))
    plataforma_raw = valor(t, "Plataforma eletrónica utilizada pela entidade adjudicante:")
    # "Link para acesso às peças do concurso (URL):" — às vezes vem com letras espaçadas,
    # e o URL pode partir-se em várias linhas até à secção seguinte ("20 - ...")
    link = re.search(r"c\s*o\s*n\s*c\s*u\s*r\s*s\s*o\s*\(\s*U\s*R\s*L\s*\)\s*:(.*?)(?=^\d{1,2} - [A-Z]|\Z)",
                     t, re.MULTILINE | re.DOTALL)
    link = re.match(r"https?://\S+", "".join(link.group(1).split())) if link else None
    # Alvará: texto livre da secção 12; às vezes só remete para o Programa do Procedimento
    alvara = re.search(r"Descrição:\s*(.*)", seccao(t, r"12 - DOCUMENTOS DE HABILITA", r"1[3-9] - "), re.DOTALL)
    alvara = " ".join(alvara.group(1).split()) if alvara else ALVARA_EM_FALTA
    if len(alvara) > ALVARA_MAX:
        alvara = alvara[:ALVARA_MAX].rsplit(" ", 1)[0] + " …"

    return {
        "Nº anúncio": valor(t, r"Anúncio de procedimento n\.º").split()[0] if valor(t, r"Anúncio de procedimento n\.º") else "",
        "Entidade adjudicante": valor(t, "Designação da entidade adjudicante:"),
        "Objeto": texto_multilinha(t, "Designação do contrato:", r"Descrição:|Tipo de Contrato"),
        "Tipo contrato": valor(t, "Tipo de Contrato Principal:") or valor(t, "Tipo de Contrato:"),
        "CPV": cpv.group(1) if cpv else "",           # código principal: só para o filtro
        "Tipo de obra (CPV)": "; ".join(f"{c} - {d}" if d else c for c, d in cpvs),
        "Preço base (€)": euros(preco),
        "Prazo propostas": valor(t, "Prazo para apresentação das propostas:"),
        "Prazo execução": valor(t, "Prazo de execução do contrato:"),
        "Local": juntar_local("; ".join(distritos), "; ".join(concelhos)),
        "Alvará exigido": alvara,
        "Distrito": "; ".join(distritos),           # Distrito/Concelho/NUT III: só para os filtros
        "Concelho": "; ".join(concelhos),
        "NUT III": "; ".join(nuts),
        "Plataforma": PLATAFORMAS.get(plataforma_raw.upper().replace(" ", ""),
                                      f"?? {plataforma_raw}" if plataforma_raw else ""),
        "Link peças": link.group(0) if link else "",
    }


# ------------------------------- filtros -------------------------------
def e_obra(r: dict) -> bool:
    return "empreitada" in sem_acentos(r["Tipo contrato"]) or r["CPV"].startswith(tuple(CPV_PREFIXOS))


def na_zona(r: dict) -> bool:
    if not (DISTRITOS or CONCELHOS or NUTS3):
        return True
    d = {sem_acentos(x) for x in r["Distrito"].split(";")}
    c = {sem_acentos(x) for x in r["Concelho"].split(";")}
    n = {x.strip().split(" ")[0].upper() for x in r["NUT III"].split(";")}
    return (bool(d & {sem_acentos(x) for x in DISTRITOS})
            or bool(c & {sem_acentos(x) for x in CONCELHOS})
            or bool(n & {x.upper() for x in NUTS3}))


def no_intervalo_preco(r: dict) -> bool:
    p = r["Preço base (€)"]
    if p is None:                       # preço não lido: manter, para não perder a obra
        return True
    return (PRECO_MIN is None or p >= PRECO_MIN) and (PRECO_MAX is None or p <= PRECO_MAX)


# -------------------------------- rede ---------------------------------
def ler_rss(sessao: requests.Session):
    resp = sessao.get(RSS_URL, timeout=30)
    resp.raise_for_status()
    canal = ET.fromstring(resp.content).find("channel")
    data_diario = canal.findtext("description", "").strip()[-10:]   # "... - 2026-10-30"
    itens = []
    for it in canal.findall("item"):
        titulo = (it.findtext("title") or "").strip()
        if titulo.startswith("Anúncio de procedimento"):
            itens.append({"titulo": titulo, "link": (it.findtext("link") or "").strip()})
    return data_diario, itens


def texto_pdf(sessao: requests.Session, url: str) -> str:
    resp = sessao.get(url, timeout=60)
    resp.raise_for_status()
    with pdfplumber.open(io.BytesIO(resp.content)) as pdf:
        return "\n".join(p.extract_text() or "" for p in pdf.pages)


# ------------------------------- base ----------------------------------
def ler_base() -> pd.DataFrame:
    """Lê a base; na primeira vez importa os Excel diários antigos (concursos_AAAA-MM-DD*.xlsx)."""
    def ler(f):
        df = pd.read_excel(f, dtype={"Nº anúncio": str, "CPV": str, "Tipo de obra (CPV)": str})
        if "Local" not in df and "Distrito" in df:     # ficheiros antigos: Distrito + Concelho
            df["Local"] = [juntar_local(d, c) for d, c in
                           zip(df["Distrito"].fillna(""), df["Concelho"].fillna(""))]
        return df.reindex(columns=COLUNAS)             # descarta colunas antigas

    if BASE.exists():
        return ler(BASE)
    antigos = sorted(PASTA_SAIDA.glob("concursos_20*.xlsx"))
    if not antigos:
        return pd.DataFrame(columns=COLUNAS)
    print(f"A importar {len(antigos)} ficheiro(s) diário(s) antigo(s) para a base")
    return pd.concat([ler(f) for f in antigos])


def gravar_base(base: pd.DataFrame) -> bool:
    if BASE.exists():
        shutil.copy(BASE, BASE.with_suffix(".bak.xlsx"))
    try:
        with pd.ExcelWriter(BASE, engine="openpyxl") as w:
            base.to_excel(w, index=False)
            folha = w.sheets["Sheet1"]
            coluna = get_column_letter(COLUNAS.index("Preço base (€)") + 1)
            for cel in folha[coluna][1:]:              # única formatação: preço em contabilidade
                cel.number_format = FORMATO_CONTABILIDADE
    except PermissionError:
        print(f"\nNão foi possível gravar {BASE.name}: feche o ficheiro no Excel e corra de novo.")
        return False
    return True


def atualizar_base(base: pd.DataFrame, novos: pd.DataFrame):
    """Junta os novos à base (sem repetidos), grava e devolve a base; None se não gravou."""
    if novos.empty:
        return base
    base = novos if base.empty else pd.concat([base, novos])
    base = (base.drop_duplicates("PDF DR", keep="last")
                .sort_values("Data publicação", ascending=False, kind="stable"))
    return base if gravar_base(base) else None


# -------------------------------- main ---------------------------------
def main():
    sem_filtros = "--todas" in sys.argv
    sessao = requests.Session()
    sessao.headers.update(HEADERS)

    data_diario, itens = ler_rss(sessao)
    base = ler_base()
    vistos = set(base["PDF DR"].dropna())
    novos_itens = [it for it in itens if it["link"] not in vistos]

    print(f"Diário de {data_diario}: {len(itens)} anúncios de procedimento na Parte L "
          f"({len(itens) - len(novos_itens)} já na base)")
    linhas, erros = [], []
    for i, it in enumerate(novos_itens, 1):
        try:
            r = parse_anuncio(texto_pdf(sessao, it["link"]))
        except Exception as e:                         # PDF ilegível, timeout, etc.
            erros.append(f"{it['link']}  ({e})")
            continue
        finally:
            time.sleep(PAUSA_ENTRE_PEDIDOS)
        print(f"  {i}/{len(novos_itens)}", end="\r")
        r["Data publicação"] = data_diario
        r["PDF DR"] = it["link"]
        if not e_obra(r):
            continue
        if sem_filtros or (na_zona(r) and no_intervalo_preco(r)):
            linhas.append(r)

    df = pd.DataFrame(linhas, columns=COLUNAS)
    base = atualizar_base(base, df)
    if base is None:
        return
    print(f"\n{len(df)} empreitadas novas; {BASE.name} tem agora {len(base)} concursos")
    if erros:
        print(f"{len(erros)} anúncios não foram lidos:")
        for e in erros:
            print("  ", e)
    estranhas = df[df["Plataforma"].str.startswith("??")]
    if len(estranhas):
        print("Plataforma fora da lista IMPIC (verificar parsing):", estranhas["Nº anúncio"].tolist())


if __name__ == "__main__":
    main()
