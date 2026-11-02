"""
recolha_base.py — recolha ÚNICA das obras ainda com prazo de propostas aberto,
publicadas em Diários anteriores (o RSS usado em concursos_dr.py só traz o último).

A lista de anúncios vem da pesquisa do Portal BASE (base.gov.pt), que dá o prazo
de propostas e o link do PDF no DR. Cada PDF é lido com o mesmo parsing e os
mesmos filtros de concursos_dr.py e acrescentado a concursos_base.xlsx.

Instalar: pip install truststore     (o site do BASE não envia a cadeia de certificados
                                      completa; usa-se a do Windows em vez de desligar o SSL)
Correr:  python recolha_base.py      (na mesma pasta que concursos_dr.py)
"""

import json
import re
import time
from datetime import date, datetime, timedelta

import truststore
truststore.inject_into_ssl()

import pandas as pd
import requests

import concursos_dr as dr

BASE_URL = "https://www.base.gov.pt/Base4/pt/resultados/"
DIAS_ATRAS = 180            # janela de publicação a pesquisar (prazos raramente passam disto)
POR_PAGINA = 50             # o BASE devolve no máximo 50 por página
FILTROS = [f"cpv={p}" for p in dr.CPV_PREFIXOS] + ["tipocontrato=5"]   # 5 = Empreitadas de obras públicas
CACHE = dr.PASTA_SAIDA / "recolha_base_cache.json"   # candidatos do dia + ids já consultados
PAUSA_BASE = 5.0           # segundos — a firewall do BASE bloqueia com 1 s entre pedidos


class BloqueadoBASE(Exception):
    pass


def pedido_base(sessao: requests.Session, dados: dict):
    resp = sessao.post(BASE_URL, data={"version": "136.0", **dados}, timeout=60,
                       headers={"X-Requested-With": "XMLHttpRequest"})
    time.sleep(PAUSA_BASE)
    if resp.status_code == 999 or "Firewall" in resp.text[:300]:
        raise BloqueadoBASE("a firewall do BASE bloqueou os pedidos; tentar mais tarde")
    resp.raise_for_status()
    return resp.json()


def pesquisar(sessao: requests.Session, filtro: str, desde: date) -> list:
    query = f"tipoacto=0&{filtro}&desdedatapublicacao={desde:%Y-%m-%d}"
    itens, pagina = [], 0
    while True:
        lote = pedido_base(sessao, {"type": "search_anuncios", "query": query,
                                    "sort": "-drPublicationDate",
                                    "page": pagina, "size": POR_PAGINA})["items"]
        itens += lote
        if len(lote) < POR_PAGINA:
            return itens
        pagina += 1


def prazo_aberto(item: dict, hoje: date) -> bool:
    prazo = item.get("proposalDeadline") or ""
    if not re.fullmatch(r"\d{2}-\d{2}-\d{4}", prazo):   # prazo desconhecido: manter
        return True
    return datetime.strptime(prazo, "%d-%m-%Y").date() >= hoje


def chave(entidade, objeto) -> tuple:
    """Para reconhecer na base um anúncio do BASE sem pedir o detalhe."""
    return " ".join(dr.sem_acentos(str(entidade)).split()), " ".join(dr.sem_acentos(str(objeto)).split())


def ler_cache(hoje: date) -> dict:
    try:
        cache = json.loads(CACHE.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        cache = {}
    # a lista de candidatos só serve no próprio dia e com os mesmos filtros
    if cache.get("data") != hoje.isoformat() or cache.get("filtros") != FILTROS:
        cache = {"data": hoje.isoformat(), "filtros": FILTROS, "candidatos": None,
                 "vistos": cache.get("vistos", [])}
    return cache


def main():
    sessao = requests.Session()
    sessao.headers.update(dr.HEADERS)
    hoje = date.today()
    cache = ler_cache(hoje)

    if cache["candidatos"] is None:
        candidatos = {}
        try:
            for filtro in FILTROS:
                for it in pesquisar(sessao, filtro, hoje - timedelta(days=DIAS_ATRAS)):
                    candidatos[it["id"]] = it
        except BloqueadoBASE as e:
            print(f"Parado: {e}")
            return
        cache["candidatos"] = list(candidatos.values())
        CACHE.write_text(json.dumps(cache), encoding="utf-8")
    candidatos = cache["candidatos"]
    vistos_ids = set(cache["vistos"])

    base = dr.ler_base()
    vistos = set(base["PDF DR"].dropna())
    na_base = {chave(e, o) for e, o in zip(base["Entidade adjudicante"], base["Objeto"])}
    abertos = [it for it in candidatos if prazo_aberto(it, hoje)]
    falta = [it for it in abertos if it["id"] not in vistos_ids
             and chave(it.get("contractingEntity"), it.get("contractDesignation")) not in na_base]
    print(f"BASE: {len(candidatos)} anúncios de obras nos últimos {DIAS_ATRAS} dias, "
          f"{len(abertos)} com prazo aberto, {len(falta)} por consultar")

    linhas, erros = [], []
    for i, it in enumerate(falta, 1):
        print(f"  {i}/{len(falta)}", end="\r")
        try:
            link = pedido_base(sessao, {"type": "detail_anuncios", "id": it["id"]}).get("reference") or ""
            cache["vistos"].append(it["id"])
            if not link or link in vistos:
                continue
            r = dr.parse_anuncio(dr.texto_pdf(sessao, link))
            time.sleep(dr.PAUSA_ENTRE_PEDIDOS)
        except BloqueadoBASE as e:                     # guarda o que já leu; correr de novo mais tarde
            print(f"\nParado em {i}/{len(falta)}: {e}")
            break
        except Exception as e:                         # PDF ilegível, timeout, etc.
            erros.append(f"BASE id {it['id']}  ({e})")
            continue
        r["Data publicação"] = datetime.strptime(it["drPublicationDate"], "%d-%m-%Y").strftime("%Y-%m-%d")
        r["PDF DR"] = link
        if dr.e_obra(r) and dr.na_zona(r) and dr.no_intervalo_preco(r):
            linhas.append(r)

    CACHE.write_text(json.dumps(cache), encoding="utf-8")
    base = dr.atualizar_base(base, pd.DataFrame(linhas, columns=dr.COLUNAS))
    if base is None:
        return
    print(f"\n{len(linhas)} empreitadas novas; {dr.BASE.name} tem agora {len(base)} concursos")
    if erros:
        print(f"{len(erros)} anúncios não foram lidos:")
        for e in erros:
            print("  ", e)


if __name__ == "__main__":
    main()
