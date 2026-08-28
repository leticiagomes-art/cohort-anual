#!/usr/bin/env python3
"""
merge_csvs.py
-------------
Roda DEPOIS do analise_ano_tigeroffers.py.

Os dados ficam PARTICIONADOS POR MÊS do pedido:

    data/pedidos/2026-01.csv.gz
    data/reembolsos/2026-01.csv.gz
    data/chargebacks/2026-01.csv.gz

Vantagens sobre um arquivo único:
  - cada arquivo tem ~200 KB, dá para subir pelo navegador do GitHub
  - o merge só reescreve os meses que receberam dado novo
  - o diff do commit fica legível
"""
import sys
import pandas as pd
from pathlib import Path

ROOT     = Path(__file__).resolve().parent.parent
DATA_DIR = ROOT / "data"
NEW_DIR  = DATA_DIR / "new"

# base → (subpasta da partição, colunas que identificam uma linha única)
BASES = {
    "base_pedidos":     ("pedidos",     ["order_item_id", "order_id"]),
    # NAO usar order_item_id aqui: no export de reembolso/chargeback da
    # BuyGoods essa coluna vem VAZIA em 100% das linhas, e a chave inteira
    # viraria nula. order_id + data + valor e o que identifica a linha.
    "base_reembolsos":  ("reembolsos",  ["order_id", "refund_date", "amount"]),
    "base_chargebacks": ("chargebacks", ["order_id", "cb_date", "amount"]),
}

def log(msg):
    print(f"[merge] {msg}", flush=True)


def chave_estavel(df, colunas):
    """
    Monta a chave de deduplicação como STRING normalizada.

    O histórico e o export novo serializam data de formas diferentes.
    Comparar objetos datetime fazia linhas idênticas virarem chaves
    distintas — foi o que apagou 19.325 reembolsos num run anterior.
    Normalizar para texto resolve de vez.
    """
    cols = [c for c in colunas if c in df.columns]
    if not cols:
        return None
    partes = []
    for c in cols:
        col = df[c]
        if "date" in c or "data" in c:
            col = pd.to_datetime(col, errors="coerce").dt.strftime("%Y-%m-%d")
        elif col.dtype.kind == "f":
            col = col.round(2)
        partes.append(col.astype(str).str.strip())
    chave = partes[0]
    for p in partes[1:]:
        chave = chave + "|" + p
    return chave


total_novas = 0

for nome, (subpasta, chave_cols) in BASES.items():
    csv_novo = NEW_DIR / f"{nome}.csv"
    destino  = DATA_DIR / subpasta
    destino.mkdir(parents=True, exist_ok=True)

    if not csv_novo.exists():
        log(f"{nome}: sem arquivo novo — pulando")
        continue

    novo = pd.read_csv(csv_novo, low_memory=False)
    if not len(novo):
        log(f"{nome}: arquivo novo vazio — pulando")
        continue

    novo["data_pedido"] = pd.to_datetime(novo["data_pedido"], errors="coerce")
    novo["_part"] = novo["data_pedido"].dt.strftime("%Y-%m")
    novo = novo[novo["_part"].notna()]
    log(f"{nome}: {len(novo):,} linhas novas em {novo['_part'].nunique()} mes(es)")
    total_novas += len(novo)

    for mes, bloco in novo.groupby("_part"):
        bloco = bloco.drop(columns="_part")
        arq   = destino / f"{mes}.csv.gz"

        antigo = None
        if arq.exists() and arq.stat().st_size > 20:
            try:
                antigo = pd.read_csv(arq, low_memory=False)
            except Exception as e:
                log(f"  {subpasta}/{mes}: historico ilegivel ({e}) - recriando")

        if antigo is not None and len(antigo):
            juntos = pd.concat([antigo, bloco], ignore_index=True)
        else:
            juntos = bloco

        antes = len(juntos)
        k = chave_estavel(juntos, chave_cols)
        if k is not None:
            juntos = (juntos.assign(_k=k)
                            .drop_duplicates("_k", keep="last")
                            .drop(columns="_k"))
        depois = len(juntos)

        # Trava contra perda de historico.
        #
        # A comparacao tem que ser com o historico JA DEDUPLICADO, nao com a
        # contagem bruta: o export da BuyGoods repete linhas (um pedido com
        # produto e frete reembolsados no mesmo dia vinha duplicado), e essas
        # repeticoes existiam dentro do proprio arquivo antigo. Comparar com
        # o bruto fazia a trava disparar por uma limpeza legitima.
        if antigo is not None and len(antigo):
            k_ant = chave_estavel(antigo, chave_cols)
            base_ant = antigo.assign(_k=k_ant)["_k"].nunique() if k_ant is not None else len(antigo)
            if depois < base_ant:
                log(f"  {subpasta}/{mes}: ABORTADO — {depois} < {base_ant} unicos do historico")
                sys.exit(
                    f"Merge de {subpasta}/{mes} perderia dado "
                    f"({base_ant} unicos -> {depois}). Verifique a chave de deduplicacao."
                )

        juntos.to_csv(arq, index=False, compression="gzip")
        kb = arq.stat().st_size / 1024
        marca = "novo" if antigo is None else f"+{depois - len(antigo)}"
        log(f"  {subpasta}/{mes}: {depois:,} linhas ({marca}, {antes - depois} dup) - {kb:.0f} KB")

log(f"Concluido. {total_novas:,} linhas novas processadas.")
