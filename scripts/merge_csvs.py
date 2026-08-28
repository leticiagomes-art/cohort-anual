#!/usr/bin/env python3
"""
merge_csvs.py
-------------
Roda DEPOIS do analise_ano_tigeroffers.py.

O pipeline gera novos CSVs em data/new/ com os dados dos últimos 60 dias.
Este script:
  1. Lê os CSVs acumulados em data/ (histórico completo)
  2. Lê os CSVs novos em data/new/ (últimos 60 dias)
  3. Concatena e deduplica por order_id (pedidos) ou order_id+data_evento (refund/CB)
  4. Salva de volta em data/ (substitui com o histórico atualizado)

Resultado: data/ sempre tem o acumulado desde jan até hoje.
"""
import sys
import pandas as pd
from pathlib import Path

ROOT     = Path(__file__).resolve().parent.parent
DATA_DIR = ROOT / "data"
NEW_DIR  = DATA_DIR / "new"

BASES = {
    "base_pedidos":     "order_id",
    "base_reembolsos":  ["order_id", "data_evento"],
    "base_chargebacks": ["order_id", "data_evento"],
}

def log(msg): print(f"[merge] {msg}")

total_novos = 0

for nome, chave_dedup in BASES.items():
    csv_acum = DATA_DIR / f"{nome}.csv"
    csv_novo = NEW_DIR   / f"{nome}.csv"

    if not csv_novo.exists():
        log(f"{nome}: arquivo novo não encontrado — pulando")
        continue

    novo = pd.read_csv(csv_novo, low_memory=False)
    log(f"{nome}: {len(novo)} linhas novas")

    if csv_acum.exists():
        acum = pd.read_csv(csv_acum, low_memory=False)
        log(f"{nome}: {len(acum)} linhas acumuladas")
        combined = pd.concat([acum, novo], ignore_index=True)
    else:
        log(f"{nome}: sem histórico anterior — usando só os novos")
        combined = novo

    # deduplicar — manter a linha mais recente de cada order_id
    antes = len(combined)
    if isinstance(chave_dedup, list):
        # para datas: converter antes de deduplicar
        for col in chave_dedup:
            if col in combined.columns:
                combined[col] = pd.to_datetime(combined[col], errors='coerce')
        combined = combined.sort_values(
            [c for c in chave_dedup if c in combined.columns],
            ascending=False
        ).drop_duplicates(
            subset=[c for c in chave_dedup if c in combined.columns],
            keep='first'
        )
    else:
        if chave_dedup in combined.columns:
            combined = combined.drop_duplicates(subset=chave_dedup, keep='last')

    depois = len(combined)
    duplicatas = antes - depois
    log(f"{nome}: {duplicatas} duplicatas removidas → {depois} linhas finais")

    # salvar de volta no acumulado
    combined.to_csv(csv_acum, index=False)
    log(f"{nome}: salvo em {csv_acum}")
    total_novos += len(novo)

log(f"Merge concluído. Total de linhas novas processadas: {total_novos}")
