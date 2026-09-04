# cohort-anual — Guia de configuração

Dashboard de **cohort semanal de reembolso e chargeback**, atualizado
automaticamente todo dia via GitHub Actions + GitHub Pages. O site mostra
uma única visão: quanto de cada semana de vendas volta como reembolso (ou
chargeback), semana a semana desde a compra — geral e por produto.

---

## Estrutura do repositório

```
cohort-anual/
├── .github/workflows/
│   ├── update_dashboard.yml    cron diário (09:00 UTC = 06:00 BRT)
│   └── diagnostico.yml         roda 1 conta só, com vídeo + trace
├── scripts/
│   ├── baixar_exports.py       Playwright — baixa os exports da BuyGoods
│   ├── analise_ano_tigeroffers.py   classifica e enriquece os exports
│   ├── merge_csvs.py           acumula o novo no histórico particionado
│   ├── gerar_dashboard.py      monta o payload de cohort e gera o HTML
│   └── dashboard_template.html template do dashboard (marcador __PAYLOAD__)
├── data/
│   ├── pedidos/AAAA-MM.csv.gz       particionado por mês do PEDIDO
│   ├── reembolsos/AAAA-MM.csv.gz
│   ├── chargebacks/AAAA-MM.csv.gz
│   └── new/                         saída temporária do pipeline (gitignored)
├── entrada/.gitkeep                 xlsx baixados da BuyGoods (gitignored)
├── index.html                       dashboard publicado (raiz — GitHub Pages)
└── requirements.txt
```

`scripts/accounts.json` (IDs das 31 contas) **não vai para o repositório** —
está no `.gitignore`. Em produção os IDs entram pelo Secret `BG_ACCOUNTS`.

---

## O que cada script faz

| Arquivo | Função |
|---|---|
| `baixar_exports.py` | Playwright — loga na BuyGoods, verifica os jobs agendados e baixa Order Items / Customers Refunds / Customers Chargebacks de cada conta pra `entrada/` |
| `analise_ano_tigeroffers.py` | Lê tudo em `entrada/`, classifica cada aba (pedidos/reembolso/chargeback), enriquece com data do pedido/semana/produto e grava `data/new/base_*.csv` |
| `merge_csvs.py` | Acumula `data/new/base_*.csv` no histórico particionado por mês em `data/{pedidos,reembolsos,chargebacks}/` — nunca deixa o histórico encolher (trava de segurança) |
| `gerar_dashboard.py` | Lê os CSVs particionados, monta o cohort semanal (geral e por produto) e injeta no template → `index.html` |
| `dashboard_template.html` | HTML/CSS/JS do dashboard, com `__PAYLOAD__` como marcador do JSON |

---

## Configurar os Secrets do GitHub

**Settings → Secrets and variables → Actions → New repository secret**

- **`BG_EMAIL`** — email de login na BuyGoods
- **`BG_PASSWORD`** — senha de login na BuyGoods
- **`BG_ACCOUNTS`** — JSON com os IDs de todas as contas, mesmo formato do
  `scripts/accounts.json` local: `{"BreathEaseX": 12457, "NervoLyn": 11283, ...}`

---

## Ativar o GitHub Pages

**Settings → Pages** → Source: *Deploy from a branch* → Branch: `main`,
pasta `/ (root)` (o `index.html` fica na raiz, não em `public/`).

---

## Rodar manualmente

**Actions → Atualizar Dashboard → Run workflow.** Leva ~30-50 min pras 31
contas. Acompanhe o log; ao final o `index.html` é commitado e publicado
automaticamente.

Pra rodar o extrator localmente (fora do GitHub Actions) e conferir os
exports antes de subir, use o `ExtratorBuyGoods.exe` (empacotado do
`baixar_exports.py` com PyInstaller) — veja o `LEIA-ME.txt` que acompanha
o executável.

---

## Frequência e horário

Cron em `.github/workflows/update_dashboard.yml`:
```yaml
- cron: '0 9 * * *'   # 09:00 UTC = 06:00 BRT
```
Referência: https://crontab.guru

---

## Solução de problemas

| Problema | Causa provável | Solução |
|---|---|---|
| Login falhou | Senha errada no Secret | Atualize `BG_PASSWORD` em Settings → Secrets |
| Elemento não encontrado | BuyGoods atualizou o layout | Rode `Actions → Diagnóstico` (grava vídeo + trace) e ajuste os seletores |
| Timeout no download | BuyGoods demorou demais numa conta de baixo volume | `esperar_tabela_pronta()` já reconhece o estado "tabela vazia" do DataTables como pronto — se persistir, normal ocasionalmente, o próximo run resolve |
| Cohort com números muito baixos/zerados | `merge_csvs.py`/`gerar_dashboard.py` sem `format='mixed'` no parse de data, ou join pela semana errada (`semana_ini` do evento em vez de `semana_pedido_ini` do pedido) | Já corrigido — se voltar a acontecer, confira essas duas coisas primeiro |
| `accounts.json` subiu por acidente | `.gitignore` não cobriu | Delete o arquivo no GitHub e troque a senha da BuyGoods |

---

## Por que só cohort

Este repositório existe especificamente para acompanhar o cohort semanal
de reembolso/chargeback. Ele já teve uma versão com várias abas (visão
geral, evolução mensal, quem decide, motivos, produtos, afiliados, MaxWeb,
reconciliação contra o Master Overview) — todo esse código e a lógica de
cálculo correspondente em `gerar_dashboard.py` foram removidos pra manter
o pipeline simples e o payload pequeno. O histórico continua no `git log`
se algum dia fizer sentido trazer alguma dessas visões de volta.
