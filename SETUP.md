# TigerOffers Dashboard — Guia de configuração

Dashboard de reembolso e chargeback atualizado automaticamente todo dia via GitHub Actions + GitHub Pages.

---

## Estrutura do repositório

```
tigeroffers-dashboard/
├── .github/
│   └── workflows/
│       └── update_dashboard.yml   ← GitHub Actions (roda todo dia às 06h BRT)
├── scripts/
│   ├── baixar_exports.py          ← baixa os xlsx da BuyGoods (Playwright)
│   ├── analise_ano_tigeroffers.py ← pipeline de análise (já existe)
│   ├── gerar_dashboard.py         ← injeta dados no template e gera o HTML
│   ├── dashboard_template.html    ← template do dashboard (sem payload)
│   └── accounts.json              ← IDs das 31 contas (NÃO commitar)
├── data/
│   ├── base_pedidos.csv           ← base acumulada — nunca deletar
│   ├── base_reembolsos.csv
│   ├── base_chargebacks.csv
│   └── aov/                       ← relatórios de AOV dos afiliados
│       ├── white_2026-08-27.xlsx  ← período mais recente (White)
│       └── black_2026-06-01.xlsx  ← período anterior (Black)
├── entrada/                       ← vazia; preenchida pelo script de download
├── public/
│   └── index.html                 ← dashboard final (atualizado automaticamente)
├── requirements.txt
└── .gitignore                     ← accounts.json está aqui — nunca sobe
```

---

## Passo 1 — Criar o repositório no GitHub

1. Acesse https://github.com/new
2. Nome: `tigeroffers-dashboard`
3. Visibilidade: **Private** (contém dados sensíveis)
4. Clique **Create repository**

---

## Passo 2 — Subir os arquivos

Faça upload de todos os arquivos respeitando a estrutura acima.
Atenção: o `accounts.json` **não vai para o GitHub** (está no `.gitignore`).

Os arquivos que sobem:
- `.github/workflows/update_dashboard.yml`
- `scripts/baixar_exports.py`
- `scripts/analise_ano_tigeroffers.py`
- `scripts/gerar_dashboard.py`
- `scripts/dashboard_template.html`
- `data/base_pedidos.csv`
- `data/base_reembolsos.csv`
- `data/base_chargebacks.csv`
- `data/aov/white_YYYY-MM-DD.xlsx`
- `data/aov/black_YYYY-MM-DD.xlsx`
- `public/index.html` (versão atual do dashboard)
- `requirements.txt`
- `.gitignore`

---

## Passo 3 — Configurar os Secrets do GitHub

Os Secrets ficam em **Settings → Secrets and variables → Actions → New repository secret**.

### Secret 1: `BG_EMAIL`
Seu email de login na BuyGoods.

### Secret 2: `BG_PASSWORD`
Sua senha de login na BuyGoods.

### Secret 3: `BG_ACCOUNTS`
JSON com os IDs de todas as contas. Copie o conteúdo do `accounts.json`:

```json
{
  "BreathEaseX": 12457,
  "NervoLyn": 11283,
  "AudiLeaf": 11867,
  "Prostafense": 10917,
  "VisiumPro": 11292,
  "FloraNew": 12638,
  "NailsCleanPro": 12481,
  "AlphaErec": 12836,
  "MaroBrain": 11498,
  "PressureCalmX": 13135,
  "IronPulseX": 12830,
  "BoosterXT": 9373,
  "GlucoRecover": 9374,
  "AudiQuiet": 12834,
  "LipoPeak": 12621,
  "VigorLong": 10825,
  "MounjaMelt": 12888,
  "LipoBliss": 12069,
  "GlycoMelt": 13023,
  "LipoVive": 10146,
  "LipoMeltX": 12793,
  "Lipotrine": 12367,
  "JointReliveX": 12833,
  "VigorTitan": 12837,
  "BoostSteelX": 12838,
  "EyeVitalis": 12840,
  "Erectrozil": 10256,
  "RedurBurn": 11265,
  "NailsClean": 10185,
  "BoosterXTGerman": 10367,
  "Oaztem": 10843
}
```

> **NervoCalmX e NeuroRemind:** IDs ainda desconhecidos. Quando achar, adicione ao Secret.

---

## Passo 4 — Ativar o GitHub Pages

1. No repositório, vá em **Settings → Pages**
2. Source: **Deploy from a branch**
3. Branch: `main` | Folder: `/public`
4. Clique **Save**

Em ~2 minutos o dashboard estará em:
```
https://SEU-USUARIO.github.io/tigeroffers-dashboard/
```

---

## Passo 5 — Testar antes de deixar automático

1. Vá em **Actions → Atualizar Dashboard**
2. Clique **Run workflow → Run workflow**
3. Acompanhe os logs — deve completar em ~30-50 minutos (31 contas)
4. Acesse a URL do GitHub Pages e confirme que o dashboard atualizou

---

## Como o script de download funciona

Para cada conta, o script:

1. Lê a aba **Scheduled Export Jobs** e verifica o Date Range de cada tipo de relatório
2. **Se Date Range = "All"** (congela na data de criação do agendamento):
   - Vai para a seção correta (Orders Items / Customers Refunds / Customers Chargebacks)
   - Clica em **Export**
   - No modal: seleciona todos os checkboxes em todas as páginas → **Add Selected** → **Create Export**
   - Aguarda em Export Downloads até o arquivo ficar pronto (até 5 min)
   - Baixa o arquivo
3. **Se Date Range = "Last 60 Days"** (já configurado corretamente):
   - Seleciona o job agendado → clica **Run Selected**
   - Aguarda em Export Downloads → baixa

Páginas de colunas no modal:
- **Order Items:** 9 páginas
- **Customers Refunds:** 7 páginas
- **Customers Chargebacks:** 7 páginas

---

## Atualizar o AOV dos afiliados (mensal)

Quando gerar os relatórios de afiliados novos:

1. Renomeie os arquivos como:
   - `white_YYYY-MM-DD.xlsx` → período mais recente (agosto em diante)
   - `black_YYYY-MM-DD.xlsx` → período anterior (junho–julho)
2. Faça upload para a pasta `data/aov/` no repositório
3. O próximo workflow vai usar automaticamente o arquivo mais recente de cada tipo

O script pega sempre o arquivo mais recente pelo nome (`sorted()` + `[-1]`).

---

## Frequência e horário

O workflow roda **todo dia às 06:00 BRT** (09:00 UTC).

Para alterar, edite o cron em `.github/workflows/update_dashboard.yml`:
```yaml
- cron: '0 9 * * *'   # 09:00 UTC = 06:00 BRT
```
Referência: https://crontab.guru

---

## Solução de problemas

| Problema | Causa provável | Solução |
|---|---|---|
| Login falhou | Senha errada no Secret | Atualize `BG_PASSWORD` em Settings → Secrets |
| Elemento não encontrado | BuyGoods atualizou o layout | Me mande o log de erro para ajustar os seletores |
| Timeout no download | BuyGoods demorou >5 min | Normal ocasionalmente — o Action vai repetir no dia seguinte |
| Dashboard não atualizou | Nenhum arquivo baixado | Verifique os logs em Actions → último run |
| `accounts.json` subiu por acidente | `.gitignore` não foi configurado | Delete o arquivo no GitHub e revogue as credenciais da BuyGoods |

---

## O que cada script faz

| Arquivo | Função |
|---|---|
| `baixar_exports.py` | Playwright headless — loga, verifica agendados, baixa exports |
| `analise_ano_tigeroffers.py` | Processa os xlsx baixados e atualiza os 3 CSVs acumulados |
| `gerar_dashboard.py` | Lê os CSVs + AOV → monta payload JSON → injeta no template HTML |
| `dashboard_template.html` | O HTML do dashboard com `__PAYLOAD__` como marcador |
| `accounts.json` | IDs das 31 contas — **não commitar, usar o Secret BG_ACCOUNTS** |
