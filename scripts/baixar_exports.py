#!/usr/bin/env python3
"""
baixar_exports.py
-----------------
Fluxo por conta:

1. Acessa Scheduled Export Jobs
2. Para cada job (Order Items, Customers Refunds, Customers Chargebacks):
   - Se Date Range == "All"  → cria export novo do zero (modal com todas as colunas)
   - Se Date Range == "Last 60 Days" → seleciona o job → Run Selected
3. Aguarda em Export Downloads e baixa o arquivo mais recente de cada tipo

Uso local:
    BG_EMAIL=x BG_PASSWORD=y python scripts/baixar_exports.py

GitHub Actions: Secrets BG_EMAIL, BG_PASSWORD, BG_ACCOUNTS (JSON com os ids)
"""
import os, sys, json, time, shutil
from pathlib import Path
from datetime import datetime, timedelta

try:
    from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout
except ImportError:
    print("pip install playwright && playwright install chromium")
    sys.exit(1)

# ── config ────────────────────────────────────────────────────────────────────
ROOT    = Path(__file__).resolve().parent.parent
ENTRADA = ROOT / "entrada"
ENTRADA.mkdir(exist_ok=True)

EMAIL    = os.environ.get('BG_EMAIL', '')
PASSWORD = os.environ.get('BG_PASSWORD', '')

_env = os.environ.get('BG_ACCOUNTS', '')
if _env:
    CONTAS: dict[str, int] = {k: v for k, v in json.loads(_env).items() if v and not str(k).startswith('_')}
else:
    _f = ROOT / "scripts" / "accounts.json"
    if _f.exists():
        raw = json.loads(_f.read_text())
        CONTAS = {k: v for k, v in raw.items() if v and not k.startswith('_')}
    else:
        print("ERRO: configure BG_ACCOUNTS ou crie scripts/accounts.json")
        sys.exit(1)

DATA_FIM = datetime.utcnow()
DATA_INI = DATA_FIM - timedelta(days=60)
FMT_FILE = "%Y-%m-%d"
BASE     = "https://admin.buygoods.com"

# tipos de job e suas propriedades
TIPOS = [
    {"nome": "Order Items",          "sufixo": "orders",      "n_paginas": 9,
     "url_secao": "orders#items",    "keywords": ["order item"]},
    {"nome": "Customers Refunds",    "sufixo": "refunds",     "n_paginas": 7,
     "url_secao": "customers#refunds",    "keywords": ["customer refund", "refund"]},
    {"nome": "Customers Chargebacks","sufixo": "chargebacks", "n_paginas": 7,
     "url_secao": "customers#chargebacks","keywords": ["customer chargeback", "chargeback"]},
]

def log(msg): print(f"[{datetime.utcnow().strftime('%H:%M:%S')}] {msg}")


# ── login ─────────────────────────────────────────────────────────────────────
def login(page):
    log("→ Login...")
    page.goto(f"{BASE}/login", wait_until="networkidle")
    page.fill("input[type='email'], input[name='email']", EMAIL)
    page.fill("input[type='password'], input[name='password']", PASSWORD)
    page.click("button[type='submit']")
    # aguarda sair da página de login (URL muda para /{account_id} ou /shared)
    page.wait_for_function(
        "() => !window.location.href.includes('/login')",
        timeout=30_000
    )
    page.wait_for_load_state("networkidle", timeout=15_000)
    log(f"  OK — {page.url}")


# ── ler Scheduled Export Jobs e classificar por tipo de date range ────────────
def ler_scheduled_jobs(page, account_id):
    """
    Retorna dict:
      { "Order Items": "All" | "Last 60 Days" | None,
        "Customers Refunds": ...,
        "Customers Chargebacks": ... }
    None = job não encontrado (tratar como "All" para forçar criação)
    """
    url = f"{BASE}/{account_id}#exports"
    page.goto(url, wait_until="networkidle", timeout=20_000)
    page.wait_for_timeout(1_500)

    # clicar na aba "Scheduled Export Jobs"
    tab = page.locator("button:has-text('Scheduled Export Jobs'), a:has-text('Scheduled Export Jobs')").first
    if tab.count():
        tab.click()
        page.wait_for_timeout(1_000)

    result = {t["nome"]: None for t in TIPOS}

    rows = page.locator("table tbody tr").all()
    for row in rows:
        txt = row.inner_text()
        cells = [td.inner_text().strip() for td in row.locator("td").all()]
        if not cells:
            continue
        template_name = cells[0] if cells else ""
        date_range    = cells[1] if len(cells) > 1 else ""

        for tipo in TIPOS:
            if tipo["nome"].lower() in template_name.lower():
                # se já tem um registro, preferir "All" (mais restritivo)
                atual = result[tipo["nome"]]
                if atual != "All":
                    result[tipo["nome"]] = date_range
                break

    log(f"  Scheduled jobs: {result}")
    return result


# ── selecionar todas as colunas no modal e criar export ───────────────────────

def abrir_modal_export(page, n_paginas):
    """
    O botão Export na BuyGoods é um <a action="export"> dentro de um dropdown.
    Estratégia:
    1. Tentar clicar direto via JavaScript (ignora hidden)
    2. Se não abrir o modal, tentar abrir o dropdown primeiro
    """
    # Tentativa 1: clicar via JS (ignora visibilidade)
    clicou = page.evaluate("""
        () => {
            // procura <a action="export"> ou qualquer elemento com texto Export
            const candidates = [
                ...document.querySelectorAll('a[action="export"]'),
                ...document.querySelectorAll('[class*="export"]:not([class*="exports"])'),
                ...[...document.querySelectorAll('a,button')].filter(
                    el => el.textContent.trim().toLowerCase() === 'export'
                )
            ];
            for (const el of candidates) {
                if (el.offsetParent !== null || true) {
                    el.click();
                    return true;
                }
            }
            return false;
        }
    """)
    page.wait_for_timeout(1_200)

    # verificar se o modal abriu
    if page.locator("text=Available Columns").first.is_visible():
        log("    Modal aberto via JS click ✓")
        return criar_export_modal(page, n_paginas)

    # Tentativa 2: abrir dropdown primeiro
    log("    Tentando via dropdown...")
    dropdown_btn = page.locator(
        "button.dropdown-toggle, "
        "[data-toggle='dropdown'], "
        "[data-bs-toggle='dropdown']"
    ).first
    if dropdown_btn.count():
        dropdown_btn.click()
        page.wait_for_timeout(600)

    # agora clicar no item Export dentro do dropdown aberto
    export_item = page.locator(
        "a[action='export'], "
        ".dropdown-menu a:has-text('Export'), "
        ".dropdown-item:has-text('Export')"
    ).first
    if export_item.count():
        export_item.evaluate("el => el.click()")
        page.wait_for_timeout(1_200)

    if page.locator("text=Available Columns").first.is_visible():
        log("    Modal aberto via dropdown ✓")
        return criar_export_modal(page, n_paginas)

    log("    Modal não abriu")
    return False

def criar_export_modal(page, n_paginas):
    """
    Dentro do modal de Export:
    - Páginas 1..n_paginas: seleciona todos os checkboxes → Add Selected
    - Última página: clica Create Export
    Retorna True se sucesso.
    """
    # aguardar modal
    try:
        page.locator("text=Available Columns").first.wait_for(state="visible", timeout=10_000)
    except PWTimeout:
        log("    Modal não apareceu")
        return False

    for pagina in range(1, n_paginas + 1):
        log(f"    Página {pagina}/{n_paginas}...")

        # selecionar todos os checkboxes não marcados (Available Columns)
        # checkboxes podem estar hidden — usar JS click que ignora visibilidade
        page.evaluate("""
            () => {
                document.querySelectorAll('input[type=\'checkbox\']:not(:checked)').forEach(cb => {
                    cb.checked = true;
                    cb.dispatchEvent(new Event('change', { bubbles: true }));
                    cb.dispatchEvent(new Event('input', { bubbles: true }));
                });
            }
        """)
        page.wait_for_timeout(300)

        # clicar Add Selected
        add_btn = page.locator(
            "button:has-text('Add Selected'), [class*='add-selected']"
        ).first
        if add_btn.count() and add_btn.is_visible():
            add_btn.click()
            page.wait_for_timeout(500)

        # ir para próxima página (se não for a última)
        if pagina < n_paginas:
            prox_btn = page.locator(
                f"[class*='pagination'] button:has-text('{pagina + 1}'), "
                f"[class*='pagination'] a:has-text('{pagina + 1}')"
            ).first
            if prox_btn.count() and prox_btn.is_visible():
                prox_btn.click()
                page.wait_for_timeout(600)
            else:
                log(f"    Página {pagina + 1} não encontrada — criando assim mesmo")
                break

    # clicar Create Export
    create_btn = page.locator(
        "button:has-text('Create Export'), "
        "button[class*='primary']:has-text('Export')"
    ).first
    try:
        create_btn.wait_for(state="visible", timeout=8_000)
        create_btn.click()
        page.wait_for_timeout(1_500)
        log("    Create Export ✓")
        return True
    except PWTimeout:
        log("    Botão 'Create Export' não encontrado")
        return False


# ── caminho A: criar export do zero (date range "All") ───────────────────────
def criar_export_zero(page, account_id, tipo):
    """
    Navega para a seção do tipo (orders/refunds/chargebacks),
    clica Export, seleciona todas as colunas e cria.
    """
    url = f"{BASE}/{account_id}/{tipo['url_secao']}"
    log(f"  [ALL → criar do zero] {tipo['nome']} → {url}")
    try:
        page.goto(url, wait_until="networkidle", timeout=20_000)
        page.wait_for_timeout(1_500)

        # para orders: selecionar "All" no filtro de status antes de exportar
        if tipo["sufixo"] == "orders":
            status_btn = page.locator(
                "button:has-text('Accepted'), select:has(option:has-text('All'))"
            ).first
            if status_btn.count():
                status_btn.click()
                page.wait_for_timeout(400)
                all_opt = page.locator(
                    "li:has-text('All'), option[value='all'], "
                    "[role='option']:has-text('All')"
                ).first
                if all_opt.count():
                    all_opt.click()
                    page.wait_for_timeout(600)
                    log("    Status filtro: All ✓")

        # o botão Export fica dentro de um dropdown — clicar para abrir
        ok = abrir_modal_export(page, tipo["n_paginas"])
        if not ok:
            return False
        return True

    except Exception as e:
        log(f"    Erro: {e}")
        return False


# ── caminho B: rodar job agendado (date range "Last 60 Days") ─────────────────
def rodar_scheduled_job(page, account_id, tipo):
    """
    Na aba Scheduled Export Jobs, seleciona o job do tipo e clica Run Selected.
    """
    url = f"{BASE}/{account_id}#exports"
    log(f"  [Last 60 Days → Run Selected] {tipo['nome']}")
    try:
        page.goto(url, wait_until="networkidle", timeout=20_000)
        page.wait_for_timeout(1_500)

        # ir para aba de agendados
        tab = page.locator("button:has-text('Scheduled Export Jobs'), a:has-text('Scheduled Export Jobs')").first
        if tab.count():
            tab.click()
            page.wait_for_timeout(1_000)

        # encontrar a linha do job e selecionar o checkbox
        rows = page.locator("table tbody tr").all()
        marcou = False
        for row in rows:
            txt = row.inner_text().lower()
            if any(k in txt for k in [tipo["nome"].lower(), tipo["sufixo"]]):
                # verificar se é o "Last 60 Days" (não o "All")
                if "all" in txt:
                    continue  # pula — esse deveria estar no caminho A
                cb = row.locator("input[type='checkbox']").first
                if cb.count():
                    # checkbox pode estar hidden — usar JS
                    cb.evaluate("el => { el.checked = true; el.dispatchEvent(new Event('change', {bubbles:true})); }")
                    marcou = True
                    log(f"    Job selecionado ✓")
                    break

        if not marcou:
            log(f"    Job 'Last 60 Days' não encontrado — tentando Run Selected na primeira linha")
            primeiro_cb = page.locator("table tbody tr").first.locator("input[type='checkbox']").first
            if primeiro_cb.count():
                primeiro_cb.evaluate("el => { el.checked = true; el.dispatchEvent(new Event('change', {bubbles:true})); }")
                marcou = True

        if not marcou:
            log(f"    Nenhum job selecionável — pulando")
            return False

        # clicar Run Selected
        run_btn = page.locator("button:has-text('Run Selected')").first
        run_btn.wait_for(state="visible", timeout=8_000)
        run_btn.click()
        page.wait_for_timeout(1_500)
        log("    Run Selected ✓")
        return True

    except Exception as e:
        log(f"    Erro: {e}")
        return False


# ── aguardar e baixar o export mais recente ───────────────────────────────────
def aguardar_e_baixar(page, account_id, nome_conta, tipo, timeout_s=300):
    exports_url = f"{BASE}/{account_id}#exports"
    deadline = time.time() + timeout_s
    keywords = tipo["keywords"]

    while time.time() < deadline:
        page.goto(exports_url, wait_until="networkidle", timeout=20_000)
        page.wait_for_timeout(2_000)

        # garantir que está na aba Export Downloads
        dl_tab = page.locator(
            "button:has-text('Export Downloads'), a:has-text('Export Downloads')"
        ).first
        if dl_tab.count():
            dl_tab.click()
            page.wait_for_timeout(1_000)

        rows = page.locator("table tbody tr").all()
        for row in rows[:10]:
            txt = row.inner_text().lower()
            if not any(k in txt for k in keywords):
                continue

            # verificar se o arquivo já está pronto (tem link de download)
            link = row.locator(
                "a[href*='download'], a[download], "
                "a:has-text('Download'), "
                "a[href*='.xlsx'], a[href*='.csv']"
            ).first

            if not link.count() or not link.is_visible():
                log(f"    Ainda processando ({tipo['sufixo']})...")
                break

            # baixar
            try:
                with page.expect_download(timeout=90_000) as dl_info:
                    link.click()
                dl = dl_info.value
                dest = ENTRADA / f"{nome_conta}_{tipo['sufixo']}_{DATA_FIM.strftime(FMT_FILE)}.xlsx"
                if dest.exists():
                    dest.unlink()
                dl.save_as(dest)
                log(f"    ✓ {dest.name}")
                return dest
            except Exception as e:
                log(f"    Erro ao baixar: {e}")
                return None

        restante = int(deadline - time.time())
        if restante > 0:
            log(f"    Aguardando... ({restante}s restantes)")
            time.sleep(12)

    log(f"    ✗ Timeout — {tipo['sufixo']} não ficou pronto em {timeout_s}s")
    return None


# ── processar uma conta ───────────────────────────────────────────────────────
def processar_conta(page, nome_conta, account_id):
    log(f"\n── {nome_conta} (id={account_id}) ──")

    # 1. ler os scheduled jobs e classificar
    scheduled = ler_scheduled_jobs(page, account_id)

    resultados = {}
    for tipo in TIPOS:
        nome_tipo = tipo["nome"]
        date_range = scheduled.get(nome_tipo)  # "All", "Last 60 Days" ou None

        log(f"\n  {nome_tipo}: date_range={date_range!r}")

        # decidir o caminho
        if date_range and "60" in date_range.lower():
            # Caminho B: Last 60 Days → Run Selected
            ok = rodar_scheduled_job(page, account_id, tipo)
        else:
            # Caminho A: "All", data fixa antiga, ou não encontrado → criar do zero
            ok = criar_export_zero(page, account_id, tipo)

        if not ok:
            log(f"  Falha ao disparar {nome_tipo} — pulando download")
            resultados[tipo["sufixo"]] = None
            continue

        # aguardar e baixar
        dest = aguardar_e_baixar(page, account_id, nome_conta, tipo)
        resultados[tipo["sufixo"]] = dest

    return resultados


# ── main ──────────────────────────────────────────────────────────────────────
def main():
    if not EMAIL or not PASSWORD:
        print("ERRO: BG_EMAIL e BG_PASSWORD não definidos")
        sys.exit(1)

    log(f"BuyGoods download | {DATA_INI.strftime(FMT_FILE)} → {DATA_FIM.strftime(FMT_FILE)}")
    log(f"{len(CONTAS)} contas configuradas")

    ok_total, erro_total = [], []

    with sync_playwright() as pw:
        browser = pw.chromium.launch(
            headless=True,
            args=["--no-sandbox", "--disable-setuid-sandbox", "--disable-dev-shm-usage"]
        )
        ctx  = browser.new_context(accept_downloads=True)
        page = ctx.new_page()
        page.set_default_timeout(30_000)

        login(page)

        for nome_conta, account_id in CONTAS.items():
            resultados = processar_conta(page, nome_conta, account_id)
            for sufixo, dest in resultados.items():
                chave = f"{nome_conta}/{sufixo}"
                (ok_total if dest else erro_total).append(chave)

        browser.close()

    log(f"\n✓ Concluído: {len(ok_total)} OK | {len(erro_total)} erro(s)")
    if erro_total:
        log(f"  Falhas: {erro_total}")

    (ENTRADA / f"log_{DATA_FIM.strftime(FMT_FILE)}.json").write_text(
        json.dumps({"ok": ok_total, "erro": erro_total,
                    "data": DATA_FIM.strftime(FMT_FILE)}, indent=2)
    )
    sys.exit(0 if ok_total else 1)


if __name__ == "__main__":
    main()
