#!/usr/bin/env python3
"""
baixar_exports.py
-----------------
Baixa os exports da BuyGoods (Order Items, Customer Refunds, Customer Chargebacks)
para todas as contas configuradas.

Fluxo por conta e por tipo:
  - Scheduled job "Last 60 Days"  → seleciona o job → Run Selected
  - Qualquer outro date range     → cria export do zero (modal de colunas)
  - Depois: aguarda em Exports e baixa o arquivo pronto

MODO DIAGNÓSTICO
  Rode com DEBUG_DOM=1 para despejar todos os botões/links da página quando
  um clique falhar. Use isso para me mandar o HTML real e eu ajustar os
  seletores com precisão em vez de adivinhar.

    DEBUG_DOM=1 BG_EMAIL=... BG_PASSWORD=... python scripts/baixar_exports.py
"""
import os, sys, json, time
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

EMAIL     = os.environ.get('BG_EMAIL', '')
PASSWORD  = os.environ.get('BG_PASSWORD', '')
DEBUG_DOM = os.environ.get('DEBUG_DOM', '') == '1'
GRAVAR    = os.environ.get('GRAVAR', '') == '1'      # vídeo + trace
SO_UMA    = os.environ.get('SO_UMA', '')             # testar 1 conta só
VIDEO_DIR = ROOT / "diagnostico" / "video"
TRACE_DIR = ROOT / "diagnostico"

_env = os.environ.get('BG_ACCOUNTS', '')
if _env:
    CONTAS = {k: v for k, v in json.loads(_env).items()
              if v and not str(k).startswith('_')}
else:
    _f = ROOT / "scripts" / "accounts.json"
    if _f.exists():
        CONTAS = {k: v for k, v in json.loads(_f.read_text()).items()
                  if v and not str(k).startswith('_')}
    else:
        print("ERRO: defina BG_ACCOUNTS ou crie scripts/accounts.json")
        sys.exit(1)

DATA_FIM = datetime.utcnow()
DATA_INI = DATA_FIM - timedelta(days=60)
FMT_FILE = "%Y-%m-%d"
BASE     = "https://admin.buygoods.com"

TIPOS = [
    {"nome": "Order Items",           "sufixo": "orders",      "n_paginas": 9,
     "url": "orders#orders-items",       "keywords": ["order item"]},
    {"nome": "Customers Refunds",     "sufixo": "refunds",     "n_paginas": 7,
     "url": "customers#refunds",         "keywords": ["customer refund", "refund"]},
    {"nome": "Customers Chargebacks", "sufixo": "chargebacks", "n_paginas": 7,
     "url": "customers#chargebacks",     "keywords": ["customer chargeback", "chargeback"]},
]

def log(m): print(f"[{datetime.utcnow().strftime('%H:%M:%S')}] {m}", flush=True)


# ── helpers de clique robusto ─────────────────────────────────────────────────
JS_CLICK_BY_TEXT = """
(texts) => {
    // A BuyGoods duplica a barra de ações no DOM (layout responsivo): existe
    // uma cópia dentro de .d-none e outra visível. Clicar na oculta não
    // dispara handler nenhum, então o elemento VISÍVEL tem prioridade.
    const norm = s => (s || '').replace(/\\s+/g, ' ').trim().toLowerCase();
    const visivel = el => el.offsetParent !== null &&
                          el.getClientRects().length > 0;
    const wanted = texts.map(norm);
    const sel = 'button, a, input[type=submit], input[type=button], [role=button], .btn';
    const els = [...document.querySelectorAll(sel)];
    const casa = el => {
        const t = norm(el.textContent) || norm(el.value) ||
                  norm(el.getAttribute('aria-label'));
        return wanted.some(w => t === w || t.includes(w)) ? t : null;
    };
    for (const el of els) {                       // 1ª passada: só visíveis
        const t = casa(el);
        if (t && visivel(el)) { el.scrollIntoView({block:'center'}); el.click(); return t; }
    }
    for (const el of els) {                       // 2ª passada: qualquer um
        const t = casa(el);
        if (t) { el.scrollIntoView({block:'center'}); el.click(); return t + ' (oculto)'; }
    }
    return null;
}
"""

JS_CLICK_BY_ATTR = """
(selector) => {
    const visivel = el => el.offsetParent !== null && el.getClientRects().length > 0;
    const els = [...document.querySelectorAll(selector)];
    if (!els.length) return false;
    const alvo = els.find(visivel) || els[0];
    alvo.scrollIntoView({block: 'center'});
    alvo.click();
    return visivel(alvo) ? 'visivel' : 'oculto';
}
"""

JS_DUMP = """
() => {
    const hiddenAncestor = el => {
        let n = el;
        while (n && n !== document.body) {
            const st = getComputedStyle(n);
            if (st.display === 'none' || st.visibility === 'hidden')
                return (n.className || n.tagName).toString().slice(0, 50);
            n = n.parentElement;
        }
        return '';
    };
    const sel = 'button, a, input[type=submit], [role=button], .btn';
    return [...document.querySelectorAll(sel)].map(el => ({
        tag: el.tagName,
        text: (el.textContent || '').replace(/\\s+/g, ' ').trim().slice(0, 34),
        cls: (el.className || '').toString().slice(0, 40),
        attrs: [...el.attributes].map(a => a.name + '=' + a.value).join(' ').slice(0, 60),
        visible: el.offsetParent !== null,
        hiddenBy: hiddenAncestor(el)
    })).filter(o => o.text || o.attrs.includes('export')).slice(0, 40);
}
"""

def click_text(page, *texts):
    """Clica no primeiro elemento cujo texto bate. Retorna o texto clicado ou None."""
    try:
        return page.evaluate(JS_CLICK_BY_TEXT, list(texts))
    except Exception as e:
        log(f"    click_text erro: {e}")
        return None

def dump_dom(page, contexto):
    """Despeja os botões/links da página para diagnóstico."""
    if not DEBUG_DOM:
        return
    try:
        els = page.evaluate(JS_DUMP)
        log(f"    ── DOM DUMP [{contexto}] — {len(els)} elementos ──")
        for e in els[:35]:
            log(f"      {e['tag']:7} vis={str(e['visible']):5} text={e['text']!r:36} "
                f"attrs={e['attrs'][:45]!r} hiddenBy={e['hiddenBy'][:30]!r}")
        log("    ── fim do dump ──")
    except Exception as e:
        log(f"    dump falhou: {e}")



JS_ATIVAR_ABA = """
(hash) => {
    // A BuyGoods usa abas Bootstrap: o painel só fica visível quando o link
    // da aba é clicado. Navegar pelo hash carrega o DOM mas deixa o painel
    // com display:none, e todo clique dentro dele vira no-op.
    const alvo = hash.replace('#', '');
    const link = document.querySelector(
        `a[href="#${alvo}"], a[href$="#${alvo}"], [data-bs-target="#${alvo}"], [data-target="#${alvo}"]`
    );
    if (link) { link.click(); return 'link'; }

    // fallback: ativar o painel na mão
    const pane = document.getElementById(alvo);
    if (pane) {
        document.querySelectorAll('.tab-pane.active, .tab-pane.show')
                .forEach(p => p.classList.remove('active', 'show'));
        pane.classList.add('active', 'show');
        pane.style.display = '';
        return 'pane';
    }
    return null;
}
"""


def ativar_aba(page, url_secao):
    """Ativa a aba correspondente ao hash da URL e espera o painel aparecer."""
    if '#' not in url_secao:
        return
    hash_ = '#' + url_secao.split('#', 1)[1]
    for tentativa in range(3):
        via = page.evaluate(JS_ATIVAR_ABA, hash_)
        page.wait_for_timeout(1_200)
        # confirmar que algo do painel ficou visível
        visivel = page.evaluate("""
            () => {
                const els = [...document.querySelectorAll('table, .card, a[action=export]')];
                return els.some(e => e.offsetParent !== null);
            }
        """)
        if visivel:
            log(f"    Aba ativada via {via} ✓")
            return
    log("    Aviso: painel da aba continua oculto")


# ── login ─────────────────────────────────────────────────────────────────────
def login(page):
    log("→ Login...")
    page.goto(f"{BASE}/login", wait_until="networkidle")
    page.fill("input[type='email'], input[name='email']", EMAIL)
    page.fill("input[type='password'], input[name='password']", PASSWORD)
    page.click("button[type='submit']")
    page.wait_for_function("() => !location.href.includes('/login')", timeout=30_000)
    page.wait_for_load_state("networkidle", timeout=15_000)
    log(f"  OK — {page.url}")


# ── ler scheduled jobs ────────────────────────────────────────────────────────
def ler_scheduled(page, account_id):
    page.goto(f"{BASE}/{account_id}#exports", wait_until="networkidle", timeout=20_000)
    page.wait_for_timeout(1_500)
    click_text(page, "Scheduled Export Jobs")
    page.wait_for_timeout(1_200)

    result = {t["nome"]: None for t in TIPOS}
    for row in page.locator("table tbody tr").all():
        cells = [td.inner_text().strip() for td in row.locator("td").all()]
        if len(cells) < 2:
            continue
        template, date_range = cells[0], cells[1]
        for t in TIPOS:
            if t["nome"].lower() in template.lower():
                if result[t["nome"]] != "All":
                    result[t["nome"]] = date_range
                break
    log(f"  Scheduled: {result}")
    return result


# ── modal de colunas ──────────────────────────────────────────────────────────
def modal_visivel(page):
    for sel in ["text=Available Columns", "text=Selected Columns",
                "text=Report Title", "button:has-text('Create Export')",
                ".modal.show", "[role=dialog]"]:
        try:
            if page.locator(sel).first.is_visible(timeout=1_000):
                return True
        except Exception:
            continue
    return False


def preencher_modal(page, n_paginas):
    """
    Para cada página de colunas: marca o checkbox mestre do cabeçalho
    "AVAILABLE COLUMNS" (que seleciona todas as da página), clica
    "Add Selected" e avança. No fim, clica "Create Export".
    """
    for pagina in range(1, n_paginas + 1):
        marcou = page.evaluate("""
            () => {
                // 1. checkbox mestre no cabeçalho AVAILABLE COLUMNS
                const cabecalhos = [...document.querySelectorAll('*')].filter(
                    e => (e.textContent || '').trim().toUpperCase() === 'AVAILABLE COLUMNS'
                );
                for (const cab of cabecalhos) {
                    const bloco = cab.closest('div');
                    const mestre = bloco && bloco.querySelector('input[type=checkbox]');
                    if (mestre && !mestre.checked) {
                        mestre.click();
                        return 'mestre';
                    }
                }
                // 2. fallback: marcar um a um dentro do painel de disponíveis
                let n = 0;
                document.querySelectorAll('input[type=checkbox]').forEach(cb => {
                    if (!cb.checked) { cb.click(); n++; }
                });
                return n ? ('individual:' + n) : 'nada';
            }
        """)
        page.wait_for_timeout(450)

        add = click_text(page, "Add Selected")
        page.wait_for_timeout(650)
        log(f"    pág {pagina}/{n_paginas}: {marcou} | Add Selected={bool(add)}")

        if pagina < n_paginas:
            foi = page.evaluate("""
                (proxima) => {
                    const links = [...document.querySelectorAll(
                        '.pagination a, .pagination button, .page-link, li a, li button')];
                    const alvo = links.find(l => (l.textContent || '').trim() === String(proxima));
                    if (!alvo) return false;
                    alvo.click();
                    return true;
                }
            """, pagina + 1)
            if not foi:
                log(f"    página {pagina + 1} não encontrada — encerrando seleção")
                break
            page.wait_for_timeout(700)

    if click_text(page, "Create Export", "Create"):
        page.wait_for_timeout(1_500)
        log("    Create Export ✓")
        return True

    dump_dom(page, "modal sem Create Export")
    log("    ✗ Create Export não encontrado")
    return False


def abrir_modal(page, n_paginas):
    """Abre o modal de export tentando várias estratégias."""
    estrategias = [
        ("a[action=export] visível",
         lambda: page.evaluate(JS_CLICK_BY_ATTR, "a[action='export'], [action='export']")),
        ("texto Export",          lambda: click_text(page, "Export")),
        ("dropdown + Export",     lambda: (
            page.evaluate(JS_CLICK_BY_ATTR,
                          "[data-bs-toggle='dropdown'], [data-toggle='dropdown'], .dropdown-toggle"),
            page.wait_for_timeout(700),
            click_text(page, "Export"))[-1]),
    ]

    for nome, fn in estrategias:
        try:
            fn()
        except Exception as e:
            log(f"    estratégia '{nome}' erro: {e}")
            continue
        page.wait_for_timeout(1_500)
        if modal_visivel(page):
            log(f"    Modal aberto via {nome} ✓")
            return preencher_modal(page, n_paginas)

    dump_dom(page, "modal não abriu")
    log("    ✗ Modal não abriu por nenhuma estratégia")
    return False


# ── caminho A: criar export do zero ───────────────────────────────────────────
def criar_do_zero(page, account_id, tipo):
    url = f"{BASE}/{account_id}/{tipo['url']}"
    log(f"  [criar do zero] {tipo['nome']}")
    try:
        page.goto(url, wait_until="networkidle", timeout=20_000)
        page.wait_for_timeout(1_800)
        ativar_aba(page, tipo["url"])

        # a tabela é montada por AJAX; enquanto não termina, o container fica
        # com d-none e nenhum clique dentro dele funciona
        try:
            page.wait_for_function("""
                () => {
                    const linhas = document.querySelectorAll('table tbody tr');
                    if (!linhas.length) return false;
                    return [...linhas].some(l => l.offsetParent !== null);
                }
            """, timeout=20_000)
            log("    Tabela renderizada e visível ✓")
        except PWTimeout:
            log("    ⚠ Tabela não ficou visível em 20s")
            dump_dom(page, "tabela invisivel")

        if tipo["sufixo"] == "orders":
            trocou = page.evaluate("""
                () => {
                    // o filtro de status é um dropdown cujo gatilho mostra "Accepted"
                    const trig = [...document.querySelectorAll('button, a, .dropdown-toggle')]
                        .find(e => (e.textContent || '').trim().startsWith('Accepted'));
                    if (!trig) return 'gatilho nao encontrado';
                    trig.click();
                    return 'aberto';
                }
            """)
            page.wait_for_timeout(700)
            escolheu = page.evaluate("""
                () => {
                    // "All" do filtro de status — precisa ser o VISÍVEL, senão
                    // pega o botão "All" do seletor de Colunas, que fica oculto
                    const visivel = el => el.offsetParent !== null &&
                                          el.getClientRects().length > 0;
                    const opts = [...document.querySelectorAll(
                        '.dropdown-menu a, .dropdown-item, li a, [role=option], option')];
                    const cands = opts.filter(o => (o.textContent || '').trim() === 'All');
                    const alvo = cands.find(visivel);
                    if (!alvo) return false;
                    alvo.click();
                    return true;
                }
            """)
            page.wait_for_timeout(1_200)
            log(f"    Status → All: {escolheu} ({trocou})")

        return abrir_modal(page, tipo["n_paginas"])
    except Exception as e:
        log(f"    Erro: {e}")
        return False


# ── caminho B: rodar job agendado ─────────────────────────────────────────────
def rodar_agendado(page, account_id, tipo):
    log(f"  [Run Selected] {tipo['nome']}")
    try:
        page.goto(f"{BASE}/{account_id}#exports", wait_until="networkidle", timeout=20_000)
        page.wait_for_timeout(1_500)
        click_text(page, "Scheduled Export Jobs")
        page.wait_for_timeout(1_200)

        # marcar o checkbox da linha certa via JS
        marcou = page.evaluate("""
            (args) => {
                const [nome, sufixo] = args;
                const rows = [...document.querySelectorAll('table tbody tr')];
                for (const row of rows) {
                    const txt = row.innerText.toLowerCase();
                    if (!txt.includes(nome.toLowerCase()) && !txt.includes(sufixo)) continue;
                    if (txt.includes('all')) continue;          // esse vai pelo caminho A
                    const cb = row.querySelector('input[type=checkbox]');
                    if (cb) {
                        cb.checked = true;
                        cb.dispatchEvent(new Event('change', {bubbles: true}));
                        cb.dispatchEvent(new Event('click',  {bubbles: true}));
                        return true;
                    }
                }
                return false;
            }
        """, [tipo["nome"], tipo["sufixo"]])

        if not marcou:
            log("    Nenhuma linha correspondente para marcar")
            dump_dom(page, "linha do job não encontrada")
            return False
        log("    Job selecionado ✓")
        page.wait_for_timeout(600)

        if click_text(page, "Run Selected", "Run"):
            page.wait_for_timeout(1_500)
            log("    Run Selected ✓")
            return True

        dump_dom(page, "Run Selected não encontrado")
        log("    ✗ Run Selected não encontrado")
        return False

    except Exception as e:
        log(f"    Erro: {e}")
        return False


# ── aguardar e baixar ─────────────────────────────────────────────────────────
def aguardar_e_baixar(page, account_id, nome_conta, tipo, timeout_s=300):
    deadline = time.time() + timeout_s
    while time.time() < deadline:
        page.goto(f"{BASE}/{account_id}#exports", wait_until="networkidle", timeout=20_000)
        page.wait_for_timeout(2_000)
        click_text(page, "Export Downloads")
        page.wait_for_timeout(1_200)

        for row in page.locator("table tbody tr").all()[:10]:
            txt = row.inner_text().lower()
            if not any(k in txt for k in tipo["keywords"]):
                continue
            link = row.locator(
                "a[href*='download'], a[download], a:has-text('Download'), "
                "a[href*='.xlsx'], a[href*='.csv']"
            ).first
            if not link.count():
                break
            try:
                with page.expect_download(timeout=90_000) as dl_info:
                    link.evaluate("el => el.click()")
                dl   = dl_info.value
                dest = ENTRADA / f"{nome_conta}_{tipo['sufixo']}_{DATA_FIM.strftime(FMT_FILE)}.xlsx"
                if dest.exists():
                    dest.unlink()
                dl.save_as(dest)
                log(f"    ✓ {dest.name}")
                return dest
            except Exception as e:
                log(f"    Erro no download: {e}")
                return None

        log(f"    Aguardando processamento... ({int(deadline - time.time())}s)")
        time.sleep(12)

    log(f"    ✗ Timeout aguardando {tipo['sufixo']}")
    return None


# ── main ──────────────────────────────────────────────────────────────────────
def main():
    if not EMAIL or not PASSWORD:
        print("ERRO: BG_EMAIL / BG_PASSWORD não definidos")
        sys.exit(1)

    log(f"BuyGoods | {DATA_INI.strftime(FMT_FILE)} → {DATA_FIM.strftime(FMT_FILE)} | "
        f"{len(CONTAS)} contas | DEBUG_DOM={DEBUG_DOM}")

    ok, erro, pendentes = [], [], []

    with sync_playwright() as pw:
        browser = pw.chromium.launch(
            headless=True,
            args=["--no-sandbox", "--disable-setuid-sandbox", "--disable-dev-shm-usage"]
        )
        # viewport grande: a BuyGoods esconde a barra de ações em telas estreitas,
        # que é o motivo de vários botões aparecerem como "not visible"
        ctx_args = {
            "accept_downloads": True,
            "viewport": {"width": 1920, "height": 1080},
        }
        if GRAVAR:
            VIDEO_DIR.mkdir(parents=True, exist_ok=True)
            ctx_args["record_video_dir"]  = str(VIDEO_DIR)
            ctx_args["record_video_size"] = {"width": 1280, "height": 720}
        ctx = browser.new_context(**ctx_args)

        if GRAVAR:
            # trace = timeline navegável com screenshot e DOM de cada passo
            ctx.tracing.start(screenshots=True, snapshots=True, sources=True)

        page = ctx.new_page()
        page.set_default_timeout(30_000)

        login(page)

        itens = list(CONTAS.items())
        if SO_UMA:
            itens = [(k, v) for k, v in itens if k.lower() == SO_UMA.lower()] or itens[:1]
            log(f"  MODO DIAGNÓSTICO: só a conta {itens[0][0]}")

        for nome_conta, account_id in itens:
            log(f"\n── {nome_conta} (id={account_id}) ──")
            sched = ler_scheduled(page, account_id)

            for tipo in TIPOS:
                dr = sched.get(tipo["nome"])
                log(f"\n  {tipo['nome']}: date_range={dr!r}")
                if dr and "60" in dr.lower():
                    disparou = rodar_agendado(page, account_id, tipo)
                else:
                    disparou = criar_do_zero(page, account_id, tipo)
                    if not disparou:
                        pendentes.append(f"{nome_conta}/{tipo['sufixo']} (date_range={dr})")

                dest = aguardar_e_baixar(page, account_id, nome_conta, tipo) if disparou else None
                (ok if dest else erro).append(f"{nome_conta}/{tipo['sufixo']}")

        if GRAVAR:
            TRACE_DIR.mkdir(parents=True, exist_ok=True)
            trace_path = TRACE_DIR / "trace.zip"
            ctx.tracing.stop(path=str(trace_path))
            log(f"  Trace salvo em {trace_path}")

        ctx.close()          # necessário para o vídeo ser finalizado
        browser.close()

        if GRAVAR:
            vids = list(VIDEO_DIR.glob("*.webm"))
            log(f"  {len(vids)} vídeo(s) em {VIDEO_DIR}")

    log(f"\n✓ {len(ok)} baixados | {len(erro)} falha(s) | {len(pendentes)} pendente(s) de ajuste")
    if erro:
        log(f"  Falhas: {erro[:20]}{' ...' if len(erro) > 20 else ''}")
    if pendentes:
        log("\n  ─── AJUSTE ESTES AGENDAMENTOS UMA ÚNICA VEZ NA BUYGOODS ───")
        log("  Scheduled Export Jobs → editar o job → Date Range = Last 60 Days")
        for p in pendentes:
            log(f"    • {p}")

    (ENTRADA / f"log_{DATA_FIM.strftime(FMT_FILE)}.json").write_text(
        json.dumps({"ok": ok, "erro": erro, "pendentes_ajuste_manual": pendentes},
                   indent=2, ensure_ascii=False))

    sys.exit(0 if ok else 1)


if __name__ == "__main__":
    main()
