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
import os, sys, json, time, re
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
STORAGE   = ROOT / "entrada" / ".sessao.json"        # cookies salvos
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
     "hash": "orders-items",       "keywords": ["order item"]},
    {"nome": "Customers Refunds",     "sufixo": "refunds",     "n_paginas": 7,
     "hash": "customers-refunds",         "keywords": ["customer refund", "refund"]},
    {"nome": "Customers Chargebacks", "sufixo": "chargebacks", "n_paginas": 7,
     "hash": "customers-chargebacks",     "keywords": ["customer chargeback", "chargeback"]},
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


def ativar_aba(page, hash_alvo):
    """
    Ativa a aba do hash e espera o painel aparecer.

    A navegação é sempre admin.buygoods.com/{account_id}#{hash} — o hash
    fica na raiz da conta, sem segmento de caminho. Se o painel não for
    ativado, ele fica com display:none e nenhum clique dentro dele
    dispara handler.
    """
    for hash_ in [f'#{hash_alvo}']:
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
            log(f"    Aba ativada via {via} ({hash_}) ✓")
            return
    log(f"    Aviso: painel #{hash_alvo} continua oculto")


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
    """
    Lê a aba Scheduled Export Jobs.

    Duas armadilhas que já nos pegaram:
      1. Ler "table tbody tr" pega a tabela de Export Downloads, que tem
         linhas antigas com date range fixo — foi assim que o NervoLyn
         apareceu como 'January 01, 2026 - August 26, 2026' quando o
         agendamento dele é 'Last 60 Days'. As caixas dos jobs têm prefixo
         próprio, então filtramos por ele.
      2. O nome do template varia: 'Order Items' vs 'Orders Items',
         'Customers Refunds' vs 'Customer Refunds'. Casamos por regex.
    """
    page.goto(f"{BASE}/{account_id}#exports", wait_until="networkidle", timeout=20_000)
    page.wait_for_timeout(1_500)
    click_text(page, "Scheduled Export Jobs")
    page.wait_for_timeout(1_500)

    result = {t["nome"]: None for t in TIPOS}
    linhas = []
    for _ in range(8):                       # as linhas não renderizam juntas
        linhas = _linhas_com_caixa(page, CAIXA_JOB)
        if len(linhas) >= 2:
            break
        page.wait_for_timeout(700)

    for cid, texto in linhas:
        for t in TIPOS:
            if re.search(KEYWORDS[t["sufixo"]], texto, re.I):
                # extrai o date range da linha (2a coluna)
                m = re.search(r'(All|Last\s+\d+\s+Days|[A-Z][a-z]+ \d{1,2}, \d{4}.*)',
                              texto)
                valor = m.group(1).strip() if m else texto
                if result[t["nome"]] is None:
                    result[t["nome"]] = valor
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



def esperar_tabela_pronta(page, timeout_ms=30_000):
    """
    Espera a tabela terminar de carregar.

    O trace mostrou o script mexendo no filtro enquanto o spinner ainda
    rodava: a tabela existe, mas o widget de status ainda não foi montado,
    então o gatilho "Accepted" não responde ao clique.
    Critério: sem spinner visível E (com linhas de dados na tabela OU a
    tabela declarou explicitamente que está vazia).

    Contas de baixo volume (NailsClean, BoosterXTGerman, EyeVitalis, etc.)
    batiam o timeout de 30s em TODA aba, sempre — não porque a tabela
    estivesse lenta, mas porque ela terminava de carregar vazia e a
    condição só aceitava "tem linha com dado". Isso empurrava runs de
    31 contas de ~35min pra mais de 1h. DataTables marca vazio com uma
    linha única <td class="dataTables_empty">, ou o texto
    "No data available in table" nem chega a existir se a chamada AJAX
    devolveu 0 registros — os dois casos agora contam como "pronta".
    """
    try:
        page.wait_for_function("""
            () => {
                const visivel = el => el.offsetParent !== null &&
                                      el.getClientRects().length > 0;

                // 1. nenhum spinner/loader visível
                const spinners = [...document.querySelectorAll(
                    '.spinner-border, .spinner-grow, .loading, .loader, ' +
                    '[class*="spinner"], [class*="loading"]')];
                if (spinners.some(visivel)) return false;

                // 2a. tabela com linhas de dados visíveis
                const linhas = [...document.querySelectorAll('table tbody tr')];
                const comDados = linhas.filter(l =>
                    visivel(l) && l.querySelectorAll('td').length > 1);
                if (comDados.length > 0) return true;

                // 2b. tabela terminou de carregar e declarou vazia (DataTables)
                const vazia = linhas.some(l => visivel(l) &&
                    l.querySelector('.dataTables_empty, [class*="empty"]'));
                if (vazia) return true;
                const infoTexto = document.querySelector('.dataTables_info');
                if (infoTexto && visivel(infoTexto) &&
                    /\\b0\\b.*entries|showing 0/i.test(infoTexto.textContent)) return true;

                return false;
            }
        """, timeout=timeout_ms)
        page.wait_for_timeout(800)   # respiro para o JS ligar os handlers
        return True
    except PWTimeout:
        log("    ⚠ tabela não terminou de carregar no tempo esperado")
        return False


def definir_status_all(page):
    """
    Troca o filtro de status de "Accepted" para "All" na tela de Orders Items.

    O gatilho não é um <button>: pode ser um <select> nativo ou um dropdown
    customizado montado com <div>/<span>. Procurar só por 'button, a' devolve
    "gatilho nao encontrado" — foi o que acontecia antes.

    Sem esse filtro o export sai só com pedidos aceitos e perde os refunded,
    declined e canceled, que são justamente o que a análise precisa.
    """
    # ── caso 1: <select> nativo ─────────────────────────────────────────────
    via_select = page.evaluate("""
        () => {
            for (const sel of document.querySelectorAll('select')) {
                const opts = [...sel.options].map(o => o.textContent.trim());
                if (!opts.includes('All')) continue;
                const alvo = [...sel.options].find(o => o.textContent.trim() === 'All');
                sel.value = alvo.value;
                sel.dispatchEvent(new Event('change', {bubbles: true}));
                sel.dispatchEvent(new Event('input',  {bubbles: true}));
                return true;
            }
            return false;
        }
    """)
    if via_select:
        page.wait_for_timeout(1_500)
        log("    Status → All (select nativo) ✓")
        return True

    # ── caso 2: dropdown customizado ────────────────────────────────────────
    abriu = page.evaluate("""
        () => {
            const visivel = el => el.offsetParent !== null &&
                                  el.getClientRects().length > 0;
            // qualquer elemento visível cujo texto seja exatamente "Accepted"
            const cands = [...document.querySelectorAll('*')].filter(el => {
                if (el.children.length > 2) return false;      // evita containers
                const t = (el.textContent || '').trim();
                return t === 'Accepted' && visivel(el);
            });
            if (!cands.length) return null;
            // clicar no elemento clicável mais próximo
            const el = cands[0];
            const clicavel = el.closest('button, a, [role=button], .dropdown-toggle, select') || el;
            clicavel.scrollIntoView({block: 'center'});
            clicavel.click();
            return clicavel.tagName + '.' + (clicavel.className || '').toString().slice(0, 30);
        }
    """)
    if not abriu:
        log("    ⚠ gatilho 'Accepted' não encontrado")
        dump_dom(page, "filtro de status")
        return False

    page.wait_for_timeout(800)

    escolheu = page.evaluate("""
        () => {
            const visivel = el => el.offsetParent !== null &&
                                  el.getClientRects().length > 0;
            // "All" dentro do menu que acabou de abrir — precisa ser visível,
            // senão pega o botão "All" do seletor de Colunas, que fica oculto
            const opts = [...document.querySelectorAll(
                '.dropdown-menu *, .dropdown-item, li, [role=option], option, a, button')];
            const alvo = opts.find(o =>
                (o.textContent || '').trim() === 'All' && visivel(o));
            if (!alvo) return false;
            alvo.click();
            return true;
        }
    """)
    page.wait_for_timeout(1_800)

    if not escolheu:
        log("    ⚠ opção 'All' não encontrada no menu aberto")
        dump_dom(page, "menu de status aberto")
        return False

    # ── confirmar que realmente mudou ───────────────────────────────────────
    agora = page.evaluate("""
        () => {
            const visivel = el => el.offsetParent !== null &&
                                  el.getClientRects().length > 0;
            const el = [...document.querySelectorAll('*')].find(e =>
                e.children.length <= 2 &&
                ['All', 'Accepted'].includes((e.textContent || '').trim()) &&
                visivel(e));
            return el ? (el.textContent || '').trim() : '?';
        }
    """)
    ok = (agora == 'All')
    log(f"    Status → All: {ok} (gatilho agora mostra {agora!r}) [{abriu}]")
    return ok


def definir_status_all_retry(page, tentativas=3):
    """O widget às vezes só responde depois que o AJAX assenta — tenta de novo."""
    for i in range(1, tentativas + 1):
        if definir_status_all(page):
            return True
        if i < tentativas:
            log(f"    tentativa {i}/{tentativas} falhou — aguardando e repetindo")
            esperar_tabela_pronta(page, timeout_ms=15_000)
            page.wait_for_timeout(1_500)
    return False


# ── caminho A: criar export do zero ───────────────────────────────────────────
def criar_do_zero(page, account_id, tipo):
    url = f"{BASE}/{account_id}#{tipo['hash']}"
    log(f"  [criar do zero] {tipo['nome']} → {url}")
    try:
        page.goto(url, wait_until="networkidle", timeout=20_000)
        page.wait_for_timeout(1_800)
        ativar_aba(page, tipo["hash"])

        # a tabela é montada por AJAX; enquanto não termina, o container fica
        # com d-none e nenhum clique dentro dele funciona
        if esperar_tabela_pronta(page):
            log("    Tabela carregada ✓")
        else:
            dump_dom(page, "tabela nao carregou")

        if tipo["sufixo"] == "orders":
            esperar_tabela_pronta(page)
            if not definir_status_all_retry(page):
                log("    ⚠ seguindo mesmo assim — export pode vir só com Accepted")
            # a troca do filtro dispara novo AJAX — esperar terminar,
            # senão o modal abre com o estado antigo
            esperar_tabela_pronta(page)

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
        # casar pela PALAVRA-CHAVE, não pelo nome completo: o template muda de
        # nome entre contas ('Order Items' vs 'Orders Items', 'Customers
        # Refunds' vs 'Customer Refunds'). E as linhas não renderizam todas
        # juntas, então tentamos algumas vezes antes de desistir.
        padrao = KEYWORDS.get(tipo["sufixo"], tipo["sufixo"])
        marcou = False
        for _ in range(8):
            marcou = page.evaluate("""
                (args) => {
                    const [prefixo, padraoStr] = args;
                    const re_ = new RegExp(padraoStr, 'i');
                    const caixas = [...document.querySelectorAll(`input[id^="${prefixo}"]`)];
                    for (const cx of caixas) {
                        const tr = cx.closest('tr');
                        const txt = tr ? tr.innerText : '';
                        if (!re_.test(txt)) continue;
                        if (/\ball\b/i.test(txt)) continue;
                        cx.checked = true;
                        cx.dispatchEvent(new Event('change', {bubbles: true}));
                        cx.dispatchEvent(new Event('click',  {bubbles: true}));
                        return true;
                    }
                    return false;
                }
            """, [CAIXA_JOB, padrao])
            if marcou:
                break
            page.wait_for_timeout(600)

        if not marcou:
            log("    Nenhuma linha correspondente para marcar")
            dump_dom(page, "linha do job não encontrada")
            return False
        log("    Job selecionado ✓")
        page.wait_for_timeout(600)

        if not click_text(page, "Run Selected", "Run"):
            dump_dom(page, "Run Selected não encontrado")
            log("    ✗ Run Selected não encontrado")
            return False

        page.wait_for_timeout(1_000)

        # A BuyGoods abre "Are you sure? You are about to run the scheduled
        # jobs." — sem clicar em Yes o job NÃO roda, e o download fica
        # esperando para sempre um arquivo que nunca é gerado.
        confirmou = page.evaluate("""
            () => {
                const visivel = el => el.offsetParent !== null &&
                                      el.getClientRects().length > 0;
                const btns = [...document.querySelectorAll(
                    'button, a, [role=button], .btn')];
                const yes = btns.find(b =>
                    (b.textContent || '').trim() === 'Yes' && visivel(b));
                if (!yes) return false;
                yes.scrollIntoView({block: 'center'});
                yes.click();
                return true;
            }
        """)

        if confirmou:
            log("    Confirmado 'Are you sure?' → Yes ✓")
            page.wait_for_timeout(2_500)
        else:
            # alguns fluxos não pedem confirmação — só registra
            log("    (sem diálogo de confirmação)")
            page.wait_for_timeout(1_500)

        log("    Run Selected ✓")
        return True

        dump_dom(page, "Run Selected não encontrado")
        log("    ✗ Run Selected não encontrado")
        return False

    except Exception as e:
        log(f"    Erro: {e}")
        return False



def rodar_jobs_em_lote(page, account_id, sufixos):
    """
    Marca todos os jobs indicados de uma vez e clica Run Selected uma única
    vez — é assim que se faz na mão e economiza uma volta inteira por tipo.
    """
    if not sufixos:
        return False

    log(f"  [Run Selected em lote] {', '.join(sufixos)}")
    page.goto(f"{BASE}/{account_id}#exports", wait_until="networkidle", timeout=20_000)
    page.wait_for_timeout(1_500)
    click_text(page, "Scheduled Export Jobs")
    page.wait_for_timeout(1_500)

    padroes = [KEYWORDS[s_] for s_ in sufixos]
    marcados = 0
    for _ in range(8):
        marcados = page.evaluate("""
            (args) => {
                const [prefixo, padroes] = args;
                const res = padroes.map(p => new RegExp(p, 'i'));
                let n = 0;
                for (const cx of document.querySelectorAll(`input[id^="${prefixo}"]`)) {
                    const tr = cx.closest('tr');
                    const txt = tr ? tr.innerText : '';
                    if (!res.some(r => r.test(txt))) continue;
                    if (/\ball\b/i.test(txt)) continue;     // esse vai pelo modal
                    if (!cx.checked) {
                        cx.checked = true;
                        cx.dispatchEvent(new Event('change', {bubbles: true}));
                        cx.dispatchEvent(new Event('click',  {bubbles: true}));
                    }
                    n++;
                }
                return n;
            }
        """, [CAIXA_JOB, padroes])
        if marcados:
            break
        page.wait_for_timeout(700)

    if not marcados:
        log("    Nenhum job marcado")
        return False
    log(f"    {marcados} job(s) marcado(s) ✓")
    page.wait_for_timeout(600)

    if not click_text(page, "Run Selected", "Run"):
        dump_dom(page, "Run Selected não encontrado")
        return False
    page.wait_for_timeout(1_000)

    confirmou = page.evaluate("""
        () => {
            const visivel = el => el.offsetParent !== null &&
                                  el.getClientRects().length > 0;
            const yes = [...document.querySelectorAll('button, a, [role=button], .btn')]
                .find(b => (b.textContent || '').trim() === 'Yes' && visivel(b));
            if (!yes) return false;
            yes.click();
            return true;
        }
    """)
    log(f"    Confirmado 'Are you sure?' → Yes: {confirmou}")
    page.wait_for_timeout(2_500)
    return True


# ── aguardar e baixar ──────────────────────────────────────────────────────

# IDs das caixas de seleção das duas tabelas — é por elas que identificamos
# as LINHAS de verdade. Buscar link por texto na página inteira acha link de
# menu que contém a mesma palavra ("refund", "chargeback") e nunca dispara
# download: o script fica repetindo a mesma tentativa inútil por minutos.
CAIXA_JOB      = 'dashboard-jobs-table-v2-check-id'
CAIXA_DOWNLOAD = 'dashboard-exports-table-v2-check-id'

# palavra-chave por tipo — o nome do template varia entre contas
# ('Order Items' vs 'Orders Items', 'Customers Refunds' vs 'Customer Refunds',
# 'Customers Chargebacks' vs só 'Chargebacks')
KEYWORDS = {
    # "Order Items" e "Orders Items" — o s varia entre contas
    'orders':      r'orders?\s+items?|items?',
    'refunds':     r'refunds?',
    'chargebacks': r'chargebacks?',
}

TAMANHO_MINIMO = 80   # bytes — abaixo disso veio vazio/cortado


def _linhas_com_caixa(page, prefixo_id):
    """[(id_da_caixa, texto_da_linha)] da tabela indicada."""
    return page.evaluate("""
        (prefixo) => [...document.querySelectorAll(`input[id^="${prefixo}"]`)]
            .map(cx => {
                const tr = cx.closest('tr');
                return [cx.id || '', (tr ? tr.innerText : '').trim()];
            })
    """, prefixo_id)


def _link_de_download(page, sufixo):
    """
    Acha o <a> de download DENTRO da linha da tabela de Export Downloads.
    Devolve o id da caixa da linha, ou None se ainda não estiver pronto.
    """
    padrao = re.compile(KEYWORDS.get(sufixo, sufixo), re.I)
    for cid, texto in _linhas_com_caixa(page, CAIXA_DOWNLOAD):
        if not padrao.search(texto):
            continue
        tem_link = page.evaluate("""
            (cid) => {
                const cx = document.getElementById(cid);
                const tr = cx && cx.closest('tr');
                return !!(tr && tr.querySelector('a'));
            }
        """, cid)
        if tem_link:
            return cid
    return None


def aguardar_e_baixar(page, account_id, nome_conta, tipo, timeout_s=300):
    """
    Aguarda o arquivo aparecer em Export Downloads e baixa.

    Três coisas que o script anterior errava e que causavam a espera infinita:
      1. procurava o link por texto na página toda (achava link de menu)
      2. não capturava o POPUP que a BuyGoods abre para disparar o download
      3. forçava extensão .xlsx em vez de usar o nome sugerido pelo download
    """
    sufixo   = tipo["sufixo"]
    deadline = time.time() + timeout_s

    while time.time() < deadline:
        page.goto(f"{BASE}/{account_id}#exports", wait_until="networkidle", timeout=20_000)
        page.wait_for_timeout(1_500)
        click_text(page, "Export Downloads")
        page.wait_for_timeout(1_200)

        cid = _link_de_download(page, sufixo)
        if cid:
            link = page.locator(f"#{cid}").locator("xpath=ancestor::tr").locator("a").first
            popup = None
            try:
                with page.expect_download(timeout=60_000) as dl_info:
                    try:
                        # o clique costuma abrir um popup que dispara o download
                        with page.expect_popup(timeout=15_000) as pop:
                            link.click(timeout=10_000)
                        popup = pop.value
                    except Exception:
                        pass   # alguns downloads vêm direto, sem popup
                dl = dl_info.value

                sugerido = Path(dl.suggested_filename or "")
                ext  = sugerido.suffix.lower() or ".xlsx"
                dest = ENTRADA / f"{nome_conta}_{sufixo}_{DATA_FIM.strftime(FMT_FILE)}{ext}"
                if dest.exists():
                    dest.unlink()
                dl.save_as(str(dest))

                if popup:
                    try: popup.close()
                    except Exception: pass

                tam = dest.stat().st_size
                if tam < TAMANHO_MINIMO:
                    log(f"    ⚠ {dest.name} veio com {tam} bytes — provavelmente incompleto")
                else:
                    log(f"    ✓ {dest.name} ({tam // 1024} KB)")
                return dest

            except Exception as e:
                log(f"    Erro no download: {e}")
                return None

        log(f"    Export ainda não listado ({int(deadline - time.time())}s restantes)")
        time.sleep(10)

    log(f"    ✗ Timeout — {sufixo} não apareceu em Export Downloads")
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
        # sessão salva: evita refazer login e reduz muito o tempo total
        if STORAGE.exists():
            ctx_args["storage_state"] = str(STORAGE)
        ctx = browser.new_context(**ctx_args)

        if GRAVAR:
            # trace = timeline navegável com screenshot e DOM de cada passo
            ctx.tracing.start(screenshots=True, snapshots=True, sources=True)

        page = ctx.new_page()
        page.set_default_timeout(30_000)

        # se a sessão salva ainda vale, pula o login
        page.goto(f"{BASE}/10495#default", wait_until="networkidle", timeout=20_000)
        if "/login" in page.url:
            login(page)
            try:
                ctx.storage_state(path=str(STORAGE))
                log("  Sessão salva para as próximas execuções")
            except Exception:
                pass
        else:
            log("→ Sessão reaproveitada (sem novo login)")

        itens = list(CONTAS.items())
        if SO_UMA:
            itens = [(k, v) for k, v in itens if k.lower() == SO_UMA.lower()] or itens[:1]
            log(f"  MODO DIAGNÓSTICO: só a conta {itens[0][0]}")

        for nome_conta, account_id in itens:
            log(f"\n── {nome_conta} (id={account_id}) ──")
            sched = ler_scheduled(page, account_id)

            # separar: quem roda por agendamento x quem precisa do modal
            via_lote, via_modal = [], []
            for tipo in TIPOS:
                dr = sched.get(tipo["nome"])
                if dr and "60" in dr.lower():
                    via_lote.append(tipo)
                else:
                    via_modal.append((tipo, dr))

            # 1) dispara todos os agendados de uma vez só
            if via_lote:
                rodar_jobs_em_lote(page, account_id,
                                   [t["sufixo"] for t in via_lote])

            # 2) os que estão travados em "All" precisam do modal, um a um
            for tipo, dr in via_modal:
                log(f"\n  {tipo['nome']}: date_range={dr!r} → modal")
                if not criar_do_zero(page, account_id, tipo):
                    pendentes.append(f"{nome_conta}/{tipo['sufixo']} (date_range={dr})")
                    erro.append(f"{nome_conta}/{tipo['sufixo']}")
                    continue
                dest = aguardar_e_baixar(page, account_id, nome_conta, tipo)
                (ok if dest else erro).append(f"{nome_conta}/{tipo['sufixo']}")

            # 3) baixa os que foram disparados em lote
            for tipo in via_lote:
                log(f"\n  {tipo['nome']}: baixando (agendado)")
                dest = aguardar_e_baixar(page, account_id, nome_conta, tipo)
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
