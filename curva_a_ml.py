# -*- coding: utf-8 -*-
"""
App GUI (Tkinter) para rodar o scraper do Mercado Livre (Curva A)
- Usuário seleciona um Excel (coluna A) diretamente no app
- Faz buscas no padrão "Marca + Modelo + Capacidade" (ou usa consultas cruas)
- Captura só os N primeiros resultados por termo
- Abre a PDP para consolidar vendedor/vendidos/preço/avaliações
- Salva parciais e, ao final, o consolidado
- Comparação com lojas próprias (campo editável)

Como empacotar (resumo):
1) Baixe o Chromium do Playwright para uma pasta local para incluir no build:
   set PLAYWRIGHT_BROWSERS_PATH=ms-playwright
   python -m playwright install chromium

2) Compile (recomendado onedir):
   pyinstaller --noconfirm --onedir --windowed ^
     --name "CurvaA-ML" ^
     --add-data "ms-playwright;ms-playwright" ^
     --hidden-import=playwright.sync_api --hidden-import=pyee ^
     ml_curvaA_app.py

3) Distribua a pasta dist/CurvaA-ML/. O executável usará o Chromium embutido.

Observações:
- Se preferir onefile, mantenha o --add-data e os hidden-imports. O app extrai os
  dados para uma pasta temporária (PyInstaller) e ajusta PLAYWRIGHT_BROWSERS_PATH.
- Se não incluir ms-playwright no build, o app tentará usar o cache do usuário.
"""
from __future__ import annotations
import os, sys, re, time, random, threading, queue, urllib.parse, traceback
from dataclasses import dataclass
from datetime import datetime
from typing import List, Optional

import pandas as pd

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# -------- Playwright --------
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

# ========================= Scraper Core (reutilizado/adaptado) =========================
SEARCH_BASE = "https://lista.mercadolivre.com.br/"

# UI defaults
DEFAULT_FIRST_N = 5
DEFAULT_HEADLESS = True
DEFAULT_RAW_QUERIES = False
DEFAULT_MINI_PAUSAS = True
DEFAULT_SCROLL = True

UA_POOL = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:125.0) Gecko/20100101 Firefox/125.0",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 13_5) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.0 Safari/605.1.15",
]
LOCALES = ["pt-BR", "pt-PT", "es-AR", "en-US"]

PDP_READY_SELECTORS = [
    "h1.ui-pdp-title",
    "span.ui-pdp-price__second-line .andes-money-amount__fraction",
    "button[aria-label*='Comprar' i]",
    "button[title*='Comprar' i]",
    "div.ui-pdp-price__second-line",
]
ANTI_BOT_HINTS = [
    "verifique que você não é um robô",
    "não somos um robô",
    "acessar o mercado livre",
    "acesso temporariamente bloqueado",
    "verificação de segurança",
    "captcha",
]

BRANDS = ["LIQUI MOLY", "ALPINESTARS", "CASTROL", "TIRRENO", "MOTUL", "FRAM", "DID"]


def base_dir() -> str:
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        return sys._MEIPASS  # tipo: ignore[attr-defined]
    return os.path.dirname(os.path.abspath(__file__))


def ensure_playwright_browsers_path():
    """Configura PLAYWRIGHT_BROWSERS_PATH para usar a pasta embutida no build, se existir."""
    # 1) Tenta pasta colada ao executável (add-data: ms-playwright)
    candidates = []
    meipass = base_dir()
    candidates.append(os.path.join(meipass, "ms-playwright"))
    # 2) Tenta ao lado do executável (caso onedir)
    exe_dir = os.path.dirname(sys.executable) if getattr(sys, "frozen", False) else os.path.dirname(os.path.abspath(__file__))
    candidates.append(os.path.join(exe_dir, "ms-playwright"))

    for path in candidates:
        if os.path.isdir(path) and os.listdir(path):
            os.environ["PLAYWRIGHT_BROWSERS_PATH"] = path
            return path
    # Caso não exista embutido, Playwright usará o cache padrão do usuário
    return None


def rand(a, b):
    return random.uniform(a, b)


def rand_sleep(a, b):
    time.sleep(rand(a, b))


def escolher_user_context():
    ua = random.choice(UA_POOL)
    loc = random.choice(LOCALES)
    vw = random.randint(1200, 1600)
    vh = random.randint(700, 950)
    return ua, loc, {"width": vw, "height": vh}


def aceitar_cookies(page):
    try:
        page.get_by_role("button", name=re.compile("Aceitar|Entendi|Accept|Concordo|OK", re.I)).click(timeout=2000)
    except Exception:
        pass


def looks_like_antibot(page) -> bool:
    try:
        txt = (page.inner_text("body") or "").lower()
        return any(h in txt for h in ANTI_BOT_HINTS)
    except Exception:
        return False


def wait_pdp_ready(page, timeout_ms=25000) -> bool:
    try:
        page.wait_for_selector("h1.ui-pdp-title", state="visible", timeout=8000)
        return True
    except PWTimeout:
        pass
    for sel in PDP_READY_SELECTORS:
        try:
            page.wait_for_selector(sel, state="visible", timeout=6000)
            return True
        except PWTimeout:
            continue
    for sel in PDP_READY_SELECTORS:
        try:
            page.wait_for_selector(sel, state="attached", timeout=5000)
            return True
        except PWTimeout:
            continue
    return False


def open_pdp(detail_page, link: str, shot_prefix="fails_pdp", attempt_max=3, log=None) -> bool:
    for attempt in range(1, attempt_max + 1):
        try:
            try:
                detail_page.goto(link, wait_until="load", timeout=30000)
            except PWTimeout:
                detail_page.goto(link, wait_until="networkidle", timeout=30000)

            try:
                aceitar_cookies(detail_page)
                detail_page.mouse.wheel(0, 300); time.sleep(0.2); detail_page.mouse.wheel(0, 600)
            except Exception:
                pass

            if looks_like_antibot(detail_page):
                try:
                    path = f"{shot_prefix}_antibot_{int(time.time())}.png"
                    detail_page.screenshot(path=path, full_page=True)
                except Exception:
                    pass
                if log:
                    log("🧱 Anti-bot na PDP; pulando link.")
                return False

            if wait_pdp_ready(detail_page, timeout_ms=25000):
                return True

            try:
                detail_page.reload(wait_until="load", timeout=15000)
            except PWTimeout:
                pass
            if wait_pdp_ready(detail_page, timeout_ms=12000):
                return True

            try:
                path = f"{shot_prefix}_{attempt}_{int(time.time())}.png"
                detail_page.screenshot(path=path, full_page=True)
            except Exception:
                pass
            if log:
                log(f"   🚫 PDP não ficou pronta (tentativa {attempt}/{attempt_max}).")

        except Exception as e:
            if log:
                log(f"   ⚠️ Erro ao abrir PDP (tentativa {attempt}/{attempt_max}): {e}")
    return False


def human_scroll(page, total_px=4000, step_px=(120, 300), jitter_px=30, top_pause=(0.2, 0.6)):
    if total_px <= 0:
        return
    scrolled = 0
    while scrolled < total_px:
        step = int(rand(*step_px)) + random.randint(-jitter_px, jitter_px)
        try:
            page.mouse.wheel(0, step)
        except Exception:
            break
        scrolled += step
        rand_sleep(*top_pause)


def human_move_mouse(page):
    try:
        w = page.viewport_size["width"]; h = page.viewport_size["height"]
        for _ in range(random.randint(2, 5)):
            x = random.randint(30, w - 30); y = random.randint(80, h - 80)
            page.mouse.move(x, y, steps=random.randint(10, 30))
            rand_sleep(0.05, 0.25)
    except Exception:
        pass


def mini_pausas():
    if random.random() < 0.25:
        rand_sleep(0.8, 2.0)


def to_int(s: str, default=0):
    if not s:
        return default
    s2 = s.replace("\xa0", " ").replace("\u202f", " ").lower()
    m = re.search(r"(\d[\d\.,]*)", s2)
    if not m:
        return default
    num = m.group(1).replace(".", "").replace(",", "")
    try:
        return int(num)
    except:
        return default


def parse_preco_texto_to_float(preco_txt: str):
    if not preco_txt:
        return None
    try:
        return float(preco_txt.replace(".", "").replace(",", "."))
    except:
        return None


def get_cards(list_page):
    cards = list_page.query_selector_all("li.poly-card, li.ui-search-layout__item")
    if not cards:
        cards = list_page.query_selector_all("a.ui-search-item__group__element.ui-search-link")
    return cards or []


def extrair_dados_card(card):
    title_el = card.query_selector("a.poly-component__title, a.ui-search-link")
    titulo = (title_el.inner_text().strip() if title_el else "").strip()
    link = title_el.get_attribute("href") if title_el else None

    patrocinado = False
    try:
        if card.query_selector(".poly-component__ads-promotions"):
            patrocinado = True
    except Exception:
        pass

    nota_media = None
    total_avaliacoes = None
    try:
        rv = card.query_selector("div.poly-component__reviews")
        if rv:
            nota_txt = rv.query_selector(".poly-reviews__rating")
            if nota_txt:
                try:
                    nota_media = float((nota_txt.inner_text() or "").strip().replace(",", "."))
                except:
                    pass
            tot_txt = rv.query_selector(".poly-reviews__total")
            if tot_txt:
                total_avaliacoes = to_int(tot_txt.inner_text(), default=None)
    except Exception:
        pass

    preco_txt = None
    try:
        frac = card.query_selector("span.andes-money-amount__fraction, span.price-tag-fraction")
        cents = card.query_selector("span.andes-money-amount__cents, span.price-tag-cents")
        if frac:
            preco_txt = frac.inner_text().strip()
            if cents:
                preco_txt = f"{preco_txt},{cents.inner_text().strip()}"
    except Exception:
        pass
    preco_num = parse_preco_texto_to_float(preco_txt)

    desconto_pct_txt = None
    try:
        desc_el = card.query_selector(".ui-search-price__discount, .poly-price__discount, .andes-money-amount__discount")
        if desc_el:
            desconto_pct_txt = (desc_el.inner_text() or "").strip()
    except Exception:
        pass

    tipo_anuncio = "Clássico"
    try:
        inst = card.query_selector("span.poly-price__installments")
        if inst and "sem juros" in (inst.inner_text() or "").lower():
            tipo_anuncio = "Premium"
    except Exception:
        pass

    return {
        "Título": titulo,
        "Link": link,
        "Patrocinado": patrocinado,
        "Nota média (lista)": nota_media,
        "Nº avaliações (lista)": total_avaliacoes,
        "Preço (lista)": preco_txt,
        "Preço (lista num)": preco_num,
        "Preço promo (lista)": preco_txt,
        "% desc (lista)": desconto_pct_txt,
        "Tipo (lista)": tipo_anuncio,
    }


def parse_preco_pdp(page):
    candidatos = [
        "span.andes-money-amount__fraction",
        "span.ui-pdp-price__second-line .andes-money-amount__fraction",
        "span.price-tag-fraction",
    ]
    cents_sel = [
        "span.andes-money-amount__cents",
        "span.ui-pdp-price__second-line .andes-money-amount__cents",
        "span.price-tag-cents",
    ]
    preco_txt = None
    for sel in candidatos:
        el = page.query_selector(sel)
        if el:
            frac = (el.inner_text() or "").strip()
            cents = None
            for cs in cents_sel:
                cel = page.query_selector(cs)
                if cel:
                    cents = (cel.inner_text() or "").strip()
                    break
            preco_txt = f"{frac},{cents}" if cents else frac
            break
    return preco_txt, parse_preco_texto_to_float(preco_txt)


def extrair_vendedor_pdp(page):
    try:
        el = page.query_selector("button.ui-pdp-seller__link-trigger-button")
        if el:
            spans = el.query_selector_all("span")
            if len(spans) > 1:
                return spans[1].inner_text().strip()
        el = page.query_selector("a.ui-pdp-media__action")
        if el:
            return (el.inner_text() or "").strip()
    except:
        pass
    return "Não encontrado"


def extrair_vendidos_pdp(page):
    el = page.query_selector("span.ui-pdp-subtitle")
    txt = el.inner_text().strip() if el else ""
    qtd = to_int(txt, default=0)
    return str(qtd)


def extrair_avaliacoes_pdp(page):
    nota = None
    total = None
    try:
        nota_el = page.query_selector("span.ui-review-summary__rating, .ui-pdp-review__rating__summary")
        if nota_el:
            ntxt = (nota_el.inner_text() or "").strip().replace(",", ".")
            try:
                nota = float(re.search(r"(\d+(\.\d+)?)", ntxt).group(1))
            except:
                pass
        tot_el = page.query_selector("span.ui-review-summary__average, .ui-review-capabilities__count")
        if tot_el:
            total = to_int(tot_el.inner_text(), default=None)
    except:
        pass
    return nota, total


def title_to_user_query(title: str) -> str:
    s = (title or "").strip()
    if not s:
        return ""

    up = s.upper()

    # Marca
    brand = None
    for b in sorted(BRANDS, key=len, reverse=True):
        if b in up:
            brand = b.title()
            up = up.replace(b, " ").strip()
            break
    if not brand:
        m = re.search(r"\b[A-ZÁ-Ú]{3,}\b", up)
        brand = (m.group(0).title() if m else "").strip()

    # Viscosidade
    vis = None
    m = re.search(r"\b(\d{1,2})\s*[W]\s*(\d{2})\b", up, flags=re.I)
    if m:
        vis = f"{m.group(1)}w{m.group(2)}"

    # Capacidade
    cap = None
    m = re.search(r"\b(\d+[.,]?\d*)\s*(LITROS?|L|ML|M[L])\b", up, flags=re.I)
    if m:
        qty = m.group(1).replace(",", ".")
        unit = m.group(2).upper()
        try:
            val = float(qty)
        except:
            val = None
        if unit in ("L", "LITRO", "LITROS"):
            if val is not None and abs(val - 1.0) < 1e-6:
                cap = "1 litro"
            elif val is not None and float(val).is_integer():
                cap = f"{int(val)} litros"
            elif val is not None:
                cap = f"{qty} litros"
            else:
                cap = f"{qty} litro"
        else:
            # ML
            try:
                cap = f"{int(float(qty))} ml"
            except:
                cap = f"{qty} ml"

    # Tokens (modelos/códigos)
    tokens = []
    for pat in [r"\b\d{3,4}\+?\b"]:
        for m in re.finditer(pat, up):
            tokens.append(m.group(0))

    if re.search(r"\bX[- ]?CESS\b", up):
        tokens.append("X-Cess")
    if re.search(r"\bGEN2\b", up):
        tokens.append("Gen2")
    if re.search(r"\b2T\b", up):
        tokens.append("2t")
    if re.search(r"\bC2\s*PLUS\b", up):
        tokens.append("C2 Plus")
    elif re.search(r"\bC2\b", up):
        tokens.append("C2")

    # Filtros: códigos tipo PH6017A
    m = re.search(r"\b[A-Z]{1,3}\d{3,6}[A-Z]?\b", up)
    if m:
        code = m.group(0)
        base = " ".join([p for p in [brand, code] if p])
        return base.strip()

    if "CHAIN" in up and "LUBE" in up:
        tokens += ["Chain", "Lube"]
    if "ROAD" in up:
        tokens.append("Road")

    parts = [brand] + tokens
    if vis:
        parts.append(vis)
    if cap:
        parts.append(cap)

    consulta = " ".join([p for p in parts if p]).strip()
    if not consulta or consulta.lower() == (brand or "").lower():
        cleaned = re.sub(r"\b(oleo|óleo|lubrificante|spray|de|do|da|para|off|road|sint[ée]tico|4t)\b", " ", s, flags=re.I)
        cleaned = re.sub(r"\s+", " ", cleaned).strip()
        return cleaned
    return consulta


def load_terms_from_excel(xlsx_path: str, sheet_name: Optional[str] = None) -> List[str]:
    import os
    ext = os.path.splitext(xlsx_path)[1].lower()

    # CSV direto
    if ext == ".csv":
        df = pd.read_csv(xlsx_path)
    else:
        # Se NÃO informarem aba, use a primeira (0).
        # Isso evita o retorno como dict quando sheet_name=None.
        sheet_arg = 0 if (sheet_name is None or str(sheet_name).strip() == "") else sheet_name
        df = pd.read_excel(xlsx_path, sheet_name=sheet_arg)

        # Em casos raros alguém passa None e o pandas devolve dict; garanta DataFrame:
        if isinstance(df, dict):
            # pega a primeira aba disponível
            first_key = next(iter(df.keys()))
            df = df[first_key]

    # Se chegou até aqui e ainda não for DataFrame, falhe de forma amigável
    if not hasattr(df, "shape"):
        raise ValueError("Não foi possível ler a planilha como DataFrame. Verifique o arquivo/aba.")

    if df.shape[1] == 0:
        return []

    # Coluna A (primeira coluna) é a fonte dos termos
    col = df.columns[0]
    termos = []
    for x in df[col].tolist():
        if pd.isna(x):
            continue
        s = str(x).strip()
        if s and s.lower() != "nan":
            termos.append(s)
    return termos


def comparar_precos_por_consulta(registros, nossas_lojas: set[str]):
    if not registros:
        return registros
    nossas_upper = {s.upper() for s in nossas_lojas}
    nossos = [r for r in registros if (r.get("Vendedor") or "").upper() in nossas_upper]
    nosso_preco = None
    if nossos:
        nossos_precos = [r.get("Preço (num)") for r in nossos if isinstance(r.get("Preço (num)"), (int, float)) and r.get("Preço (num)") is not None]
        if nossos_precos:
            nosso_preco = min(nossos_precos)
    for r in registros:
        vendedor_up = (r.get("Vendedor") or "").upper()
        r["É nossa loja?"] = "Sim" if vendedor_up in nossas_upper else "Não"
        if nosso_preco is not None and r.get("Preço (num)") is not None and r["É nossa loja?"] == "Não":
            r["Concorrente abaixo de nós?"] = "Sim" if r["Preço (num)"] < nosso_preco else "Não"
            r["Nosso menor preço (num)"] = nosso_preco
        else:
            r["Concorrente abaixo de nós?"] = ""
            r["Nosso menor preço (num)"] = nosso_preco if nosso_preco is not None else ""
    return registros


# ========================= Thread do Scraper =========================
@dataclass
class JobConfig:
    xlsx_path: str
    sheet_name: Optional[str]
    first_n: int
    headless: bool
    raw_queries: bool
    mini_pausas: bool
    scroll_pages: bool
    nossas_lojas: set
    out_dir: str


class ScraperThread(threading.Thread):
    def __init__(self, cfg: JobConfig, log_q: queue.Queue, progress_q: queue.Queue, stop_evt: threading.Event):
        super().__init__(daemon=True)
        self.cfg = cfg
        self.log_q = log_q
        self.progress_q = progress_q
        self.stop_evt = stop_evt

    def log(self, msg: str):
        ts = datetime.now().strftime("%H:%M:%S")
        self.log_q.put(f"[{ts}] {msg}")

    def run(self):
        try:
            ensure_playwright_browsers_path()

            termos = load_terms_from_excel(self.cfg.xlsx_path, self.cfg.sheet_name)
            total = len(termos)
            if total == 0:
                self.log("Excel sem termos na coluna A.")
                self.progress_q.put((0, 0))
                return

            todos = []
            parcial_path = os.path.join(self.cfg.out_dir, "resultado_curvaA_parcial.xlsx")
            final_path = os.path.join(self.cfg.out_dir, "resultado_curvaA.xlsx")
            os.makedirs(self.cfg.out_dir, exist_ok=True)

            with sync_playwright() as p:
                ua, loc, viewport = escolher_user_context()
                browser = p.chromium.launch(headless=self.cfg.headless)
                context = browser.new_context(locale=loc, user_agent=ua, viewport=viewport)
                list_page = context.new_page()
                detail_page = context.new_page()

                self.log(f"🌐 UA: {ua[:55]}… | locale: {loc} | viewport: {viewport}")

                for idx, termo in enumerate(termos, start=1):
                    if self.stop_evt.is_set():
                        self.log("🛑 Interrompido pelo usuário.")
                        break

                    consulta = termo.strip() if self.cfg.raw_queries else title_to_user_query(termo)
                    if not consulta:
                        self.log(f"⚠️ Termo vazio/inalcançável: {termo}")
                        self.progress_q.put((idx, total))
                        continue

                    url = SEARCH_BASE + urllib.parse.quote_plus(consulta)
                    self.log(f"🔎 {termo}  →  {consulta}")
                    self.log(f"    {url}")

                    try:
                        list_page.goto(url, wait_until="domcontentloaded", timeout=30000)
                        aceitar_cookies(list_page)
                        list_page.wait_for_selector("main", timeout=15000)
                    except PWTimeout:
                        self.log("⏱️ Timeout na busca; pulando termo.")
                        self.progress_q.put((idx, total))
                        continue

                    if self.cfg.scroll_pages:
                        human_move_mouse(list_page)
                        human_scroll(list_page, total_px=random.randint(1500, 3000))
                    if self.cfg.mini_pausas:
                        mini_pausas()

                    cards = get_cards(list_page)
                    self.log(f"🧩 {len(cards)} cards encontrados")
                    if not cards:
                        self.progress_q.put((idx, total))
                        continue

                    selecionados = []
                    vistos_links = set()
                    for i, card in enumerate(cards, start=1):
                        if self.stop_evt.is_set():
                            break
                        try:
                            box = None
                            try:
                                box = card.bounding_box()
                            except Exception:
                                pass
                            if box:
                                list_page.mouse.move(
                                    box["x"] + box["width"]/2 + random.randint(-15, 15),
                                    box["y"] + box["height"]/2 + random.randint(-8, 8),
                                    steps=random.randint(8, 20)
                                )
                                time.sleep(0.1)

                            base = extrair_dados_card(card)
                            link = base.get("Link")
                            if not link or link in vistos_links:
                                continue
                            vistos_links.add(link)
                            selecionados.append(base)
                            if len(selecionados) >= self.cfg.first_n:
                                break
                        except Exception as e:
                            self.log(f"❌ Erro no card {i}: {e}")

                    if not selecionados:
                        self.progress_q.put((idx, total))
                        continue

                    self.log(f"➡️ Processando os {len(selecionados)} primeiros…")

                    try:
                        for j, base in enumerate(selecionados, 1):
                            if self.stop_evt.is_set():
                                break
                            self.log(f"   → ({j}/{len(selecionados)}) {base.get('Título','')[:90]}")
                            ok = open_pdp(detail_page, base["Link"], shot_prefix=f"pdp_{int(time.time())}", attempt_max=3, log=self.log)
                            if not ok:
                                self.log("     ❌ Falha ao abrir PDP; seguindo.")
                                continue

                            if self.cfg.scroll_pages:
                                human_move_mouse(detail_page)
                                human_scroll(detail_page, total_px=random.randint(900, 1600))
                            if self.cfg.mini_pausas:
                                mini_pausas()

                            preco_pdp_txt, preco_pdp_num = parse_preco_pdp(detail_page)
                            vendedor = extrair_vendedor_pdp(detail_page)
                            vendidos = extrair_vendidos_pdp(detail_page)
                            nota_pdp, total_pdp = extrair_avaliacoes_pdp(detail_page)

                            preco_txt_final = base["Preço (lista)"] or preco_pdp_txt
                            preco_num_final = base["Preço (lista num)"] if base["Preço (lista num)"] is not None else preco_pdp_num
                            nota_final = base["Nota média (lista)"] if base["Nota média (lista)"] is not None else nota_pdp
                            total_av_final = base["Nº avaliações (lista)"] if base["Nº avaliações (lista)"] is not None else total_pdp

                            todos.append({
                                "Termo original": termo,
                                "Consulta": consulta,
                                "Patrocinado": "Sim" if base["Patrocinado"] else "Não",
                                "Título": base["Título"],
                                "Preço": preco_txt_final,
                                "Preço (num)": preco_num_final,
                                "Preço promo (lista)": base["Preço promo (lista)"],
                                "% desc (lista)": base["% desc (lista)"],
                                "Tipo anúncio": base["Tipo (lista)"],
                                "Vendedor": vendedor,
                                "Vendidos (PDP)": vendidos,
                                "Nota média": nota_final,
                                "Nº avaliações": total_av_final,
                                "Link": base["Link"],
                            })

                            time.sleep(rand(0.6, 1.6))

                        # parcial por termo
                        if todos:
                            df_parcial = pd.DataFrame(todos)
                            df_parcial.to_excel(parcial_path, index=False)
                            self.log(f"💾 Parcial salva → {parcial_path}")

                    except Exception as e:
                        self.log(f"⚠️ Erro ao processar PDPs: {e}")

                    self.progress_q.put((idx, total))
                    # pausa entre termos
                    time.sleep(rand(2.5, 5.5))

                browser.close()

            if not todos:
                self.log("📭 Nenhum resultado coletado.")
                self.progress_q.put((total, total))
                return

            # Ordenar e comparar
            df = pd.DataFrame(todos)
            df["Nota média (ord)"] = pd.to_numeric(df["Nota média"], errors="coerce")
            df["Nº avaliações (ord)"] = pd.to_numeric(df["Nº avaliações"], errors="coerce")
            df = df.sort_values(
                by=["Termo original", "Patrocinado", "Nota média (ord)", "Nº avaliações (ord)"],
                ascending=[True, True, False, False]
            ).drop(columns=["Nota média (ord)", "Nº avaliações (ord)"])

            # Comparação de preços por consulta
            registros_cmp = []
            for termo in df["Termo original"].unique():
                bloco = df[df["Termo original"] == termo].to_dict("records")
                bloco2 = comparar_precos_por_consulta(bloco, self.cfg.nossas_lojas)
                registros_cmp.extend(bloco2)
            df_final = pd.DataFrame(registros_cmp)

            cols = [
                "Termo original","Consulta","Patrocinado","Tipo anúncio","É nossa loja?","Concorrente abaixo de nós?",
                "Nosso menor preço (num)","Título","Vendedor","Preço","Preço (num)","Preço promo (lista)","% desc (lista)",
                "Nota média","Nº avaliações","Vendidos (PDP)","Link"
            ]
            cols = [c for c in cols if c in df_final.columns] + [c for c in df_final.columns if c not in cols]
            df_final = df_final[cols]

            df_final.to_excel(final_path, index=False)
            self.log(f"🏁 Finalizado. Arquivo salvo: {final_path}")
            self.progress_q.put((total, total))

        except Exception as e:
            self.log("❌ Erro fatal:\n" + traceback.format_exc())
            self.progress_q.put((0, 0))


# ========================= GUI (Tkinter) =========================
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Curva A – Scraper ML")
        self.geometry("780x620")

        self.log_q: queue.Queue = queue.Queue()
        self.progress_q: queue.Queue = queue.Queue()
        self.stop_evt = threading.Event()
        self.worker: Optional[ScraperThread] = None

        self._build_ui()
        self.after(150, self._poll_queues)

    def _build_ui(self):
        pad = {"padx": 8, "pady": 6}

        frm = ttk.Frame(self)
        frm.pack(fill=tk.BOTH, expand=False, **pad)

        # Linha 1: Seleção do Excel e Aba
        ttk.Label(frm, text="Arquivo Excel (coluna A)").grid(row=0, column=0, sticky="w")
        self.var_excel = tk.StringVar()
        ent_excel = ttk.Entry(frm, textvariable=self.var_excel, width=70)
        ent_excel.grid(row=1, column=0, columnspan=2, sticky="we")
        ttk.Button(frm, text="Selecionar…", command=self._choose_excel).grid(row=1, column=2, sticky="we")

        ttk.Label(frm, text="Aba (opcional)").grid(row=0, column=3, sticky="w")
        self.var_sheet = tk.StringVar()
        ttk.Entry(frm, textvariable=self.var_sheet, width=20).grid(row=1, column=3, sticky="we")

        # Linha 2: Opções
        opt = ttk.LabelFrame(self, text="Opções")
        opt.pack(fill=tk.X, expand=False, **pad)

        self.var_firstn = tk.IntVar(value=DEFAULT_FIRST_N)
        ttk.Label(opt, text="Capturar primeiros N:").grid(row=0, column=0, sticky="w", **pad)
        ttk.Spinbox(opt, from_=1, to=10, width=6, textvariable=self.var_firstn).grid(row=0, column=1, sticky="w")

        self.var_headless = tk.BooleanVar(value=DEFAULT_HEADLESS)
        ttk.Checkbutton(opt, text="Headless (sem abrir navegador)", variable=self.var_headless).grid(row=0, column=2, sticky="w")

        self.var_raw = tk.BooleanVar(value=DEFAULT_RAW_QUERIES)
        ttk.Checkbutton(opt, text="Usar consultas cruas (não transformar)", variable=self.var_raw).grid(row=0, column=3, sticky="w")

        self.var_scroll = tk.BooleanVar(value=DEFAULT_SCROLL)
        ttk.Checkbutton(opt, text="Scroll e mouse humanizados", variable=self.var_scroll).grid(row=1, column=0, sticky="w", **pad)

        self.var_minipause = tk.BooleanVar(value=DEFAULT_MINI_PAUSAS)
        ttk.Checkbutton(opt, text="Pausas aleatórias", variable=self.var_minipause).grid(row=1, column=1, sticky="w")

        ttk.Label(opt, text="Monitorar Lojas Especificas (separar por ponto e vírgula)").grid(row=1, column=2, sticky="w")
        self.var_lojas = tk.StringVar(value="Lojas que deseja monitorar")
        ttk.Entry(opt, textvariable=self.var_lojas, width=30).grid(row=1, column=3, sticky="we")

        # Linha 3: Saída
        outf = ttk.LabelFrame(self, text="Saída")
        outf.pack(fill=tk.X, expand=False, **pad)
        ttk.Label(outf, text="Pasta de saída").grid(row=0, column=0, sticky="w", **pad)
        self.var_outdir = tk.StringVar(value=os.path.join(os.path.expanduser("~"), "Desktop", "saida_ml"))
        ttk.Entry(outf, textvariable=self.var_outdir, width=60).grid(row=0, column=1, sticky="we")
        ttk.Button(outf, text="Escolher…", command=self._choose_outdir).grid(row=0, column=2, sticky="we")
        ttk.Button(outf, text="Abrir pasta", command=self._open_outdir).grid(row=0, column=3, sticky="we")

        # Controles: Iniciar / Parar
        ctrlf = ttk.Frame(self)
        ctrlf.pack(fill=tk.X, expand=False, **pad)
        self.btn_start = ttk.Button(ctrlf, text="Iniciar", command=self._start)
        self.btn_start.pack(side=tk.LEFT)
        self.btn_stop = ttk.Button(ctrlf, text="Parar", command=self._stop, state=tk.DISABLED)
        self.btn_stop.pack(side=tk.LEFT, padx=8)

        # Progress bar
        self.prog = ttk.Progressbar(self, mode="determinate")
        self.prog.pack(fill=tk.X, expand=False, **pad)

        # Log
        logf = ttk.LabelFrame(self, text="Log")
        logf.pack(fill=tk.BOTH, expand=True, **pad)
        self.txt = tk.Text(logf, height=18, wrap="word")
        self.txt.pack(fill=tk.BOTH, expand=True)
        self.txt.config(state=tk.DISABLED)

    # ----- helpers UI -----
    def _choose_excel(self):
        path = filedialog.askopenfilename(title="Selecione o Excel", filetypes=[("Excel", "*.xlsx;*.xlsm;*.xls"), ("Todos", "*.*")])
        if path:
            self.var_excel.set(path)

    def _choose_outdir(self):
        path = filedialog.askdirectory(title="Escolher pasta de saída")
        if path:
            self.var_outdir.set(path)

    def _open_outdir(self):
        path = self.var_outdir.get().strip()
        if path and os.path.isdir(path):
            if sys.platform.startswith("win"):
                os.startfile(path)
            elif sys.platform == "darwin":
                os.system(f"open '{path}'")
            else:
                os.system(f"xdg-open '{path}'")
        else:
            messagebox.showinfo("Pasta", "Pasta de saída inexistente.")

    def _toggle_controls(self, running: bool):
        self.btn_start.config(state=tk.DISABLED if running else tk.NORMAL)
        self.btn_stop.config(state=tk.NORMAL if running else tk.DISABLED)

    def _start(self):
        xlsx = self.var_excel.get().strip()
        if not xlsx or not os.path.isfile(xlsx):
            messagebox.showerror("Arquivo", "Selecione um Excel válido.")
            return
        outdir = self.var_outdir.get().strip()
        if not outdir:
            messagebox.showerror("Saída", "Defina a pasta de saída.")
            return

        lojas = set([s.strip() for s in self.var_lojas.get().split(";") if s.strip()])
        cfg = JobConfig(
            xlsx_path=xlsx,
            sheet_name=(self.var_sheet.get().strip() or None),
            first_n=max(1, int(self.var_firstn.get() or 5)),
            headless=bool(self.var_headless.get()),
            raw_queries=bool(self.var_raw.get()),
            mini_pausas=bool(self.var_minipause.get()),
            scroll_pages=bool(self.var_scroll.get()),
            nossas_lojas=lojas,
            out_dir=outdir,
        )

        self.stop_evt.clear()
        self._toggle_controls(True)
        self._log_clear()
        self._log("Iniciando…")
        self.prog.config(value=0, maximum=100)

        self.worker = ScraperThread(cfg, self.log_q, self.progress_q, self.stop_evt)
        self.worker.start()

    def _stop(self):
        if self.worker and self.worker.is_alive():
            self.stop_evt.set()
            self._log("Solicitada parada. Aguardando o lote atual…")
        self._toggle_controls(False)

    def _log(self, msg: str):
        self.txt.config(state=tk.NORMAL)
        self.txt.insert(tk.END, msg + "\n")
        self.txt.see(tk.END)
        self.txt.config(state=tk.DISABLED)

    def _log_clear(self):
        self.txt.config(state=tk.NORMAL)
        self.txt.delete("1.0", tk.END)
        self.txt.config(state=tk.DISABLED)

    def _poll_queues(self):
        try:
            while True:
                m = self.log_q.get_nowait()
                self._log(m)
        except queue.Empty:
            pass

        try:
            while True:
                cur, total = self.progress_q.get_nowait()
                val = 0 if total == 0 else int((cur / total) * 100)
                self.prog.config(value=val)
                if cur >= total and total > 0:
                    self._toggle_controls(False)
        except queue.Empty:
            pass

        self.after(150, self._poll_queues)


if __name__ == "__main__":
    ensure_playwright_browsers_path()
    app = App()
    app.mainloop()
