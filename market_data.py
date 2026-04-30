"""
Busca cotações em tempo real da B3 via yfinance e taxa Selic via BCB.
Cache de 15 minutos para evitar throttle.
"""
import time
import json
import urllib.request
from concurrent.futures import ThreadPoolExecutor, as_completed
from typing import Optional

try:
    import yfinance as yf
    YFINANCE_OK = True
except ImportError:
    YFINANCE_OK = False

# Mapeamento SKU (como aparece no Excel) → ticker Yahoo Finance
TICKER_MAP: dict[str, str] = {
    'bbas3':  'BBAS3.SA',
    'bbsa3':  'BBAS3.SA',
    'bbasa3': 'BBAS3.SA',
    'b3sa3':  'B3SA3.SA',
    'recr11': 'RECR11.SA',
    'trxf11': 'TRXF11.SA',
    'tgar11': 'TGAR11.SA',
    'tgar1':  'TGAR11.SA',
    'petr4':  'PETR4.SA',
    'hapv3':  'HAPV3.SA',
    'aure3':  'AURE3.SA',
    'aure3f': 'AURE3.SA',
    'simh3':  'SIMH3.SA',
    'wrld11': 'WRLD11.SA',
}

# Tipo de ativo para UI
ASSET_TYPE: dict[str, str] = {
    'RECR11.SA': 'FII', 'TRXF11.SA': 'FII', 'TGAR11.SA': 'FII',
    'BBAS3.SA': 'Ação', 'B3SA3.SA': 'Ação', 'PETR4.SA': 'Ação',
    'HAPV3.SA': 'Ação', 'AURE3.SA': 'Ação', 'SIMH3.SA': 'Ação',
    'WRLD11.SA': 'ETF',
}

_cache: dict[str, tuple[float, dict]] = {}
CACHE_TTL = 900   # 15 minutos


# ── Cotações ─────────────────────────────────────────────────────────────────

def _fetch_one(sku: str, ticker_sym: str) -> tuple[str, Optional[dict]]:
    try:
        fi = yf.Ticker(ticker_sym).fast_info
        price = fi.last_price
        prev  = fi.previous_close or fi.regular_market_previous_close
        if not price or not prev:
            return sku, None
        change = (price - prev) / prev * 100
        return sku, {
            'ticker':    ticker_sym.replace('.SA', ''),
            'type':      ASSET_TYPE.get(ticker_sym, 'Ativo'),
            'price':     round(price, 2),
            'prev':      round(prev, 2),
            'change':    round(change, 2),
            'day_high':  round(fi.day_high or price, 2),
            'day_low':   round(fi.day_low or price, 2),
            'year_high': round(fi.year_high or price, 2),
            'year_low':  round(fi.year_low or price, 2),
            'year_change': round((fi.year_change or 0) * 100, 2),
        }
    except Exception as e:
        print(f'[market] {ticker_sym}: {e}')
        return sku, None


def get_quotes(skus: list[str]) -> dict[str, dict]:
    """Retorna cotações para uma lista de SKUs. Usa cache de 15 min."""
    if not YFINANCE_OK:
        return {}

    result: dict[str, dict] = {}
    to_fetch: list[tuple[str, str]] = []

    for sku in skus:
        ticker = TICKER_MAP.get(sku.lower())
        if not ticker:
            continue
        cached = _cache.get(ticker)
        if cached and (time.time() - cached[0]) < CACHE_TTL:
            result[sku.lower()] = cached[1]
        else:
            to_fetch.append((sku.lower(), ticker))

    if to_fetch:
        with ThreadPoolExecutor(max_workers=8) as ex:
            futures = {ex.submit(_fetch_one, s, t): s for s, t in to_fetch}
            for fut in as_completed(futures):
                sku, data = fut.result()
                if data:
                    _cache[TICKER_MAP[sku]] = (time.time(), data)
                    result[sku] = data

    return result


def get_all_quotes() -> dict[str, dict]:
    """Cotações de todos os ativos da carteira, deduplicated por ticker."""
    raw = get_quotes(list(TICKER_MAP.keys()))
    # Deduplicar: usar o ticker limpo (sem .SA) como chave
    seen: dict[str, dict] = {}
    for data in raw.values():
        key = data['ticker']   # e.g. 'RECR11'
        if key not in seen:
            seen[key] = data
    return seen


# ── Taxa Selic (BCB) ──────────────────────────────────────────────────────────

_selic_cache: tuple[float, Optional[float]] = (0, None)
SELIC_TTL = 3600  # 1 hora


def get_selic_rate() -> Optional[float]:
    """Retorna a taxa Selic anual atual (em %) via API do Banco Central."""
    global _selic_cache
    ts, val = _selic_cache
    if val is not None and (time.time() - ts) < SELIC_TTL:
        return val

    try:
        url = 'https://api.bcb.gov.br/dados/serie/bcdata.sgs.11/dados/ultimos/1?formato=json'
        with urllib.request.urlopen(url, timeout=5) as resp:
            data = json.loads(resp.read())
            # API retorna taxa diária; anualizar: (1 + d/100)^252 - 1
            daily = float(data[0]['valor'].replace(',', '.')) / 100
            annual = ((1 + daily) ** 252 - 1) * 100
            _selic_cache = (time.time(), round(annual, 2))
            return _selic_cache[1]
    except Exception as e:
        print(f'[selic] {e}')
        return None
