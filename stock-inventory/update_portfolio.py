#!/usr/bin/env python3
"""
Daily portfolio price updater — Notion Portfolio Tracker
Fetches live prices via yfinance and updates:
  1. Each row in the Holdings database (Current Price, FX Rate, Market Value TWD, P&L, etc.)
  2. The Portfolio Tracker page header (heading date, total value, allocation table)

Run via GitHub Actions at 13:30 Taipei time (05:30 UTC).
Requires env var: NOTION_API_KEY
"""
import os
import sys
import time
from datetime import datetime, date
import pytz
import yfinance as yf
import requests

# ── Config ─────────────────────────────────────────────────────────────────────
NOTION_TOKEN  = os.environ['NOTION_API_KEY']
DATABASE_ID   = 'd68d0bda-f13b-45a9-98ed-0930f3bf36be'   # Holdings database
TRACKER_PAGE  = '36c94a3e-d2ab-80f9-8b8e-fc7ff3efffcf'   # Portfolio Tracker page
TZ_TAIPEI     = pytz.timezone('Asia/Taipei')

NOTION_HEADERS = {
    'Authorization': f'Bearer {NOTION_TOKEN}',
    'Content-Type': 'application/json',
    'Notion-Version': '2022-06-28',
}

# ── Notion API helpers ──────────────────────────────────────────────────────────

def _req(method, path, **kwargs):
    url = f'https://api.notion.com/v1{path}'
    for attempt in range(3):
        r = requests.request(method, url, headers=NOTION_HEADERS, **kwargs)
        if r.status_code == 429:          # rate limited
            time.sleep(2 ** attempt)
            continue
        r.raise_for_status()
        return r.json()
    raise RuntimeError(f'Notion API failed after retries: {path}')

def notion_get(path):   return _req('GET',    path)
def notion_post(path, body): return _req('POST',   path, json=body)
def notion_patch(path, body): return _req('PATCH',  path, json=body)

def query_database(db_id):
    results, body = [], {'page_size': 100}
    while True:
        data = notion_post(f'/databases/{db_id}/query', body)
        results.extend(data['results'])
        if not data.get('has_more'):
            break
        body['start_cursor'] = data['next_cursor']
        time.sleep(0.3)
    return results

# ── FX rates ───────────────────────────────────────────────────────────────────

def fetch_fx_rates():
    """Returns {USD: twd_per_usd, JPY: twd_per_jpy, TWD: 1.0}"""
    rates = {'TWD': 1.0}
    fallbacks = {'USD': 31.5, 'JPY': 0.218}
    for pair, key in [('USDTWD=X', 'USD'), ('JPYTWD=X', 'JPY')]:
        try:
            rate = yf.Ticker(pair).fast_info.last_price
            if rate and rate > 0:
                rates[key] = float(rate)
                print(f'  FX {key}/TWD = {rate:.4f}')
            else:
                raise ValueError('zero/null rate')
        except Exception as e:
            rates[key] = fallbacks[key]
            print(f'  ⚠️  FX {key} using fallback {fallbacks[key]}: {e}')
    return rates

# ── Stock price fetching ────────────────────────────────────────────────────────

def to_yf_ticker(ticker, market):
    """Convert Notion Ticker + Market to Yahoo Finance symbol."""
    t = ticker.strip().replace('TPE:', '').replace('TAI:', '')
    if market == 'TW':
        if not t.endswith('.TW'):
            t += '.TW'
    elif market == 'JP':
        if '.' not in t:
            t += '.T'
    return t

def fetch_price(ticker, market):
    """Returns (current_price, today_change_pct_decimal) or (None, None)."""
    symbol = to_yf_ticker(ticker, market)
    for attempt in range(2):
        try:
            fi = yf.Ticker(symbol).fast_info
            price = fi.last_price
            prev  = fi.previous_close
            if not price or price <= 0:
                raise ValueError(f'invalid price {price}')
            change = ((price - prev) / prev) if (prev and prev > 0) else 0.0
            return float(price), float(change)
        except Exception as e:
            if attempt == 0:
                time.sleep(1)
            else:
                print(f'    ✗  {symbol}: {e}')
    return None, None

# ── Update Holdings rows ────────────────────────────────────────────────────────

def update_holdings(holdings, fx):
    today_iso = date.today().isoformat()

    for page in holdings:
        props   = page['properties']
        page_id = page['id']

        def get_str(field, sub='plain_text'):
            items = props.get(field, {}).get('rich_text', []) or []
            return items[0].get(sub, '') if items else ''

        def get_num(field):
            return props.get(field, {}).get('number') or 0

        def get_sel(field):
            sel = props.get(field, {}).get('select') or {}
            return sel.get('name', '')

        def get_title():
            items = props.get('Name', {}).get('title', []) or []
            return items[0].get('plain_text', '') if items else ''

        ticker   = get_str('Ticker')
        market   = get_sel('Market') or 'US'
        currency = get_sel('Currency') or 'USD'
        shares   = get_num('Shares')
        cost_twd = get_num('Cost TWD')
        name     = get_title() or ticker

        if not ticker or not shares:
            print(f'  –  {name}: skipped (no ticker or shares)')
            continue

        price, change_pct = fetch_price(ticker, market)
        if price is None:
            continue

        fx_rate   = fx.get(currency, 1.0)
        mkt_val   = price * shares * fx_rate
        pnl       = mkt_val - cost_twd
        ret_pct   = (pnl / cost_twd) if cost_twd else 0.0
        today_pnl = mkt_val * change_pct

        notion_patch(f'/pages/{page_id}', {'properties': {
            'Current Price':    {'number': round(price, 4)},
            'FX Rate':          {'number': round(fx_rate, 4)},
            'Market Value TWD': {'number': round(mkt_val)},
            'P&L TWD':          {'number': round(pnl)},
            'Return %':         {'number': round(ret_pct, 6)},
            'Today Change %':   {'number': round(change_pct, 6)},
            'Today P&L TWD':    {'number': round(today_pnl)},
            'Last Updated':     {'date': {'start': today_iso}},
        }})
        print(f'  ✓  {name} ({to_yf_ticker(ticker, market)}): '
              f'{price:.4f} × {fx_rate:.2f} = TWD {mkt_val:,.0f} | '
              f'P&L {pnl:+,.0f} ({ret_pct*100:+.2f}%) | '
              f'Today {change_pct*100:+.2f}%')
        time.sleep(0.2)  # be polite to Notion API

# ── Update Portfolio Tracker page ───────────────────────────────────────────────

def rich_text(content, bold=False):
    obj = {'type': 'text', 'text': {'content': str(content)}}
    if bold:
        obj['annotations'] = {'bold': True}
    return obj

def update_tracker_page(holdings, fx):
    today_str = datetime.now(TZ_TAIPEI).strftime('%Y-%m-%d')

    # Aggregate by market from updated Holdings data
    totals = {m: {'mkt': 0.0, 'cost': 0.0} for m in ('US', 'TW', 'JP')}
    for page in holdings:
        props  = page['properties']
        market = (props.get('Market', {}).get('select') or {}).get('name', '')
        mkt_v  = props.get('Market Value TWD', {}).get('number') or 0
        cost_v = props.get('Cost TWD', {}).get('number') or 0
        if market in totals:
            totals[market]['mkt']  += mkt_v
            totals[market]['cost'] += cost_v

    grand_mkt  = sum(v['mkt']  for v in totals.values())
    grand_cost = sum(v['cost'] for v in totals.values())
    grand_pnl  = grand_mkt - grand_cost
    flags = {'US': '🇺🇸', 'TW': '🇹🇼', 'JP': '🇯🇵'}

    # ── Update heading_2 and paragraph in-place; update table rows in-place ──
    children = notion_get(f'/blocks/{TRACKER_PAGE}/children')['results']

    for blk in children:
        btype = blk['type']
        bid   = blk['id']

        if btype == 'heading_2':
            notion_patch(f'/blocks/{bid}', {'heading_2': {
                'rich_text': [rich_text(f'📊 市場配置 Portfolio Allocation — Updated {today_str}')]
            }})

        elif btype == 'paragraph':
            text = blk.get('paragraph', {}).get('rich_text', [])
            raw = ''.join(t.get('plain_text', '') for t in text)
            if '總市值' in raw or 'Total Portfolio' in raw:
                notion_patch(f'/blocks/{bid}', {'paragraph': {
                    'rich_text': [rich_text(
                        f'總市值 Total Portfolio Value: TWD {grand_mkt:,.0f}',
                        bold=True
                    )]
                }})

        elif btype == 'table':
            rows = notion_get(f'/blocks/{bid}/children')['results']
            mkt_order = ['US', 'TW', 'JP']
            for i, mkt in enumerate(mkt_order):
                if i + 1 >= len(rows):
                    break
                v   = totals[mkt]
                pct = (v['mkt'] / grand_mkt * 100) if grand_mkt else 0
                pnl = v['mkt'] - v['cost']
                notion_patch(f'/blocks/{rows[i+1]["id"]}', {'table_row': {'cells': [
                    [rich_text(f'{flags[mkt]} {mkt}')],
                    [rich_text(f'{v["mkt"]:,.0f}')],
                    [rich_text(f'{pct:.1f}%')],
                    [rich_text(f'{v["cost"]:,.0f}')],
                    [rich_text(f'{pnl:+,.0f}')],
                ]}})
            if len(rows) >= 5:
                notion_patch(f'/blocks/{rows[4]["id"]}', {'table_row': {'cells': [
                    [rich_text('Total', bold=True)],
                    [rich_text(f'{grand_mkt:,.0f}', bold=True)],
                    [rich_text('100%', bold=True)],
                    [rich_text(f'{grand_cost:,.0f}', bold=True)],
                    [rich_text(f'{grand_pnl:+,.0f}', bold=True)],
                ]}})

    print(f'\n📊 Tracker page updated:')
    for mkt in ('US', 'TW', 'JP'):
        v = totals[mkt]
        pnl = v['mkt'] - v['cost']
        pct = (v['mkt'] / grand_mkt * 100) if grand_mkt else 0
        print(f'   {flags[mkt]} {mkt}: TWD {v["mkt"]:>14,.0f}  ({pct:5.1f}%)  P&L {pnl:+,.0f}')
    print(f'   {"Total":>6}: TWD {grand_mkt:>14,.0f}  (100.0%)  P&L {grand_pnl:+,.0f}')

# ── Main ────────────────────────────────────────────────────────────────────────

def main():
    now = datetime.now(TZ_TAIPEI)
    print(f'=== Portfolio Update {now.strftime("%Y-%m-%d %H:%M")} Taipei ===\n')

    print('Fetching FX rates...')
    fx = fetch_fx_rates()

    print('\nQuerying Holdings database...')
    holdings = query_database(DATABASE_ID)
    print(f'Found {len(holdings)} holding(s)\n')

    print('Updating Holdings rows...')
    update_holdings(holdings, fx)

    print('\nRe-fetching Holdings for tracker totals...')
    holdings = query_database(DATABASE_ID)

    print('\nUpdating Portfolio Tracker page...')
    update_tracker_page(holdings, fx)

    print('\n✅ Done!')

if __name__ == '__main__':
    main()
