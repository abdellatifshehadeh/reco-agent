"""
Core reconciliation matching engine.
Matches custodian transactions against ZagTrader netted transactions.

Matching passes:
  1. Exact amount + same settlement date
  2. Amount match + type-aware date tolerance
     - Income / dividend: 15 days (posting delay)
     - Settlement / transfer: 5 days
"""
import re
import numpy as np
import pandas as pd


# ── Narrative helpers ──────────────────────────────────────────────────────

def _clean(text: str) -> str:
    t = str(text or '').upper()
    for pat in [
        r'\bARK CAPITAL MANAGEMENT \(DUBAI\)\b', r'\bARK CAPITAL MANAGEMENT\b',
        r'\bLTD\b', r'\bLIMITED\b', r'\bT24\b', r'\bSETTLEMENT OF TRADEID\b',
        r'\bCREDIT COUPON TICKET FOR\b', r'\bCASH DIVIDEND FOR\b',
        r'\bCA INCOME \(INTR\)\b', r'\bCA DIVIDEND\b', r'\bCA INCOME\b',
        r'#\s*\d+\s*EXD[:\s]+[\d\.]+',
        r'[#\-_/\.]', r'\s+',
    ]:
        t = re.sub(pat, ' ', t)
    return t.strip()


_TYPE_MAP = {
    'RECEIVE VS PAYMENT': ['SETTLEMENT', 'TRADEID', 'RECEIVE'],
    'DELIVER VS PAYMENT': ['SETTLEMENT', 'TRADEID', 'DELIVER'],
    'CASH DEPOSIT':       ['TRANSFER FROM', 'WIRE IN', 'DEPOSIT'],
    'CASH WITHDRAWAL':    ['WIRE OUT', 'TRANSFER FROM FAB', 'TRANSFER TO'],
    'CA PAY DUE':         ['BOND MATURED', 'FI REDM', 'FI MCAL', 'MATURITY'],
    'CA DIVIDEND':        ['CASH DIVIDEND', 'DIVIDEND', 'DIV'],
    'CA INCOME':          ['CREDIT COUPON', 'COUPON', 'INTR', 'INCOME'],
}

_INCOME_TYPES    = {'CA DIVIDEND', 'CA INCOME', 'CA INCOME (INTR)', 'CA PAY DUE'}
_INCOME_ZAG_KEYS = ('DIVIDEND', 'COUPON', 'INTR', 'INCOME', 'CREDIT COUPON')


def _narrative_score(fab_desc, fab_type, zag_desc, zag_ref, zag_comp=''):
    fab_key = _clean(fab_desc + ' ' + fab_type)
    zag_key = _clean(str(zag_desc) + ' ' + str(zag_ref or '') + ' ' + str(zag_comp or ''))

    fab_type_key = _clean(fab_type)
    for key, hints in _TYPE_MAP.items():
        if key in fab_type_key and any(h in zag_key for h in hints):
            return 'HIGH', 'type_map'

    fab_words = {w for w in fab_key.split() if len(w) > 3}
    zag_words = {w for w in zag_key.split() if len(w) > 3}
    overlap = fab_words & zag_words
    if len(overlap) >= 2:
        return 'HIGH', f'words:{",".join(sorted(overlap)[:3])}'
    if len(overlap) == 1:
        return 'MEDIUM', f'words:{",".join(overlap)}'
    return 'LOW', 'no_match'


def _date_tolerance(fab_txn_type, zag_desc):
    fab_t = str(fab_txn_type or '').upper()
    zag_d = str(zag_desc or '').upper()
    is_income = (
        any(t in fab_t for t in _INCOME_TYPES) or
        any(k in zag_d for k in _INCOME_ZAG_KEYS)
    )
    return 15 if is_income else 5


# ── Signed amount helper ────────────────────────────────────────────────────

def _signed(row) -> float:
    d = row.get('debit');  c = row.get('credit')
    if d is not None and not (isinstance(d, float) and np.isnan(d)): return -abs(float(d))
    if c is not None and not (isinstance(c, float) and np.isnan(c)): return  abs(float(c))
    return 0.0


# ── Main reconcile function ─────────────────────────────────────────────────

def reconcile(custodian_df: pd.DataFrame, zag_netted_df: pd.DataFrame) -> pd.DataFrame:
    """
    Match custodian transactions against ZagTrader netted rows.
    Returns a DataFrame with one row per item (matched or unmatched).
    """
    results = []
    currencies = sorted(
        set(custodian_df['currency'].dropna()) |
        set(zag_netted_df['currency'].dropna())
    )

    for ccy in currencies:
        cust = custodian_df[custodian_df['currency'] == ccy].copy().reset_index(drop=True)
        zag  = zag_netted_df[zag_netted_df['currency'] == ccy].copy().reset_index(drop=True)
        cust_used = [False] * len(cust)
        zag_used  = [False] * len(zag)

        # Pass 1 — exact amount + same settlement date
        for ci, cr in cust.iterrows():
            c_amt  = _signed(cr)
            c_date = cr['settlement_date']
            for zi, zr in zag.iterrows():
                if zag_used[zi]: continue
                z_amt  = _signed(zr)
                z_date = zr['settlement_date']
                amt_ok  = abs(abs(c_amt) - abs(z_amt)) < 0.05
                date_ok = (
                    pd.notna(c_date) and pd.notna(z_date) and
                    pd.Timestamp(c_date).date() == pd.Timestamp(z_date).date()
                )
                if amt_ok and date_ok:
                    conf, method = _narrative_score(
                        cr.get('description', ''), cr.get('txn_type', ''),
                        zr.get('description', ''), zr.get('ref', ''),
                        zr.get('component_descs', ''))
                    results.append(_match_row(ccy, cr, zr, c_amt, z_amt, 'MATCHED',
                                              conf, f'exact_date+amt | {method}', ''))
                    cust_used[ci] = True; zag_used[zi] = True; break

        # Pass 2 — amount match + type-aware date tolerance
        for ci, cr in cust.iterrows():
            if cust_used[ci]: continue
            c_amt  = _signed(cr)
            c_date = cr.get('settlement_date') or cr.get('trade_date')
            for zi, zr in zag.iterrows():
                if zag_used[zi]: continue
                z_amt  = _signed(zr)
                z_date = zr.get('settlement_date') or zr.get('trade_date')
                amt_ok = abs(abs(c_amt) - abs(z_amt)) < 0.05
                try:
                    diff_days = abs((pd.Timestamp(c_date) - pd.Timestamp(z_date)).days) \
                                if pd.notna(c_date) and pd.notna(z_date) else 99
                except Exception:
                    diff_days = 99
                tol   = _date_tolerance(cr.get('txn_type', ''), zr.get('description', ''))
                label = 'income_posting_delay' if tol == 15 else 'trade_settle_timing'
                if amt_ok and diff_days <= tol:
                    conf, method = _narrative_score(
                        cr.get('description', ''), cr.get('txn_type', ''),
                        zr.get('description', ''), zr.get('ref', ''),
                        zr.get('component_descs', ''))
                    results.append(_match_row(ccy, cr, zr, c_amt, z_amt, 'MATCHED',
                                              conf, f'amt+{diff_days}d ({label}) | {method}',
                                              f'Posting date diff = {diff_days} days'))
                    cust_used[ci] = True; zag_used[zi] = True; break

        # Unmatched custodian
        for ci, cr in cust.iterrows():
            if not cust_used[ci]:
                results.append(_unmatched_cust(ccy, cr, _signed(cr)))

        # Unmatched ZAG
        for zi, zr in zag.iterrows():
            if not zag_used[zi]:
                results.append(_unmatched_zag(ccy, zr, _signed(zr)))

    return pd.DataFrame(results)


def balance_summary(custodian_bal: dict, zag_bal: dict) -> pd.DataFrame:
    all_ccys = sorted(set(list(custodian_bal.keys()) + list(zag_bal.keys())))
    rows = []
    for ccy in all_ccys:
        c_open  = custodian_bal.get(ccy, {}).get('starting')
        c_close = custodian_bal.get(ccy, {}).get('ending')
        z_open  = zag_bal.get(ccy, {}).get('starting')
        z_close = zag_bal.get(ccy, {}).get('ending')
        diff    = round(c_close - z_close, 4) if c_close is not None and z_close is not None else None
        status  = ('OK'               if diff is not None and abs(diff) < 0.05 else
                   'BREAK'            if diff is not None else
                   'ONE SIDE MISSING')
        rows.append(dict(currency=ccy, cust_opening=c_open, cust_closing=c_close,
                         zag_opening=z_open, zag_closing=z_close,
                         difference=diff, status=status))
    return pd.DataFrame(rows)


def apply_manual_matches(recon_df: pd.DataFrame,
                         match_groups: list[dict]) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    Apply user-confirmed match groups from the side-by-side review sheet.
    Returns (updated_recon_df, manual_matched_df).
    match_groups is a list of {'ccy', 'match_id', 'fab_rows', 'zag_rows'} dicts.
    """
    manual_fab_sigs = set()
    manual_zag_sigs = set()
    for g in match_groups:
        for r in g.get('fab_rows', []):
            if r.get('fab_desc') and r.get('fab_amt') is not None:
                manual_fab_sigs.add((str(r['fab_desc'])[:50], float(r['fab_amt'])))
        for r in g.get('zag_rows', []):
            if r.get('zag_desc') and r.get('zag_amt') is not None:
                manual_zag_sigs.add((str(r['zag_desc'])[:50], float(r['zag_amt'])))

    def is_resolved(row):
        if row['match_status'] == 'UNMATCHED - CUSTODIAN ONLY':
            return (str(row.get('cust_description') or '')[:50],
                    float(row.get('cust_amount') or 0)) in manual_fab_sigs
        else:
            return (str(row.get('zag_description') or '')[:50],
                    float(row.get('zag_amount') or 0)) in manual_zag_sigs

    breaks = recon_df[recon_df['match_status'].str.startswith('UNMATCHED')].copy()
    breaks['_res'] = breaks.apply(is_resolved, axis=1)
    remaining = breaks[~breaks['_res']].drop(columns=['_res'])
    matched   = recon_df[recon_df['match_status'] == 'MATCHED']

    manual_rows = []
    for g in match_groups:
        fab_t = sum(float(r['fab_amt']) for r in g.get('fab_rows', []) if r.get('fab_amt') is not None)
        zag_t = sum(float(r['zag_amt']) for r in g.get('zag_rows', []) if r.get('zag_amt') is not None)
        fab_b = g['fab_rows'][0] if g.get('fab_rows') else {}
        zag_b = g['zag_rows'][0] if g.get('zag_rows') else {}
        n_f   = len(g.get('fab_rows', []))
        n_z   = len(g.get('zag_rows', []))
        net   = round(abs(fab_t) - abs(zag_t), 2) if g.get('fab_rows') and g.get('zag_rows') else None
        is_flagged = net is not None and abs(net) > 100 and abs(net) / max(abs(fab_t), abs(zag_t), 0.01) > 0.01
        notes = f"Match ID {g['match_id']} | {n_f} custodian | {n_z} ZAG"
        if is_flagged:
            notes += f" | ⚠ LARGE DIFF {abs(net):,.2f} — VERIFY"
        manual_rows.append({
            'match_status': 'MATCHED',
            'confidence': 'REVIEW REQUIRED' if is_flagged else 'HIGH',
            'match_method': f"manual_id_{g['match_id']} ({n_f}C:{n_z}Z)",
            'currency': g['ccy'],
            'cust_settle_date': fab_b.get('fab_date'),
            'cust_type': fab_b.get('fab_type', ''),
            'cust_description': ' + '.join(str(r.get('fab_desc', ''))[:35] for r in g.get('fab_rows', [])),
            'cust_amount': fab_t if g.get('fab_rows') else None,
            'cust_balance': None,
            'zag_date': zag_b.get('zag_date'),
            'zag_ref': zag_b.get('zag_ref', ''),
            'zag_description': ' + '.join(str(r.get('zag_desc', ''))[:35] for r in g.get('zag_rows', [])),
            'zag_amount': zag_t if g.get('zag_rows') else None,
            'zag_balance': None,
            'zag_journal': zag_b.get('zag_jnl', ''),
            'zag_netted': False,
            'nett_type': '',
            'zag_gross': None,
            'zag_tax': None,
            'difference': net,
            'notes': notes,
        })

    manual_df = pd.DataFrame(manual_rows)
    combined  = pd.concat([matched, manual_df, remaining], ignore_index=True)
    return combined, manual_df


# ── Row builders ────────────────────────────────────────────────────────────

def _match_row(ccy, cr, zr, c_amt, z_amt, status, conf, method, note):
    return dict(
        currency=ccy, match_status=status, confidence=conf, match_method=method,
        cust_settle_date=cr.get('settlement_date'), cust_trade_date=cr.get('trade_date'),
        cust_type=cr.get('txn_type'), cust_description=cr.get('description'),
        cust_amount=round(c_amt, 2), cust_balance=cr.get('balance'),
        custodian=cr.get('custodian', ''),
        zag_date=zr.get('settlement_date'), zag_ref=zr.get('ref'),
        zag_description=zr.get('description'), zag_amount=round(z_amt, 2),
        zag_balance=zr.get('balance'), zag_journal=zr.get('journal_ref'),
        zag_netted=zr.get('netted', False), nett_type=zr.get('nett_type', ''),
        zag_gross=zr.get('zag_gross'), zag_tax=zr.get('zag_tax'),
        difference=round(abs(c_amt) - abs(z_amt), 4), notes=note,
    )


def _unmatched_cust(ccy, cr, c_amt):
    return dict(
        currency=ccy, match_status='UNMATCHED - CUSTODIAN ONLY',
        confidence='', match_method='',
        cust_settle_date=cr.get('settlement_date'), cust_trade_date=cr.get('trade_date'),
        cust_type=cr.get('txn_type'), cust_description=cr.get('description'),
        cust_amount=round(c_amt, 2), cust_balance=cr.get('balance'),
        custodian=cr.get('custodian', ''),
        zag_date=None, zag_ref=None, zag_description=None, zag_amount=None,
        zag_balance=None, zag_journal=None, zag_netted=False, nett_type='',
        zag_gross=None, zag_tax=None, difference=None,
        notes='In custodian statement only — not found in ZagTrader',
    )


def _unmatched_zag(ccy, zr, z_amt):
    return dict(
        currency=ccy, match_status='UNMATCHED - ZAG ONLY',
        confidence='', match_method='',
        cust_settle_date=None, cust_trade_date=None,
        cust_type=None, cust_description=None,
        cust_amount=None, cust_balance=None, custodian='',
        zag_date=zr.get('settlement_date'), zag_ref=zr.get('ref'),
        zag_description=zr.get('description'), zag_amount=round(z_amt, 2),
        zag_balance=zr.get('balance'), zag_journal=zr.get('journal_ref'),
        zag_netted=zr.get('netted', False), nett_type=zr.get('nett_type', ''),
        zag_gross=zr.get('zag_gross'), zag_tax=zr.get('zag_tax'),
        difference=None,
        notes='In ZagTrader only — not found in custodian statement',
    )
