"""
ZagTrader statement parser.
Handles multi-currency Excel exports (one sheet per currency + Cash Balance sheet).
Includes gross+tax netting, coupon split aggregation, transfer+charge collapsing.
"""
import numpy as np
import pandas as pd


def parse_raw(filepath: str) -> tuple[pd.DataFrame, dict]:
    """
    Parse ZagTrader Excel export.
    Returns (raw_df, balance_info).
    """
    xl = pd.ExcelFile(filepath)
    currency_sheets = [s for s in xl.sheet_names if s != 'Cash Balance']
    rows = []
    balance_info = {}

    for sheet in currency_sheets:
        df = pd.read_excel(xl, sheet_name=sheet, header=None)
        for i, row in df.iterrows():
            if i < 4:
                continue
            raw = list(row)

            # Open balance row
            if str(raw[3]).strip() == 'Open Balance':
                try:
                    balance_info.setdefault(sheet, {})['starting'] = float(raw[7])
                except Exception:
                    pass
                continue

            date_val = pd.to_datetime(str(raw[0]).strip(), format='%d %b %Y', errors='coerce')
            if pd.isnull(date_val):
                continue

            def to_num(v):
                try:
                    f = float(v)
                    return None if np.isnan(f) else f
                except Exception:
                    return None

            ref     = str(raw[2]).strip() if str(raw[2]).strip() not in ['nan', ''] else None
            desc    = str(raw[3]).strip() if str(raw[3]).strip() not in ['nan', ''] else ''
            debit   = to_num(raw[5])
            credit  = to_num(raw[6])
            balance = to_num(raw[7])
            jref    = str(raw[8]).strip() if str(raw[8]).strip() not in ['nan', ''] else None

            if debit is not None or credit is not None:
                rows.append({
                    'source': 'ZAG',
                    'currency': sheet,
                    'trade_date': date_val,
                    'settlement_date': date_val,
                    'ref': ref,
                    'description': desc,
                    'debit': debit,
                    'credit': credit,
                    'balance': balance,
                    'journal_ref': jref,
                    'is_tax_line': False,
                    'netted': False,
                })

    # Ending balances from Cash Balance sheet
    try:
        df_cb = pd.read_excel(xl, sheet_name='Cash Balance', header=None, skiprows=1)
        for _, row in df_cb.iterrows():
            try:
                ccy = str(row[0]).strip()
                bal = float(row[1])
                if len(ccy) == 3:
                    balance_info.setdefault(ccy, {})['ending'] = bal
            except Exception:
                pass
    except Exception:
        pass

    return pd.DataFrame(rows), balance_info


def net(zag_raw_df: pd.DataFrame) -> pd.DataFrame:
    """
    Collapse multi-line ZagTrader journal entries into single net rows before matching.

    Handles:
      - Pattern A: gross income/dividend + WHT tax lines (same journal_ref)
      - Pattern B: coupon notional splits (same security, same direction)
      - Pattern C: transfer + bank charge (one large + one small ≤100)
      - Pattern D: FD sale + interest components
    """
    df = zag_raw_df.copy()
    df['is_tax_line'] = (
        df['description'].str.startswith('TAX for:', na=False) |
        df['description'].str.startswith('Tax for:', na=False)
    )

    netted_rows = []
    used = set()

    grouped = df[df['journal_ref'].notna()].groupby(['currency', 'journal_ref'])

    for (ccy, jref), grp in grouped:
        idx_list = list(grp.index)
        if len(grp) < 2:
            continue

        gross_rows = grp[~grp['is_tax_line']]
        tax_rows   = grp[grp['is_tax_line']]

        # ── Pattern A: gross + WHT tax ──
        if len(tax_rows) > 0:
            gross_debit  = gross_rows['debit'].fillna(0).sum()
            gross_credit = gross_rows['credit'].fillna(0).sum()
            tax_credit   = tax_rows['credit'].fillna(0).sum()
            tax_debit    = tax_rows['debit'].fillna(0).sum()
            net_debit    = gross_debit  - tax_credit
            net_credit   = gross_credit - tax_debit
            base = gross_rows.iloc[0]
            netted_rows.append({
                **_base_row(base, ccy, jref),
                'debit':  round(net_debit,  4) if net_debit  > 0.001 else None,
                'credit': round(net_credit, 4) if net_credit > 0.001 else None,
                'balance': _last_balance(grp),
                'zag_gross': gross_debit if gross_debit else gross_credit,
                'zag_tax':   tax_credit  if tax_credit  else tax_debit,
                'nett_type': 'gross+tax',
                'component_count': len(grp),
                'component_descs': ' | '.join(grp['description'].tolist()),
            })
            for i in idx_list:
                used.add(i)
            continue

        # ── Patterns B / C / D (no tax lines) ──
        all_debits  = gross_rows['debit'].fillna(0).sum()
        all_credits = gross_rows['credit'].fillna(0).sum()

        def row_amt(r):
            d = r['debit']; c = r['credit']
            if d is not None and not (isinstance(d, float) and np.isnan(d)): return abs(float(d))
            if c is not None and not (isinstance(c, float) and np.isnan(c)): return abs(float(c))
            return 0.0

        # Pattern B — coupon notional split (same description prefix)
        descriptions = gross_rows['description'].tolist()
        first_desc   = str(descriptions[0]).upper()[:50]
        all_same_sec = all(str(d).upper()[:50] == first_desc for d in descriptions)

        # Pattern C — transfer + bank charge
        sorted_amts = sorted([row_amt(r) for _, r in gross_rows.iterrows()], reverse=True)
        is_charge = (
            len(sorted_amts) == 2 and
            sorted_amts[0] > 0 and
            sorted_amts[1] <= 100 and
            sorted_amts[0] / max(sorted_amts[1], 0.01) > 100
        )

        # Pattern D — FD/bond: principal + interest
        descs_upper = [str(d).upper() for d in gross_rows['description']]
        is_fd = (
            any('SALE OF SECURITY' in d or 'BOND MATURED' in d or
                'BOND EARLY REDEMPTION' in d for d in descs_upper) and
            any('INTEREST' in d for d in descs_upper)
        )

        if all_same_sec or is_charge or is_fd:
            base = gross_rows.iloc[0]
            net_debit  = round(all_debits,  4) if all_debits  > 0.001 else None
            net_credit = round(all_credits, 4) if all_credits > 0.001 else None

            if is_charge:
                main_amt = sorted_amts[0]
                main_row = max(gross_rows.iterrows(), key=lambda x: row_amt(x[1]))[1]
                if main_row['debit'] is not None and not (isinstance(main_row['debit'], float) and np.isnan(main_row['debit'])):
                    net_debit  = round(main_amt, 4)
                    net_credit = None
                else:
                    net_credit = round(main_amt, 4)
                    net_debit  = None

            nett_type = 'coupon_split' if all_same_sec else 'transfer_charge' if is_charge else 'fd_split'
            netted_rows.append({
                **_base_row(base, ccy, jref),
                'debit':  net_debit,
                'credit': net_credit,
                'balance': _last_balance(grp),
                'zag_gross': all_debits if all_debits else all_credits,
                'zag_tax': None,
                'nett_type': nett_type,
                'component_count': len(grp),
                'component_descs': ' | '.join(gross_rows['description'].tolist()),
            })
            for i in idx_list:
                used.add(i)

    # Add all non-netted rows
    for idx, row in df.iterrows():
        if idx in used:
            continue
        r = row.to_dict()
        r.setdefault('netted', False)
        r.setdefault('zag_gross', None)
        r.setdefault('zag_tax', None)
        r.setdefault('nett_type', None)
        r.setdefault('component_count', 1)
        r.setdefault('component_descs', r.get('description', ''))
        netted_rows.append(r)

    return pd.DataFrame(netted_rows).reset_index(drop=True)


def _base_row(base_series, ccy, jref):
    return {
        'source': 'ZAG',
        'currency': ccy,
        'trade_date': base_series['trade_date'],
        'settlement_date': base_series['settlement_date'],
        'ref': base_series.get('ref'),
        'description': base_series['description'],
        'journal_ref': jref,
        'is_tax_line': False,
        'netted': True,
    }


def _last_balance(grp):
    vals = grp['balance'].dropna()
    return vals.iloc[-1] if len(vals) else None
