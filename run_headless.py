"""
Headless runner — use this for scheduled/automated runs without the Streamlit UI.

Usage:
    python run_headless.py \\
        --cust  /path/to/fab_jan2026.xls \\
        --zag   /path/to/zag_jan2026.xlsx \\
        --period "Jan 2026" \\
        --out   /path/to/output/

Multiple custodian files:
    python run_headless.py \\
        --cust fab_jan.xls fab_feb.xls fab_mar.xls \\
        --zag  zag_q1.xlsx \\
        --period "Q1 2026"

With previous review file:
    python run_headless.py \\
        --cust fab_may.xls \\
        --zag  zag_may.xlsx \\
        --period "May 2026" \\
        --review prev_report_annotated.xlsx
"""
import argparse
import sys
import os
from datetime import datetime
from pathlib import Path

sys.path.insert(0, os.path.dirname(__file__))

import pandas as pd
from parsers.registry import detect_parser
from core.zag_parser import parse_raw as zag_parse_raw, net as zag_net
from core.engine import reconcile, balance_summary, apply_manual_matches
from core.report import build_report
from core.review_parser import parse_review_file


def run(cust_files: list[str], zag_file: str, period: str,
        out_dir: str = ".", review_files: list[str] = None) -> str:

    print(f"[{datetime.now():%H:%M:%S}] Starting reconciliation — {period}")

    # ── Parse custodian files ──
    all_cust_dfs  = []
    all_cust_bals = {}
    custodian_name = "Custodian"

    for path in cust_files:
        parser = detect_parser(Path(path).name)
        if not parser:
            print(f"  ⚠ No parser for {path} — skipping")
            continue
        custodian_name = parser.name
        df, bal = parser.parse(path)
        df = df[df['currency'] != 'ZAR'].reset_index(drop=True)
        all_cust_dfs.append(df)
        for ccy, info in bal.items():
            existing = all_cust_bals.get(ccy, {})
            if 'starting' not in existing: existing['starting'] = info.get('starting')
            existing['ending'] = info.get('ending')
            all_cust_bals[ccy] = existing
        print(f"  ✓ Parsed {Path(path).name}  [{parser.name}]  {len(df)} transactions")

    if not all_cust_dfs:
        raise ValueError("No custodian files could be parsed")

    cust_df = pd.concat(all_cust_dfs, ignore_index=True)

    # ── Parse ZagTrader ──
    zag_raw, zag_bal = zag_parse_raw(zag_file)
    zag_netted = zag_net(zag_raw)
    print(f"  ✓ Parsed ZagTrader  {len(zag_raw)} raw → {len(zag_netted)} netted rows")

    # ── Reconcile ──
    recon_df   = reconcile(cust_df, zag_netted)
    balance_df = balance_summary(all_cust_bals, zag_bal)

    # ── Apply manual matches ──
    all_match_groups = []
    for rf in (review_files or []):
        groups = parse_review_file(rf)
        offset = len(all_match_groups) * 1000
        for g in groups: g['match_id'] = g['match_id'] + offset
        all_match_groups.extend(groups)
        print(f"  ✓ Loaded {len(groups)} match groups from {Path(rf).name}")

    manual_df = None
    if all_match_groups:
        recon_df, manual_df = apply_manual_matches(recon_df, all_match_groups)

    # ── Stats ──
    matched  = recon_df[recon_df['match_status'] == 'MATCHED']
    breaks   = recon_df[recon_df['match_status'].str.startswith('UNMATCHED')]
    total    = len(recon_df)
    rate     = round(len(matched) / total * 100, 1) if total else 0
    bal_ok   = len(balance_df[balance_df['status'] == 'OK'])
    bal_brk  = len(balance_df[balance_df['status'] == 'BREAK'])

    print(f"\n  📊 Results:")
    print(f"     Total items:  {total}")
    print(f"     Matched:      {len(matched)}  ({rate}%)")
    print(f"     Breaks:       {len(breaks)}")
    print(f"     Balance OK:   {bal_ok}  |  Break: {bal_brk}")

    # ── Build and save report ──
    report_bytes = build_report(
        recon_df=recon_df,
        balance_df=balance_df,
        custodian_df=cust_df,
        zag_raw_df=zag_raw,
        zag_netted_df=zag_netted,
        manual_groups=all_match_groups,
        period_label=period,
        custodian_name=custodian_name,
        manual_df=manual_df,
    )

    fname = f"Recon_{period.replace(' ', '_')}_{datetime.now():%Y%m%d_%H%M}.xlsx"
    out_path = os.path.join(out_dir, fname)
    with open(out_path, 'wb') as f:
        f.write(report_bytes)

    print(f"\n  ✅ Report saved: {out_path}")
    return out_path


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Cash reconciliation headless runner')
    parser.add_argument('--cust',   nargs='+', required=True, help='Custodian file(s)')
    parser.add_argument('--zag',    required=True, help='ZagTrader file')
    parser.add_argument('--period', required=True, help='Period label e.g. "Jan 2026"')
    parser.add_argument('--out',    default='.', help='Output directory')
    parser.add_argument('--review', nargs='*', help='Previous annotated report(s)')
    args = parser.parse_args()

    run(
        cust_files=args.cust,
        zag_file=args.zag,
        period=args.period,
        out_dir=args.out,
        review_files=args.review or [],
    )
