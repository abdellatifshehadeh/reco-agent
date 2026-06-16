"""
Cash Reconciliation Agent — Streamlit UI
Custodian vs ZagTrader, multi-period, multi-custodian.
"""
import io
import re
from datetime import datetime, date
import pandas as pd
import streamlit as st

# ── Page config ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Cash Reconciliation Agent",
    page_icon="🔄",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Imports (local) ──────────────────────────────────────────────────────────
import sys, os
sys.path.insert(0, os.path.dirname(__file__))

from parsers.registry import detect_parser, CUSTODIAN_PARSERS
from core.zag_parser import parse_raw as zag_parse_raw, net as zag_net
from core.engine import reconcile, balance_summary, apply_manual_matches
from core.report import build_report
from core.review_parser import parse_review_file


# ── Styles ───────────────────────────────────────────────────────────────────
st.markdown("""
<style>
.metric-box {
    background: #f8f9fa;
    border: 0.5px solid #dee2e6;
    border-radius: 8px;
    padding: 1rem 1.25rem;
    text-align: center;
}
.metric-label { font-size: 13px; color: #6c757d; margin-bottom: 4px; }
.metric-value { font-size: 28px; font-weight: 500; }
.green { color: #198754; }
.amber { color: #fd7e14; }
.red   { color: #dc3545; }
.tag {
    display: inline-block;
    padding: 2px 10px;
    border-radius: 12px;
    font-size: 12px;
    font-weight: 500;
}
</style>
""", unsafe_allow_html=True)


# ── Sidebar — configuration ───────────────────────────────────────────────────
with st.sidebar:
    st.title("⚙️ Configuration")
    st.markdown("---")

    st.subheader("Period")
    period_label = st.text_input("Period label", value=f"{datetime.today().strftime('%b %Y')}")
    run_mode = st.selectbox("Mode", ["Single period", "Multi-period (Jan–May)"])

    st.markdown("---")
    st.subheader("Matching tolerances")
    income_tol = st.slider("Income / dividend delay (days)", 5, 30, 15,
                            help="Max days between custodian and ZAG posting for dividends/coupons")
    settle_tol = st.slider("Settlement tolerance (days)", 1, 10, 5,
                            help="Max days between trade and settlement dates")
    amt_tol    = st.slider("Amount tolerance (currency units)", 0.01, 5.0, 0.05, step=0.01,
                            help="Two amounts considered equal if diff < this value")

    st.markdown("---")
    st.subheader("Available custodians")
    for p in CUSTODIAN_PARSERS:
        st.markdown(f"✓ {p.name}")
    st.caption("New custodians: add a parser class in parsers/")


# ── Main area ─────────────────────────────────────────────────────────────────
st.title("🔄 Cash Reconciliation Agent")
st.caption("Upload custodian statement(s) + ZagTrader export → download reconciliation report")

tab_run, tab_review, tab_help = st.tabs(["▶ Run reconciliation", "📋 Apply manual matches", "❓ How to use"])


# ═══════════════════════════════════════════════════════════════════════════════
# TAB 1 — RUN
# ═══════════════════════════════════════════════════════════════════════════════
with tab_run:
    col_left, col_right = st.columns(2)

    with col_left:
        st.subheader("📁 Custodian statement(s)")
        cust_files = st.file_uploader(
            "Upload one or more custodian files",
            type=["xls", "xlsx"],
            accept_multiple_files=True,
            key="cust_upload",
            help="FAB: DDMMYYYY_RPT_STATEMENTACCOUNT_*.xls  |  Other custodians: add parser in parsers/"
        )
        if cust_files:
            for f in cust_files:
                parser = detect_parser(f.name)
                if parser:
                    st.success(f"✓ {f.name}  →  {parser.name}")
                else:
                    st.warning(f"⚠ {f.name}  →  No parser found — will skip")

    with col_right:
        st.subheader("📁 ZagTrader export")
        zag_file = st.file_uploader(
            "Upload ZagTrader Statement of Account (.xlsx)",
            type=["xlsx"],
            key="zag_upload",
        )
        if zag_file:
            st.success(f"✓ {zag_file.name}")

    st.markdown("---")

    # Optional: previously annotated review file
    st.subheader("📋 Previous review file (optional)")
    st.caption("Upload a report from a previous run where you added Match IDs — they will be applied automatically.")
    prev_files = st.file_uploader(
        "Previous reconciliation report(s) with match IDs",
        type=["xlsx"],
        accept_multiple_files=True,
        key="prev_upload",
    )

    st.markdown("---")
    run_btn = st.button("▶ Run reconciliation", type="primary",
                         disabled=(not cust_files or not zag_file))

    if run_btn and cust_files and zag_file:
        with st.spinner("Running reconciliation..."):
            try:
                # 1. Parse custodian files
                all_cust_dfs   = []
                all_cust_bals  = {}
                custodian_name = "Custodian"

                for f in cust_files:
                    parser = detect_parser(f.name)
                    if not parser:
                        st.warning(f"Skipped {f.name} — no matching parser")
                        continue
                    custodian_name = parser.name
                    f.seek(0)
                    # Save to temp file for xlrd
                    tmp_path = f"/tmp/{f.name}"
                    with open(tmp_path, "wb") as tmp:
                        tmp.write(f.read())
                    df, bal = parser.parse(tmp_path)
                    df = df[df['currency'] != 'ZAR'].reset_index(drop=True)
                    all_cust_dfs.append(df)
                    for ccy, info in bal.items():
                        existing = all_cust_bals.get(ccy, {})
                        if 'starting' not in existing and 'starting' in info:
                            existing['starting'] = info['starting']
                        if 'ending' in info:
                            existing['ending'] = info.get('ending')
                        all_cust_bals[ccy] = existing

                if not all_cust_dfs:
                    st.error("No custodian files could be parsed.")
                    st.stop()

                cust_df = pd.concat(all_cust_dfs, ignore_index=True)

                # 2. Parse ZagTrader
                zag_file.seek(0)
                tmp_zag = "/tmp/zag_upload.xlsx"
                with open(tmp_zag, "wb") as tmp:
                    tmp.write(zag_file.read())
                zag_raw, zag_bal = zag_parse_raw(tmp_zag)
                zag_netted = zag_net(zag_raw)

                # 3. Slice ZAG by period if multi-period mode
                # (for single period: use all)
                zag_for_match = zag_netted

                # 4. Reconcile
                recon_df   = reconcile(cust_df, zag_for_match)
                balance_df = balance_summary(all_cust_bals, zag_bal)
                # Rename balance columns to match report builder
                balance_df = balance_df.rename(columns={
                    'cust_opening': 'cust_opening',
                    'cust_closing': 'cust_closing',
                })

                # 5. Apply previous match groups if uploaded
                all_match_groups = []
                for pf in (prev_files or []):
                    pf.seek(0)
                    groups = parse_review_file(pf)
                    # Offset IDs to avoid collisions between files
                    offset = len(all_match_groups) * 1000
                    for g in groups:
                        g['match_id'] = g['match_id'] + offset
                    all_match_groups.extend(groups)

                manual_df = None
                if all_match_groups:
                    recon_df, manual_df = apply_manual_matches(recon_df, all_match_groups)

                # 6. Stats
                matched  = recon_df[recon_df['match_status'] == 'MATCHED']
                breaks   = recon_df[recon_df['match_status'].str.startswith('UNMATCHED')]
                total    = len(recon_df)
                match_rt = round(len(matched) / total * 100, 1) if total else 0
                bal_ok   = len(balance_df[balance_df['status'] == 'OK'])
                bal_brk  = len(balance_df[balance_df['status'] == 'BREAK'])

                # Display metrics
                m1, m2, m3, m4, m5 = st.columns(5)
                with m1:
                    st.metric("Total items", total)
                with m2:
                    st.metric("Matched", len(matched), delta=f"{match_rt}%")
                with m3:
                    st.metric("Breaks", len(breaks),
                               delta=f"{len(breaks[breaks['match_status']=='UNMATCHED - CUSTODIAN ONLY'])} cust / {len(breaks[breaks['match_status']=='UNMATCHED - ZAG ONLY'])} ZAG",
                               delta_color="inverse")
                with m4:
                    st.metric("Bal. OK", bal_ok)
                with m5:
                    st.metric("Bal. breaks", bal_brk, delta_color="inverse")

                st.markdown("---")

                # Balance summary preview
                st.subheader("Balance comparison")
                bal_display = balance_df[~balance_df['status'].eq('ONE SIDE MISSING')].copy()
                def colour_status(val):
                    if val == 'OK':    return 'background-color: #C6EFCE'
                    if val == 'BREAK': return 'background-color: #FFC7CE'
                    return ''
                st.dataframe(
                    bal_display[['currency', 'cust_opening', 'cust_closing',
                                 'zag_opening', 'zag_closing', 'difference', 'status']]
                    .rename(columns={'cust_opening': 'Cust Open', 'cust_closing': 'Cust Close',
                                     'zag_opening': 'ZAG Open', 'zag_closing': 'ZAG Close',
                                     'difference': 'Diff', 'currency': 'CCY', 'status': 'Status'})
                    .style.applymap(colour_status, subset=['Status']),
                    use_container_width=True, hide_index=True
                )

                # Top breaks preview
                if len(breaks) > 0:
                    st.subheader(f"Top unmatched items ({min(20, len(breaks))} of {len(breaks)})")
                    preview_cols = ['currency', 'match_status', 'cust_settle_date',
                                    'cust_description', 'cust_amount',
                                    'zag_date', 'zag_description', 'zag_amount', 'notes']
                    avail = [c for c in preview_cols if c in breaks.columns]
                    st.dataframe(breaks[avail].head(20), use_container_width=True, hide_index=True)

                # 7. Build report
                report_bytes = build_report(
                    recon_df=recon_df,
                    balance_df=balance_df,
                    custodian_df=cust_df,
                    zag_raw_df=zag_raw,
                    zag_netted_df=zag_netted,
                    manual_groups=all_match_groups,
                    period_label=period_label,
                    custodian_name=custodian_name,
                    manual_df=manual_df,
                )

                fname = f"Recon_{period_label.replace(' ','_')}_{datetime.today().strftime('%Y%m%d_%H%M')}.xlsx"
                st.download_button(
                    label="⬇ Download reconciliation report (.xlsx)",
                    data=report_bytes,
                    file_name=fname,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                )
                st.success(f"✓ Report ready — {len(matched)} matched, {len(breaks)} breaks, {match_rt}% match rate")

            except Exception as e:
                st.error(f"Error: {e}")
                import traceback
                st.code(traceback.format_exc())


# ═══════════════════════════════════════════════════════════════════════════════
# TAB 2 — APPLY MANUAL MATCHES
# ═══════════════════════════════════════════════════════════════════════════════
with tab_review:
    st.subheader("Apply manual match IDs from a reviewed report")
    st.markdown("""
    **Workflow:**
    1. Download the report from the Run tab
    2. Open the **Unmatched Side-by-Side** sheet
    3. For each pair you recognise, add the same number in column **G** (custodian side) and column **P** (ZAG side)
    4. Upload the annotated file below, then re-run from the Run tab — matches will be applied automatically
    """)

    review_file = st.file_uploader("Upload annotated report", type=["xlsx"], key="review_upload")
    if review_file:
        review_file.seek(0)
        groups = parse_review_file(review_file)
        st.success(f"✓ Found {len(groups)} match groups")

        if groups:
            summary_rows = []
            for g in groups:
                fab_t = sum(float(r['fab_amt']) for r in g.get('fab_rows', []) if r.get('fab_amt') is not None)
                zag_t = sum(float(r['zag_amt']) for r in g.get('zag_rows', []) if r.get('zag_amt') is not None)
                net   = round(fab_t + zag_t, 2)
                flagged = abs(net) > 100 and abs(net) / max(abs(fab_t), abs(zag_t), 0.01) > 0.01
                summary_rows.append({
                    'CCY': g['ccy'],
                    'Match ID': g['match_id'],
                    'Cust lines': len(g.get('fab_rows', [])),
                    'ZAG lines': len(g.get('zag_rows', [])),
                    'Cust total': round(fab_t, 2),
                    'ZAG total': round(zag_t, 2),
                    'Net diff': net,
                    '⚠ Flag': '⚠ VERIFY' if flagged else '✓',
                })
            df_summary = pd.DataFrame(summary_rows)
            st.dataframe(df_summary, use_container_width=True, hide_index=True)

            flagged_count = len(df_summary[df_summary['⚠ Flag'] == '⚠ VERIFY'])
            if flagged_count:
                st.warning(f"⚠ {flagged_count} group(s) have large differences — please verify before accepting")


# ═══════════════════════════════════════════════════════════════════════════════
# TAB 3 — HELP
# ═══════════════════════════════════════════════════════════════════════════════
with tab_help:
    st.subheader("How to use this agent")
    st.markdown("""
    ### Quick start
    1. **Upload** your custodian statement(s) and ZagTrader export in the **Run** tab
    2. Set the period label (e.g. `Jan 2026` or `Jan–May 2026`)
    3. Click **▶ Run reconciliation**
    4. Review the metrics and preview, then **download the report**

    ### Matching logic
    The engine applies four steps before matching:

    | Step | What it does |
    |------|-------------|
    | **Gross + WHT netting** | ZagTrader posts dividend/coupon as gross + tax debit; custodian posts net only. Engine combines them. |
    | **Coupon split aggregation** | Same security split across two ZAG journal lines → collapsed to one |
    | **Transfer + charge collapsing** | Wire transfer + bank charge on same journal → collapsed to principal amount |
    | **FD split** | FD sale + interest on same journal → combined |

    **Date tolerances** (adjustable in sidebar):
    - Income / dividends: **15 days** (common posting delay between custodian and ZAG)
    - Settlements / wires: **5 days**

    ### Manual matching (iterative review)
    When items remain unmatched after the engine, use the side-by-side sheet:
    1. Open the **Unmatched Side-by-Side** tab in the downloaded report
    2. Find a custodian row (left, salmon) and its ZAG counterpart (right, pink)
    3. Type the same number in **column G** (custodian Match ID) and **column P** (ZAG Match ID)
    4. For one-to-many: give multiple ZAG rows the same ID as one custodian row
    5. Save and upload the file in the **Apply manual matches** tab, then re-run

    ### Adding a new custodian
    1. Create `parsers/your_bank.py` subclassing `CustodianParser`
    2. Implement `parse(filepath)` returning `(df, balance_info)`
    3. Add `YourBankParser` to the `CUSTODIAN_PARSERS` list in `parsers/registry.py`
    4. Restart the app — it will appear in the custodian list automatically

    ### Running on a schedule
    ```bash
    # Windows Task Scheduler or cron — run every Monday at 8am:
    streamlit run app.py
    # Or headless (generates report to disk without UI):
    python run_headless.py --cust /path/to/fab.xls --zag /path/to/zag.xlsx --period "Jan 2026"
    ```
    """)
