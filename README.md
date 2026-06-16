# Cash Reconciliation Agent

Custodian vs ZagTrader cash reconciliation — Streamlit web app.

## What it does

- Parses custodian statements (FAB today, others by adding a parser)
- Parses ZagTrader multi-currency exports
- Nets ZAG gross+WHT tax lines, coupon splits, transfer+charge, FD components
- Matches transactions using amount + date with type-aware tolerances
- Generates a full Excel report with matched, breaks, side-by-side, netting detail
- Accepts manual match IDs from annotated reports (iterative review loop)

## Setup

```bash
# 1. Clone / copy this folder
cd recon_agent

# 2. Install dependencies
pip install -r requirements.txt

# 3. Launch the app
streamlit run app.py
```

Open http://localhost:8501 in your browser.

## Scheduled / automated runs (headless)

```bash
# Single month
python run_headless.py \
    --cust  /data/fab/01022026_RPT_STATEMENTACCOUNT_Report_7938.xls \
    --zag   /data/zag/Statement_Jan2026.xlsx \
    --period "Jan 2026" \
    --out   /data/reports/

# Multi-month (one ZAG file covering the period)
python run_headless.py \
    --cust  fab_jan.xls fab_feb.xls fab_mar.xls \
    --zag   zag_q1.xlsx \
    --period "Q1 2026" \
    --out   /data/reports/

# With previous annotated review file
python run_headless.py \
    --cust  fab_may.xls \
    --zag   zag_may.xlsx \
    --period "May 2026" \
    --review prev_Recon_Apr_2026_annotated.xlsx \
    --out   /data/reports/
```

### Windows Task Scheduler (weekly, every Monday 8am)

```
Program: python
Arguments: C:\recon_agent\run_headless.py --cust C:\data\fab.xls --zag C:\data\zag.xlsx --period "Weekly" --out C:\reports
```

### Linux/Mac cron (first day of each month, 7am)

```cron
0 7 1 * * cd /opt/recon_agent && python run_headless.py --cust /data/fab.xls --zag /data/zag.xlsx --period "Monthly" --out /data/reports/
```

## Adding a new custodian

1. Copy `parsers/fab.py` as `parsers/your_bank.py`
2. Change `name`, `file_extensions`, `can_parse()`, and `parse()` to match your bank's format
3. Add `YourBankParser` to `CUSTODIAN_PARSERS` in `parsers/registry.py`
4. Restart — the new parser appears automatically

## Report structure (Excel tabs)

| Tab | Content |
|-----|---------|
| Period Summary | Match stats, balance comparison, flagged manual groups |
| Matched | All matched items (engine + manual, colour-coded) |
| Breaks | Unmatched items with user notes |
| Unmatched Side-by-Side | FAB left / ZAG right — add Match IDs here |
| Manual Match Detail | All manual groups with net difference |
| ZAG Netting Detail | All journal entries collapsed before matching |
| Raw - Custodian | Raw parsed custodian transactions |
| Raw - ZagTrader | Raw ZagTrader transactions |

## Matching logic

| Rule | Detail |
|------|--------|
| Pass 1 | Exact amount (±0.05) + same settlement date |
| Pass 2 | Exact amount + income delay tolerance (default 15 days) |
| Pass 3 | Exact amount + settlement tolerance (default 5 days) |
| Netting | Gross+WHT, coupon splits, transfer+charge, FD principal+interest |
| Manual | Match IDs added in side-by-side sheet — applied on next run |
