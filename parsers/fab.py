"""
FAB (First Abu Dhabi Bank) custodian statement parser.
Handles the PortfolioValuation XLS format with Currency: XXX section headers.
"""
import re
import pandas as pd
from .base import CustodianParser


class FABParser(CustodianParser):
    name = "FAB (First Abu Dhabi Bank)"
    file_extensions = [".xls", ".xlsx"]

    DEBIT_TYPES  = ['Receive vs Payment', 'Cash Withdrawal', 'Cash Out']
    CREDIT_TYPES = ['Deliver vs Payment', 'Cash Deposit', 'CA Dividend',
                    'CA Income', 'CA Pay due']

    def parse(self, filepath: str) -> tuple[pd.DataFrame, dict]:
        xls = pd.ExcelFile(filepath, engine="xlrd")
        df_raw = pd.read_excel(xls, sheet_name=xls.sheet_names[0], header=None)

        rows = []
        current_currency = None
        in_txn_block = False
        balance_info = {}

        for _, row in df_raw.iterrows():
            vals = [str(v).strip() for v in row if str(v).strip() not in ['nan', 'NaN', '']]
            if not vals:
                continue

            # Detect currency section header
            for v in vals:
                m = re.match(r'^Currency:\s*([A-Z]{3})$', v)
                if m:
                    current_currency = m.group(1)
                    in_txn_block = False
                    break

            # Detect transaction block header
            if len(vals) >= 3 and vals[0] == 'Trade Date' and vals[1] == 'Settlement Date':
                in_txn_block = True
                continue

            if not in_txn_block or not current_currency:
                continue

            # Balance rows
            if 'Starting Balance' in vals:
                try:
                    balance_info.setdefault(current_currency, {})['starting'] = \
                        float(str(vals[-1]).replace(',', ''))
                except Exception:
                    pass
                continue

            if 'Ending Balance' in vals or 'Current Balance' in vals:
                try:
                    balance_info.setdefault(current_currency, {})['ending'] = \
                        float(str(vals[-1]).replace(',', ''))
                except Exception:
                    pass
                continue

            # Transaction rows
            if len(vals) >= 5:
                try:
                    trade_date = pd.to_datetime(vals[0], dayfirst=True, errors='coerce')
                    if pd.isnull(trade_date):
                        continue
                    settlement_date = pd.to_datetime(vals[1], dayfirst=True, errors='coerce')
                    txn_type    = vals[2]
                    description = vals[3]
                    amount      = float(str(vals[4]).replace(',', ''))
                    balance     = float(str(vals[-1]).replace(',', ''))

                    if any(d in txn_type for d in self.DEBIT_TYPES):
                        debit, credit = amount, None
                    elif any(c in txn_type for c in self.CREDIT_TYPES):
                        debit, credit = None, amount
                    else:
                        debit, credit = amount, None

                    rows.append({
                        'source': 'CUSTODIAN',
                        'custodian': self.name,
                        'currency': current_currency,
                        'trade_date': trade_date,
                        'settlement_date': settlement_date,
                        'txn_type': txn_type,
                        'description': description,
                        'debit': debit,
                        'credit': credit,
                        'balance': balance,
                    })
                except Exception:
                    continue

        df = pd.DataFrame(rows)
        return df, balance_info

    @classmethod
    def can_parse(cls, filename: str) -> bool:
        """FAB files follow naming pattern: DDMMYYYY_RPT_STATEMENTACCOUNT_Report_NNNN.xls"""
        name = filename.lower()
        return ('rpt_statementaccount' in name or 'portfoliovaluation' in name) and \
               any(name.endswith(ext) for ext in cls.file_extensions)
