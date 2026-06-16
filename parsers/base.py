"""
Base class for custodian statement parsers.
To add a new custodian, subclass CustodianParser and implement parse().
"""
from abc import ABC, abstractmethod
import pandas as pd


class CustodianParser(ABC):
    """
    Base class all custodian parsers must implement.
    parse() must return:
        df          - DataFrame with columns:
                      currency, trade_date, settlement_date, txn_type,
                      description, debit, credit, balance
        balance_info - dict {ccy: {'starting': float, 'ending': float}}
    """
    name = "Unknown Custodian"
    file_extensions = [".xls", ".xlsx"]

    @abstractmethod
    def parse(self, filepath: str) -> tuple[pd.DataFrame, dict]:
        pass

    @classmethod
    def can_parse(cls, filename: str) -> bool:
        return any(filename.lower().endswith(ext) for ext in cls.file_extensions)
