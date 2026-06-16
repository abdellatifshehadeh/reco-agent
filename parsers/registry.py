"""
Parser registry.
Add new custodian parsers here — they are auto-detected from filename patterns.
"""
from .fab import FABParser

# ── Register all available custodian parsers here ──
CUSTODIAN_PARSERS = [
    FABParser,
    # IBParser,        # Interactive Brokers — add when ready
    # SucdenParser,    # Sucden Financial
    # MarexParser,     # Marex
]


def detect_parser(filename: str):
    """Return the first parser that claims it can handle this filename, or None."""
    for parser_cls in CUSTODIAN_PARSERS:
        if parser_cls.can_parse(filename):
            return parser_cls()
    return None


def list_custodians() -> list[str]:
    return [p.name for p in CUSTODIAN_PARSERS]
