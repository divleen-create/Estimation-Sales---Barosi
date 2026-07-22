"""Indian-convention number formatting helpers."""
from __future__ import annotations
from typing import Optional


def indian_group(n) -> str:
    """12345678 -> '1,23,45,678' (Indian digit grouping)."""
    if n is None:
        return "–"
    s = str(int(round(n)))
    neg = s.startswith("-")
    s = s.lstrip("-")
    if len(s) <= 3:
        out = s
    else:
        last3 = s[-3:]
        rest = s[:-3]
        parts = []
        while len(rest) > 2:
            parts.insert(0, rest[-2:])
            rest = rest[:-2]
        if rest:
            parts.insert(0, rest)
        out = ",".join(parts) + "," + last3
    return ("-" if neg else "") + out


def gmv_auto(v: Optional[float]) -> str:
    """Auto ₹Cr / ₹L for headline figures."""
    if v is None:
        return "–"
    if abs(v) >= 1e7:
        return f"₹{v / 1e7:.2f} Cr"
    return f"₹{v / 1e5:.1f} L"


def lakh(v: Optional[float], dp: int = 1) -> str:
    """Compact ₹ Lakh number for dense grids (no symbol; header carries it)."""
    if v is None:
        return "–"
    return f"{v / 1e5:.{dp}f}"


def rupees(v: Optional[float]) -> str:
    if v is None:
        return "–"
    return "₹" + indian_group(v)


def money_l(v: Optional[float], dp: int = 1) -> str:
    """Compact rupee value in lakh with the ₹ symbol, e.g. '₹7.6 L'."""
    if v is None:
        return "–"
    return f"₹{v / 1e5:.{dp}f} L"


def pct(g: Optional[float]) -> str:
    if g is None:
        return "–"
    return f"{g * 100:+.1f}%"


def pct_plain(g: Optional[float]) -> str:
    """Unsigned percentage, e.g. ad contribution '26.5%'."""
    if g is None:
        return "–"
    return f"{g * 100:.1f}%"


def growth_class(g: Optional[float]) -> str:
    """CSS class buckets for growth colouring."""
    if g is None:
        return "g-na"
    p = g * 100
    if p >= 25:
        return "g-pos2"
    if p >= 5:
        return "g-pos1"
    if p > -5:
        return "g-neu"
    if p > -25:
        return "g-neg1"
    return "g-neg2"
