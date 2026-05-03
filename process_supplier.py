import io
import math
import re
import tempfile
from pathlib import Path
from typing import Optional

import openpyxl

from suppliers_config import SUPPLIERS, AVRORA


# ── Markup formulas (from AVRORA Steel pricing policy) ───────────────────────

def _retail_pct(price: float) -> float:
    if price < 5_000:      return 0.30
    if price < 10_000:     return 0.25
    if price < 18_000:     return 0.18
    if price < 40_000:     return 0.15
    if price < 70_000:     return 0.11
    if price < 120_000:    return 0.08
    if price < 250_000:    return 0.06
    if price < 500_000:    return 0.05
    if price < 1_000_000:  return 0.035
    return 0.025


def _round_step(value: float) -> int:
    if value < 5_000:      step = 10
    elif value < 50_000:   step = 50
    elif value < 500_000:  step = 100
    else:                  step = 1_000
    return int(math.ceil(value / step) * step)


def _apply_retail(base: float) -> Optional[int]:
    if not base or base <= 0:
        return None
    markup = max(base * _retail_pct(base), 100)
    return _round_step(base + markup)


def _apply_wholesale(base: float) -> Optional[int]:
    # Wholesale = 60% of retail markup (below retail, above base)
    if not base or base <= 0:
        return None
    markup = max(base * _retail_pct(base) * 0.60, 50)
    return _round_step(base + markup)


# ── Supplier detection ────────────────────────────────────────────────────────

def detect_supplier(filename: str) -> Optional[dict]:
    name_lower = Path(filename).stem.lower()
    for cfg in SUPPLIERS.values():
        for kw in cfg["keywords"]:
            if kw in name_lower:
                return cfg
    return None


# ── Price column detection ────────────────────────────────────────────────────

def _find_price_cols(ws, skip_rows: int, price_keywords: list[str]) -> list[int]:
    found = []
    scan_until = min(skip_rows + 2, ws.max_row or 1)
    for row in ws.iter_rows(max_row=scan_until, values_only=True):
        for col_idx, val in enumerate(row):
            if val and isinstance(val, str):
                if any(kw in val.lower() for kw in price_keywords):
                    found.append(col_idx)
    return list(set(found))


# ── Contact replacement ───────────────────────────────────────────────────────

_PHONE_RE = re.compile(r'(\+7|8)[\s\-]?\(?\d{3}\)?[\s\-]?\d{3}[\s\-]?\d{2}[\s\-]?\d{2}')
_EMAIL_RE = re.compile(r'[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}')


def _replace_contacts(ws):
    contact_line1 = (
        f"{AVRORA['company']}  |  "
        f"Тел: {AVRORA['phone']}  |  "
        f"WhatsApp: {AVRORA['whatsapp']}"
    )
    contact_line2 = (
        f"{AVRORA['email']}  |  "
        f"{AVRORA['website']}  |  "
        f"{AVRORA['address']}  |  "
        f"{AVRORA['schedule']}"
    )

    replaced = 0
    for row in ws.iter_rows(max_row=6):
        for cell in row:
            if not cell.value:
                continue
            val = str(cell.value)
            if _PHONE_RE.search(val) or _EMAIL_RE.search(val) or "www." in val.lower():
                cell.value = contact_line1 if replaced == 0 else contact_line2
                replaced += 1
                if replaced >= 2:
                    return

    # No contacts found — overwrite cell A1
    if replaced == 0:
        ws.cell(row=1, column=1).value = contact_line1


# ── Main processing function ──────────────────────────────────────────────────

def process_file(filepath: str, supplier_cfg: dict) -> tuple[bytes, bytes]:
    """
    Returns (retail_xlsx_bytes, wholesale_xlsx_bytes).
    """
    results = []

    for markup_fn in (_apply_retail, _apply_wholesale):
        wb = openpyxl.load_workbook(filepath, data_only=True)
        ws = wb.active

        skip_rows  = supplier_cfg["skip_rows"]
        kws        = supplier_cfg["price_col_keywords"]
        fallback   = supplier_cfg.get("price_col_fallback", 5)

        price_cols = _find_price_cols(ws, skip_rows, kws)
        if not price_cols:
            price_cols = [fallback]

        for row_idx, row in enumerate(ws.iter_rows()):
            if row_idx < skip_rows:
                continue
            for col_idx, cell in enumerate(row):
                if col_idx not in price_cols:
                    continue
                if not isinstance(cell.value, (int, float)):
                    continue
                base = float(cell.value)
                if base <= 0:
                    continue
                new_price = markup_fn(base)
                if new_price is not None:
                    cell.value = new_price

        _replace_contacts(ws)

        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)
        results.append(buf.read())

    return results[0], results[1]
