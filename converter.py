"""
Tally DayBook XML → Sales Report XLSX converter.

Parses a Tally ERP/Prime DayBook export (XML) and produces a flat
sales report in the same column layout as the user's template.
"""
from __future__ import annotations

import os
import re
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from datetime import datetime
from typing import List, Optional

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


# Output column order — matches the user-supplied template exactly.
COLUMNS = [
    "Invoice No",
    "Invoice Date",
    "Client Name",
    "Client GST No",
    "Item",
    "HSN Code",
    "Quantity",
    "Rate",
    "Amount (Excl. GST)",
    "GST %",
    "CGST Amount",
    "SGST Amount",
    "IGST Amount",
    "Total Invoice Value",
]


# ---------------------------------------------------------------------------
# Data containers
# ---------------------------------------------------------------------------
@dataclass
class LineItem:
    item_name: str
    hsn_code: str
    amount: float          # taxable value (always positive in XML)
    gst_pct: float         # e.g. 18.0, 5.0, 0.0


@dataclass
class Voucher:
    voucher_no: str
    date: Optional[datetime]
    party_name: str
    party_gstin: str
    is_credit: bool        # True for Sales Credit Notes
    is_interstate: bool    # True if IGST applies
    items: List[LineItem] = field(default_factory=list)


# ---------------------------------------------------------------------------
# Parsing
# ---------------------------------------------------------------------------
def _text(elem: Optional[ET.Element], tag: str) -> str:
    """Return stripped text of first matching child, or '' if absent."""
    if elem is None:
        return ""
    found = elem.find(tag)
    if found is None or found.text is None:
        return ""
    return found.text.strip()


def _parse_date(yyyymmdd: str) -> Optional[datetime]:
    if not yyyymmdd or len(yyyymmdd) != 8 or not yyyymmdd.isdigit():
        return None
    try:
        return datetime.strptime(yyyymmdd, "%Y%m%d")
    except ValueError:
        return None


# Pull GST percentage out of the stock-item name suffix.
# Tally items in this dataset look like "3924 - BUCKET, PLASTIC GOODS - 18%".
_GST_PCT_RE = re.compile(r"-\s*(\d+(?:\.\d+)?)\s*%\s*$")


def _extract_gst_pct(stock_item_name: str) -> float:
    if not stock_item_name:
        return 0.0
    m = _GST_PCT_RE.search(stock_item_name)
    if m:
        try:
            return float(m.group(1))
        except ValueError:
            return 0.0
    return 0.0


def _clean_item_name(stock_item_name: str) -> str:
    """
    Strip leading HSN code and trailing GST suffix from a Tally stock item name
    so the spreadsheet shows just the goods description.

    Examples:
        '3924  - BUCKET, PLASTIC GOODS - 18%' -> 'BUCKET, PLASTIC GOODS'
        '9603 - Brooms,Wars - 0%'             -> 'Brooms,Wars'
        '3401- DETERGENT - 18%'               -> 'DETERGENT'
    """
    if not stock_item_name:
        return ""
    name = _GST_PCT_RE.sub("", stock_item_name).strip()
    # Strip leading HSN-like prefix "1234 -" or "1234-"
    name = re.sub(r"^\s*\d{4,8}\s*-\s*", "", name).strip()
    # Trim a trailing dash if any
    name = re.sub(r"-\s*$", "", name).strip()
    return name


def _is_credit_voucher(vchtype: str) -> bool:
    if not vchtype:
        return False
    upper = vchtype.upper()
    return "CREDIT" in upper or "RETURN" in upper


def _parse_voucher(v_elem: ET.Element) -> Optional[Voucher]:
    vchtype = (v_elem.get("VCHTYPE") or "").strip()

    voucher_no = _text(v_elem, "VOUCHERNUMBER")
    if not voucher_no:
        # Skip vouchers with no number (orphan / system entries)
        return None

    date = _parse_date(_text(v_elem, "DATE"))

    # Party / client name — prefer the human-friendly mailing name, then ledger
    party_name = (
        _text(v_elem, "PARTYMAILINGNAME")
        or _text(v_elem, "PARTYLEDGERNAME")
        or _text(v_elem, "PARTYNAME")
        or _text(v_elem, "BASICBUYERNAME")
    )
    party_gstin = _text(v_elem, "PARTYGSTIN")

    cmp_state = (_text(v_elem, "CMPGSTSTATE") or "").strip().lower()
    place_of_supply = (
        _text(v_elem, "PLACEOFSUPPLY")
        or _text(v_elem, "STATENAME")
        or ""
    ).strip().lower()
    is_interstate = bool(cmp_state and place_of_supply and cmp_state != place_of_supply)

    is_credit = _is_credit_voucher(vchtype)

    items: List[LineItem] = []
    for inv in v_elem.findall("ALLINVENTORYENTRIES.LIST"):
        stock_name = _text(inv, "STOCKITEMNAME")
        if not stock_name:
            continue
        amount_text = _text(inv, "AMOUNT")
        try:
            amount = float(amount_text) if amount_text else 0.0
        except ValueError:
            amount = 0.0
        # Tally exports sales taxable amounts as negative numbers
        # (because party ledger is debit-positive). Take absolute value;
        # we'll re-sign at output time based on credit/debit.
        amount = abs(amount)

        hsn = _text(inv, "GSTHSNNAME")
        gst_pct = _extract_gst_pct(stock_name)

        items.append(
            LineItem(
                item_name=_clean_item_name(stock_name),
                hsn_code=hsn,
                amount=amount,
                gst_pct=gst_pct,
            )
        )

    if not items:
        return None

    return Voucher(
        voucher_no=voucher_no,
        date=date,
        party_name=party_name,
        party_gstin=party_gstin,
        is_credit=is_credit,
        is_interstate=is_interstate,
        items=items,
    )


_BAD_CHARREF_RE = re.compile(rb"&#(?:x0*[0-8bcefBCEF]|x1[0-9a-fA-F]|0*[0-8]|11|12|1[4-9]|2[0-9]|3[01]);")
_BAD_CTRL_RE = re.compile(rb"[\x00-\x08\x0B\x0C\x0E-\x1F]")


def parse_tally_xml(xml_path: str) -> List[Voucher]:
    """Parse a Tally DayBook XML file and return a list of vouchers.

    Tally exports often embed XML-illegal numeric character references
    (e.g. ``&#4;``) and raw control bytes inside text nodes. We strip
    both before feeding the bytes to ElementTree.
    """
    with open(xml_path, "rb") as f:
        data = f.read()
    data = _BAD_CHARREF_RE.sub(b"", data)
    data = _BAD_CTRL_RE.sub(b"", data)
    try:
        tree = ET.ElementTree(ET.fromstring(data))
    except ET.ParseError as exc:
        raise ValueError(f"Could not parse Tally XML: {exc}") from exc

    root = tree.getroot()
    vouchers: List[Voucher] = []
    for v in root.iter("VOUCHER"):
        parsed = _parse_voucher(v)
        if parsed is not None:
            vouchers.append(parsed)
    return vouchers


# ---------------------------------------------------------------------------
# Tax math
# ---------------------------------------------------------------------------
def _round2(x: float) -> float:
    return round(x + 1e-9, 2)


def build_rows(vouchers: List[Voucher]):
    """
    Convert parsed vouchers into flat output rows in the column order
    defined by COLUMNS. Returns a list of dicts.
    """
    rows = []
    for v in vouchers:
        sign = -1.0 if v.is_credit else 1.0
        for it in v.items:
            taxable = sign * it.amount
            half_rate = it.gst_pct / 2.0 / 100.0
            full_rate = it.gst_pct / 100.0

            if v.is_interstate:
                cgst = 0.0
                sgst = 0.0
                igst = _round2(taxable * full_rate)
            else:
                cgst = _round2(taxable * half_rate)
                sgst = _round2(taxable * half_rate)
                igst = 0.0

            total = _round2(taxable + cgst + sgst + igst)

            rows.append({
                "Invoice No": v.voucher_no,
                "Invoice Date": v.date,
                "Client Name": v.party_name,
                "Client GST No": v.party_gstin or "",
                "Item": it.item_name,
                "HSN Code": it.hsn_code,
                # Quantity & Rate left blank — XML does not contain them
                "Quantity": "",
                "Rate": "",
                "Amount (Excl. GST)": _round2(taxable),
                "GST %": it.gst_pct,
                "CGST Amount": cgst,
                "SGST Amount": sgst,
                "IGST Amount": igst,
                "Total Invoice Value": total,
            })
    return rows


# ---------------------------------------------------------------------------
# Excel writer
# ---------------------------------------------------------------------------
def write_xlsx(rows, out_path: str, sheet_name: str = "Sales Report (from Tally)"):
    wb = Workbook()
    ws = wb.active
    ws.title = sheet_name[:31]  # Excel sheet name limit

    header_font = Font(name="Calibri", bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", start_color="305496")
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    body_font = Font(name="Calibri")
    left_align = Alignment(horizontal="left", vertical="center")
    right_align = Alignment(horizontal="right", vertical="center")
    center_align = Alignment(horizontal="center", vertical="center")

    # Header
    for col_idx, name in enumerate(COLUMNS, start=1):
        cell = ws.cell(row=1, column=col_idx, value=name)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_align
    ws.row_dimensions[1].height = 32
    ws.freeze_panes = "A2"

    # Body
    for r_idx, row in enumerate(rows, start=2):
        for c_idx, name in enumerate(COLUMNS, start=1):
            value = row.get(name, "")
            cell = ws.cell(row=r_idx, column=c_idx, value=value)
            cell.font = body_font
            if name == "Invoice Date" and isinstance(value, datetime):
                cell.number_format = "dd-mmm-yyyy"
                cell.alignment = center_align
            elif name in ("Amount (Excl. GST)", "CGST Amount", "SGST Amount",
                          "IGST Amount", "Total Invoice Value", "Rate"):
                cell.number_format = "#,##0.00;(#,##0.00);-"
                cell.alignment = right_align
            elif name == "GST %":
                cell.number_format = '0"%"'
                cell.alignment = center_align
            elif name == "Quantity":
                cell.alignment = right_align
            else:
                cell.alignment = left_align

    # Column widths tuned to template
    widths = {
        "Invoice No": 16,
        "Invoice Date": 14,
        "Client Name": 30,
        "Client GST No": 18,
        "Item": 32,
        "HSN Code": 10,
        "Quantity": 9,
        "Rate": 10,
        "Amount (Excl. GST)": 16,
        "GST %": 8,
        "CGST Amount": 12,
        "SGST Amount": 12,
        "IGST Amount": 12,
        "Total Invoice Value": 18,
    }
    for col_idx, name in enumerate(COLUMNS, start=1):
        ws.column_dimensions[get_column_letter(col_idx)].width = widths.get(name, 14)

    # Add an autofilter over the data range
    if rows:
        last_col = get_column_letter(len(COLUMNS))
        ws.auto_filter.ref = f"A1:{last_col}{len(rows) + 1}"

    wb.save(out_path)


# ---------------------------------------------------------------------------
# Public entry-point
# ---------------------------------------------------------------------------
def convert(xml_path: str, out_path: Optional[str] = None) -> str:
    """
    Convert a Tally DayBook XML file to a sales-report XLSX.

    Output is written next to the input by default, with the suffix
    "_sales_report.xlsx".

    Returns the absolute path of the written file.
    """
    xml_path = os.path.abspath(xml_path)
    if not os.path.isfile(xml_path):
        raise FileNotFoundError(xml_path)

    vouchers = parse_tally_xml(xml_path)
    rows = build_rows(vouchers)

    if out_path is None:
        base = os.path.splitext(os.path.basename(xml_path))[0]
        out_path = os.path.join(os.path.dirname(xml_path), f"{base}_sales_report.xlsx")

    write_xlsx(rows, out_path)
    return os.path.abspath(out_path)


if __name__ == "__main__":
    import sys
    if len(sys.argv) < 2:
        print("Usage: python converter.py <tally_daybook.xml> [output.xlsx]")
        sys.exit(1)
    inp = sys.argv[1]
    outp = sys.argv[2] if len(sys.argv) > 2 else None
    written = convert(inp, outp)
    print(f"Wrote: {written}")
