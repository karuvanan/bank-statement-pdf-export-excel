from __future__ import annotations

import json
import re
from dataclasses import dataclass
from datetime import date
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any


ROOT = Path.cwd()
OCR_DIR = ROOT / "tmp" / "pdfs" / "ocr"
OUTPUT_JSON = ROOT / "tmp" / "pdfs" / "extracted_transactions.json"


DATE_RE = re.compile(r"(\d{1,2})\s*/\s*(\d{1,2})")
MONEY_RE = re.compile(r"(\d{1,3}(?:,\d{3})*(?:\.\d{2})|\d+\.\d{2}|\.\d{2})")


@dataclass
class Line:
    text: str
    x: float
    y: float
    w: float
    h: float

    @property
    def cx(self) -> float:
        return self.x + self.w / 2


def load_json(path: Path) -> Any:
    return json.loads(path.read_text(encoding="utf-8-sig"))


def norm_text(text: str) -> str:
    return " ".join(str(text).replace("\u00a0", " ").split())


def statement_period(pdf_name: str) -> tuple[int | None, int | None]:
    match = re.search(r"(\d{2})(\d{2})", pdf_name)
    if not match:
        return None, None
    return 2000 + int(match.group(1)), int(match.group(2))


def normalize_date(text: str) -> str | None:
    cleaned = (
        norm_text(text)
        .upper()
        .replace("O", "0")
        .replace("I", "1")
        .replace("L", "1")
        .replace("|", "/")
        .replace("\\", "/")
    )
    match = DATE_RE.search(cleaned)
    if not match:
        return None
    day, month = int(match.group(1)), int(match.group(2))
    if not (1 <= day <= 31 and 1 <= month <= 12):
        return None
    return f"{day:02d}/{month:02d}"


def normalize_money_text(text: str) -> str:
    return (
        norm_text(text)
        .replace("O", "0")
        .replace("o", "0")
        .replace("，", ",")
        .replace("–", "-")
        .replace("—", "-")
        .replace("−", "-")
        .replace(" ", "")
    )


def parse_money(text: str) -> tuple[Decimal | None, int | None, str]:
    cleaned = normalize_money_text(text)
    sign = None
    tail = cleaned[-3:]
    if "+" in tail:
        sign = 1
    elif "-" in tail:
        sign = -1

    match = MONEY_RE.search(cleaned)
    if not match:
        return None, sign, cleaned

    raw = match.group(1)
    if raw.startswith("."):
        raw = "0" + raw
    try:
        value = Decimal(raw.replace(",", ""))
    except InvalidOperation:
        return None, sign, cleaned
    return value, sign, cleaned


def find_header_bottom(lines: list[Line], height: float) -> float:
    hits = []
    for line in lines:
        upper = line.text.upper()
        if any(token in upper for token in ("ENTRY DATE", "TRANSACTION DESCRIPTION", "STATEMENT BALANCE")):
            hits.append(line.y + line.h)
    if hits:
        return max(hits) + 20
    return height * 0.29


def line_in_band(line: Line, start: float, end: float) -> bool:
    return start <= line.y <= end


def closest_money_line(lines: list[Line], y: float, tolerance: float = 48) -> Line | None:
    money_lines = [line for line in lines if parse_money(line.text)[0] is not None and abs(line.y - y) <= tolerance]
    if not money_lines:
        return None
    return min(money_lines, key=lambda line: abs(line.y - y))


def full_date(entry: str | None, pdf_name: str) -> str | None:
    if not entry:
        return None
    year, _month = statement_period(pdf_name)
    if year is None:
        return None
    day, month = [int(part) for part in entry.split("/")]
    try:
        return date(year, month, day).isoformat()
    except ValueError:
        return None


def parse_page(payload: dict[str, Any]) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    width = float(payload["width"])
    height = float(payload["height"])
    lines = [
        Line(norm_text(item.get("text", "")), float(item["x"]), float(item["y"]), float(item["w"]), float(item["h"]))
        for item in payload.get("lines", [])
        if norm_text(item.get("text", ""))
    ]
    body_start = find_header_bottom(lines, height)
    body_end = height * 0.84

    date_lines = []
    value_date_lines = []
    desc_lines = []
    amount_lines = []
    balance_lines = []

    for line in lines:
        if not line_in_band(line, body_start, body_end):
            continue
        rx = line.x / width
        rc = line.cx / width
        if 0.10 <= rc <= 0.24:
            if normalize_date(line.text):
                date_lines.append(line)
        elif 0.24 < rc <= 0.32:
            if normalize_date(line.text):
                value_date_lines.append(line)
        elif 0.28 <= rx <= 0.62:
            desc_lines.append(line)
        elif 0.60 <= rc <= 0.74:
            amount_lines.append(line)
        elif 0.74 < rc <= 0.90:
            balance_lines.append(line)

    starts: list[dict[str, Any]] = []
    for line in date_lines:
        starts.append({"kind": "transaction", "y": line.y, "entry_date": normalize_date(line.text), "line": line})
    for line in desc_lines:
        if "BEGINNING BALANCE" in line.text.upper():
            starts.append({"kind": "beginning", "y": line.y, "entry_date": None, "line": line})

    starts.sort(key=lambda item: item["y"])

    summary_keywords = ("ENDING BALANCE", "LEDGER BALANCE", "TOTAL DEBIT", "TOTAL CREDIT", "PROFIT OUTSTANDING")
    summary_ys = [line.y for line in desc_lines if any(keyword in line.text.upper() for keyword in summary_keywords)]
    first_summary_y = min(summary_ys) if summary_ys else body_end

    rows = []
    for idx, start in enumerate(starts):
        y0 = float(start["y"])
        y1 = float(starts[idx + 1]["y"]) - 6 if idx + 1 < len(starts) else first_summary_y - 6
        if y1 <= y0:
            y1 = y0 + 44

        desc = [line for line in desc_lines if y0 - 10 <= line.y < y1]
        # Keep the beginning-balance row tidy.
        if start["kind"] == "beginning":
            desc = [line for line in desc if "BEGINNING BALANCE" in line.text.upper()]

        value_date = None
        nearby_value_dates = [line for line in value_date_lines if abs(line.y - y0) <= 40]
        if nearby_value_dates:
            value_date = normalize_date(min(nearby_value_dates, key=lambda line: abs(line.y - y0)).text)

        amount_line = None if start["kind"] == "beginning" else closest_money_line(amount_lines, y0)
        balance_line = closest_money_line(balance_lines, y0)
        amount_abs, amount_sign, amount_clean = parse_money(amount_line.text if amount_line else "")
        balance_abs, _balance_sign, balance_clean = parse_money(balance_line.text if balance_line else "")

        rows.append(
            {
                "source_pdf": payload["pdf"],
                "page": payload["page"],
                "row_kind": start["kind"],
                "entry_date": start["entry_date"],
                "entry_date_full": full_date(start["entry_date"], payload["pdf"]),
                "value_date": value_date,
                "description": "\n".join(line.text for line in sorted(desc, key=lambda line: (line.y, line.x))),
                "amount_text": amount_line.text if amount_line else "",
                "amount_abs": float(amount_abs) if amount_abs is not None else None,
                "amount_sign_ocr": amount_sign,
                "amount_clean": amount_clean,
                "balance_text": balance_line.text if balance_line else "",
                "balance": float(balance_abs) if balance_abs is not None else None,
                "balance_clean": balance_clean,
                "y": y0,
            }
        )

    summaries = []
    for line in desc_lines:
        upper = line.text.upper()
        label = None
        for keyword in summary_keywords:
            if keyword in upper:
                label = keyword.title()
                break
        if not label:
            continue
        balance_line = closest_money_line(balance_lines, line.y, tolerance=56)
        value, sign, clean = parse_money(balance_line.text if balance_line else "")
        summaries.append(
            {
                "source_pdf": payload["pdf"],
                "page": payload["page"],
                "label": label,
                "value_text": balance_line.text if balance_line else "",
                "value": float(value) if value is not None else None,
                "sign": sign,
                "clean": clean,
            }
        )

    return rows, summaries


def derive_signed_amounts(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    by_pdf: dict[str, list[dict[str, Any]]] = {}
    for row in rows:
        by_pdf.setdefault(row["source_pdf"], []).append(row)

    for pdf_rows in by_pdf.values():
        pdf_rows.sort(key=lambda row: (row["page"], row["y"]))
        previous_balance: Decimal | None = None
        for row in pdf_rows:
            notes = []
            balance = Decimal(str(row["balance"])) if row["balance"] is not None else None
            amount_abs = Decimal(str(row["amount_abs"])) if row["amount_abs"] is not None else None
            sign_ocr = row.get("amount_sign_ocr")
            signed_amount: Decimal | None = None

            if row["row_kind"] == "beginning":
                previous_balance = balance
                row["amount"] = None
                row["debit"] = None
                row["credit"] = None
                row["notes"] = ""
                continue

            if amount_abs is not None and previous_balance is not None and balance is not None:
                delta = balance - previous_balance
                if abs(abs(delta) - amount_abs) <= Decimal("0.02"):
                    signed_amount = amount_abs if delta >= 0 else -amount_abs
                    if sign_ocr is not None and signed_amount != amount_abs * Decimal(sign_ocr):
                        notes.append("OCR sign adjusted by balance movement")
                elif sign_ocr is not None:
                    signed_amount = amount_abs * Decimal(sign_ocr)
                    notes.append("Balance movement mismatch")
                else:
                    signed_amount = amount_abs
                    notes.append("Amount sign inferred as positive; check row")
            elif amount_abs is not None:
                signed_amount = amount_abs * Decimal(sign_ocr if sign_ocr is not None else 1)
                if sign_ocr is None:
                    notes.append("Amount sign not detected")
            elif previous_balance is not None and balance is not None:
                delta = balance - previous_balance
                if abs(delta) > Decimal("0.02"):
                    signed_amount = delta
                    notes.append("OCR amount missing; inferred from balance movement")

            row["amount"] = float(signed_amount) if signed_amount is not None else None
            row["debit"] = float(-signed_amount) if signed_amount is not None and signed_amount < 0 else None
            row["credit"] = float(signed_amount) if signed_amount is not None and signed_amount > 0 else None
            row["notes"] = "; ".join(notes)

            if balance is not None:
                previous_balance = balance

    return rows


def build_statement_summaries(rows: list[dict[str, Any]], summaries: list[dict[str, Any]]) -> list[dict[str, Any]]:
    by_pdf: dict[str, list[dict[str, Any]]] = {}
    by_pdf_summary: dict[str, dict[str, float | None]] = {}
    for row in rows:
        by_pdf.setdefault(row["source_pdf"], []).append(row)
    for item in summaries:
        by_pdf_summary.setdefault(item["source_pdf"], {})[item["label"]] = item["value"]

    out = []
    for pdf_name, pdf_rows in sorted(by_pdf.items()):
        ordered = sorted(pdf_rows, key=lambda row: (row["page"], row["y"]))
        transaction_rows = [row for row in ordered if row["row_kind"] == "transaction"]
        beginning = next((row["balance"] for row in ordered if row["row_kind"] == "beginning"), None)
        ending = by_pdf_summary.get(pdf_name, {}).get("Ending Balance")
        if ending is None:
            ending = next((row["balance"] for row in reversed(ordered) if row["balance"] is not None), None)
        calc_debit = round(sum(row["debit"] or 0 for row in transaction_rows), 2)
        calc_credit = round(sum(row["credit"] or 0 for row in transaction_rows), 2)
        total_debit = by_pdf_summary.get(pdf_name, {}).get("Total Debit")
        total_credit = by_pdf_summary.get(pdf_name, {}).get("Total Credit")

        out.append(
            {
                "source_pdf": pdf_name,
                "statement_period": "-".join(str(part).zfill(2) for part in statement_period(pdf_name) if part is not None),
                "beginning_balance": beginning,
                "ending_balance": ending,
                "transaction_count": len(transaction_rows),
                "calculated_debit": calc_debit,
                "statement_total_debit": total_debit,
                "calculated_credit": calc_credit,
                "statement_total_credit": total_credit,
                "rows_with_notes": sum(1 for row in transaction_rows if row.get("notes")),
            }
        )
    return out


def main() -> None:
    rows: list[dict[str, Any]] = []
    summaries: list[dict[str, Any]] = []
    for path in sorted(OCR_DIR.glob("*.json")):
        page_rows, page_summaries = parse_page(load_json(path))
        rows.extend(page_rows)
        summaries.extend(page_summaries)

    rows = derive_signed_amounts(rows)
    statement_summaries = build_statement_summaries(rows, summaries)
    payload = {
        "rows": sorted(rows, key=lambda row: (row["source_pdf"], row["page"], row["y"])),
        "summaries": summaries,
        "statement_summaries": statement_summaries,
    }
    OUTPUT_JSON.parent.mkdir(parents=True, exist_ok=True)
    OUTPUT_JSON.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"Parsed {len(payload['rows'])} rows across {len(statement_summaries)} statements.")
    notes = sum(1 for row in payload["rows"] if row.get("notes"))
    print(f"Rows with review notes: {notes}")
    print(OUTPUT_JSON)


if __name__ == "__main__":
    main()
