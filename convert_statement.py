import argparse
import re
from decimal import Decimal, InvalidOperation
from typing import List, Optional, Tuple

import pandas as pd
import pdfplumber

DATE_RE = re.compile(r"\d{2}/\d{2}/\d{2}")
TIME_RE = re.compile(r"\d{2}:\d{2}")

# Column boundaries (x0 positions) tuned for the supplied bank statement layout.
WITHDRAWAL_RANGE: Tuple[float, float] = (340, 420)
DEPOSIT_RANGE: Tuple[float, float] = (420, 470)
BALANCE_RANGE: Tuple[float, float] = (470, 540)
BRANCH_MIN_X: float = 530


def _normalise_whitespace(text: str) -> str:
    text = text.replace("~", " ")
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def _parse_amount(value: str) -> Optional[float]:
    value = value.strip()
    if not value:
        return None
    value = value.replace(",", "")
    try:
        return float(Decimal(value))
    except (InvalidOperation, ValueError):
        return None


def extract_statement_records(pdf_path: str, password: Optional[str] = None) -> pd.DataFrame:
    """Extract transaction rows from the password-protected bank statement PDF."""

    records: List[dict] = []
    with pdfplumber.open(pdf_path, password=password) as pdf:
        for page_idx, page in enumerate(pdf.pages, start=1):
            table_region = page.crop((36, 226, 558, 751))
            words = table_region.extract_words(x_tolerance=1.5, y_tolerance=2)
            words.sort(key=lambda w: (w["top"], w["x0"]))

            # Group words into text lines using the Y position tolerance
            lines = []
            for word in words:
                if not lines or abs(word["top"] - lines[-1]["top"]) > 1.5:
                    lines.append({"top": word["top"], "words": [word]})
                else:
                    lines[-1]["words"].append(word)

            current = None
            for line in lines:
                line_words = sorted(line["words"], key=lambda w: w["x0"])
                texts = [w["text"] for w in line_words]
                if not texts:
                    continue

                first_text = texts[0]
                if DATE_RE.fullmatch(first_text):
                    if current:
                        records.append(current)
                    current = {
                        "Date": first_text,
                        "Time": "",
                        "Description": "",
                        "Withdrawal": "",
                        "Deposit": "",
                        "Balance": "",
                        "Branch": "",
                        "Page": page_idx,
                    }
                    description_parts: List[str] = []
                    for word in line_words[1:]:
                        x0 = word["x0"]
                        text = word["text"]
                        if WITHDRAWAL_RANGE[0] <= x0 <= WITHDRAWAL_RANGE[1]:
                            current["Withdrawal"] = text
                        elif DEPOSIT_RANGE[0] < x0 <= DEPOSIT_RANGE[1]:
                            current["Deposit"] = text
                        elif BALANCE_RANGE[0] < x0 <= BALANCE_RANGE[1]:
                            current["Balance"] = text
                        elif x0 > BRANCH_MIN_X:
                            current["Branch"] = (current["Branch"] + " " + text).strip()
                        else:
                            description_parts.append(text)
                    current["Description"] = _normalise_whitespace(" ".join(description_parts))
                elif first_text == "Total":
                    if current:
                        records.append(current)
                        current = None
                    summary = {
                        "Date": "",
                        "Time": "",
                        "Description": "",
                        "Withdrawal": "",
                        "Deposit": "",
                        "Balance": "",
                        "Branch": "",
                        "Page": page_idx,
                    }
                    if len(texts) >= 4 and texts[1] in {"Withdrawal", "Deposit"}:
                        count = texts[2]
                        amount = texts[3]
                        summary["Description"] = f"Total {texts[1]} (count: {count})"
                        if texts[1] == "Withdrawal":
                            summary["Withdrawal"] = amount
                        else:
                            summary["Deposit"] = amount
                    else:
                        summary["Description"] = _normalise_whitespace(" ".join(texts))
                    records.append(summary)
                elif current is not None:
                    extra_words = line_words
                    if TIME_RE.fullmatch(first_text):
                        current["Time"] = first_text
                        extra_words = line_words[1:]
                    extra_desc: List[str] = []
                    for word in extra_words:
                        x0 = word["x0"]
                        text = word["text"]
                        if WITHDRAWAL_RANGE[0] <= x0 <= WITHDRAWAL_RANGE[1] and not current["Withdrawal"]:
                            current["Withdrawal"] = text
                        elif DEPOSIT_RANGE[0] < x0 <= DEPOSIT_RANGE[1] and not current["Deposit"]:
                            current["Deposit"] = text
                        elif BALANCE_RANGE[0] < x0 <= BALANCE_RANGE[1] and not current["Balance"]:
                            current["Balance"] = text
                        elif x0 > BRANCH_MIN_X:
                            current["Branch"] = (current["Branch"] + " " + text).strip()
                        else:
                            extra_desc.append(text)
                    if extra_desc:
                        description = current["Description"]
                        if description:
                            description += " "
                        description += " ".join(extra_desc)
                        current["Description"] = _normalise_whitespace(description)
            if current:
                records.append(current)

    df = pd.DataFrame(records, columns=[
        "Date",
        "Time",
        "Description",
        "Withdrawal",
        "Deposit",
        "Balance",
        "Branch",
        "Page",
    ])

    for column in ["Withdrawal", "Deposit", "Balance"]:
        df[column] = (
            df[column]
            .fillna("")
            .apply(lambda v: _parse_amount(v) if isinstance(v, str) else None)
        )

    return df


def main() -> None:
    parser = argparse.ArgumentParser(description="Convert a password-protected bank statement PDF to CSV/XLSX.")
    parser.add_argument("pdf_path", help="Path to the PDF statement")
    parser.add_argument("--password", help="PDF password", default=None)
    parser.add_argument("--csv", help="Path to write CSV output")
    parser.add_argument("--excel", help="Path to write Excel output")
    parser.add_argument("--sheet-name", default="Statement", help="Excel worksheet name (default: Statement)")
    args = parser.parse_args()

    if not args.csv and not args.excel:
        parser.error("At least one of --csv or --excel must be provided")

    df = extract_statement_records(args.pdf_path, password=args.password)

    if args.csv:
        df.to_csv(args.csv, index=False)
    if args.excel:
        df.to_excel(args.excel, index=False, sheet_name=args.sheet_name)


if __name__ == "__main__":
    main()
