
from __future__ import annotations

import io
import re
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable

import pandas as pd
import pdfplumber
from openpyxl import load_workbook


DATE_RE = re.compile(r"^\d{1,2}/\d{1,2}/\d{2,4}$")


# ---------------------------------------------------------------------------
# Known MOR template: KUB / Fourth Creek page 1
#
# These are not character offsets. They are PDF x-coordinate anchors taken
# from the printed table geometry. This avoids the problem where pdftotext
# spacing changes from month to month.
# ---------------------------------------------------------------------------

KUB_PAGE1_FIELDS = [
    ("Date", 43.8),
    ("Rainfall (in)", 108.3),
    ("Average Flow (MGD)", 141.8),
    ("Maximum Flow (MGD)", 175.4),
    ("Minimum Flow (MGD)", 209.0),
    ("Bypass Hours", 242.6),
    ("Influent Temperature (°C)", 276.1),
    ("BOD Raw (mg/L)", 309.7),
    ("BOD Intermediate (mg/L)", 343.3),
    ("BOD Final (mg/L)", 376.8),
    ("BOD Reduction (%)", 410.4),
    ("Suspended Solids Raw (mg/L)", 444.0),
    ("Suspended Solids Intermediate (mg/L)", 477.6),
    ("Suspended Solids Final (mg/L)", 511.1),
    ("Suspended Solids Reduction (%)", 544.7),
    ("Set Solids Raw", 578.3),
    ("Set Solids Intermediate", 611.9),
    ("Set Solids Final", 645.4),
    ("DO (mg/L)", 679.0),
    ("pH Raw", 712.6),
    ("pH Intermediate", 746.1),
    ("pH Final", 779.7),
    ("NH3-N Influent (ppm)", 813.3),
    ("Grease Influent (ppm)", 846.9),
    ("Lbs Cl2", 880.4),
    ("Total N Influent (mg/L)", 914.0),
    ("Total P Influent (mg/L)", 947.6),
    ("E. coli", 981.2),
    ("Cl2 Residual", 1014.7),
    ("NH3-N Compliance Final Effluent (ppm)", 1048.3),
    ("Grease Final Effluent (ppm)", 1081.9),
    ("Effluent Flow (MGD)", 1115.4),
    ("Total N Effluent (mg/L)", 1149.0),
    ("Total P Effluent (mg/L)", 1182.6),
]

KUB_PAGE2_FIELDS = [
    ("Date", 43.8),
    ("Secondary System DO (ppm)", 164.3),
    ("MLSS (ppm)", 220.2),
    ("SVI", 276.2),
    ("30 Minute Set Solids", 332.2),
    ("Waste (GPD)", 388.1),
    ("Cadmium Influent (mg/L)", 444.1),
    ("Cadmium Effluent (mg/L)", 500.0),
    ("Chromium Influent (mg/L)", 556.0),
    ("Chromium Effluent (mg/L)", 611.9),
    ("Copper Influent (mg/L)", 667.9),
    ("Copper Effluent (mg/L)", 723.8),
    ("Nickel Influent (mg/L)", 779.8),
    ("Nickel Effluent (mg/L)", 835.7),
    ("Zinc Influent (mg/L)", 891.7),
    ("Zinc Effluent (mg/L)", 947.7),
    ("Silver Influent (mg/L)", 1003.6),
    ("Silver Effluent (mg/L)", 1059.6),
    ("Lead Influent (mg/L)", 1115.5),
    ("Lead Effluent (mg/L)", 1171.5),
]


@dataclass
class DetectedDataset:
    name: str
    source_name: str
    dataframe: pd.DataFrame
    confidence: str
    notes: list[str]


def clean_text(value) -> str:
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value).strip())


def parse_value(text: str):
    text = clean_text(text)
    if text == "":
        return None

    # Keep common non-numeric lab reporting strings as text.
    if text.startswith("<") or text.startswith(">"):
        return text

    candidate = text.replace(",", "")
    try:
        return float(candidate)
    except Exception:
        return text


def word_center(word) -> float:
    return (float(word["x0"]) + float(word["x1"])) / 2.0


def group_words_by_row(words, tolerance=2.2):
    """Group PDF words into visual rows based on y position."""
    ordered = sorted(words, key=lambda w: (float(w["top"]), float(w["x0"])))
    rows = []
    for w in ordered:
        top = float(w["top"])
        if not rows or abs(top - rows[-1]["top"]) > tolerance:
            rows.append({"top": top, "words": [w]})
        else:
            rows[-1]["words"].append(w)
            # Keep stable row top while allowing minor text baseline variation.
            rows[-1]["top"] = sum(float(x["top"]) for x in rows[-1]["words"]) / len(rows[-1]["words"])
    for row in rows:
        row["words"] = sorted(row["words"], key=lambda w: float(w["x0"]))
    return rows


def row_has_date(row) -> bool:
    return any(DATE_RE.match(clean_text(w["text"])) for w in row["words"])


def extract_date_rows(page):
    words = page.extract_words(
        x_tolerance=1,
        y_tolerance=2,
        keep_blank_chars=False,
        use_text_flow=False,
    )
    rows = group_words_by_row(words)
    return [r for r in rows if row_has_date(r)], rows


def nearest_field(x: float, field_defs, tolerance: float):
    name, anchor = min(field_defs, key=lambda item: abs(x - item[1]))
    distance = abs(x - anchor)
    if distance <= tolerance:
        return name
    return None


def extract_known_page(page, field_defs, tolerance=16.5) -> pd.DataFrame:
    date_rows, _ = extract_date_rows(page)
    records = []

    for row in date_rows:
        record = {name: None for name, _ in field_defs}

        for word in row["words"]:
            text = clean_text(word["text"])
            x = word_center(word)

            if DATE_RE.match(text):
                record["Date"] = pd.to_datetime(text, errors="coerce")
                continue

            field = nearest_field(x, field_defs, tolerance)
            if field and field != "Date":
                # A valid data cell is normally one token. If there are two
                # tokens at the same anchor, preserve them rather than guessing.
                value = parse_value(text)
                if record[field] is None:
                    record[field] = value
                else:
                    record[field] = f"{record[field]} {text}"

        if record.get("Date") is not None and not pd.isna(record["Date"]):
            records.append(record)

    return pd.DataFrame(records, columns=[f[0] for f in field_defs])


def is_kub_fourth_creek(pdf) -> bool:
    try:
        text = (pdf.pages[0].extract_text() or "").lower()
    except Exception:
        return False

    required = [
        "report of operation of wastewater treatment plant",
        "influent flows",
        "5-day bod",
        "set solids",
        "final effluent parameters",
    ]
    return sum(term in text for term in required) >= 4


def extract_kub_pdf(data: bytes, source_name: str) -> list[DetectedDataset]:
    out = []
    with pdfplumber.open(io.BytesIO(data)) as pdf:
        if not pdf.pages:
            return out

        p1 = extract_known_page(pdf.pages[0], KUB_PAGE1_FIELDS)
        if not p1.empty:
            out.append(
                DetectedDataset(
                    name="Daily MOR - Page 1",
                    source_name=source_name,
                    dataframe=p1,
                    confidence="High",
                    notes=[
                        "Recognized Fourth Creek/KUB MOR layout.",
                        "Columns are assigned using PDF x-coordinates, not character spacing.",
                    ],
                )
            )

        if len(pdf.pages) >= 2:
            p2 = extract_known_page(pdf.pages[1], KUB_PAGE2_FIELDS)
            if not p2.empty:
                out.append(
                    DetectedDataset(
                        name="Secondary System - Page 2",
                        source_name=source_name,
                        dataframe=p2,
                        confidence="High",
                        notes=[
                            "Recognized Fourth Creek/KUB secondary-system table.",
                            "This DO is separate from the page-1 DO column.",
                        ],
                    )
                )
    return out


# ---------------------------------------------------------------------------
# Generic positional PDF fallback
# ---------------------------------------------------------------------------

def cluster_positions(values: list[float], tolerance=11.0):
    if not values:
        return []

    values = sorted(values)
    clusters = [[values[0]]]

    for x in values[1:]:
        center = sum(clusters[-1]) / len(clusters[-1])
        if abs(x - center) <= tolerance:
            clusters[-1].append(x)
        else:
            clusters.append([x])

    return [sum(c) / len(c) for c in clusters]


def find_header_rows(all_rows, first_data_top, max_rows=4):
    prior = [r for r in all_rows if r["top"] < first_data_top - 1]
    return prior[-max_rows:]


def generic_header_for_anchor(anchor, header_rows, boundaries):
    left, right = boundaries
    pieces = []

    # Use lower-level header rows first; upper group headings are only added
    # if the leaf label is too generic.
    for row in header_rows:
        words = [
            clean_text(w["text"])
            for w in row["words"]
            if left <= word_center(w) < right
        ]
        if words:
            pieces.append(" ".join(words))

    if not pieces:
        return ""

    leaf = pieces[-1]
    generic_leafs = {
        "raw", "inter.", "inter", "final", "ppm", "mg/l", "mgd",
        "in %", "%", "hours", "inches", "total n", "total p"
    }

    if leaf.lower() in generic_leafs and len(pieces) >= 2:
        return f"{pieces[-2]} {leaf}".strip()

    return leaf


def make_unique(names):
    counts = {}
    out = []
    for i, n in enumerate(names, start=1):
        name = clean_text(n) or f"Detected Column {i}"
        counts[name] = counts.get(name, 0) + 1
        if counts[name] > 1:
            out.append(f"{name} ({counts[name]})")
        else:
            out.append(name)
    return out


def extract_generic_pdf(data: bytes, source_name: str) -> list[DetectedDataset]:
    datasets = []

    with pdfplumber.open(io.BytesIO(data)) as pdf:
        for page_no, page in enumerate(pdf.pages, start=1):
            date_rows, all_rows = extract_date_rows(page)
            if len(date_rows) < 2:
                continue

            # Date is its own field. Numeric/text columns are inferred from x
            # positions that recur across daily rows.
            value_x = []
            date_anchor = None

            for row in date_rows:
                for w in row["words"]:
                    text = clean_text(w["text"])
                    x = word_center(w)
                    if DATE_RE.match(text):
                        if date_anchor is None:
                            date_anchor = x
                    else:
                        value_x.append(x)

            anchors = cluster_positions(value_x, tolerance=10.5)
            if not anchors:
                continue

            first_top = min(r["top"] for r in date_rows)
            header_rows = find_header_rows(all_rows, first_top, max_rows=4)

            # Cell boundaries are halfway between neighboring anchors.
            full_anchors = ([date_anchor] if date_anchor is not None else []) + anchors
            full_anchors = sorted(full_anchors)

            labels = []
            for idx, anchor in enumerate(full_anchors):
                if date_anchor is not None and abs(anchor - date_anchor) < 2:
                    labels.append("Date")
                    continue

                left = -1e9 if idx == 0 else (full_anchors[idx - 1] + anchor) / 2
                right = 1e9 if idx == len(full_anchors) - 1 else (anchor + full_anchors[idx + 1]) / 2
                labels.append(generic_header_for_anchor(anchor, header_rows, (left, right)))

            labels = make_unique(labels)

            records = []
            for row in date_rows:
                record = {label: None for label in labels}
                for w in row["words"]:
                    text = clean_text(w["text"])
                    x = word_center(w)

                    if DATE_RE.match(text) and "Date" in record:
                        record["Date"] = pd.to_datetime(text, errors="coerce")
                        continue

                    idx = min(range(len(full_anchors)), key=lambda i: abs(x - full_anchors[i]))
                    if abs(x - full_anchors[idx]) <= 16.0:
                        label = labels[idx]
                        if label == "Date":
                            continue
                        value = parse_value(text)
                        if record[label] is None:
                            record[label] = value
                        else:
                            record[label] = f"{record[label]} {text}"

                if record.get("Date") is not None and not pd.isna(record["Date"]):
                    records.append(record)

            df = pd.DataFrame(records, columns=labels)
            if len(df) >= 2 and len(df.columns) >= 2:
                datasets.append(
                    DetectedDataset(
                        name=f"Generic daily table - Page {page_no}",
                        source_name=source_name,
                        dataframe=df,
                        confidence="Review required",
                        notes=[
                            "This file did not match a saved MOR template.",
                            "Columns were inferred from repeated PDF x-positions and nearby header text.",
                            "Confirm sample values before exporting.",
                        ],
                    )
                )

    return datasets


def extract_pdf(data: bytes, source_name: str) -> list[DetectedDataset]:
    with pdfplumber.open(io.BytesIO(data)) as pdf:
        known = is_kub_fourth_creek(pdf)

    if known:
        return extract_kub_pdf(data, source_name)

    return extract_generic_pdf(data, source_name)


# ---------------------------------------------------------------------------
# Excel
# ---------------------------------------------------------------------------

def score_header_row(row):
    vals = [clean_text(x) for x in row if clean_text(x)]
    if not vals:
        return -999

    numeric = 0
    for v in vals:
        try:
            float(v.replace(",", ""))
            numeric += 1
        except Exception:
            pass

    text_ratio = 1 - numeric / len(vals)
    uniqueness = len(set(vals)) / len(vals)
    keywords = (
        "date", "flow", "bod", "tss", "solid", "ph", "nh3", "ammonia",
        "grease", "do", "mlss", "svi", "chlor", "nitrogen", "phosph"
    )
    keyword_score = min(sum(k in " ".join(vals).lower() for k in keywords) * 0.4, 2.0)
    return text_ratio * 2 + uniqueness + keyword_score


def extract_excel(data: bytes, source_name: str) -> list[DetectedDataset]:
    wb = load_workbook(io.BytesIO(data), data_only=True, read_only=True)
    datasets = []

    for ws in wb.worksheets:
        rows = [list(r) for r in ws.iter_rows(values_only=True)]
        rows = [r for r in rows if any(clean_text(v) for v in r)]

        if len(rows) < 2:
            continue

        max_scan = min(12, len(rows))
        header_idx = max(range(max_scan), key=lambda i: score_header_row(rows[i]))

        raw_headers = [clean_text(v) for v in rows[header_idx]]
        headers = make_unique(raw_headers)

        body = rows[header_idx + 1:]
        width = len(headers)
        normalized = []
        for row in body:
            row = list(row) + [None] * max(0, width - len(row))
            normalized.append(row[:width])

        df = pd.DataFrame(normalized, columns=headers)
        df = df.dropna(how="all")

        # Remove completely blank detected columns.
        keep = []
        for c in df.columns:
            if c.startswith("Detected Column"):
                series = df[c].fillna("").astype(str).str.strip()
                if not series.ne("").any():
                    continue
            keep.append(c)
        df = df[keep]

        if not df.empty and len(df.columns) >= 2:
            datasets.append(
                DetectedDataset(
                    name=f"Excel sheet: {ws.title}",
                    source_name=source_name,
                    dataframe=df,
                    confidence="High",
                    notes=["Read directly from Excel cell structure."],
                )
            )

    return datasets


# ---------------------------------------------------------------------------
# Upload / ZIP handling
# ---------------------------------------------------------------------------

def unpack_upload(filename: str, data: bytes):
    suffix = Path(filename).suffix.lower()

    if suffix in {".pdf", ".xlsx", ".xlsm"}:
        return [(filename, data)]

    if suffix == ".zip":
        items = []
        with zipfile.ZipFile(io.BytesIO(data)) as z:
            for info in z.infolist():
                if info.is_dir():
                    continue
                inner = Path(info.filename)
                ext = inner.suffix.lower()
                if ext in {".pdf", ".xlsx", ".xlsm"}:
                    items.append((inner.name, z.read(info)))
        return items

    return []


def detect_file(filename: str, data: bytes) -> list[DetectedDataset]:
    suffix = Path(filename).suffix.lower()
    if suffix == ".pdf":
        return extract_pdf(data, filename)
    if suffix in {".xlsx", ".xlsm"}:
        return extract_excel(data, filename)
    return []


def combine_same_named_datasets(datasets: list[DetectedDataset]):
    """
    Combine repeated monthly PDFs when they resolve to the same saved-template
    dataset and have the same columns. Generic tables remain separate.
    """
    groups = {}

    for ds in datasets:
        key = (ds.name, tuple(ds.dataframe.columns), ds.confidence)
        groups.setdefault(key, []).append(ds)

    out = []
    for key, members in groups.items():
        if len(members) == 1:
            out.append(members[0])
            continue

        combined = pd.concat([m.dataframe for m in members], ignore_index=True)
        if "Date" in combined.columns:
            combined = combined.sort_values("Date", kind="stable").reset_index(drop=True)

        out.append(
            DetectedDataset(
                name=members[0].name,
                source_name=f"{len(members)} files",
                dataframe=combined,
                confidence=members[0].confidence,
                notes=members[0].notes + [f"Combined {len(members)} matching files."],
            )
        )

    return out
