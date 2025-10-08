"""Tools for categorising financial metrics from PDF reports into Excel workbooks.

The module supports the following workflow:

1. Create an Excel template that contains a sheet named ``Mapping``.  Each row of
   this sheet should describe how a specific financial metric is identified and
   where the extracted value has to be written.  The sheet supports the
   following columns (case insensitive):

   - ``Category`` *(required)* – a human readable name for the metric.
   - ``Pattern`` *(required)* – a semicolon separated list of keywords that are
     expected to appear in the PDF before the target number.  To use a custom
     regular expression wrap it in forward slashes, e.g. ``/Revenue\s+\$?([\d,.]+)/``.
   - ``Sheet`` *(optional)* – the worksheet where the value should be written.
     When omitted the values are written to a sheet called ``AutoFilled``.
   - ``Cell`` *(optional)* – the Excel cell address (e.g. ``B3``).  When this is
     missing the writer will look for the first cell in the target sheet whose
     value matches ``Category`` and place the numeric value in the adjacent
     column.
   - ``Scale`` *(optional)* – multiplier that should be applied to the parsed
     value (e.g. ``1e6`` when the PDF reports the figure in millions).
   - ``Aggregation`` *(optional)* – one of ``first`` (default), ``sum`` or
     ``average`` describing how multiple matches should be combined.

2. Place the PDF files that should be analysed inside a directory.

3. Run :func:`extract_values_from_directory` to parse the PDFs and
   :func:`write_values_to_template` to populate a copy of the template with the
   discovered figures.

The design aims to keep the logic deterministic so that the filled Excel sheet
can be reviewed and adjusted by analysts.  Where automated extraction fails to
find a value, the output sheet clearly indicates that the number is missing and
provides the PDF page references for any partial matches.  This makes it easy
to manually fill in the blanks without repeating the full extraction process.
"""

from __future__ import annotations

import argparse
import dataclasses
import logging
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Iterator, List, Mapping, Optional, Sequence, Tuple

import pandas as pd
import pdfplumber
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter


LOGGER = logging.getLogger(__name__)


@dataclass
class FinancialCategory:
    """Definition of a financial metric that should be extracted."""

    name: str
    patterns: Sequence[str]
    sheet: str = "AutoFilled"
    cell: Optional[str] = None
    scale: float = 1.0
    aggregation: str = "first"

    def __post_init__(self) -> None:
        self.patterns = [p for pat in self.patterns for p in _normalise_pattern(pat)]
        if not self.patterns:
            raise ValueError(f"Category '{self.name}' must have at least one pattern")
        if self.aggregation not in {"first", "sum", "average"}:
            raise ValueError(
                "Aggregation must be one of 'first', 'sum', or 'average', got "
                f"{self.aggregation!r} for category {self.name!r}."
            )


@dataclass
class ParsedValue:
    """Represents a numeric value that was extracted from a PDF."""

    value: float
    source_file: Path
    page_number: int
    matched_text: str


def _normalise_pattern(pattern: str) -> List[str]:
    """Split a pattern string into individual regular expressions.

    Users can provide a semicolon or newline separated list of patterns in the
    template.  Empty items are removed and the resulting list is stripped of
    whitespace.
    """

    parts = re.split(r"[\n;]+", pattern)
    return [part.strip() for part in parts if part.strip()]


def _is_missing(value: object) -> bool:
    if value in (None, ""):
        return True
    try:
        return bool(pd.isna(value))
    except TypeError:
        return False


def _normalise_optional_string(value: object) -> Optional[str]:
    if _is_missing(value):
        return None
    text = str(value).strip()
    return text or None


def _load_category_from_row(row: Mapping[str, object]) -> FinancialCategory:
    name = str(row.get("category") or row.get("Category") or "").strip()
    pattern = str(row.get("pattern") or row.get("Pattern") or "").strip()
    if not name or not pattern:
        raise ValueError("Each mapping row must define both 'Category' and 'Pattern'.")

    sheet = str(row.get("sheet") or row.get("Sheet") or "AutoFilled").strip() or "AutoFilled"
    cell_raw = row.get("cell") or row.get("Cell")
    cell = _normalise_optional_string(cell_raw)
    scale_raw = row.get("scale") or row.get("Scale")
    scale = float(scale_raw) if not _is_missing(scale_raw) else 1.0
    aggregation = (
        str(row.get("aggregation") or row.get("Aggregation") or "first").strip().lower() or "first"
    )

    return FinancialCategory(
        name=name,
        patterns=[pattern],
        sheet=sheet,
        cell=cell,
        scale=scale,
        aggregation=aggregation,
    )


def load_categories(template_path: Path, mapping_sheet: str = "Mapping") -> List[FinancialCategory]:
    """Load financial categories from the Excel template.

    Parameters
    ----------
    template_path:
        Path to the Excel workbook that serves as the template.
    mapping_sheet:
        Name of the worksheet that contains the mapping definition.
    """

    try:
        mapping_df = pd.read_excel(template_path, sheet_name=mapping_sheet)
    except ValueError as exc:  # sheet not found
        raise ValueError(
            f"Sheet {mapping_sheet!r} could not be found in template {template_path}."
        ) from exc

    categories: List[FinancialCategory] = []
    for row in mapping_df.to_dict(orient="records"):
        if all(_is_missing(value) for value in row.values()):
            continue
        categories.append(_load_category_from_row(row))
    if not categories:
        raise ValueError("No categories were loaded from the template mapping sheet.")
    return categories


def extract_values_from_directory(
    pdf_directory: Path, categories: Sequence[FinancialCategory]
) -> Dict[str, List[ParsedValue]]:
    """Extract values for the categories from all PDFs in a directory."""

    if not pdf_directory.exists():
        raise FileNotFoundError(f"Directory {pdf_directory} does not exist.")

    results: Dict[str, List[ParsedValue]] = {category.name: [] for category in categories}
    pdf_files = sorted(
        [path for path in pdf_directory.iterdir() if path.suffix.lower() in {".pdf"}]
    )
    if not pdf_files:
        LOGGER.warning("No PDF files found in %s", pdf_directory)
        return results

    compiled_patterns: Dict[str, List[re.Pattern[str]]] = {
        category.name: [_compile_pattern(pattern) for pattern in category.patterns]
        for category in categories
    }

    for pdf_path in pdf_files:
        LOGGER.info("Processing %s", pdf_path)
        with pdfplumber.open(pdf_path) as pdf:
            for page_number, page in enumerate(pdf.pages, start=1):
                text = page.extract_text() or ""
                for category in categories:
                    matches = _search_patterns(text, compiled_patterns[category.name])
                    for value, snippet in matches:
                        parsed = ParsedValue(
                            value=value * category.scale,
                            source_file=pdf_path,
                            page_number=page_number,
                            matched_text=snippet,
                        )
                        results[category.name].append(parsed)

    return results


def _compile_pattern(pattern: str) -> re.Pattern[str]:
    if pattern.startswith("/") and pattern.endswith("/") and len(pattern) > 2:
        pattern_body = pattern[1:-1]
    else:
        escaped = re.escape(pattern)
        pattern_body = rf"{escaped}[\s:;\-]*([\$€£]?[-+]?\(?\d[\d,\.]*\)?(?:\s*(?:thousand|million|billion))?)"
    return re.compile(pattern_body, re.IGNORECASE)


def _search_patterns(original_text: str, patterns: Sequence[re.Pattern[str]]) -> Iterator[Tuple[float, str]]:
    for pattern in patterns:
        for match in pattern.finditer(original_text):
            value_group = match.group(match.lastindex or 1)
            if not value_group:
                continue
            value = _parse_numeric(value_group)
            if value is None:
                continue
            span_start, span_end = match.span()
            snippet_start = max(0, span_start - 40)
            snippet_end = min(len(original_text), span_end + 40)
            snippet = original_text[snippet_start:snippet_end].strip()
            yield value, snippet


def _parse_numeric(raw_value: str) -> Optional[float]:
    text = raw_value.strip()
    multiplier = 1.0
    lower = text.lower()
    for keyword, factor in {
        "thousand": 1_000,
        "million": 1_000_000,
        "billion": 1_000_000_000,
    }.items():
        if keyword in lower:
            multiplier = factor
            text = re.sub(keyword, "", text, flags=re.IGNORECASE)
            break

    is_percentage = "%" in text
    text = text.replace("%", "").strip()
    text = text.replace(",", "")
    text = text.replace("$", "")
    negative = text.startswith("(") and text.endswith(")")
    text = text.strip("() ")
    if not text:
        return None
    try:
        value = float(text)
    except ValueError:
        return None

    if negative:
        value = -value
    value *= multiplier
    if is_percentage:
        value /= 100.0
    return value


def aggregate_category_values(
    category: FinancialCategory, values: Sequence[ParsedValue]
) -> Optional[ParsedValue]:
    if not values:
        return None
    if category.aggregation == "first":
        return values[0]
    if category.aggregation == "sum":
        total = sum(value.value for value in values)
        return dataclasses.replace(values[0], value=total)
    if category.aggregation == "average":
        total = sum(value.value for value in values)
        avg = total / len(values)
        return dataclasses.replace(values[0], value=avg)
    raise ValueError(f"Unknown aggregation method: {category.aggregation}")


def write_values_to_template(
    template_path: Path,
    output_path: Path,
    categories: Sequence[FinancialCategory],
    extracted_values: Mapping[str, Sequence[ParsedValue]],
) -> None:
    """Populate a copy of the template with the extracted values."""

    workbook = load_workbook(template_path)
    notes_sheet = workbook.create_sheet("ExtractionNotes")
    notes_sheet.append([
        "Category",
        "Value",
        "Source File",
        "Page",
        "Context",
    ])

    for category in categories:
        values = list(extracted_values.get(category.name, []))
        aggregated = aggregate_category_values(category, values)

        target_sheet = workbook[category.sheet] if category.sheet in workbook.sheetnames else workbook.create_sheet(category.sheet)
        target_cell = category.cell or _find_cell_for_category(target_sheet, category.name)

        if aggregated is not None and target_cell:
            target_sheet[target_cell].value = aggregated.value
        elif target_cell:
            target_sheet[target_cell].value = None

        for value in values:
            notes_sheet.append(
                [
                    category.name,
                    value.value,
                    value.source_file.name,
                    value.page_number,
                    value.matched_text,
                ]
            )
        if not values:
            notes_sheet.append([category.name, None, None, None, "No match found"])

    workbook.save(output_path)


def _find_cell_for_category(sheet, category_name: str) -> Optional[str]:
    for row in sheet.iter_rows():
        for cell in row:
            if isinstance(cell.value, str) and cell.value.strip().lower() == category_name.lower():
                column_index = cell.column + 1
                return f"{get_column_letter(column_index)}{cell.row}"
    # If not found, append at bottom
    next_row = sheet.max_row + 1
    sheet.cell(row=next_row, column=1, value=category_name)
    return f"B{next_row}"


def main(argv: Optional[Sequence[str]] = None) -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("pdf_directory", type=Path, help="Directory containing PDF reports")
    parser.add_argument("template", type=Path, help="Excel template with Mapping sheet")
    parser.add_argument("output", type=Path, help="Output Excel file path")
    parser.add_argument(
        "--mapping-sheet",
        default="Mapping",
        help="Name of the worksheet that defines the extraction mapping",
    )
    parser.add_argument(
        "--log-level", default="INFO", help="Logging level (DEBUG, INFO, WARNING, ERROR)"
    )

    args = parser.parse_args(argv)
    logging.basicConfig(level=getattr(logging, args.log_level.upper(), logging.INFO))

    categories = load_categories(args.template, mapping_sheet=args.mapping_sheet)
    values = extract_values_from_directory(args.pdf_directory, categories)
    write_values_to_template(args.template, args.output, categories, values)


if __name__ == "__main__":
    main()
