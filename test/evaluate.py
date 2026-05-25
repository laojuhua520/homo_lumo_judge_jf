#!/usr/bin/env python3
"""Score the HOMO/LUMO submission tables.

The evaluator intentionally uses only the Python standard library.  The local
submission files may be ordinary CSV files, or Excel workbooks saved with a
.csv extension, so both formats are supported here.
"""

from __future__ import annotations

import argparse
import csv
import json
import re
import subprocess
import sys
import xml.etree.ElementTree as ET
from dataclasses import asdict, dataclass
from decimal import Decimal, InvalidOperation
from pathlib import Path, PurePosixPath
from typing import Iterable
from zipfile import BadZipFile, ZipFile


MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
OFFICE_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PACKAGE_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
NS = {"main": MAIN_NS}


@dataclass
class ScoreItem:
    name: str
    score: int
    max_score: int
    passed: bool
    detail: str


@dataclass
class EvaluationResult:
    total_score: int
    max_score: int
    items: list[ScoreItem]
    gate_detail: str | None


def normalize_header(value: str) -> str:
    return "".join(ch for ch in str(value).lower().strip() if ch.isalnum())


def normalize_id(value: object) -> str:
    text = str(value).strip()
    try:
        number = Decimal(text)
    except InvalidOperation:
        return text
    if number == number.to_integral_value():
        return str(int(number))
    return text


def parse_float(value: object, label: str) -> float:
    text = str(value).strip()
    if not text:
        raise ValueError(f"{label} is empty")
    match = re.search(r"[-+]?(?:\d+(?:\.\d*)?|\.\d+)(?:[Ee][-+]?\d+)?", text)
    if not match:
        raise ValueError(f"{label} is not numeric: {value!r}")
    return float(match.group(0))


def is_xlsx_file(path: Path) -> bool:
    with path.open("rb") as handle:
        return handle.read(4) == b"PK\x03\x04"


def column_index(cell_ref: str) -> int:
    letters = "".join(ch for ch in cell_ref if ch.isalpha())
    index = 0
    for letter in letters:
        index = index * 26 + ord(letter.upper()) - ord("A") + 1
    return max(index - 1, 0)


def read_shared_strings(zf: ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in zf.namelist():
        return []
    root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
    strings: list[str] = []
    for item in root.findall("main:si", NS):
        strings.append("".join(text.text or "" for text in item.findall(".//main:t", NS)))
    return strings


def workbook_sheets(zf: ZipFile) -> list[tuple[str, str]]:
    workbook = ET.fromstring(zf.read("xl/workbook.xml"))
    rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
    rel_map = {
        rel.attrib["Id"]: rel.attrib["Target"]
        for rel in rels.findall(f"{{{PACKAGE_REL_NS}}}Relationship")
    }

    sheets: list[tuple[str, str]] = []
    for sheet in workbook.findall("main:sheets/main:sheet", NS):
        name = sheet.attrib["name"]
        rel_id = sheet.attrib[f"{{{OFFICE_REL_NS}}}id"]
        target = rel_map[rel_id]
        if target.startswith("/"):
            path = target.lstrip("/")
        elif target.startswith("xl/"):
            path = target
        else:
            path = str(PurePosixPath("xl") / target)
        sheets.append((name, path))
    return sheets


def text_from_cell(cell: ET.Element, shared_strings: list[str]) -> str:
    cell_type = cell.attrib.get("t")
    if cell_type == "inlineStr":
        return "".join(text.text or "" for text in cell.findall(".//main:t", NS)).strip()

    value_elem = cell.find("main:v", NS)
    if value_elem is None or value_elem.text is None:
        return ""

    value = value_elem.text
    if cell_type == "s":
        try:
            return shared_strings[int(value)].strip()
        except (IndexError, ValueError):
            return ""
    if cell_type == "b":
        return "TRUE" if value == "1" else "FALSE"
    return value.strip()


def read_xlsx_rows(path: Path) -> list[tuple[str, int, list[str]]]:
    rows: list[tuple[str, int, list[str]]] = []
    try:
        with ZipFile(path) as zf:
            shared_strings = read_shared_strings(zf)
            for sheet_name, sheet_path in workbook_sheets(zf):
                root = ET.fromstring(zf.read(sheet_path))
                for row_elem in root.findall("main:sheetData/main:row", NS):
                    row_number = int(row_elem.attrib.get("r", len(rows) + 1))
                    cells: list[str] = []
                    for cell in row_elem.findall("main:c", NS):
                        idx = column_index(cell.attrib.get("r", "A1"))
                        while len(cells) <= idx:
                            cells.append("")
                        cells[idx] = text_from_cell(cell, shared_strings)
                    rows.append((sheet_name, row_number, cells))
    except (BadZipFile, KeyError, ET.ParseError) as exc:
        raise ValueError(f"{path} is not a readable XLSX workbook") from exc
    return rows


def read_csv_rows(path: Path) -> list[tuple[str, int, list[str]]]:
    with path.open(newline="", encoding="utf-8-sig") as handle:
        sample = handle.read(4096)
        handle.seek(0)
        try:
            dialect = csv.Sniffer().sniff(sample, delimiters=",\t;")
        except csv.Error:
            dialect = csv.excel
        return [
            ("csv", line_number, [cell.strip() for cell in row])
            for line_number, row in enumerate(csv.reader(handle, dialect), start=1)
        ]


def read_table(path: Path) -> list[dict[str, str]]:
    if not path.is_file():
        raise FileNotFoundError(path)

    raw_rows = read_xlsx_rows(path) if is_xlsx_file(path) else read_csv_rows(path)
    rows_by_sheet: dict[str, list[tuple[int, list[str]]]] = {}
    for sheet_name, row_number, cells in raw_rows:
        rows_by_sheet.setdefault(sheet_name, []).append((row_number, cells))

    data: list[dict[str, str]] = []
    for sheet_name, sheet_rows in rows_by_sheet.items():
        non_empty = [
            (row_number, cells)
            for row_number, cells in sheet_rows
            if any(str(cell).strip() for cell in cells)
        ]
        if not non_empty:
            continue

        header_row_number, header = non_empty[0]
        header = [cell.strip() for cell in header]
        if not any(header):
            continue

        for row_number, cells in non_empty[1:]:
            padded = cells + [""] * max(0, len(header) - len(cells))
            row = {
                header[index]: padded[index].strip()
                for index in range(len(header))
                if header[index].strip()
            }
            row["_sheet"] = sheet_name
            row["_row"] = str(row_number)
            row["_header_row"] = str(header_row_number)
            data.append(row)
    return data


def find_file(root: Path, names: Iterable[str]) -> Path | None:
    for name in names:
        direct = root / name
        if direct.is_file():
            return direct
    for name in names:
        matches = sorted(root.rglob(name))
        for match in matches:
            if match.is_file():
                return match
    return None


def value_by_alias(row: dict[str, str], aliases: Iterable[str]) -> str | None:
    lookup = {normalize_header(key): value for key, value in row.items()}
    for alias in aliases:
        value = lookup.get(normalize_header(alias))
        if value is not None and str(value).strip():
            return str(value).strip()
    return None


def column_name_by_alias(rows: list[dict[str, str]], aliases: Iterable[str]) -> str | None:
    alias_set = {normalize_header(alias) for alias in aliases}
    for row in rows:
        for key in row:
            if normalize_header(key) in alias_set:
                return key
    return None


def load_smiles_library(path: Path) -> tuple[str, set[str]]:
    text = path.read_text(encoding="utf-8", errors="ignore")
    smiles_set: set[str] = set()
    for line in text.splitlines():
        stripped = line.strip()
        if not stripped:
            continue
        parts = re.split(r"[\t,]", stripped)
        for part in parts:
            item = part.strip()
            if item and normalize_header(item) not in {"file", "smiles"}:
                smiles_set.add(item)
    return text, smiles_set


def score_smiles_membership(root: Path) -> ScoreItem:
    product_path = find_file(root, ["product_stiles.txt", "product_smiles.txt"])
    if product_path is None:
        return ScoreItem(
            "duplicate_unit_smiles",
            0,
            10,
            False,
            "missing product_stiles.txt/product_smiles.txt library",
        )

    library_text, library_smiles = load_smiles_library(product_path)
    table_specs = [
        ("substitute.csv", ["substitute.csv"]),
        ("conjugate.csv", ["conjugate.csv", "concatenate.csv"]),
        ("minEg.csv", ["minEg.csv", "mineg.csv"]),
    ]

    checked = 0
    failures: list[str] = []
    for label, names in table_specs:
        path = find_file(root, names)
        if path is None:
            failures.append(f"{label}: file missing")
            continue
        try:
            rows = read_table(path)
        except Exception as exc:  # noqa: BLE001 - keep scoring resilient.
            failures.append(f"{label}: cannot read file ({exc})")
            continue
        if not rows:
            failures.append(f"{label}: no data rows")
            continue
        for row in rows:
            checked += 1
            row_no = row.get("_row", "?")
            smiles = value_by_alias(row, ["SMILES", "SMILE", "smiles"])
            if smiles is None:
                failures.append(f"{label} row {row_no}: missing SMILES")
                continue
            if smiles not in library_smiles and smiles not in library_text:
                failures.append(f"{label} row {row_no}: SMILES not found in library")

    if failures:
        detail = "; ".join(failures[:8])
        if len(failures) > 8:
            detail += f"; plus {len(failures) - 8} more"
        return ScoreItem("duplicate_unit_smiles", 0, 10, False, detail)

    return ScoreItem(
        "duplicate_unit_smiles",
        10,
        10,
        True,
        f"all {checked} SMILES entries are present in {product_path.name}",
    )


def score_orca_config(root: Path) -> ScoreItem:
    path = find_file(root, ["orca_config.csv"])
    if path is None:
        return ScoreItem("orca_config", 0, 10, False, "missing orca_config.csv")

    try:
        rows = read_table(path)
    except Exception as exc:  # noqa: BLE001
        return ScoreItem("orca_config", 0, 10, False, f"cannot read orca_config.csv ({exc})")
    if not rows:
        return ScoreItem("orca_config", 0, 10, False, "orca_config.csv has no data rows")

    basis_aliases = ["set_mase", "set_base", "set_basis", "basis", "basis_set", "base_set", "base"]
    method_aliases = ["dft_method", "ft_method", "functional", "method", "dft"]
    basis_col = column_name_by_alias(rows, basis_aliases)
    method_col = column_name_by_alias(rows, method_aliases)
    if basis_col is None or method_col is None:
        return ScoreItem(
            "orca_config",
            0,
            10,
            False,
            "missing set_mase/set_base or dft_method column",
        )

    basis_values = {row.get(basis_col, "").strip() for row in rows if row.get(basis_col, "").strip()}
    method_values = {row.get(method_col, "").strip() for row in rows if row.get(method_col, "").strip()}
    if len(basis_values) != 1 or len(method_values) != 1:
        return ScoreItem(
            "orca_config",
            0,
            10,
            False,
            f"{basis_col} values={sorted(basis_values)}, {method_col} values={sorted(method_values)}",
        )

    basis = next(iter(basis_values))
    method = next(iter(method_values))
    basis_lower = basis.lower()
    method_lower = method.lower()

    strict_ok = method_lower == "b3lyp" and "def2" in basis_lower
    swapped_ok = basis_lower == "b3lyp" and "def2" in method_lower
    if strict_ok:
        return ScoreItem(
            "orca_config",
            10,
            10,
            True,
            f"consistent {method_col}={method} and {basis_col}={basis}",
        )
    if swapped_ok:
        return ScoreItem(
            "orca_config",
            10,
            10,
            True,
            f"consistent B3LYP/def2 settings found; columns appear swapped ({basis_col}={basis}, {method_col}={method})",
        )

    return ScoreItem(
        "orca_config",
        0,
        10,
        False,
        f"expected B3LYP method and def2 basis, got {method_col}={method}, {basis_col}={basis}",
    )


def percentage_values_from_manifest(path: Path) -> list[float]:
    if not path.is_file():
        return []
    with path.open(newline="", encoding="utf-8-sig") as handle:
        reader = csv.DictReader(handle)
        if not reader.fieldnames:
            return []
        fields = {normalize_header(name): name for name in reader.fieldnames}
        candidates = [
            fields[name]
            for name in [
                "percentdifferenceabs",
                "percentdifference",
                "differencepercent",
                "finalresult",
            ]
            if name in fields
        ]
        values: list[float] = []
        for row in reader:
            for field in candidates:
                value = (row.get(field) or "").strip()
                if value:
                    try:
                        values.append(parse_float(value, field))
                    except ValueError:
                        pass
        return values


def percentage_values_from_text(text: str) -> list[float]:
    values: list[float] = []
    for match in re.finditer(
        r"([-+]?(?:\d+(?:\.\d*)?|\.\d+)(?:[Ee][-+]?\d+)?)\s*%",
        text,
    ):
        values.append(float(match.group(1)))
    return values


def default_calculation_script(root: Path) -> Path | None:
    return find_file(root, ["judge_calculate.py", "judge_recalculate.py"])


def calculation_command(root: Path, script: Path) -> list[str]:
    command = [sys.executable, str(script)]
    if script.name == "judge_recalculate.py":
        command.extend(["--root", str(root), "--run-orca"])
    return command


def effective_timeout(timeout: int | None) -> int | None:
    if timeout is None or timeout <= 0:
        return None
    return timeout


def score_reliability(root: Path, script: Path | None, timeout: int | None) -> ScoreItem:
    script = script or default_calculation_script(root)
    if script is None:
        return ScoreItem(
            "calculation_reliability",
            0,
            10,
            False,
            "judge_calculate.py/judge_recalculate.py was not found, so the <=20% check could not be run",
        )

    if not script.is_absolute():
        script = (root / script).resolve()
    if not script.is_file():
        return ScoreItem("calculation_reliability", 0, 10, False, f"{script} does not exist")

    command = calculation_command(root, script)
    timeout_seconds = effective_timeout(timeout)
    try:
        result = subprocess.run(
            command,
            cwd=root,
            text=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            timeout=timeout_seconds,
            check=False,
        )
    except subprocess.TimeoutExpired:
        return ScoreItem(
            "calculation_reliability",
            0,
            10,
            False,
            f"{script.name} timed out after {timeout_seconds} seconds",
        )
    except OSError as exc:
        return ScoreItem(
            "calculation_reliability",
            0,
            10,
            False,
            f"failed to run {script.name}: {exc}",
        )

    output = result.stdout or ""
    percentages = percentage_values_from_text(output)
    for manifest_name in ["manifest.csv", "generated_inps/manifest.csv"]:
        percentages.extend(percentage_values_from_manifest(root / manifest_name))

    if result.returncode != 0:
        tail = " ".join(output.splitlines()[-3:])[:240]
        return ScoreItem(
            "calculation_reliability",
            0,
            10,
            False,
            f"{script.name} exited with code {result.returncode}; {tail}",
        )

    if not percentages:
        return ScoreItem(
            "calculation_reliability",
            0,
            10,
            False,
            f"{script.name} ran, but no percentage result was found",
        )

    worst = max(abs(value) for value in percentages)
    if worst <= 20.0:
        return ScoreItem(
            "calculation_reliability",
            10,
            10,
            True,
            f"maximum reported difference is {worst:.3g}%",
        )
    return ScoreItem(
        "calculation_reliability",
        0,
        10,
        False,
        f"maximum reported difference is {worst:.3g}%, above 20%",
    )


def score_substitute(root: Path) -> ScoreItem:
    path = find_file(root, ["substitute.csv"])
    if path is None:
        return ScoreItem("substitute_fixed_units", 0, 20, False, "missing substitute.csv")

    try:
        rows = read_table(path)
    except Exception as exc:  # noqa: BLE001
        return ScoreItem("substitute_fixed_units", 0, 20, False, f"cannot read substitute.csv ({exc})")

    required = ["3", "6", "12", "13", "19"]
    selected: dict[str, dict[str, str]] = {}
    duplicates: set[str] = set()
    for row in rows:
        x = value_by_alias(row, ["X"])
        if x is None:
            continue
        x_norm = normalize_id(x)
        if x_norm in required:
            if x_norm in selected:
                duplicates.add(x_norm)
            else:
                selected[x_norm] = row

    missing = [x for x in required if x not in selected]
    if missing:
        return ScoreItem(
            "substitute_fixed_units",
            0,
            20,
            False,
            f"missing required X values: {', '.join(missing)}",
        )
    if duplicates:
        return ScoreItem(
            "substitute_fixed_units",
            0,
            20,
            False,
            f"duplicate required X values: {', '.join(sorted(duplicates))}",
        )

    y_values = {normalize_id(value_by_alias(row, ["Y"]) or "") for row in selected.values()}
    if len(y_values) != 1 or "" in y_values:
        return ScoreItem(
            "substitute_fixed_units",
            0,
            20,
            False,
            f"required X rows do not share one Y value: {sorted(y_values)}",
        )

    try:
        homos = [parse_float(value_by_alias(selected[x], ["HOMO"]) or "", f"HOMO for X={x}") for x in required]
    except ValueError as exc:
        return ScoreItem("substitute_fixed_units", 0, 20, False, str(exc))

    if all(later < earlier for earlier, later in zip(homos, homos[1:])):
        return ScoreItem(
            "substitute_fixed_units",
            20,
            20,
            True,
            f"X={required} share Y={next(iter(y_values))}; HOMO decreases as X increases",
        )
    return ScoreItem(
        "substitute_fixed_units",
        0,
        20,
        False,
        f"HOMO is not strictly decreasing for X={required}: {homos}",
    )


def score_conjugate(root: Path) -> ScoreItem:
    path = find_file(root, ["conjugate.csv", "concatenate.csv"])
    if path is None:
        return ScoreItem("conjugate_fixed_units", 0, 20, False, "missing conjugate.csv/concatenate.csv")

    try:
        rows = read_table(path)
    except Exception as exc:  # noqa: BLE001
        return ScoreItem("conjugate_fixed_units", 0, 20, False, f"cannot read {path.name} ({exc})")

    required = ["4", "5", "8", "11"]
    selected: dict[str, dict[str, str]] = {}
    duplicates: set[str] = set()
    for row in rows:
        y = value_by_alias(row, ["Y"])
        if y is None:
            continue
        y_norm = normalize_id(y)
        if y_norm in required:
            if y_norm in selected:
                duplicates.add(y_norm)
            else:
                selected[y_norm] = row

    missing = [y for y in required if y not in selected]
    if missing:
        return ScoreItem(
            "conjugate_fixed_units",
            0,
            20,
            False,
            f"missing required Y values: {', '.join(missing)}",
        )
    if duplicates:
        return ScoreItem(
            "conjugate_fixed_units",
            0,
            20,
            False,
            f"duplicate required Y values: {', '.join(sorted(duplicates))}",
        )

    x_values = {normalize_id(value_by_alias(row, ["X"]) or "") for row in selected.values()}
    if len(x_values) != 1 or "" in x_values:
        return ScoreItem(
            "conjugate_fixed_units",
            0,
            20,
            False,
            f"required Y rows do not share one X value: {sorted(x_values)}",
        )

    try:
        egs = [parse_float(value_by_alias(selected[y], ["Eg", "E_g", "gap"]) or "", f"Eg for Y={y}") for y in required]
    except ValueError as exc:
        return ScoreItem("conjugate_fixed_units", 0, 20, False, str(exc))

    if all(later < earlier for earlier, later in zip(egs, egs[1:])):
        return ScoreItem(
            "conjugate_fixed_units",
            20,
            20,
            True,
            f"Y={required} share X={next(iter(x_values))}; Eg decreases as Y increases",
        )
    return ScoreItem(
        "conjugate_fixed_units",
        0,
        20,
        False,
        f"Eg is not strictly decreasing for Y={required}: {egs}",
    )


def score_min_eg(root: Path) -> ScoreItem:
    path = find_file(root, ["minEg.csv", "mineg.csv"])
    if path is None:
        return ScoreItem("minEg_highest_Eg", 0, 30, False, "missing minEg.csv")

    try:
        rows = read_table(path)
    except Exception as exc:  # noqa: BLE001
        return ScoreItem("minEg_highest_Eg", 0, 30, False, f"cannot read minEg.csv ({exc})")

    values: list[float] = []
    for row in rows:
        value = value_by_alias(row, ["Eg", "E_g", "gap"])
        if value is None:
            continue
        try:
            values.append(parse_float(value, f"Eg row {row.get('_row', '?')}"))
        except ValueError:
            continue
    if not values:
        return ScoreItem("minEg_highest_Eg", 0, 30, False, "no numeric Eg values found")

    highest = max(values)
    if 5.5 < highest < 6.5:
        score = 30
    elif 4.2 < highest < 5.5:
        score = 20
    elif 3.5 < highest < 4.2:
        score = 10
    else:
        score = 0

    return ScoreItem(
        "minEg_highest_Eg",
        score,
        30,
        score > 0,
        f"highest Eg is {highest:g} eV",
    )


def evaluate(root: Path, calculation_script: Path | None, calculation_timeout: int) -> EvaluationResult:
    items = [
        score_smiles_membership(root),
        score_orca_config(root),
        score_reliability(root, calculation_script, calculation_timeout),
        score_substitute(root),
        score_conjugate(root),
        score_min_eg(root),
    ]
    max_score = sum(item.max_score for item in items)

    gate_detail: str | None = None
    if not items[0].passed:
        total_score = 0
        gate_detail = "SMILES library membership check failed, so the final score is 0."
    elif not items[1].passed:
        total_score = 10
        gate_detail = "ORCA configuration consistency check failed, so the final score is 10."
    elif not items[2].passed:
        total_score = 30
        gate_detail = "Calculation reliability check failed, so the final score is 30."
    else:
        total_score = sum(item.score for item in items)

    return EvaluationResult(total_score, max_score, items, gate_detail)


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--root",
        type=Path,
        default=Path.cwd(),
        help="Directory containing the submission CSV files. Default: current directory.",
    )
    parser.add_argument(
        "--calculation-script",
        type=Path,
        help=(
            "Path to judge_calculate.py or judge_recalculate.py. "
            "Default: auto-detect under --root."
        ),
    )
    parser.add_argument(
        "--calculation-timeout",
        type=int,
        default=0,
        help=(
            "Timeout in seconds for the calculation reliability script. "
            "Use 0 to wait indefinitely. Default: 0."
        ),
    )
    parser.add_argument(
        "--json",
        action="store_true",
        help="Print machine-readable JSON instead of the text report.",
    )
    return parser.parse_args(argv)


def print_text_report(result: EvaluationResult) -> None:
    print(f"Final score: {result.total_score}/{result.max_score}")
    if result.gate_detail:
        print(result.gate_detail)
    for item in result.items:
        status = "PASS" if item.passed else "FAIL"
        score_text = f"{item.score}/{item.max_score} " if item.score == item.max_score else ""
        print(f"{item.name}: {score_text}[{status}] - {item.detail}")


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    root = args.root.resolve()
    result = evaluate(root, args.calculation_script, args.calculation_timeout)
    if args.json:
        print(json.dumps(asdict(result), ensure_ascii=False, indent=2))
    else:
        print_text_report(result)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
