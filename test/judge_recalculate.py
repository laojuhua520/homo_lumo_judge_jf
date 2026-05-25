#!/usr/bin/env python3
"""Generate ORCA inputs and optionally compare ORCA/Multiwfn Eg with table Eg."""

from __future__ import annotations

import argparse
import csv
import random
import re
import shutil
import subprocess
import sys
import xml.etree.ElementTree as ET
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
from pathlib import Path, PurePosixPath
from zipfile import ZipFile

from mol2smiles import mol_file_to_smiles
from smiles2inp import orca_generator


MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
OFFICE_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PACKAGE_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
NS = {"main": MAIN_NS}


@dataclass(frozen=True)
class DataRow:
    source: Path
    sheet: str
    row_number: int
    values: dict[str, str]
    x: str
    y: str


@dataclass(frozen=True)
class OrcaSettings:
    functional: str
    basis: str


@dataclass(frozen=True)
class GeneratedResult:
    row: DataRow
    settings: OrcaSettings
    mol_path: Path
    smiles_path: Path
    inp_path: Path
    smiles: str
    csv_eg_ev: float | None = None
    out_path: Path | None = None
    holder_path: Path | None = None
    multiwfn_log_path: Path | None = None
    multiwfn_eg_ev: float | None = None
    percent_difference: float | None = None
    signed_percent_difference: float | None = None


def normalize_header(value: str) -> str:
    return "".join(ch for ch in value.lower().strip() if ch.isalnum())


def value_by_alias(row: dict[str, str], aliases: list[str]) -> str | None:
    lookup = {normalize_header(key): value for key, value in row.items()}
    for alias in aliases:
        value = lookup.get(normalize_header(alias))
        if value is not None and value.strip():
            return value.strip()
    return None


def normalize_id(value: str) -> str:
    text = str(value).strip()
    try:
        number = Decimal(text)
    except InvalidOperation:
        return text
    if number == number.to_integral_value():
        return str(int(number))
    return text


def parse_float(value: str | None, *, label: str) -> float | None:
    if value is None or not value.strip():
        return None

    text = value.strip()
    try:
        return float(text)
    except ValueError as exc:
        raise ValueError(f"cannot parse {label} as a number: {value}") from exc


def is_xlsx_file(path: Path) -> bool:
    with path.open("rb") as handle:
        return handle.read(4) == b"PK\x03\x04"


def column_index(cell_ref: str) -> int:
    letters = "".join(ch for ch in cell_ref if ch.isalpha())
    index = 0
    for letter in letters:
        index = index * 26 + ord(letter.upper()) - ord("A") + 1
    return index - 1


def text_from_cell(cell: ET.Element, shared_strings: list[str]) -> str:
    cell_type = cell.attrib.get("t")
    if cell_type == "inlineStr":
        return "".join(
            text.text or ""
            for text in cell.findall(".//main:t", NS)
        ).strip()

    value_elem = cell.find("main:v", NS)
    if value_elem is None or value_elem.text is None:
        return ""

    value = value_elem.text
    if cell_type == "s":
        return shared_strings[int(value)].strip()
    if cell_type == "b":
        return "TRUE" if value == "1" else "FALSE"
    return value.strip()


def read_shared_strings(zf: ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in zf.namelist():
        return []

    root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
    strings: list[str] = []
    for item in root.findall("main:si", NS):
        strings.append(
            "".join(text.text or "" for text in item.findall(".//main:t", NS))
        )
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


def read_xlsx_sheets(path: Path) -> list[tuple[str, list[tuple[int, list[str]]]]]:
    out: list[tuple[str, list[tuple[int, list[str]]]]] = []
    with ZipFile(path) as zf:
        shared_strings = read_shared_strings(zf)
        for sheet_name, sheet_path in workbook_sheets(zf):
            root = ET.fromstring(zf.read(sheet_path))
            rows: list[tuple[int, list[str]]] = []
            for row_elem in root.findall("main:sheetData/main:row", NS):
                row_number = int(row_elem.attrib.get("r", len(rows) + 1))
                cells: list[str] = []
                for cell in row_elem.findall("main:c", NS):
                    idx = column_index(cell.attrib.get("r", "A1"))
                    while len(cells) <= idx:
                        cells.append("")
                    cells[idx] = text_from_cell(cell, shared_strings)
                rows.append((row_number, cells))
            out.append((sheet_name, rows))
    return out


def read_csv_sheet(path: Path) -> list[tuple[str, list[tuple[int, list[str]]]]]:
    with path.open(newline="", encoding="utf-8-sig") as handle:
        sample = handle.read(4096)
        handle.seek(0)
        try:
            dialect = csv.Sniffer().sniff(sample)
        except csv.Error:
            dialect = csv.excel
        rows = [
            (line_number, [cell.strip() for cell in row])
            for line_number, row in enumerate(csv.reader(handle, dialect), start=1)
        ]
    return [("csv", rows)]


def read_data_rows(path: Path) -> list[tuple[str, int, dict[str, str]]]:
    sheets = read_xlsx_sheets(path) if is_xlsx_file(path) else read_csv_sheet(path)
    data_rows: list[tuple[str, int, dict[str, str]]] = []

    for sheet_name, rows in sheets:
        non_empty = [
            (row_number, cells)
            for row_number, cells in rows
            if any(cell.strip() for cell in cells)
        ]
        if not non_empty:
            continue

        _, header = non_empty[0]
        if value_by_alias({name: name for name in header}, ["X"]) is None:
            continue
        if value_by_alias({name: name for name in header}, ["Y"]) is None:
            continue

        for row_number, cells in non_empty[1:]:
            padded = cells + [""] * max(0, len(header) - len(cells))
            row = {
                header[index].strip(): padded[index].strip()
                for index in range(len(header))
                if header[index].strip()
            }
            data_rows.append((sheet_name, row_number, row))

    return data_rows


def find_existing_file(root: Path, names: list[str]) -> Path:
    for name in names:
        direct = root / name
        if direct.is_file():
            return direct

        matches = sorted(path for path in root.rglob(name) if path.is_file())
        if matches:
            return matches[0]

    joined = ", ".join(names)
    raise FileNotFoundError(f"could not find any of: {joined}")


def load_source_rows(source_paths: list[Path]) -> list[tuple[Path, list[DataRow]]]:
    grouped_rows: list[tuple[Path, list[DataRow]]] = []
    for source in source_paths:
        rows: list[DataRow] = []
        for sheet, row_number, values in read_data_rows(source):
            x = value_by_alias(values, ["X"])
            y = value_by_alias(values, ["Y"])
            if x is None or y is None:
                continue
            rows.append(
                DataRow(
                    source=source,
                    sheet=sheet,
                    row_number=row_number,
                    values=values,
                    x=normalize_id(x),
                    y=normalize_id(y),
                )
            )
        grouped_rows.append((source, rows))
    return grouped_rows


def select_rows(
    grouped_rows: list[tuple[Path, list[DataRow]]],
    count: int,
    *,
    randomize: bool,
    seed: int | None,
) -> list[DataRow]:
    rng = random.Random(seed)
    selected: list[DataRow] = []
    remaining: list[DataRow] = []

    for _, rows in grouped_rows:
        if not rows:
            continue
        chosen_index = rng.randrange(len(rows)) if randomize else 0
        selected.append(rows[chosen_index])
        remaining.extend(row for index, row in enumerate(rows) if index != chosen_index)
        if len(selected) == count:
            return selected

    if randomize:
        rng.shuffle(remaining)
    selected.extend(remaining[: max(0, count - len(selected))])

    if len(selected) < count:
        raise ValueError(f"only found {len(selected)} data row(s), need {count}")
    return selected[:count]


def load_orca_config(path: Path) -> dict[tuple[str, str], dict[str, str]]:
    config: dict[tuple[str, str], dict[str, str]] = {}
    for _, _, row in read_data_rows(path):
        x = value_by_alias(row, ["X"])
        y = value_by_alias(row, ["Y"])
        if x is None or y is None:
            continue
        config[(normalize_id(x), normalize_id(y))] = row
    return config


def looks_like_basis(value: str) -> bool:
    text = value.lower().strip()
    prefixes = (
        "def2",
        "cc-",
        "aug-cc",
        "jun-cc",
        "may-cc",
        "6-",
        "3-",
        "sto-",
        "lanl",
        "sdd",
    )
    exact = {"svp", "tzvp", "tzvpp", "qzvp", "qzvpp", "minix"}
    return text.startswith(prefixes) or text in exact or "basis" in text


def looks_like_functional(value: str) -> bool:
    text = value.lower().strip()
    known = {
        "b3lyp",
        "pbe",
        "pbe0",
        "bp86",
        "blyp",
        "tpss",
        "m06",
        "m062x",
        "m06-2x",
        "wb97x",
        "wb97x-d",
        "cam-b3lyp",
        "hf",
    }
    return text in known or text.endswith("lyp") or "functional" in text


def infer_orca_settings(row: dict[str, str]) -> OrcaSettings:
    set_base = value_by_alias(
        row,
        ["set_base", "set_mase", "basis", "basis_set", "base", "base_set"],
    )
    dft_method = value_by_alias(
        row,
        ["dft_method", "ft_method", "functional", "method", "dft"],
    )

    functional: str | None = None
    basis: str | None = None
    for value in [set_base, dft_method]:
        if value is None:
            continue
        if looks_like_basis(value):
            basis = value
        if looks_like_functional(value):
            functional = value

    if basis is None and set_base and not looks_like_functional(set_base):
        basis = set_base
    if functional is None and dft_method and not looks_like_basis(dft_method):
        functional = dft_method

    if functional is None and set_base:
        functional = set_base
    if basis is None and dft_method:
        basis = dft_method

    if functional is None or basis is None:
        raise ValueError(
            "orca_config row must contain method/basis values in set_base/"
            "dft_method or their aliases"
        )

    return OrcaSettings(functional=functional, basis=basis)


def run_make_unit(root: Path, diamine: Path, dianhydride: Path, out_dir: Path) -> Path:
    command = [
        sys.executable,
        str(root / "make_unit.py"),
        "--diamines",
        str(diamine),
        "--dianhydrides",
        str(dianhydride),
        "--out",
        str(out_dir),
    ]
    subprocess.run(command, cwd=root, check=True)

    mol_path = out_dir / f"{diamine.stem}__{dianhydride.stem}.mol"
    if not mol_path.is_file():
        raise FileNotFoundError(f"make_unit.py did not create {mol_path}")
    return mol_path


def run_orca(
    inp_path: Path,
    *,
    orca_bin: str,
    force: bool,
    timeout: int | None,
) -> Path:
    out_path = inp_path.with_suffix(".out")
    err_path = inp_path.with_suffix(".orca.err")
    if out_path.is_file() and not force:
        return out_path

    with out_path.open("w") as stdout, err_path.open("w") as stderr:
        subprocess.run(
            [orca_bin, inp_path.name],
            cwd=inp_path.parent,
            stdout=stdout,
            stderr=stderr,
            text=True,
            check=True,
            timeout=timeout,
        )
    return out_path


def existing_orca_output(inp_path: Path) -> Path | None:
    out_path = inp_path.with_suffix(".out")
    return out_path if out_path.is_file() else None


def convert_orca_result_to_holder(
    out_path: Path,
    *,
    orca_2mkl_bin: str,
) -> Path:
    base = out_path.with_suffix("")
    gbw_path = base.with_suffix(".gbw")
    if not out_path.is_file():
        raise FileNotFoundError(out_path)
    if not gbw_path.is_file():
        raise FileNotFoundError(
            f"{gbw_path} is required for orca_2mkl Molden conversion. "
            "Run ORCA first and keep the .gbw file next to the .out file."
        )

    subprocess.run(
        [orca_2mkl_bin, base.name, "-molden"],
        cwd=out_path.parent,
        check=True,
    )

    molden_path = base.with_suffix(".molden.input")
    if not molden_path.is_file():
        raise FileNotFoundError(f"orca_2mkl did not create {molden_path}")

    # Multiwfn detects Molden files by extension, so keep the Molden suffix.
    holder_path = base.with_name(f"{base.name}.holder.molden.input")
    shutil.copyfile(molden_path, holder_path)
    return holder_path


def parse_multiwfn_gap(output: str) -> float:
    gap_match = re.search(
        r"HOMO-LUMO\s+gap:\s*[-+]?\d*\.?\d+(?:[Ee][-+]?\d+)?\s*a\.u\.\s*"
        r"([-+]?\d*\.?\d+(?:[Ee][-+]?\d+)?)\s*eV",
        output,
        flags=re.IGNORECASE,
    )
    if gap_match:
        return float(gap_match.group(1))

    homo_match = re.search(
        r"HOMO,\s*energy:\s*[-+]?\d*\.?\d+(?:[Ee][-+]?\d+)?\s*a\.u\.\s*"
        r"([-+]?\d*\.?\d+(?:[Ee][-+]?\d+)?)\s*eV",
        output,
        flags=re.IGNORECASE,
    )
    lumo_match = re.search(
        r"LUMO,\s*energy:\s*[-+]?\d*\.?\d+(?:[Ee][-+]?\d+)?\s*a\.u\.\s*"
        r"([-+]?\d*\.?\d+(?:[Ee][-+]?\d+)?)\s*eV",
        output,
        flags=re.IGNORECASE,
    )
    if homo_match and lumo_match:
        return float(lumo_match.group(1)) - float(homo_match.group(1))

    raise ValueError("could not find HOMO-LUMO gap in Multiwfn output")


def read_gap_with_multiwfn(
    holder_path: Path,
    *,
    multiwfn_bin: str,
    log_path: Path,
    timeout: int,
) -> float:
    result = subprocess.run(
        [multiwfn_bin, str(holder_path)],
        input="0\n",
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        timeout=timeout,
    )
    log_path.write_text(result.stdout)
    return parse_multiwfn_gap(result.stdout)


def percent_difference(
    calculated: float | None,
    reference: float | None,
) -> tuple[float | None, float | None]:
    if calculated is None or reference is None:
        return None, None
    if reference == 0:
        raise ZeroDivisionError("csv Eg is zero, cannot calculate percentage difference")

    signed = (calculated - reference) / reference * 100.0
    return abs(signed), signed


def generate_for_row(
    row: DataRow,
    *,
    root: Path,
    output_root: Path,
    config: dict[tuple[str, str], dict[str, str]],
    run_orca_calculation: bool,
    analyze_eg: bool,
    orca_bin: str,
    orca_2mkl_bin: str,
    multiwfn_bin: str,
    force_orca: bool,
    orca_timeout: int | None,
    multiwfn_timeout: int,
    nprocs: int,
) -> GeneratedResult:
    config_row = config.get((row.x, row.y))
    if config_row is None:
        raise KeyError(f"no orca_config.csv row for X={row.x}, Y={row.y}")
    settings = infer_orca_settings(config_row)

    diamine = root / "diamines" / f"A_{row.x}.mol"
    dianhydride = root / "dianhydrides" / f"H_{row.y}.mol"
    if not diamine.is_file():
        raise FileNotFoundError(diamine)
    if not dianhydride.is_file():
        raise FileNotFoundError(dianhydride)

    job_name = f"{row.source.stem}_row{row.row_number}_A{row.x}_H{row.y}"
    job_dir = output_root / job_name
    job_dir.mkdir(parents=True, exist_ok=True)

    mol_path = run_make_unit(root, diamine, dianhydride, job_dir)
    smiles = mol_file_to_smiles(mol_path)

    smiles_path = job_dir / f"{job_name}.smi"
    smiles_path.write_text(f"file\tsmiles\n{mol_path}\t{smiles}\n")

    inp_path = job_dir / f"{job_name}.inp"
    orca_generator(
        smiles,
        inp_path,
        functional=settings.functional,
        basis=settings.basis,
        nprocs=nprocs,
    )

    csv_eg_ev = parse_float(value_by_alias(row.values, ["Eg", "E_g", "gap"]), label="Eg")
    out_path: Path | None = None
    holder_path: Path | None = None
    multiwfn_log_path: Path | None = None
    multiwfn_eg_ev: float | None = None
    abs_percent: float | None = None
    signed_percent: float | None = None

    if run_orca_calculation:
        out_path = run_orca(
            inp_path,
            orca_bin=orca_bin,
            force=force_orca,
            timeout=orca_timeout,
        )
    elif analyze_eg:
        out_path = existing_orca_output(inp_path)
        if out_path is None:
            raise FileNotFoundError(
                f"no ORCA output found for {inp_path}. "
                "Run with --run-orca or place the .out/.gbw files next to the .inp file."
            )

    if analyze_eg or run_orca_calculation:
        if out_path is None:
            raise RuntimeError("internal error: ORCA output path was not set")

        holder_path = convert_orca_result_to_holder(
            out_path,
            orca_2mkl_bin=orca_2mkl_bin,
        )
        multiwfn_log_path = job_dir / f"{job_name}.multiwfn.log"
        multiwfn_eg_ev = read_gap_with_multiwfn(
            holder_path,
            multiwfn_bin=multiwfn_bin,
            log_path=multiwfn_log_path,
            timeout=multiwfn_timeout,
        )
        abs_percent, signed_percent = percent_difference(multiwfn_eg_ev, csv_eg_ev)

    return GeneratedResult(
        row=row,
        settings=settings,
        mol_path=mol_path,
        smiles_path=smiles_path,
        inp_path=inp_path,
        smiles=smiles,
        csv_eg_ev=csv_eg_ev,
        out_path=out_path,
        holder_path=holder_path,
        multiwfn_log_path=multiwfn_log_path,
        multiwfn_eg_ev=multiwfn_eg_ev,
        percent_difference=abs_percent,
        signed_percent_difference=signed_percent,
    )


def write_manifest(path: Path, results: list[GeneratedResult]) -> None:
    with path.open("w", newline="") as handle:
        writer = csv.writer(handle)
        writer.writerow(
            [
                "source",
                "sheet",
                "row_number",
                "X",
                "Y",
                "functional",
                "basis",
                "csv_eg_ev",
                "multiwfn_eg_ev",
                "percent_difference_abs",
                "percent_difference_signed",
                "mol",
                "smiles_file",
                "inp",
                "orca_out",
                "holder",
                "multiwfn_log",
            ]
        )
        for result in results:
            writer.writerow(
                [
                    result.row.source,
                    result.row.sheet,
                    result.row.row_number,
                    result.row.x,
                    result.row.y,
                    result.settings.functional,
                    result.settings.basis,
                    result.csv_eg_ev,
                    result.multiwfn_eg_ev,
                    result.percent_difference,
                    result.signed_percent_difference,
                    result.mol_path,
                    result.smiles_path,
                    result.inp_path,
                    result.out_path,
                    result.holder_path,
                    result.multiwfn_log_path,
                ]
            )


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--root",
        type=Path,
        default=Path(__file__).resolve().parent,
        help="Directory to search and use as the project root. Default: script directory.",
    )
    parser.add_argument(
        "--count",
        type=int,
        default=2,
        help="Number of data rows to process. Default: 2.",
    )
    parser.add_argument(
        "--random",
        action="store_true",
        help="Randomly choose rows instead of taking the first data row from each table.",
    )
    parser.add_argument(
        "--seed",
        type=int,
        help="Optional random seed used with --random.",
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=Path("generated_inps"),
        help="Output directory. Relative paths are resolved under --root.",
    )
    parser.add_argument(
        "--run-orca",
        action="store_true",
        help="Run ORCA for each generated .inp and then calculate the HOMO-LUMO gap.",
    )
    parser.add_argument(
        "--analyze-eg",
        action="store_true",
        help=(
            "Use existing .out/.gbw files next to each .inp to calculate the "
            "HOMO-LUMO gap without running ORCA. --run-orca implies this."
        ),
    )
    parser.add_argument(
        "--force-orca",
        action="store_true",
        help="Re-run ORCA even if the matching .out file already exists.",
    )
    parser.add_argument(
        "--orca-bin",
        default="orca",
        help="ORCA executable. Default: orca.",
    )
    parser.add_argument(
        "--orca-2mkl-bin",
        default="orca_2mkl",
        help="ORCA Molden converter executable. Default: orca_2mkl.",
    )
    parser.add_argument(
        "--multiwfn-bin",
        default="Multiwfn",
        help="Multiwfn executable. Default: Multiwfn.",
    )
    parser.add_argument(
        "--orca-timeout",
        type=int,
        help="Optional timeout in seconds for each ORCA run.",
    )
    parser.add_argument(
        "--multiwfn-timeout",
        type=int,
        default=300,
        help="Timeout in seconds for each Multiwfn gap read. Default: 300.",
    )
    parser.add_argument(
        "--nprocs",
        type=int,
        default=16,
        help="Number of CPU cores written to each ORCA .inp file. Default: 16.",
    )
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    root = args.root.resolve()
    output_root = args.output_dir
    if not output_root.is_absolute():
        output_root = root / output_root
    output_root.mkdir(parents=True, exist_ok=True)

    conjugate_table = find_existing_file(root, ["concatenate.csv", "conjugate.csv"])
    min_eg_table = find_existing_file(root, ["minEg.csv"])
    config_table = find_existing_file(root, ["orca_config.csv"])

    grouped_rows = load_source_rows([conjugate_table, min_eg_table])
    selected_rows = select_rows(
        grouped_rows,
        args.count,
        randomize=args.random,
        seed=args.seed,
    )
    config = load_orca_config(config_table)
    analyze_eg = args.analyze_eg or args.run_orca

    results = [
        generate_for_row(
            row,
            root=root,
            output_root=output_root,
            config=config,
            run_orca_calculation=args.run_orca,
            analyze_eg=analyze_eg,
            orca_bin=args.orca_bin,
            orca_2mkl_bin=args.orca_2mkl_bin,
            multiwfn_bin=args.multiwfn_bin,
            force_orca=args.force_orca,
            orca_timeout=args.orca_timeout,
            multiwfn_timeout=args.multiwfn_timeout,
            nprocs=args.nprocs,
        )
        for row in selected_rows
    ]
    write_manifest(output_root / "manifest.csv", results)

    print(f"source tables: {conjugate_table}, {min_eg_table}")
    for result in results:
        print(
            "generated "
            f"X={result.row.x} Y={result.row.y} "
            f"with {result.settings.functional}/{result.settings.basis}: "
            f"{result.inp_path}"
        )
        if result.multiwfn_eg_ev is not None:
            csv_eg_text = (
                f"{result.csv_eg_ev:.6g} eV"
                if result.csv_eg_ev is not None
                else "missing"
            )
            percent_text = (
                f"{result.percent_difference:.3f}% "
                f"(signed={result.signed_percent_difference:.3f}%)"
                if result.percent_difference is not None
                and result.signed_percent_difference is not None
                else "not calculated"
            )
            print(
                "  Eg comparison: "
                f"csv={csv_eg_text}, "
                f"Multiwfn={result.multiwfn_eg_ev:.6g} eV, "
                f"difference={percent_text}"
            )
    print(f"manifest: {output_root / 'manifest.csv'}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
