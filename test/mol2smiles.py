#!/usr/bin/env python3
"""Convert mol files to SMILES with RDKit."""

from __future__ import annotations

import argparse
import csv
import sys
from pathlib import Path

try:
    from rdkit import Chem
except ModuleNotFoundError as exc:  # pragma: no cover - depends on local env
    Chem = None
    _RDKIT_IMPORT_ERROR = exc
else:
    _RDKIT_IMPORT_ERROR = None


def _require_rdkit() -> None:
    if Chem is None:
        raise RuntimeError(
            "RDKit is required for mol-to-SMILES conversion, but it is not "
            "installed in this Python environment."
        ) from _RDKIT_IMPORT_ERROR


def mol_file_to_smiles(mol_path: str | Path, *, canonical: bool = True) -> str:
    """Read one .mol file and return its SMILES string."""
    _require_rdkit()

    path = Path(mol_path)
    mol = Chem.MolFromMolFile(str(path))
    if mol is None:
        raise ValueError(f"RDKit could not read molecule from {path}")

    return Chem.MolToSmiles(mol, canonical=canonical)


def write_smiles_table(
    mol_files: list[Path],
    output_path: Path,
    *,
    canonical: bool = True,
    include_header: bool = True,
) -> None:
    rows = [(str(path), mol_file_to_smiles(path, canonical=canonical)) for path in mol_files]

    with output_path.open("w", newline="") as handle:
        writer = csv.writer(handle, delimiter="\t", lineterminator="\n")
        if include_header:
            writer.writerow(["file", "smiles"])
        writer.writerows(rows)


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("mol_files", nargs="+", type=Path, help="Input .mol file(s).")
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        help="Optional tab-separated output txt file. Defaults to printing to stdout.",
    )
    parser.add_argument(
        "--no-canonical",
        action="store_true",
        help="Keep RDKit's non-canonical SMILES ordering.",
    )
    parser.add_argument(
        "--no-header",
        action="store_true",
        help="Do not write the file/smiles header row when using --output.",
    )
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    canonical = not args.no_canonical

    if args.output:
        write_smiles_table(
            args.mol_files,
            args.output,
            canonical=canonical,
            include_header=not args.no_header,
        )
        print(f"Wrote {len(args.mol_files)} SMILES entries to {args.output}")
        return 0

    writer = csv.writer(sys.stdout, delimiter="\t", lineterminator="\n")
    writer.writerow(["file", "smiles"])
    for mol_file in args.mol_files:
        writer.writerow([str(mol_file), mol_file_to_smiles(mol_file, canonical=canonical)])
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
