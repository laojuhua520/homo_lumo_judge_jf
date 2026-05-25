#!/usr/bin/env python3
"""Generate an ORCA .inp file from a SMILES string."""

from __future__ import annotations

import argparse
from pathlib import Path

from rdkit import Chem
from rdkit.Chem import AllChem, Descriptors


def orca_generator(
    smiles: str,
    filename: str | Path,
    *,
    functional: str = "B3LYP",
    basis: str = "def2-SVP",
    nprocs: int = 16,
    random_seed: int = 0xF00D,
) -> None:
    if nprocs < 1:
        raise ValueError("nprocs must be at least 1")

    mol = Chem.MolFromSmiles(smiles)
    if mol is None:
        raise ValueError(f"invalid SMILES: {smiles}")
    mol = Chem.AddHs(mol)

    embed_status = AllChem.EmbedMolecule(mol, randomSeed=random_seed)
    if embed_status != 0:
        embed_status = AllChem.EmbedMolecule(
            mol,
            randomSeed=random_seed,
            useRandomCoords=True,
        )
    if embed_status != 0:
        raise ValueError("RDKit could not generate 3D coordinates for the molecule")

    AllChem.MMFFOptimizeMolecule(mol)

    xyz_block = Chem.MolToXYZBlock(mol)
    lines = xyz_block.splitlines()
    xyz_coords = "\n".join(lines[2:])

    charge = Chem.GetFormalCharge(mol)
    radicals = Descriptors.NumRadicalElectrons(mol)
    spin_mult = radicals + 1

    inp_content = f"""! {functional} {basis} RIJCOSX Opt
%pal
    nprocs {nprocs}
end
%scf
    MaxIter 500
end
%geom
    MaxIter 500
end
%elprop
    Dipole true
    Polar 1
end
* xyz {charge} {spin_mult}
{xyz_coords}
*
"""

    output_path = Path(filename)
    output_path.write_text(inp_content)
    print(f"wrote {output_path}")


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("smiles", help="Input SMILES string.")
    parser.add_argument("output", type=Path, help="Output .inp file.")
    parser.add_argument(
        "--functional",
        default="B3LYP",
        help="DFT functional/method keyword. Default: B3LYP.",
    )
    parser.add_argument(
        "--basis",
        default="def2-SVP",
        help="Basis-set keyword. Default: def2-SVP.",
    )
    parser.add_argument(
        "--nprocs",
        type=int,
        default=16,
        help="Number of CPU cores for ORCA. Default: 16.",
    )
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    orca_generator(
        args.smiles,
        args.output,
        functional=args.functional,
        basis=args.basis,
        nprocs=args.nprocs,
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
