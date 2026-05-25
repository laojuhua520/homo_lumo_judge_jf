#!/usr/bin/env python3
"""
Build imide-terminated model structures from diamine and dianhydride .mol files.

Default layout:
    polyimide/
      diamines/*.mol
      dianhydrides/*.mol
      products/*.mol

The transformation is:
  - choose one terminal amine N from the diamine;
  - choose one cyclic anhydride group from the dianhydride;
  - remove the anhydride bridge O and bond the amine N to the two carbonyl C atoms,
    forming an imide ring;
  - remove all other terminal amine groups from the diamine and replace those
    removed substituents with H caps by default;
  - convert every other anhydride group in the dianhydride into an imide ring
    terminated by an explicit N-H bond.

The script intentionally avoids RDKit so it can run in minimal Python environments.
It preserves the input bond orders and writes V2000 mol output.
"""

from __future__ import annotations

import argparse
import math
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable


@dataclass
class Atom:
    symbol: str
    x: float
    y: float
    z: float


@dataclass(frozen=True)
class Bond:
    a: int
    b: int
    order: int


@dataclass
class Molecule:
    name: str
    atoms: list[Atom]
    bonds: list[Bond]

    def neighbors(self, atom_idx: int) -> list[tuple[int, Bond]]:
        out: list[tuple[int, Bond]] = []
        for bond in self.bonds:
            if bond.a == atom_idx:
                out.append((bond.b, bond))
            elif bond.b == atom_idx:
                out.append((bond.a, bond))
        return out

    def bond_between(self, a: int, b: int) -> Bond | None:
        for bond in self.bonds:
            if {bond.a, bond.b} == {a, b}:
                return bond
        return None


@dataclass(frozen=True)
class AmineSite:
    n: int
    anchor: int


@dataclass(frozen=True)
class AnhydrideGroup:
    bridge_o: int
    c1: int
    c2: int
    ox1: int
    ox2: int
    anchor1: int
    anchor2: int


@dataclass(frozen=True)
class CapRequest:
    source: str
    anchor: int
    removed: int


Vec = tuple[float, float, float]


def read_mol(path: Path) -> Molecule:
    lines = path.read_text().splitlines()
    if len(lines) < 4:
        raise ValueError(f"{path} is too short to be a mol file")

    if any("V3000" in line for line in lines[:10]):
        mol = read_v3000(path, lines)
    else:
        mol = read_v2000(path, lines)

    if not mol.name.strip():
        mol.name = path.stem
    return mol


def read_v2000(path: Path, lines: list[str]) -> Molecule:
    try:
        counts = lines[3]
        atom_count = int(counts[0:3])
        bond_count = int(counts[3:6])
    except Exception:
        parts = lines[3].split()
        if len(parts) < 2:
            raise ValueError(f"cannot parse V2000 counts line in {path}")
        atom_count = int(parts[0])
        bond_count = int(parts[1])

    atoms: list[Atom] = []
    atom_start = 4
    for i in range(atom_count):
        line = lines[atom_start + i]
        parts = line.split()
        if len(parts) < 4:
            raise ValueError(f"cannot parse atom line {i + 1} in {path}")
        atoms.append(Atom(parts[3], float(parts[0]), float(parts[1]), float(parts[2])))

    bonds: list[Bond] = []
    bond_start = atom_start + atom_count
    for i in range(bond_count):
        line = lines[bond_start + i]
        parts = line.split()
        if len(parts) < 3:
            raise ValueError(f"cannot parse bond line {i + 1} in {path}")
        a = int(parts[0]) - 1
        b = int(parts[1]) - 1
        order = int(float(parts[2]))
        bonds.append(Bond(a, b, order))

    return Molecule(lines[0].strip(), atoms, bonds)


def read_v3000(path: Path, lines: list[str]) -> Molecule:
    atoms: list[Atom] = []
    bonds: list[Bond] = []
    id_to_idx: dict[int, int] = {}
    in_atom_block = False
    in_bond_block = False

    for raw in lines:
        line = raw.strip()
        if line == "M  V30 BEGIN ATOM":
            in_atom_block = True
            continue
        if line == "M  V30 END ATOM":
            in_atom_block = False
            continue
        if line == "M  V30 BEGIN BOND":
            in_bond_block = True
            continue
        if line == "M  V30 END BOND":
            in_bond_block = False
            continue

        if not line.startswith("M  V30 "):
            continue

        parts = line.split()
        if in_atom_block:
            if len(parts) < 7:
                raise ValueError(f"cannot parse V3000 atom line in {path}: {raw}")
            atom_id = int(parts[2])
            symbol = parts[3]
            atom = Atom(symbol, float(parts[4]), float(parts[5]), float(parts[6]))
            id_to_idx[atom_id] = len(atoms)
            atoms.append(atom)
        elif in_bond_block:
            if len(parts) < 6:
                raise ValueError(f"cannot parse V3000 bond line in {path}: {raw}")
            order = int(float(parts[3]))
            a = id_to_idx[int(parts[4])]
            b = id_to_idx[int(parts[5])]
            bonds.append(Bond(a, b, order))

    if not atoms:
        raise ValueError(f"no atoms found in V3000 mol file {path}")

    return Molecule(lines[0].strip(), atoms, bonds)


def write_v2000(path: Path, mol: Molecule) -> None:
    if len(mol.atoms) > 999 or len(mol.bonds) > 999:
        raise ValueError("V2000 output supports at most 999 atoms and 999 bonds")

    lines = [
        mol.name[:80],
        "  polyimide_make_unit",
        "",
        f"{len(mol.atoms):3d}{len(mol.bonds):3d}  0  0  0  0  0  0  0  0999 V2000",
    ]

    for atom in mol.atoms:
        lines.append(
            f"{atom.x:10.4f}{atom.y:10.4f}{atom.z:10.4f} "
            f"{atom.symbol:<3s} 0  0  0  0  0  0  0  0  0  0  0  0"
        )

    for bond in mol.bonds:
        lines.append(f"{bond.a + 1:3d}{bond.b + 1:3d}{bond.order:3d}  0        0")

    lines.append("M  END")
    path.write_text("\n".join(lines) + "\n")


def find_terminal_amines(mol: Molecule) -> list[AmineSite]:
    sites: list[AmineSite] = []
    for idx, atom in enumerate(mol.atoms):
        if atom.symbol != "N":
            continue
        neighbors = mol.neighbors(idx)
        heavy = [(n, b) for n, b in neighbors if mol.atoms[n].symbol != "H"]
        if len(heavy) != 1:
            continue
        anchor, bond = heavy[0]
        if bond.order != 1:
            continue
        if any(b.order != 1 for _, b in neighbors):
            continue
        sites.append(AmineSite(idx, anchor))
    return sites


def find_anhydrides(mol: Molecule) -> list[AnhydrideGroup]:
    groups: list[AnhydrideGroup] = []
    for bridge_idx, bridge in enumerate(mol.atoms):
        if bridge.symbol != "O":
            continue

        bridge_neighbors = mol.neighbors(bridge_idx)
        heavy_neighbors = [
            (n, b)
            for n, b in bridge_neighbors
            if mol.atoms[n].symbol != "H"
        ]
        if len(heavy_neighbors) != 2:
            continue
        if any(b.order != 1 for _, b in heavy_neighbors):
            continue
        if any(mol.atoms[n].symbol != "C" for n, _ in heavy_neighbors):
            continue

        c1, c2 = heavy_neighbors[0][0], heavy_neighbors[1][0]
        info1 = carbonyl_info(mol, c1, bridge_idx)
        info2 = carbonyl_info(mol, c2, bridge_idx)
        if info1 is None or info2 is None:
            continue

        ox1, anchor1 = info1
        ox2, anchor2 = info2
        groups.append(AnhydrideGroup(bridge_idx, c1, c2, ox1, ox2, anchor1, anchor2))
    return groups


def carbonyl_info(mol: Molecule, carbon_idx: int, bridge_idx: int) -> tuple[int, int] | None:
    if mol.atoms[carbon_idx].symbol != "C":
        return None

    terminal_oxygens: list[int] = []
    anchors: list[int] = []
    for neighbor_idx, bond in mol.neighbors(carbon_idx):
        symbol = mol.atoms[neighbor_idx].symbol
        if neighbor_idx == bridge_idx:
            if bond.order != 1:
                return None
        elif symbol == "O" and bond.order == 2:
            terminal_oxygens.append(neighbor_idx)
        elif symbol != "H":
            anchors.append(neighbor_idx)

    if len(terminal_oxygens) != 1 or len(anchors) != 1:
        return None
    return terminal_oxygens[0], anchors[0]


def build_product(
    diamine: Molecule,
    dianhydride: Molecule,
    amines: list[AmineSite],
    anhydrides: list[AnhydrideGroup],
    selected_amine: AmineSite,
    selected_anhydride: AnhydrideGroup,
    name: str,
    explicit_caps: bool,
) -> Molecule:
    diamine_remove: set[int] = set()
    dianhydride_remove: set[int] = {group.bridge_o for group in anhydrides}
    diamine_caps: list[CapRequest] = []
    terminal_imides = [
        group
        for group in anhydrides
        if group.bridge_o != selected_anhydride.bridge_o
    ]

    for amine in amines:
        attached_h = [
            n
            for n, _ in diamine.neighbors(amine.n)
            if diamine.atoms[n].symbol == "H"
        ]
        diamine_remove.update(attached_h)
        if amine.n != selected_amine.n:
            diamine_remove.add(amine.n)
            diamine_caps.append(CapRequest("diamine", amine.anchor, amine.n))

    transform = make_diamine_transform(diamine, dianhydride, selected_amine, selected_anhydride)

    atoms: list[Atom] = []
    bonds: list[Bond] = []
    dianhydride_map: dict[int, int] = {}
    diamine_map: dict[int, int] = {}

    for old_idx, atom in enumerate(dianhydride.atoms):
        if old_idx in dianhydride_remove:
            continue
        dianhydride_map[old_idx] = len(atoms)
        atoms.append(Atom(atom.symbol, atom.x, atom.y, atom.z))

    for old_idx, atom in enumerate(diamine.atoms):
        if old_idx in diamine_remove:
            continue
        x, y, z = transform((atom.x, atom.y, atom.z))
        diamine_map[old_idx] = len(atoms)
        atoms.append(Atom(atom.symbol, x, y, z))

    append_kept_bonds(bonds, dianhydride.bonds, dianhydride_map)
    append_kept_bonds(bonds, diamine.bonds, diamine_map)

    n_new = diamine_map[selected_amine.n]
    c1_new = dianhydride_map[selected_anhydride.c1]
    c2_new = dianhydride_map[selected_anhydride.c2]
    bonds.append(Bond(n_new, c1_new, 1))
    bonds.append(Bond(n_new, c2_new, 1))

    add_terminal_imides(atoms, bonds, dianhydride, dianhydride_map, terminal_imides)

    if explicit_caps:
        add_caps(atoms, bonds, diamine, diamine_map, diamine_caps, transform)

    return Molecule(name, atoms, bonds)


def append_kept_bonds(bonds: list[Bond], source_bonds: Iterable[Bond], atom_map: dict[int, int]) -> None:
    for bond in source_bonds:
        if bond.a in atom_map and bond.b in atom_map:
            bonds.append(Bond(atom_map[bond.a], atom_map[bond.b], bond.order))


def add_terminal_imides(
    atoms: list[Atom],
    bonds: list[Bond],
    source_mol: Molecule,
    atom_map: dict[int, int],
    groups: Iterable[AnhydrideGroup],
) -> None:
    for group in groups:
        if group.c1 not in atom_map or group.c2 not in atom_map:
            continue

        n_source = source_mol.atoms[group.bridge_o]
        n_idx = len(atoms)
        atoms.append(Atom("N", n_source.x, n_source.y, n_source.z))
        bonds.append(Bond(n_idx, atom_map[group.c1], 1))
        bonds.append(Bond(n_idx, atom_map[group.c2], 1))

        h_x, h_y, h_z = terminal_imide_h_position(source_mol, group)
        h_idx = len(atoms)
        atoms.append(Atom("H", h_x, h_y, h_z))
        bonds.append(Bond(n_idx, h_idx, 1))


def terminal_imide_h_position(mol: Molecule, group: AnhydrideGroup) -> Vec:
    n_pos = atom_vec(mol.atoms[group.bridge_o])
    c_mid = v_scale(v_add(atom_vec(mol.atoms[group.c1]), atom_vec(mol.atoms[group.c2])), 0.5)
    direction = v_unit(v_sub(n_pos, c_mid))
    if direction is None:
        ox_mid = v_scale(v_add(atom_vec(mol.atoms[group.ox1]), atom_vec(mol.atoms[group.ox2])), 0.5)
        direction = v_unit(v_sub(n_pos, ox_mid))
    if direction is None:
        direction = (1.0, 0.0, 0.0)
    return v_add(n_pos, v_scale(direction, 1.01))


def make_diamine_transform(
    diamine: Molecule,
    dianhydride: Molecule,
    amine: AmineSite,
    group: AnhydrideGroup,
):
    n0 = atom_vec(diamine.atoms[amine.n])
    anchor0 = atom_vec(diamine.atoms[amine.anchor])
    source = v_sub(anchor0, n0)

    bridge = atom_vec(dianhydride.atoms[group.bridge_o])
    c_mid = v_scale(v_add(atom_vec(dianhydride.atoms[group.c1]), atom_vec(dianhydride.atoms[group.c2])), 0.5)
    target = v_sub(bridge, c_mid)

    rotation = rotation_from_to(source, target)

    def transform(point: Vec) -> Vec:
        shifted = v_sub(point, n0)
        rotated = mat_vec(rotation, shifted)
        return v_add(bridge, rotated)

    return transform


def add_caps(
    atoms: list[Atom],
    bonds: list[Bond],
    source_mol: Molecule,
    atom_map: dict[int, int],
    caps: Iterable[CapRequest],
    transform,
) -> None:
    for cap in caps:
        if cap.anchor not in atom_map:
            continue

        anchor_atom = source_mol.atoms[cap.anchor]
        removed_atom = source_mol.atoms[cap.removed]
        anchor = transform((anchor_atom.x, anchor_atom.y, anchor_atom.z))
        removed = transform((removed_atom.x, removed_atom.y, removed_atom.z))
        direction = v_sub(removed, anchor)
        length = v_norm(direction)
        if length < 1.0e-8:
            direction = (1.0, 0.0, 0.0)
        else:
            direction = v_scale(direction, 1.0 / length)

        h_pos = v_add(anchor, v_scale(direction, 1.09))
        h_idx = len(atoms)
        atoms.append(Atom("H", h_pos[0], h_pos[1], h_pos[2]))
        bonds.append(Bond(atom_map[cap.anchor], h_idx, 1))


def atom_vec(atom: Atom) -> Vec:
    return (atom.x, atom.y, atom.z)


def v_add(a: Vec, b: Vec) -> Vec:
    return (a[0] + b[0], a[1] + b[1], a[2] + b[2])


def v_sub(a: Vec, b: Vec) -> Vec:
    return (a[0] - b[0], a[1] - b[1], a[2] - b[2])


def v_scale(v: Vec, scale: float) -> Vec:
    return (v[0] * scale, v[1] * scale, v[2] * scale)


def v_dot(a: Vec, b: Vec) -> float:
    return a[0] * b[0] + a[1] * b[1] + a[2] * b[2]


def v_cross(a: Vec, b: Vec) -> Vec:
    return (
        a[1] * b[2] - a[2] * b[1],
        a[2] * b[0] - a[0] * b[2],
        a[0] * b[1] - a[1] * b[0],
    )


def v_norm(v: Vec) -> float:
    return math.sqrt(v_dot(v, v))


def v_unit(v: Vec) -> Vec | None:
    length = v_norm(v)
    if length < 1.0e-8:
        return None
    return v_scale(v, 1.0 / length)


def identity_matrix() -> tuple[Vec, Vec, Vec]:
    return ((1.0, 0.0, 0.0), (0.0, 1.0, 0.0), (0.0, 0.0, 1.0))


def rotation_from_to(source: Vec, target: Vec) -> tuple[Vec, Vec, Vec]:
    src = v_unit(source)
    dst = v_unit(target)
    if src is None or dst is None:
        return identity_matrix()

    cross = v_cross(src, dst)
    sin_theta = v_norm(cross)
    cos_theta = max(-1.0, min(1.0, v_dot(src, dst)))

    if sin_theta < 1.0e-8:
        if cos_theta > 0:
            return identity_matrix()
        axis = perpendicular_unit(src)
        return rotation_axis_angle(axis, math.pi)

    axis = v_scale(cross, 1.0 / sin_theta)
    return rotation_axis_angle(axis, math.atan2(sin_theta, cos_theta))


def perpendicular_unit(v: Vec) -> Vec:
    trial = (1.0, 0.0, 0.0)
    if abs(v_dot(v, trial)) > 0.9:
        trial = (0.0, 1.0, 0.0)
    out = v_cross(v, trial)
    unit = v_unit(out)
    if unit is None:
        return (0.0, 0.0, 1.0)
    return unit


def rotation_axis_angle(axis: Vec, angle: float) -> tuple[Vec, Vec, Vec]:
    x, y, z = axis
    c = math.cos(angle)
    s = math.sin(angle)
    t = 1.0 - c
    return (
        (t * x * x + c, t * x * y - s * z, t * x * z + s * y),
        (t * x * y + s * z, t * y * y + c, t * y * z - s * x),
        (t * x * z - s * y, t * y * z + s * x, t * z * z + c),
    )


def mat_vec(matrix: tuple[Vec, Vec, Vec], vector: Vec) -> Vec:
    return (
        v_dot(matrix[0], vector),
        v_dot(matrix[1], vector),
        v_dot(matrix[2], vector),
    )


def expand_mol_paths(paths: Iterable[Path]) -> list[Path]:
    out: list[Path] = []
    for path in paths:
        if path.is_dir():
            out.extend(sorted(path.glob("*.mol")))
        elif path.is_file():
            out.append(path)
        else:
            raise FileNotFoundError(path)
    return sorted(dict.fromkeys(out))


def site_label_amine(site: AmineSite) -> str:
    return f"N{site.n + 1}"


def site_label_anhydride(group: AnhydrideGroup) -> str:
    return f"O{group.bridge_o + 1}"


def report_sites(diamine_paths: list[Path], dianhydride_paths: list[Path]) -> int:
    status = 0
    for path in diamine_paths:
        mol = read_mol(path)
        sites = find_terminal_amines(mol)
        if not sites:
            status = 1
        labels = ", ".join(
            f"{site_label_amine(site)} anchored to atom {site.anchor + 1}"
            for site in sites
        )
        print(f"diamine {path.name}: {labels or 'no terminal amine found'}")

    for path in dianhydride_paths:
        mol = read_mol(path)
        groups = find_anhydrides(mol)
        if not groups:
            status = 1
        labels = ", ".join(
            f"{site_label_anhydride(group)} carbonyls {group.c1 + 1}/{group.c2 + 1}"
            for group in groups
        )
        print(f"dianhydride {path.name}: {labels or 'no cyclic anhydride found'}")
    return status


def run(args: argparse.Namespace) -> int:
    diamine_paths = expand_mol_paths(args.diamines)
    dianhydride_paths = expand_mol_paths(args.dianhydrides)

    if args.list_sites:
        return report_sites(diamine_paths, dianhydride_paths)

    args.out.mkdir(parents=True, exist_ok=True)

    written = 0
    skipped = 0
    for diamine_path in diamine_paths:
        diamine = read_mol(diamine_path)
        amines = find_terminal_amines(diamine)
        if not amines:
            print(f"skip {diamine_path.name}: no terminal amine found", file=sys.stderr)
            skipped += 1
            continue

        for dianhydride_path in dianhydride_paths:
            dianhydride = read_mol(dianhydride_path)
            anhydrides = find_anhydrides(dianhydride)
            if not anhydrides:
                print(f"skip {dianhydride_path.name}: no cyclic anhydride found", file=sys.stderr)
                skipped += 1
                continue

            amine_choices = amines if args.all_sites else amines[:1]
            anhydride_choices = anhydrides if args.all_sites else anhydrides[:1]

            for amine in amine_choices:
                for anhydride in anhydride_choices:
                    base = f"{diamine_path.stem}__{dianhydride_path.stem}"
                    if args.all_sites:
                        base += f"__{site_label_amine(amine)}_{site_label_anhydride(anhydride)}"
                    product = build_product(
                        diamine,
                        dianhydride,
                        amines,
                        anhydrides,
                        amine,
                        anhydride,
                        base,
                        explicit_caps=not args.no_explicit_caps,
                    )
                    output_path = args.out / f"{base}.mol"
                    write_v2000(output_path, product)
                    written += 1
                    print(f"wrote {output_path}")

    print(f"done: wrote {written} product(s), skipped {skipped} input combination(s)")
    return 0 if written else 1


def build_parser() -> argparse.ArgumentParser:
    script_dir = Path(__file__).resolve().parent
    parser = argparse.ArgumentParser(
        description=(
            "Batch-generate imide-terminated .mol products from diamine and "
            "dianhydride .mol files."
        )
    )
    parser.add_argument(
        "--diamines",
        nargs="+",
        type=Path,
        default=[script_dir / "diamines"],
        help="Diamine .mol file(s) or directories. Default: ./diamines",
    )
    parser.add_argument(
        "--dianhydrides",
        nargs="+",
        type=Path,
        default=[script_dir / "dianhydrides"],
        help="Dianhydride .mol file(s) or directories. Default: ./dianhydrides",
    )
    parser.add_argument(
        "--out",
        type=Path,
        default=script_dir / "products",
        help="Output directory. Default: ./products",
    )
    parser.add_argument(
        "--all-sites",
        action="store_true",
        help="Generate every terminal-amine/cyclic-anhydride site pairing instead of only the first detected pair.",
    )
    parser.add_argument(
        "--no-explicit-caps",
        action="store_true",
        help=(
            "Do not add explicit H atoms where unused diamine amine groups are "
            "removed. Terminal imide N-H groups are always explicit."
        ),
    )
    parser.add_argument(
        "--list-sites",
        action="store_true",
        help="Only print detected amine and anhydride reactive sites.",
    )
    return parser


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()
    return run(args)


if __name__ == "__main__":
    raise SystemExit(main())
