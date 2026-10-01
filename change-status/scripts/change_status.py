"""Change status: edit a fixed-width MMAN EDI .txt from a B/L / Status / POL list.

Usage:
    python change_status.py <folder>                      # folder holding Input/ (MMAN_*.txt + .xls/.xlsx/.csv)
    python change_status.py <mman.txt> <list.xls|xlsx|csv> [-o OUTPUT]
Options:
    --shed-yes 0145   shed code for B/Ls that have a Status in column B
    --shed-no  0122   shed code for all other B/Ls

The list file: only columns A-C are read -> A=B/L No., B=Status ("9  N", "7  N", blank), C=POL (KRPUS...).
Output is always MMAN_EDI.txt (next to Input/ in folder mode). Input files are never modified.
"""
import argparse
import re
import sys
from collections import Counter
from pathlib import Path

import pandas as pd

ENC = "latin-1"          # byte-for-byte safe; the file is fixed-width, never change its length
COL_POL = 35             # 1. POL (5 chars)
COL_COPY = 40            # 2. copy of col 45-49 (THBKK)
COL_PORT = 45            # 3. original THBKK - unchanged
COL_SHED = 50            # 4. shed code (4 chars)
COL_CTRY = 843           # 5. first 2 letters of POL (replaces TH)
COL_STATUS = 17          # 6. status on container lines ("8  N" -> "9  N")
BL_RE = re.compile(r"^HASL")
CNTR_RE = re.compile(r"^[A-Z]{4}\d{7} ")
UNIT_RE = re.compile(r"(\d{11}KGM\d{8})PG(\d{11}MTQ)")


def map_columns(headers: list) -> list:
    """Name columns A-C by their header text, so any order works
    (KMBK form: B/L No., Status, POL; SUR form: POL, Status, B/L No.)."""
    names = []
    for h in headers:
        k = re.sub(r"[^a-z]", "", str(h).lower())
        names.append("bl" if k.startswith("bl") else "status" if k.startswith("status") else "pol" if k.startswith("pol") else None)
    if sorted(n for n in names if n) == ["bl", "pol", "status"]:
        return names
    return ["bl", "status", "pol"]  # no recognisable headers: default order A=B/L, B=Status, C=POL


def norm_status(st: str) -> str:
    """'7 N' / '7N' / '7   N' -> '7  N' (code at col 17, flag at col 20, as in the file)."""
    t = st.split()
    if len(t) == 1 and len(t[0]) == 2:
        t = [t[0][0], t[0][1]]
    if len(t) == 2 and len(t[0]) == 1 and len(t[1]) == 1:
        return f"{t[0]}  {t[1]}".upper()
    raise SystemExit(f"ERROR: unrecognised Status {st!r} (expected like '9  N')")


def read_list(path: Path) -> dict:
    if path.suffix.lower() == ".csv":
        df = pd.read_csv(path, header=0, dtype=str)
    else:
        df = pd.read_excel(path, header=0, dtype=str)
    df = df.iloc[:, :3]
    df.columns = map_columns(list(df.columns))
    cfg = {}
    for r in df.itertuples():
        if not isinstance(r.bl, str) or not r.bl.strip():
            continue
        st = norm_status(r.status) if isinstance(r.status, str) and r.status.strip() else None
        pol = r.pol.strip().upper() if isinstance(r.pol, str) else ""
        cfg[r.bl.strip()] = (st, pol)
    return cfg


def set_at(s: str, pos: int, txt: str) -> str:
    return s[:pos] + txt + s[pos + len(txt):]


def change_status(mman: Path, lst: Path, out: Path, shed_yes: str, shed_no: str) -> dict:
    cfg = read_list(lst)
    raw = mman.read_bytes()
    parts = [p.decode(ENC) for p in re.split(rb"(\r\n|\n)", raw)]  # keep original line endings

    rep = {"bl_lines": 0, "shed_yes": [], "shed_no": [], "status_lines": Counter(),
           "pg_to_px": [], "not_in_list": [], "no_pol": [], "layout_warn": []}
    cur = None
    for i, line in enumerate(parts):
        if BL_RE.match(line):
            bl = line[:20].strip()
            base = bl[:16]
            if base not in cfg:
                rep["not_in_list"].append(bl)
                cur = None
                continue
            st, pol = cfg[base]
            cur = st
            if line[COL_PORT:COL_PORT + 5].strip() == "" or line[COL_POL:COL_COPY].strip():
                rep["layout_warn"].append(bl)
            if pol:
                line = set_at(line, COL_POL, pol.ljust(5)[:5])
                line = set_at(line, COL_CTRY, pol[:2])
            else:
                rep["no_pol"].append(bl)
            line = set_at(line, COL_COPY, line[COL_PORT:COL_PORT + 5])
            line = set_at(line, COL_SHED, shed_yes if st else shed_no)
            m = UNIT_RE.search(line)
            if m:
                line = line[:m.start()] + m.group(1) + "PX" + m.group(2) + line[m.end():]
                rep["pg_to_px"].append(bl)
            rep["shed_yes" if st else "shed_no"].append(bl)
            rep["bl_lines"] += 1
            parts[i] = line
        elif cur and CNTR_RE.match(line):
            parts[i] = set_at(line, COL_STATUS, cur)
            rep["status_lines"][cur] += 1

    data = "".join(parts).encode(ENC)
    if len(data) != len(raw):
        raise SystemExit(f"ERROR: size changed {len(raw)} -> {len(data)}; output not written")
    out.write_bytes(data)
    rep["bytes_changed"] = sum(a != b for a, b in zip(raw, data))
    rep["size"] = len(data)
    return rep


def find_inputs(folder: Path):
    inp = folder / "Input" if (folder / "Input").is_dir() else folder
    txts = [t for t in sorted(inp.glob("*.txt")) if not t.name.upper().startswith("MMAN_EDI")]
    txts = [t for t in txts if t.name.upper().startswith("MMAN")] or txts
    lists = sorted([*inp.glob("*.xls"), *inp.glob("*.xlsx"), *inp.glob("*.csv")])
    if len(txts) != 1 or len(lists) != 1:
        raise SystemExit(f"ERROR: need exactly 1 MMAN .txt and 1 list file in {inp}; "
                         f"found txt={[t.name for t in txts]} list={[l.name for l in lists]}")
    return txts[0], lists[0], inp / "MMAN_EDI.txt"


def main():
    ap = argparse.ArgumentParser(description="Change status -> MMAN_EDI.txt")
    ap.add_argument("paths", nargs="+")
    ap.add_argument("-o", "--output")
    ap.add_argument("--shed-yes", default="0145")
    ap.add_argument("--shed-no", default="0122")
    a = ap.parse_args()
    if len(a.paths) == 1:
        mman, lst, out = find_inputs(Path(a.paths[0]))
    else:
        mman, lst = Path(a.paths[0]), Path(a.paths[1])
        out = mman.parent / "MMAN_EDI.txt"
    if a.output:
        out = Path(a.output)
    rep = change_status(mman, lst, out, a.shed_yes, a.shed_no)

    print(f"MMAN : {mman}\nLIST : {lst}\nOUT  : {out}")
    print(f"B/L lines edited : {rep['bl_lines']}  (shed {a.shed_yes}: {len(rep['shed_yes'])}, "
          f"shed {a.shed_no}: {len(rep['shed_no'])})")
    print("Status set       : " + (", ".join(f"{k.strip()!r} x{v}" for k, v in rep["status_lines"].items()) or "none"))
    print(f"PG -> PX         : {rep['pg_to_px'] or 'none'}")
    print(f"Size / changed   : {rep['size']} bytes, {rep['bytes_changed']} bytes changed")
    for key, label in (("not_in_list", "B/L NOT in list (untouched)"), ("no_pol", "B/L without POL"),
                       ("layout_warn", "Layout warning (check columns)")):
        if rep[key]:
            print(f"WARNING {label}: {rep[key]}")


if __name__ == "__main__":
    sys.exit(main())
