"""Change status: edit a fixed-width MMAN EDI .txt from a B/L / Status / POL list.

Usage:
    python change_status.py <folder>                      # folder holding Input/ (MMAN_*.txt + .xls/.xlsx/.csv)
    python change_status.py <mman.txt> <list.xls|xlsx|csv> [-o OUTPUT]
Options:
    --shed-yes 0145   shed code for B/Ls that have a Status in column B
    --shed-no  0122   shed code for all other B/Ls

The list file: only columns A-C are read -> A=B/L No., B=Status ("9  N", "7  N", blank), C=POL (KRPUS...).
Output is always MMAN_EDI.txt (in output/ next to input/ in folder mode). Input files are never modified.
"""
import argparse
import re
import sys
from collections import Counter
from pathlib import Path

import pandas as pd

ENC = "latin-1"          # byte-for-byte safe; the file is fixed-width, never change its length
COL_POL = 35             # 1. POL (5 chars)
COL_COPY = 40            # 2. port of discharge = the port most B/Ls have in col 45-49 (THBKK, THLCH)
COL_PORT = 45            # 3. original THBKK - unchanged
COL_SHED = 50            # 4. shed code (4 chars)
COL_CTRY = 843           # 5. first 2 letters of POL (replaces TH)
COL_STATUS = 17          # 6. status on container lines ("8  N" -> "9  N")
BL_RE = re.compile(r"^HASL")
CNTR_RE = re.compile(r"^[A-Z]{4}\d{7} ")
UNIT_RE = re.compile(r"(\d{11}KGM\d{8})PG(\d{11}MTQ)")
# 7. names: consignee name (880), consignee name+address (973) and notify name (1671)
#    take the name written at the notify name+address field (1764), single spaces only
COL_C_NAME, COL_C_FULL, COL_N_NAME, COL_N_FULL = 880, 973, 1671, 1764
W_NAME, W_FULL = 70, 696   # name = 70 chars (880-949); 950-972 holds tax ID + branch, never touched
TRANSIT_RE = re.compile(r"TRANSIT\b|\bTRANS[\s-]?S?HIP", re.I)   # also "INTRANSIT"
# 8. TRANSIT/TRANSHIPMENT: notify country (col 1669, "TH") = destination country, found in the
#    consignee name+address (col 973), else in the description ("... TO LAOS ...")
COL_N_CTRY = 1669
# 9. TRANSIT/TRANSHIPMENT: the code before the description (col 3803, normally "O"):
#    TRANSHIPMENT = 3, TRANSIT = 7, TRANSIT to LAOS = the "CARGO MOVEMENT = L / 7" written in the B/L (7 if none)
COL_MOVE = 3803
TRANSHIP_RE = re.compile(r"\bTRANS[\s-]?S?HIP", re.I)
MOVEMENT_RE = re.compile(r"MOVEMENT\s*[=:]?\s*([L7])\b", re.I)
DEST_CTRY = [("LA", r"\bLAOS?\b|\bLAO\s*P\.?\s*D\.?\s*R|\bVIENTIANE\b"), ("MM", r"\bMYANMAR\b|\bBURMA\b|\bYANGON\b"),
             ("KH", r"\bCAMBODIA\b|\bPHNOM\s*PENH\b"), ("CN", r"\bCHINA\b"), ("PG", r"\bPAPUA\s+NEW\s+GUINEA\b"),
             ("MY", r"\bMALAYSIA\b")]
# 10. every B/L except TRANSIT/TRANSHIPMENT: blank the notify tax ID and the address fields of the
#     consignee and the notify party (the consignee tax ID at 950 and both branch codes stay)
CLEAR = [(1741, 13), (1415, 254), (2206, 254)]   # notify tax ID, consignee address, notify address
# 4. shed locked by the original port (col 45-49); THLCH with DG (UN/CLASS) = 2826;
#    THBKK and the other THLCH B/Ls use the SHED NO. given (0145 with Status / 0122 without)
PORT_SHED = {"THBMT": "0110", "THSCT": "0302", "THLKR": "0332", "THUCT": "0113"}
SHED_LCH_DG = "2826"
DG_RE = re.compile(r"\bUN\s?(?:NO\.?\s?)?\d{4}\b|\bIMO\b|\bIMDG\b|\bCLASS\s?\d", re.I)


def dest_country(line: str):
    """Country code of the transit destination, or None."""
    def hits(text):
        return sorted((m.start(), code) for code, pat in DEST_CTRY for m in re.finditer(pat, text, re.I))
    addr, desc = hits(line[COL_C_FULL:COL_C_FULL + W_FULL]), hits(line[3302:])
    if addr:
        return addr[-1][1]   # the country closes the address
    return desc[0][1] if desc else None


def transit_move(line: str, dest):
    """Code for col 3803 of a TRANSIT/TRANSHIPMENT B/L."""
    if TRANSHIP_RE.search(line):
        return "3"
    if dest == "LA":
        m = MOVEMENT_RE.search(line[3302:])
        return m.group(1).upper() if m else "7"
    return "7"


def fix_names(line: str):
    """Copy the name at col 1764 into cols 880, 973 and 1671 with single spaces only.
    Returns (line, result): 'changed', 'same', 'transit' (TRANSIT/TRANSHIPMENT B/L, names kept)
    or 'differ' (consignee is not the notify party, kept)."""
    old = line[COL_C_NAME:COL_C_NAME + W_NAME].rstrip()
    if not old:
        return line, "same"
    if TRANSIT_RE.search(line):
        code = dest_country(line)
        line = set_at(line, COL_MOVE, transit_move(line, code))
        return (set_at(line, COL_N_CTRY, code) if code else line), "transit"
    full = line[COL_N_FULL:COL_N_FULL + W_FULL]

    def prefix(target):
        """The start of the 1764 field whose words are `target` (the name before the address)."""
        for m in re.finditer(r"\S+", full):
            words = " ".join(full[:m.end()].split())
            if words == target:
                return full[:m.end()]
            if len(words) >= len(target):
                return None
        return None

    notify = " ".join(line[COL_N_NAME:COL_N_NAME + W_NAME].split())
    new = prefix(" ".join(old.split())) or (prefix(notify) if notify else None) or notify
    if not new:
        return line, "differ"
    new = " ".join(new.split())[:W_NAME]
    out = set_at(line, COL_C_NAME, new.ljust(W_NAME))
    out = set_at(out, COL_N_NAME, new.ljust(W_NAME))
    field = out[COL_C_FULL:COL_C_FULL + W_FULL]
    if field.startswith(old):  # the consignee name opens its name+address field
        nf = new + field[len(old):]
        if not nf[W_FULL:].strip():
            out = set_at(out, COL_C_FULL, nf[:W_FULL].ljust(W_FULL))
    return out, "changed" if out != line else "same"


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
    dg_col = next((c for c in df.columns if re.sub(r"[^a-z]", "", str(c).lower()) == "dg"), None)
    dg = df[dg_col].fillna("").astype(str).str.strip() != "" if dg_col is not None else pd.Series(False, index=df.index)
    df = df.iloc[:, :3].copy()
    df.columns = map_columns(list(df.columns))
    df["dg"] = dg
    cfg = {}
    for r in df.itertuples():
        if not isinstance(r.bl, str) or not r.bl.strip():
            continue
        st = norm_status(r.status) if isinstance(r.status, str) and r.status.strip() else None
        pol = r.pol.strip().upper() if isinstance(r.pol, str) else ""
        cfg[r.bl.strip()] = (st, pol, bool(r.dg))
    return cfg


def set_at(s: str, pos: int, txt: str) -> str:
    return s[:pos] + txt + s[pos + len(txt):]


def change_status(mman: Path, lst: Path, out: Path, shed_yes: str, shed_no: str) -> dict:
    cfg = read_list(lst)
    raw = mman.read_bytes()
    parts = [p.decode(ENC) for p in re.split(rb"(\r\n|\n)", raw)]  # keep original line endings

    rep = {"bl_lines": 0, "shed_yes": [], "shed_no": [], "sheds": Counter(), "dg": [], "status_lines": Counter(),
           "pg_to_px": [], "not_in_list": [], "no_pol": [], "layout_warn": [],
           "names": {"changed": [], "same": [], "transit": [], "differ": []}, "transit_ctry": [], "transit_no_ctry": [], "cleared": []}
    ports = Counter(p[COL_PORT:COL_PORT + 5] for p in parts if BL_RE.match(p) and p[COL_PORT:COL_PORT + 5].strip())
    rep["port"] = port_main = ports.most_common(1)[0][0] if ports else ""
    cur = None
    for i, line in enumerate(parts):
        if BL_RE.match(line):
            bl = line[:20].strip()
            base = bl[:16]
            if base not in cfg:
                rep["not_in_list"].append(bl)
                cur = None
                continue
            st, pol, dg = cfg[base]
            dg = dg or bool(DG_RE.search(line[3302:]))
            cur = st
            if line[COL_PORT:COL_PORT + 5].strip() == "" or line[COL_POL:COL_COPY].strip():
                rep["layout_warn"].append(bl)
            if pol:
                line = set_at(line, COL_POL, pol.ljust(5)[:5])
                line = set_at(line, COL_CTRY, pol[:2])
            else:
                rep["no_pol"].append(bl)
            line = set_at(line, COL_COPY, port_main or line[COL_PORT:COL_PORT + 5])
            port = line[COL_PORT:COL_PORT + 5]
            sh = PORT_SHED.get(port) or (SHED_LCH_DG if port == "THLCH" and dg else (shed_yes if st else shed_no))
            line = set_at(line, COL_SHED, sh)
            rep["sheds"][f"{sh} ({port})"] += 1
            if dg:
                rep["dg"].append(bl)
            m = UNIT_RE.search(line)
            if m:
                line = line[:m.start()] + m.group(1) + "PX" + m.group(2) + line[m.end():]
                rep["pg_to_px"].append(bl)
            if len(line) >= COL_N_FULL + W_FULL:
                line, res = fix_names(line)
                rep["names"][res].append(bl)
                if res == "transit":
                    code = line[COL_N_CTRY:COL_N_CTRY + 2]
                    if code != "TH":
                        rep["transit_ctry"].append(f"{bl}={code} move {line[COL_MOVE]}")
                    else:
                        rep["transit_no_ctry"].append(bl)
                else:
                    before = line
                    for pos, width in CLEAR:
                        line = set_at(line, pos, " " * width)
                    if line != before:
                        rep["cleared"].append(bl)
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
    out_dir = folder / "output"   # result goes to output/, next to input/
    out_dir.mkdir(exist_ok=True)
    return txts[0], lists[0], out_dir / "MMAN_EDI.txt"


def main():
    ap = argparse.ArgumentParser(description="Change status -> MMAN_EDI.txt")
    ap.add_argument("paths", nargs="+")
    ap.add_argument("-o", "--output")
    ap.add_argument("--shed-yes", default="0145")
    ap.add_argument("--shed-no", default="0122", help='"" = use --shed-yes for every B/L (e.g. LCH jobs without Status)')
    ap.add_argument("--shed", help="one SHED NO. for every B/L (same as --shed-yes X --shed-no X)")
    a = ap.parse_args()
    if a.shed:
        a.shed_yes = a.shed_no = a.shed
    if not a.shed_no.strip():
        a.shed_no = a.shed_yes
    if len(a.paths) == 1:
        mman, lst, out = find_inputs(Path(a.paths[0]))
    else:
        mman, lst = Path(a.paths[0]), Path(a.paths[1])
        out = mman.parent / "MMAN_EDI.txt"
    if a.output:
        out = Path(a.output)
    rep = change_status(mman, lst, out, a.shed_yes, a.shed_no)

    print(f"MMAN : {mman}\nLIST : {lst}\nOUT  : {out}")
    print(f"B/L lines edited : {rep['bl_lines']}  (port col 41-45: {rep['port']})")
    print("SHED             : " + ", ".join(f"{k} x{v}" for k, v in sorted(rep["sheds"].items())))
    if rep["dg"]:
        print(f"DG B/L           : {rep['dg']}")
    print("Status set       : " + (", ".join(f"{k.strip()!r} x{v}" for k, v in rep["status_lines"].items()) or "none"))
    print(f"PG -> PX         : {rep['pg_to_px'] or 'none'}")
    n = rep["names"]
    print(f"Names (col 1765) : changed {len(n['changed'])}, already single-spaced {len(n['same'])}")
    if n["transit"]:
        print(f"TRANSIT/TRANSHIP : names kept {n['transit']}  country/code: {rep['transit_ctry']}")
    print(f"Cleared notify tax ID + addresses: {len(rep['cleared'])} B/L")
    print(f"Size / changed   : {rep['size']} bytes, {rep['bytes_changed']} bytes changed")
    if rep["transit_no_ctry"]:
        print(f"WARNING TRANSIT without a known destination country (TH kept): {rep['transit_no_ctry']}")
    if n["differ"]:
        print(f"WARNING no name at col 1765 / notify name (kept, check): {n['differ']}")
    for key, label in (("not_in_list", "B/L NOT in list (untouched)"), ("no_pol", "B/L without POL"),
                       ("layout_warn", "Layout warning (check columns)")):
        if rep[key]:
            print(f"WARNING {label}: {rep[key]}")
    if rep["bl_lines"] == 0:
        print("ERROR: no B/L in the MMAN file matches the list - check that the list belongs to this vessel")
        return 1


if __name__ == "__main__":
    sys.exit(main())
