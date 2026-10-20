---
name: change-status
description: >
  "Change status" for a Thailand import MMAN EDI manifest (fixed-width MMAN_*.txt). Reads a
  B/L list (.xls/.xlsx/.csv, or a PDF, where only columns A-C matter: B/L No., Status, POL)
  and writes MMAN_EDI.txt: fills POL + a copy of THBKK before THBKK, a shed code (0145 for
  B/Ls with a Status, 0122 for the rest), the POL country code (KR) in place of TH, the new
  container status (e.g. 8  N -> 9  N), and PG -> PX. Use whenever the user says
  "Change status", "เปลี่ยน status mman", "แก้ MMAN", or points to a folder with Input/
  holding an MMAN_*.txt plus a status list.
---

# Change status (MMAN → MMAN_EDI.txt)

Rattana's confirmed format. Run the bundled script; don't edit the file by hand. The file is
fixed-width, latin-1 and CRLF. Never change its length or line endings (sed in Git Bash once
broke CRLF).

## Input

- A folder with `Input/` (or `input/`) containing **1** `MMAN_*.txt` + **1** list file (`.xls`, `.xlsx`, `.csv`)
- List file: **only columns A-C are used** (ignore D-F; sometimes there are only A-C).
  Columns are found **by header name**, so either form works:
  - **KMBK form**: A = B/L No., B = Status, C = POL
  - **SUR form** (e.g. `SUR.xls`): A = POL, B = Status, C = B/L No.
  - B/L No. = base B/L, e.g. `HASLK01260700462`; also applied to sub-B/Ls `…462A`, `…462B`
  - Status = the status to set, e.g. `9  N`, `7  N`; blank = keep the original status
  - POL, e.g. `KRPUS`, `KRKAN`
- If the list is a **PDF**, extract columns A-C into a CSV with the header `B/L No.,Status,POL`
  (keep the spacing in Status, e.g. `9  N`), then pass that CSV.

## Rules (implemented in the script, 0-based columns on each HASL… line)

| # | Position | Value |
|---|---|---|
| 1 | col 35-39 | POL (column C) |
| 2 | col 40-44 | port of discharge = the port most B/Ls in the file have at col 45-49 (THBKK or THLCH), e.g. `CNTAOTHLCHTHLKR0332` |
| 3 | col 45-49 | THBKK from the original, unchanged |
| 4 | col 50-53 | shed **locked by the original port**: THBMT 0110, THSCT 0302, THLKR 0332, THUCT 0113, THLCH with DG (UN/CLASS in the description or a DG column in the list) 2826. THBKK and other THLCH: `0145` if the B/L has a Status, otherwise `0122` (`--shed-yes/--shed-no` to change; `--shed X` = one code for every B/L, e.g. LCH jobs with no Status — ask the user which code) |
| 5 | col 843-844 | first 2 letters of POL (`KR`), replacing `TH` |
| 6 | container line col 17 | Status from column B (`8  N` → `9  N`) |
| – | package unit | `PG` → `PX` on every B/L |

| 7 | col 880 / 973 / 1671 | consignee name, consignee name+address, notify name = **always** the name at col 1764, **single spaces only**, for BKK and LCH (the part of 1764 matching the consignee words, else the notify-name words). Only TRANSIT/TRANSHIPMENT B/Ls keep their names |
| 8 | col 1669-1670 | TRANSIT/TRANSHIPMENT: TH before the notify name = destination country (LA, MM, KH, CN, PG, MY) |
| 9 | col 3803 | TRANSIT/TRANSHIPMENT only: TRANSIT = 7, TRANSIT to LAOS = CARGO MOVEMENT L/7 in the B/L (7 if none), TRANSHIPMENT = 3 |
| 10 | col 1741 / 1415 / 2206 | every B/L except TRANSIT/TRANSHIPMENT: blank the notify tax ID and the consignee / notify addresses |

Rules 7-10 are the same code as the 5-mman-vnsgn site. **Otherwise leave the letter `O` before the description (col 3803) unchanged.** The user cancelled the O→5 rule.

## Steps

1. Run:
   ```bash
   PYTHONIOENCODING=utf-8 python "<skill dir>/scripts/change_status.py" "<folder containing Input>"
   ```
   or specify the files: `change_status.py <mman.txt> <list.xls> [-o out.txt]`.
   The output is always `MMAN_EDI.txt` **in the `output/` folder** (created next to `input/`). The input files are never modified.
   Requires `pandas`, `xlrd` (.xls) and `openpyxl` (.xlsx).
2. Report to the user in Thai: how many B/Ls, 0145/0122 counts, how many status lines changed and to what, which B/Ls went PG→PX, and that the file size matches the original.
3. If the script prints a `WARNING` (B/L not in the list, B/L without POL, unexpected layout), tell the user and ask. Don't guess.
