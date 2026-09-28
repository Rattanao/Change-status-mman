# check-status-mman

**Change status** tool: edits the MMAN EDI file (fixed-width `MMAN_*.txt`) using the B/L list (Excel/CSV), outputs `MMAN_EDI.txt`

Can be used as a standalone Python program or as a **Claude Code skill** (`change-status`)

## What it changes (on each B/L line)

```
Before:   HASLK01260700462                             THBKK    GOLDWAY…
After:    HASLK01260700462                   KRPUSTHBKKTHBKK0145GOLDWAY…
                                             │    │    │    └ 4. shed: 0145 (B/L with Status) / 0122 (the rest)
                                             │    │    └ 3. THBKK (original)
                                             │    └ 2. THBKK (copy of #3)
                                             └ 1. POL (column C)
```

- 5. `TH` → first 2 letters of POL (`KR`), in the country field after the notify address
- 6. Status on container lines: `SKHU9545854      8  N` → `SKHU9545854      9  N` (per column B; blank = unchanged)
- Package unit `PG` → `PX` on every B/L
- The letter `O` before the description is **not changed**
- File size and line breaks (CRLF) stay the same as the original

## B/L list file

Only columns **A-C** are used (D-F ignored)

| B/L No. | Status | POL |
|---|---|---|
| HASLK01260700462 | 9  N | KRPUS |
| HASLK01260805808 |  | KRPUS |
| HASLK01260807609 |  | KRKAN |

A B/L in column A also covers its sub-B/Ls (`HASLK01260700462A`, `…B`, …)

## Usage

Install the required packages (once):

```bash
pip install pandas xlrd openpyxl
```

Folder layout:

```
check-status/
├── Input/
│   ├── MMAN_THBKK_....txt
│   └── status-kmbk.xls
└── MMAN_EDI.txt        ← output
```

Run:

```bash
python change-status/scripts/change_status.py "C:\Users\...\Downloads\check-status"
```

Or specify the files directly:

```bash
python change-status/scripts/change_status.py MMAN.txt status.xls -o MMAN_EDI.txt
```

Change the shed codes:

```bash
python change-status/scripts/change_status.py "<folder>" --shed-yes 0145 --shed-no 0122
```

## Using as a Claude Code skill

Copy the `change-status/` folder to `~/.claude/skills/`

```bash
cp -r change-status ~/.claude/skills/
```

Then tell Claude: **"Change status"** and point it at the folder containing `Input/`
