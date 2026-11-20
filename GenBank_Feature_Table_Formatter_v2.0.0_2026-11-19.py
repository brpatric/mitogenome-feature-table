#!/usr/bin/env python3
r"""
GenBank_Feature_Table_Formatter_v2.0.0_2026-11-19.py

Writes a GenBank five-column feature table (.tbl) for a mitochondrial genome
from a gene-order spreadsheet (one row per gene: Gene, Begin, End, Strand,
Start codon, Stop codon). The .tbl file is what NCBI's table2asn (or
BankIt) needs, together with the genome FASTA, to make a GenBank submission.

PROVENANCE
  Program    : GenBank_Feature_Table_Formatter
  Version    : 2.0.0
  Date       : 2026-11-19
  Author     : L. Brian Patrick, Department of Biological Sciences,
               Dakota Wesleyan University, Mitchell, SD, USA
               ORCID 0000-0003-1969-9600
  Development: v1.0.0 written by the author with the assistance of Claude
               (Anthropic). v2.0.0 drafted with the assistance of Claude
               Opus 5.5 (Anthropic), 2026-11-19, from specifications by the
               author, and reviewed by the author.
  Derived from: GenBank_Feature_Table_Formatter.py (v1.0.0, the author's
               original script).
  Tested with: Python 3.13.18 and openpyxl 3.1.5, on the two Erigone atra
               mitogenomes in example/ (GenBank PP726906 and PP726907).
               Output was compared feature by feature with both GenBank
               records and checked with NCBI tbl2asn 25.7 (see README).
  Companion  : README_GenBank_Feature_Table_Formatter_v2.0.0_2026-11-19.txt
               (README.md on GitHub)
  Version history
    1.0.0  (undated)   Original script: GeneOrder Excel file (exactly six
                       columns) -> .tbl; parameters edited at the bottom of
                       the file; pandas + openpyxl.
    2.0.0  2026-11-19  Rewrite. Columns found by header name (extra columns
                       such as Length are ignored); MITOS2-style gene names
                       accepted; features crossing the origin handled for
                       every feature type and strand, with the two intervals
                       in the correct order (v1.0.0 reversed them for
                       minus-strand tRNAs); transl_except qualifiers for
                       incomplete stop codons; footnote letters on codons
                       ignored (TAb -> TA); start codons checked against the
                       genetic code (non-standard starts marked 5' partial);
                       optional FASTA check of every protein-coding gene;
                       NCBI product names for rRNAs; control region written
                       as a valid misc_feature; sequence ID checked for
                       spaces; SETTINGS block at the top, questions in the
                       shell if it is left empty, and command-line options;
                       pandas no longer needed; CSV/TSV input accepted.

HOW TO USE (IDLE 3.13)
  1. Open this file in IDLE.
  2. Edit the SETTINGS below (at least INPUT_FILE; FASTA_FILE is recommended).
  3. Save, then press F5 (Run > Run Module).
  If INPUT_FILE is left as None, the program asks for the file names instead.
  After one F5 run, run() can be called again from the IDLE shell prompt:
    >>> run(r"C:\path\GeneOrder_male.xlsx", fasta_file=r"C:\path\male.fasta")

OPTIONAL: COMMAND LINE (bash, macOS Terminal, Windows PowerShell)
  python3 GenBank_Feature_Table_Formatter_v2.0.0_2026-11-19.py GENE_ORDER
          [--fasta FILE] [--out FILE] [--seq-id ID] [--length BP] [--code N]
          [--sheet NAME] [--no-partial-start]
  With no arguments, the SETTINGS below are used.

INPUT: gene-order table (.xlsx, .csv, .tsv or .txt), one mitogenome per file
  Columns are found by their header names (any order, any capitalisation);
  other columns (e.g. Length, Intergenic nucleotides) are ignored.
    Gene    gene name: COX1, cox1, nad5, ND5, cob, CytB, trnL1, trnL1(tag),
            16S, rrnL, 12S, rrnS, Control Region, ... (see README for the list)
    Begin   first position of the gene, counting along the + strand
    End     last position, counting along the + strand. For a gene that
            crosses the origin, Begin is larger than End (e.g. 14448 and 5).
    Strand  + or - (J/N and plus/minus are also accepted)
    Start   start codon, protein-coding genes only (e.g. ATA)
    Stop    stop codon, protein-coding genes only: TAA, TAG, or an incomplete
            stop (TA or T). Footnote letters (TAb, Tb) are ignored.
  For an .xlsx file the first sheet is read unless SHEET is given.

OUTPUT
  <OUTPUT_FILE>   five-column, tab-delimited GenBank feature table.
                  Default: the input file name with the extension .tbl.

Requires: openpyxl for .xlsx input  (pip install openpyxl). CSV/TSV input
          needs only the Python standard library.

REFERENCES
  NCBI. Feature table format (five-column format used by table2asn and
    BankIt). https://www.ncbi.nlm.nih.gov/genbank/feature_table/
    (origin-crossing features follow its minus-strand example:
     500 1 gene / 15000 14200)
  NCBI. Organelle genome submissions (checks applied to mitochondrial
    genomes). https://www.ncbi.nlm.nih.gov/genbank/organelle_submit/
  NCBI. The Genetic Codes, table 5: invertebrate mitochondrial code.
    https://www.ncbi.nlm.nih.gov/Taxonomy/Utils/wprintgc.cgi
  Ojala, D., Montoya, J., & Attardi, G. (1981). tRNA punctuation model of
    RNA processing in human mitochondria. Nature, 290(5806), 470-474.
    https://doi.org/10.1038/290470a0  (incomplete stop codons completed by
    polyadenylation)
  Patrick, L. B., & Santos, S. R. (2025). The complete mitochondrial genome
    of Erigone atra Blackwall, 1833 (Araneae, Linyphiidae). Mitochondrial DNA
    Part B, 10(12), 1243-1248. https://doi.org/10.1080/23802359.2025.2594297
    (source of the worked example)
"""

# =============================================================================
# SETTINGS  (edit these, save, then press F5 in IDLE)
# =============================================================================

INPUT_FILE    = None    # gene-order table, e.g. r"C:\Users\me\GeneOrder_male.xlsx"
                        # None = ask for the file names when the program runs
FASTA_FILE    = None    # genome FASTA (recommended): gives the genome length and
                        # sequence ID, and lets every protein-coding gene be checked
OUTPUT_FILE   = None    # None = input file name with the extension .tbl
SEQ_ID        = None    # sequence ID on the >Feature line; must match the FASTA
                        # header. None = first word of the FASTA header, or the
                        # input file name if no FASTA is given
GENOME_LENGTH = None    # bp; only needed when no FASTA is given and a gene
                        # crosses the origin
TRANSL_TABLE  = 5       # NCBI genetic code: 5 = invertebrate mitochondrial
SHEET         = None    # .xlsx sheet name; None = first sheet
PARTIAL_IF_INVALID_START = True   # mark a CDS as 5' partial (<) when its start
                                  # codon is not a start codon in TRANSL_TABLE

# =============================================================================

import csv
import os
import re
import sys

PROGRAM = "GenBank_Feature_Table_Formatter"
VERSION = "2.0.0"
DATE = "2026-11-19"

# -----------------------------------------------------------------------------
# Gene names
# -----------------------------------------------------------------------------

# GenBank gene symbol -> product name (NCBI standard names for mitochondrial genes)
PCG_PRODUCTS = {
    "COX1": "cytochrome c oxidase subunit I",
    "COX2": "cytochrome c oxidase subunit II",
    "COX3": "cytochrome c oxidase subunit III",
    "ATP6": "ATP synthase F0 subunit 6",
    "ATP8": "ATP synthase F0 subunit 8",
    "CYTB": "cytochrome b",
    "ND1": "NADH dehydrogenase subunit 1",
    "ND2": "NADH dehydrogenase subunit 2",
    "ND3": "NADH dehydrogenase subunit 3",
    "ND4": "NADH dehydrogenase subunit 4",
    "ND4L": "NADH dehydrogenase subunit 4L",
    "ND5": "NADH dehydrogenase subunit 5",
    "ND6": "NADH dehydrogenase subunit 6",
}

# accepted spellings (lower case, letters and digits only) -> gene symbol
PCG_ALIASES = {}
for _sym, _names in {
    "COX1": ["cox1", "coi", "co1", "coxi", "cytochromecoxidasesubuniti"],
    "COX2": ["cox2", "coii", "co2", "coxii", "cytochromecoxidasesubunitii"],
    "COX3": ["cox3", "coiii", "co3", "coxiii", "cytochromecoxidasesubunitiii"],
    "ATP6": ["atp6", "atpase6", "atp06"],
    "ATP8": ["atp8", "atpase8", "atp08"],
    "CYTB": ["cytb", "cob", "cyb", "cytochromeb"],
    "ND1": ["nd1", "nad1", "nadh1"],
    "ND2": ["nd2", "nad2", "nadh2"],
    "ND3": ["nd3", "nad3", "nadh3"],
    "ND4": ["nd4", "nad4", "nadh4"],
    "ND4L": ["nd4l", "nad4l", "nadh4l"],
    "ND5": ["nd5", "nad5", "nadh5"],
    "ND6": ["nd6", "nad6", "nadh6"],
}.items():
    for _n in _names:
        PCG_ALIASES[_n] = _sym

RRNA_ALIASES = {}
for _prod, _names in {
    "16S ribosomal RNA": ["16s", "16srrna", "16sribosomalrna", "rrnl", "rnl",
                          "lrrna", "largesubunitribosomalrna", "lsurrna"],
    "12S ribosomal RNA": ["12s", "12srrna", "12sribosomalrna", "rrns", "rns",
                          "srrna", "smallsubunitribosomalrna", "ssurrna"],
}.items():
    for _n in _names:
        RRNA_ALIASES[_n] = _prod

CONTROL_REGION_NAMES = {"controlregion", "putativecontrolregion", "cr", "dloop",
                        "atrichregion", "atrich", "ncr", "noncodingregion"}

AA3 = {"A": "Ala", "R": "Arg", "N": "Asn", "D": "Asp", "C": "Cys", "Q": "Gln",
       "E": "Glu", "G": "Gly", "H": "His", "I": "Ile", "L": "Leu", "K": "Lys",
       "M": "Met", "F": "Phe", "P": "Pro", "S": "Ser", "T": "Thr", "W": "Trp",
       "Y": "Tyr", "V": "Val"}
AA1 = {v.lower(): k for k, v in AA3.items()}

# -----------------------------------------------------------------------------
# Genetic codes (NCBI numbering). Each code is the standard code plus changes.
# -----------------------------------------------------------------------------

_BASES = "TCAG"
_STANDARD_AA = "FFLLSSSSYY**CC*WLLLLPPPPHHQQRRRRIIIMTTTTNNKKSSRRVVVVAAAADDEEGGGG"
STANDARD = {a + b + c: _STANDARD_AA[16 * i + 4 * j + k]
            for i, a in enumerate(_BASES) for j, b in enumerate(_BASES)
            for k, c in enumerate(_BASES)}

GENETIC_CODES = {
    # code: (name, changes from the standard code, start codons)
    2: ("vertebrate mitochondrial",
        {"AGA": "*", "AGG": "*", "ATA": "M", "TGA": "W"},
        {"ATT", "ATC", "ATA", "ATG", "GTG"}),
    4: ("mold, protozoan and coelenterate mitochondrial",
        {"TGA": "W"},
        {"TTA", "TTG", "CTG", "ATT", "ATC", "ATA", "ATG", "GTG"}),
    5: ("invertebrate mitochondrial",
        {"AGA": "S", "AGG": "S", "ATA": "M", "TGA": "W"},
        {"TTG", "ATT", "ATC", "ATA", "ATG", "GTG"}),
    9: ("echinoderm and flatworm mitochondrial",
        {"AAA": "N", "AGA": "S", "AGG": "S", "TGA": "W"},
        {"ATG", "GTG"}),
    13: ("ascidian mitochondrial",
         {"AGA": "G", "AGG": "G", "ATA": "M", "TGA": "W"},
         {"TTG", "ATA", "ATG", "GTG"}),
    14: ("alternative flatworm mitochondrial",
         {"AAA": "N", "AGA": "S", "AGG": "S", "TAA": "Y", "TGA": "W"},
         {"ATG"}),
}


def codon_table(code):
    table = dict(STANDARD)
    table.update(GENETIC_CODES[code][1])
    return table


# -----------------------------------------------------------------------------
# Small helpers
# -----------------------------------------------------------------------------

def _key(text):
    """Lower case, letters and digits only: 'Cyt-b ' -> 'cytb'."""
    return re.sub(r"[^a-z0-9]", "", str(text).lower())


def classify_gene(name):
    """
    Return (kind, symbol, product) for a gene name, kind being
    'CDS', 'tRNA', 'rRNA', 'control' or None (not recognised).
    """
    raw = str(name).strip()
    base = re.sub(r"\(.*?\)", "", raw).strip()        # trnK(ctt) -> trnK
    k = _key(base)

    if k in PCG_ALIASES:
        sym = PCG_ALIASES[k]
        return "CDS", sym, PCG_PRODUCTS[sym]
    if k in RRNA_ALIASES:
        return "rRNA", base, RRNA_ALIASES[k]
    if k in CONTROL_REGION_NAMES:
        return "control", base, "putative control region"

    # tRNAs: trnL1, trnL1(tag), trnC_0, trnC, tRNA-Leu, tRNA-Leu2 ...
    m = re.match(r"(?i)^trna[-_ ]?([a-z]{3})([12])?\b", base)
    if m and m.group(1).lower() in AA1:
        aa = AA1[m.group(1).lower()]
        return "tRNA", "trn" + aa + (m.group(2) or ""), "tRNA-" + AA3[aa]
    m = re.match(r"(?i)^trn[-_ ]?([a-z])([12])?(?:_\d+)?$", base)
    if m and m.group(1).upper() in AA3:
        aa = m.group(1).upper()
        return "tRNA", "trn" + aa + (m.group(2) or ""), "tRNA-" + AA3[aa]
    return None, raw, None


def clean_codon(value):
    """'TAb' -> 'TA', 'Tb' -> 'T', ' ata ' -> 'ATA', None/'' -> None."""
    if value is None:
        return None
    s = str(value).strip()
    if not s or s.lower() in ("nan", "none", "-", "na", "n/a"):
        return None
    stripped = re.sub(r"[a-z*\u2020\u2021]+$", "", s)  # footnote letters/marks
    s = stripped if stripped else s                     # 'taa' -> keep it
    s = s.upper().replace("U", "T")
    return s if re.fullmatch(r"[ACGT]{1,3}", s) else "?" + s


def clean_strand(value):
    s = str(value).strip().lower() if value is not None else ""
    if s in ("+", "plus", "j", "1", "+1", "forward", "f"):
        return "+"
    if s in ("-", "\u2212", "minus", "n", "-1", "reverse", "r"):
        return "-"
    if s in ("+/-", "-/+", "\u00b1", ".", ""):
        return "+"           # control region / unstranded: listed as +
    return None


def to_int(value):
    if value is None:
        return None
    try:
        f = float(str(value).strip().replace(",", ""))
    except ValueError:
        return None
    return int(f) if f == int(f) else None


def revcomp(seq):
    return seq.translate(str.maketrans("ACGTRYKMBDHVNacgtrykmbdhvn",
                                       "TGCAYRMKVHDBNtgcayrmkvhdbn"))[::-1]


# -----------------------------------------------------------------------------
# Reading the inputs
# -----------------------------------------------------------------------------

COLUMN_NAMES = {
    "gene": {"gene", "genes", "name", "feature", "locus"},
    "begin": {"begin", "from", "beg", "startposition", "startpos"},
    "end": {"end", "to", "endposition", "endpos", "stopposition", "stoppos"},
    "strand": {"strand", "str", "direction"},
    "start": {"start", "startcodon", "initiationcodon"},
    "stop": {"stop", "stopcodon", "terminationcodon"},
}


def _read_rows(path, sheet=None):
    ext = os.path.splitext(path)[1].lower()
    if ext in (".xlsx", ".xlsm"):
        try:
            import openpyxl
        except ImportError:
            raise SystemExit("Reading .xlsx files needs openpyxl: "
                             "pip install openpyxl  (or save the sheet as CSV)")
        wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
        if sheet is None:
            ws = wb.worksheets[0]
        elif sheet in wb.sheetnames:
            ws = wb[sheet]
        else:
            raise SystemExit("Sheet '{}' not found in {}. Sheets: {}".format(
                sheet, os.path.basename(path), ", ".join(wb.sheetnames)))
        rows = [list(r) for r in ws.iter_rows(values_only=True)]
        wb.close()
        return rows
    if ext == ".xls":
        raise SystemExit("Old .xls files are not supported: save the sheet "
                         "as .xlsx or CSV.")
    with open(path, newline="", encoding="utf-8-sig") as fh:
        text = fh.read()
    delim = "\t" if (ext in (".tsv", ".tab") or
                     text.split("\n", 1)[0].count("\t") >
                     text.split("\n", 1)[0].count(",")) else ","
    return [row for row in csv.reader(text.splitlines(), delimiter=delim)]


def read_gene_order(path, sheet=None):
    """Return a list of dicts with keys gene, begin, end, strand, start, stop, row."""
    rows = _read_rows(path, sheet)
    header_i, cols = None, {}
    for i, row in enumerate(rows[:15]):
        keys = [_key(c) if c is not None else "" for c in row]
        if any(k in COLUMN_NAMES["gene"] for k in keys):
            header_i = i
            for field, names in COLUMN_NAMES.items():
                for j, k in enumerate(keys):
                    # footnote letters on headers: 'Starta' -> also try 'start'
                    if k in names or (k[-1:] in "abcd" and k[:-1] in names):
                        cols.setdefault(field, j)
            break
    if header_i is None:
        raise SystemExit("No header row with a 'Gene' column found in {}"
                         .format(os.path.basename(path)))
    missing = [f for f in ("gene", "begin", "end", "strand") if f not in cols]
    if missing:
        raise SystemExit("Column(s) not found in {}: {}. Header row is: {}".format(
            os.path.basename(path), ", ".join(missing), rows[header_i]))

    def cell(row, field):
        j = cols.get(field)
        return row[j] if j is not None and j < len(row) else None

    genes = []
    for n, row in enumerate(rows[header_i + 1:], start=header_i + 2):
        gene = cell(row, "gene")
        if gene is None or str(gene).strip() == "":
            continue
        genes.append({"row": n, "gene": str(gene).strip(),
                      "begin": cell(row, "begin"), "end": cell(row, "end"),
                      "strand": cell(row, "strand"),
                      "start": cell(row, "start"), "stop": cell(row, "stop")})
    return genes


def read_fasta(path):
    """Return (sequence ID, sequence) of the first record; warn if more than one."""
    seq_id, parts, n = None, [], 0
    with open(path, encoding="utf-8-sig") as fh:
        for line in fh:
            line = line.strip()
            if not line:
                continue
            if line.startswith(">"):
                n += 1
                if n == 1:
                    seq_id = line[1:].split()[0] if line[1:].split() else ""
                continue
            if n == 1:
                parts.append(line)
    if n == 0:
        raise SystemExit("No FASTA header ('>') found in {}".format(path))
    if n > 1:
        print("  WARNING: {} holds {} records; only the first ({}) is used."
              .format(os.path.basename(path), n, seq_id))
    return seq_id, re.sub(r"\s", "", "".join(parts)).upper()


# -----------------------------------------------------------------------------
# Locations
# -----------------------------------------------------------------------------

def intervals(begin, end, strand, length):
    """
    Feature intervals in the order the feature table needs them (5' -> 3' of
    the feature). Begin > End means the feature crosses the origin.
    """
    if begin <= end:
        return [(begin, end)] if strand == "+" else [(end, begin)]
    if length is None:
        raise ValueError("crosses the origin (Begin {} > End {}) but the genome "
                         "length is unknown: give FASTA_FILE or GENOME_LENGTH"
                         .format(begin, end))
    if strand == "+":
        return [(begin, length), (1, end)]
    return [(end, 1), (length, begin)]


def location_lines(ivals, key, partial5=False, partial3=False):
    lines = []
    for n, (a, b) in enumerate(ivals):
        a_s = ("<" if partial5 and n == 0 else "") + str(a)
        b_s = (">" if partial3 and n == len(ivals) - 1 else "") + str(b)
        lines.append("{}\t{}\t{}".format(a_s, b_s, key) if n == 0
                     else "{}\t{}".format(a_s, b_s))
    return lines


def feature_sequence(seq, ivals):
    """Sequence of a feature (5' -> 3') from its feature-table intervals."""
    out = []
    for a, b in ivals:
        out.append(seq[a - 1:b] if a <= b else revcomp(seq[b - 1:a]))
    return "".join(out)


def stop_except_location(begin, end, strand, n_bases, length):
    """transl_except position of an incomplete (1- or 2-base) stop codon."""
    if strand == "+":
        lo = end - n_bases + 1
        if lo < 1:
            return None
        return str(end) if n_bases == 1 else "{}..{}".format(lo, end)
    hi = begin + n_bases - 1
    if length is not None and hi > length:
        return None
    return ("complement({})".format(begin) if n_bases == 1
            else "complement({}..{})".format(begin, hi))


def feature_length(begin, end, length):
    return end - begin + 1 if begin <= end else (length - begin + 1) + end


# -----------------------------------------------------------------------------
# Main routine
# -----------------------------------------------------------------------------

def run(input_file, fasta_file=None, output_file=None, seq_id=None,
        genome_length=None, transl_table=5, sheet=None,
        partial_if_invalid_start=True):
    r"""
    Write the feature table and return its path. Can be called from the
    IDLE shell after one F5 run, e.g.
        >>> run(r"C:\data\GeneOrder_male.xlsx", fasta_file=r"C:\data\male.fasta")
    """
    print("{} v{} ({})".format(PROGRAM, VERSION, DATE))
    if transl_table not in GENETIC_CODES:
        raise SystemExit("Genetic code {} is not supported. Supported: {}".format(
            transl_table, ", ".join(str(c) for c in sorted(GENETIC_CODES))))
    code_name, _, starts = GENETIC_CODES[transl_table]
    table = codon_table(transl_table)

    seq = None
    if fasta_file:
        fasta_id, seq = read_fasta(fasta_file)
        if genome_length and genome_length != len(seq):
            print("  WARNING: GENOME_LENGTH {} differs from the FASTA length {}; "
                  "the FASTA length is used.".format(genome_length, len(seq)))
        genome_length = len(seq)
        if seq_id is None:
            seq_id = fasta_id
        elif seq_id != fasta_id:
            print("  WARNING: SEQ_ID '{}' differs from the FASTA ID '{}'. "
                  "table2asn matches them, so they must be identical."
                  .format(seq_id, fasta_id))
    if seq_id is None:
        seq_id = re.sub(r"\s+", "_", os.path.splitext(os.path.basename(input_file))[0])
        print("  NOTE: no FASTA given; sequence ID set to '{}'. It must match "
              "the ID in your FASTA header.".format(seq_id))
    if re.search(r"\s", seq_id):
        raise SystemExit("The sequence ID '{}' contains spaces. Use one word, "
                         "matching the FASTA header (e.g. Erigone_atra_male)."
                         .format(seq_id))
    if output_file is None:
        output_file = os.path.splitext(input_file)[0] + ".tbl"

    genes = read_gene_order(input_file, sheet)
    lines = [">Feature {}".format(seq_id)]
    warnings, counts, seen = [], {"CDS": 0, "tRNA": 0, "rRNA": 0, "control": 0}, {}

    def warn(g, msg):
        warnings.append("row {} ({}): {}".format(g["row"], g["gene"], msg))

    for g in genes:
        kind, symbol, product = classify_gene(g["gene"])
        if kind is None:
            warn(g, "gene name not recognised; row skipped")
            continue
        begin, end = to_int(g["begin"]), to_int(g["end"])
        strand = clean_strand(g["strand"])
        if begin is None or end is None:
            warn(g, "Begin/End missing or not a whole number; row skipped")
            continue
        if strand is None:
            warn(g, "strand '{}' not recognised; row skipped".format(g["strand"]))
            continue
        if genome_length and (max(begin, end) > genome_length or min(begin, end) < 1):
            warn(g, "coordinates {}-{} lie outside the genome (1-{}); row skipped"
                 .format(begin, end, genome_length))
            continue
        try:
            iv = intervals(begin, end, strand, genome_length)
        except ValueError as e:
            warn(g, str(e) + "; row skipped")
            continue

        counts[kind] += 1
        if kind in ("CDS", "rRNA"):
            seen[symbol] = seen.get(symbol, 0) + 1
            if seen[symbol] == 2:
                warn(g, "{} appears more than once".format(symbol))

        if kind == "CDS":
            start = clean_codon(g["start"])
            stop = clean_codon(g["stop"])
            notes, quals, partial5 = [], [], False
            n = feature_length(begin, end, genome_length)

            if start and start.startswith("?"):
                warn(g, "start codon '{}' not understood".format(start[1:]))
                start = None
            if stop and stop.startswith("?"):
                warn(g, "stop codon '{}' not understood".format(stop[1:]))
                stop = None
            if start:
                notes.append("Start codon " + start)
                if len(start) != 3:
                    warn(g, "start codon '{}' is not 3 bases".format(start))
                elif start not in starts:
                    if partial_if_invalid_start:
                        partial5 = True
                        warn(g, "{} is not a start codon in genetic code {}; "
                             "CDS marked 5' partial (<)".format(start, transl_table))
                    else:
                        warn(g, "{} is not a start codon in genetic code {}; "
                             "table2asn will report it".format(start, transl_table))
            else:
                warn(g, "no start codon given")

            incomplete = stop is not None and len(stop) < 3
            if stop:
                notes.append("Stop codon " + stop)
                if incomplete:
                    notes.append("TAA stop codon is completed by the addition "
                                 "of 3' A residues to the mRNA")
                    loc = stop_except_location(begin, end, strand, len(stop),
                                               genome_length)
                    if loc:
                        quals.append(("transl_except", "(pos:{},aa:TERM)".format(loc)))
                    else:
                        warn(g, "could not place transl_except across the origin")
                elif table.get(stop) != "*":
                    warn(g, "{} is not a stop codon in genetic code {}"
                         .format(stop, transl_table))
            else:
                warn(g, "no stop codon given")

            codons = n + (3 - len(stop) if incomplete else 0)
            if codons % 3:
                warn(g, "length {} bp is not a whole number of codons{}"
                     .format(n, " (with the incomplete stop completed)"
                             if incomplete else ""))

            if seq is not None:
                s = feature_sequence(seq, iv)
                if start and len(start) == 3 and s[:3] != start:
                    warn(g, "sequence starts with {}, table says {}"
                         .format(s[:3], start))
                if stop and s[-len(stop):] != stop:
                    warn(g, "sequence ends with {}, table says {}"
                         .format(s[-len(stop):], stop))
                # leave out the stop codon (the last whole codon if none given)
                body = s[:len(s) - (len(stop) if stop else 3 + len(s) % 3)]
                internal = [i // 3 + 1 for i in range(3, len(body) - 2, 3)
                            if table.get(body[i:i + 3]) == "*"]
                if internal:
                    warn(g, "internal stop codon(s) at codon {} (genetic code {})"
                         .format(", ".join(str(c) for c in internal[:5]), transl_table))

            lines += location_lines(iv, "gene", partial5)
            lines.append("\t\t\tgene\t" + symbol)
            lines += location_lines(iv, "CDS", partial5)
            lines.append("\t\t\tproduct\t" + product)
            lines.append("\t\t\ttransl_table\t{}".format(transl_table))
            for q, v in quals:
                lines.append("\t\t\t{}\t{}".format(q, v))
            if notes:
                lines.append("\t\t\tnote\t" + "; ".join(notes))

        elif kind == "tRNA":
            lines += location_lines(iv, "tRNA")
            lines.append("\t\t\tproduct\t" + product)

        elif kind == "rRNA":
            lines += location_lines(iv, "rRNA")
            lines.append("\t\t\tproduct\t" + product)
            lines.append("\t\t\tnote\t" + product.split()[0])

        else:  # control region
            lines += location_lines(intervals(begin, end, "+", genome_length),
                                    "misc_feature")
            lines.append("\t\t\tnote\tputative control region")

    with open(output_file, "w", encoding="utf-8", newline="\n") as fh:
        fh.write("\n".join(lines) + "\n")

    # --- report ------------------------------------------------------------
    print("  input : {}".format(input_file))
    if fasta_file:
        print("  FASTA : {}  ({} bp, ID {})".format(fasta_file, genome_length, seq_id))
    print("  code  : {} ({})".format(transl_table, code_name))
    print("  found : {} protein-coding, {} tRNA, {} rRNA, {} control region"
          .format(counts["CDS"], counts["tRNA"], counts["rRNA"], counts["control"]))
    expected = {"CDS": 13, "tRNA": 22, "rRNA": 2}
    odd = [k for k, v in expected.items() if counts[k] != v]
    if odd:
        print("  NOTE: a typical metazoan mitogenome has 13 protein-coding, 22 tRNA "
              "and 2 rRNA genes; check the {} count(s).".format(", ".join(odd)))
    if warnings:
        print("  Messages ({}):".format(len(warnings)))
        for w in warnings:
            print("    - " + w)
    if seq is None:
        print("  (Give FASTA_FILE to check start/stop codons and reading frames "
              "against the sequence.)")
    print("Done. Feature table written to: {}".format(os.path.abspath(output_file)))
    return output_file


# -----------------------------------------------------------------------------
# Entry points: IDLE (F5 with SETTINGS) and command line
# -----------------------------------------------------------------------------

def _ask(prompt, required=False):
    while True:
        ans = input(prompt).strip().strip('"').strip("'")
        if ans or not required:
            return ans or None
        print("  (required)")


def main(argv=None):
    argv = sys.argv[1:] if argv is None else argv
    if argv:
        import argparse
        p = argparse.ArgumentParser(
            prog=os.path.basename(sys.argv[0]),
            description="Write a GenBank five-column feature table (.tbl) for a "
                        "mitochondrial genome from a gene-order spreadsheet.")
        p.add_argument("gene_order", help="gene-order table (.xlsx, .csv, .tsv)")
        p.add_argument("--fasta", help="genome FASTA (recommended)")
        p.add_argument("--out", help="output .tbl (default: input name + .tbl)")
        p.add_argument("--seq-id", help="sequence ID (default: from the FASTA header)")
        p.add_argument("--length", type=int, help="genome length in bp (if no FASTA)")
        p.add_argument("--code", type=int, default=5,
                       help="NCBI genetic code (default 5, invertebrate mitochondrial)")
        p.add_argument("--sheet", help=".xlsx sheet name (default: first sheet)")
        p.add_argument("--no-partial-start", action="store_true",
                       help="do not mark CDSs with a non-standard start codon as 5' partial")
        p.add_argument("--version", action="version",
                       version="{} {} ({})".format(PROGRAM, VERSION, DATE))
        a = p.parse_args(argv)
        return run(a.gene_order, a.fasta, a.out, a.seq_id, a.length, a.code,
                   a.sheet, not a.no_partial_start)

    if INPUT_FILE is None:
        print("{} v{} ({})  -- SETTINGS not filled in; answer the questions below."
              .format(PROGRAM, VERSION, DATE))
        inp = _ask("Gene-order file (.xlsx/.csv): ", required=True)
        fasta = _ask("Genome FASTA file (recommended; Enter to skip): ")
        length = None
        if not fasta:
            ans = _ask("Genome length in bp (Enter if no gene crosses the origin): ")
            length = int(ans) if ans else None
        return run(inp, fasta, OUTPUT_FILE, SEQ_ID, length or GENOME_LENGTH,
                   TRANSL_TABLE, SHEET, PARTIAL_IF_INVALID_START)
    return run(INPUT_FILE, FASTA_FILE, OUTPUT_FILE, SEQ_ID, GENOME_LENGTH,
               TRANSL_TABLE, SHEET, PARTIAL_IF_INVALID_START)


if __name__ == "__main__":
    main()
