README_GenBank_Feature_Table_Formatter_v2.0.1_2026-11-20.txt
==============================================================
README for GenBank_Feature_Table_Formatter_v2.0.1_2026-11-20.py
(plain-text copy of README.md on GitHub)

GenBank Feature Table Formatter

[DOI]

[DOI]

Writes the GenBank five-column feature table (.tbl) for a mitochondrial genome
from a simple gene-order spreadsheet (one row per gene: Gene, Begin, End,
Strand, Start codon, Stop codon). The .tbl file and the genome FASTA are what
NCBI's table2asn (or BankIt) needs to make a GenBank submission. It is written
to run from the IDLE 3.13 shell, and also runs from the command line.

Typing a feature table by hand is slow and error-prone. Of the two published
Erigone atra mitogenomes used as the worked example, each GenBank record has
one tRNA coordinate that differs from the paper's gene table, and both list
the two halves of the origin-spanning tRNA-Cys in the wrong order (see Step
4). The program writes every feature from the same table you publish, and can
check every protein-coding gene against the sequence before you submit.

Current version: 2.0.1 (2026-11-20)  ·  Author: L. Brian Patrick, Department
of Biological Sciences, Dakota Wesleyan University (ORCID 0000-0003-1969-9600)

Contents

    mitogenome-feature-table/
    ├── GenBank_Feature_Table_Formatter_v2.0.1_2026-11-20.py           the program
    ├── README_GenBank_Feature_Table_Formatter_v2.0.1_2026-11-20.txt   plain-text manual (this file)
    ├── README.md                                                      the same manual for GitHub
    ├── CITATION.cff                                                   citation information
    ├── LICENSE
    └── example/                                                       complete worked example
        ├── input/     gene-order spreadsheets and genome FASTA files for two Erigone atra mitogenomes
        └── output/    the feature tables the program writes for them, and the shell messages

What it does

For each row of the gene-order table, the program writes the features GenBank
expects for a mitochondrial genome:

  Gene type             Features written   Qualifiers
  --------------------- ------------------ ---------------------------------------------------------------------------------------------------------------------
  Protein-coding (13)   gene and CDS       gene, product, transl_table, note (start and stop codons), and transl_except for an incomplete stop codon (T or TA)
  tRNA (22)             tRNA               product (e.g. tRNA-Leu)
  rRNA (2)              rRNA               product (16S ribosomal RNA, 12S ribosomal RNA), note (16S, 12S)
  Control region        misc_feature       note (putative control region)

It also:

-   writes minus-strand features with the larger coordinate first, as the
    format requires;
-   handles a feature that crosses the origin (Begin larger than End) on
    either strand, writing the two intervals in the correct 5′→3′ order;
-   writes the NCBI standard product names;
-   recognises gene names in the usual styles: COX1 or cox1, ND5 or nad5, CytB
    or cob, 16S or rrnL, trnL1 or trnL1(tag) (the MITOS2 style), and so on;
-   checks each start codon against the genetic code. A start codon that the
    code does not allow (e.g. TTA in code 5) is marked 5′ partial (<), which
    is how NCBI recorded cox1 in both Erigone atra records;
-   if you give the genome FASTA, checks every protein-coding gene: the start
    and stop codons in the table must match the sequence, the gene must be a
    whole number of codons, and there must be no internal stop codons.

It reads columns by their header names, so extra columns (Length, Intergenic
nucleotides, notes) can stay in the spreadsheet.

Requirements

-   Python 3.8 or later (written and tested for IDLE 3.13)
-   openpyxl, only to read .xlsx files. CSV and tab-delimited files need
    nothing beyond Python.

Install openpyxl for the same Python that IDLE uses:

    Windows:  py -3.13 -m pip install openpyxl
    macOS:    python3.13 -m pip install openpyxl
    Linux:    python3 -m pip install --user openpyxl

Download

Click Code → Download ZIP at the top of this page and unzip it, or clone it:

    git clone https://github.com/brpatric/mitogenome-feature-table.git

------------------------------------------------------------------------------

The input: a gene-order table

One mitogenome per file, one row per gene, with a header row. The program
reads the first sheet of an .xlsx file (or a CSV / tab-delimited text file).
Columns are matched by name, in any order and any capitalisation:

  Column   Content
  -------- --------------------------------------------------------------------------------------------------------------------------------
  Gene     gene name (see the list below)
  Begin    first position of the gene, counted along the + strand
  End      last position, counted along the + strand. For a gene that crosses the origin, Begin is larger than End (e.g. trnC 14448 to 5)
  Strand   + or - (J/N and plus/minus are also accepted). For the control region +/- or blank is fine
  Start    start codon (protein-coding genes only), e.g. ATA
  Stop     stop codon (protein-coding genes only): TAA, TAG, or an incomplete stop TA or T

For every gene, Begin is the lower and End the higher coordinate, whichever
strand it is on, as in the gene tables of most mitogenome papers. Footnote
letters after a codon are ignored (TAb is read as TA). Other columns are
ignored, so you do not need to delete them.

Accepted gene names (capitalisation, spaces and hyphens do not matter):

  Gene               Accepted names
  ------------------ ---------------------------------------------------------------------------------
  COX1, COX2, COX3   COX1 cox1 COI CO1 …
  ATP6, ATP8         ATP6 atp6 ATPase6 …
  CYTB               CYTB CytB cob Cyt b
  ND1–ND6, ND4L      ND1 nad1 NADH1 … ND4L nad4l
  tRNAs              trnK, trnL1, trnL2, trnS1, trnS2, trnK(ctt), trnC_0(gca), tRNA-Lys, tRNA-Leu2 …
  16S rRNA           16S rrnL rnl lrRNA
  12S rRNA           12S rrnS rns srRNA
  Control region     Control Region CR D-loop AT-rich region

A row with a name that is not recognised is skipped, with a message.

------------------------------------------------------------------------------

Worked example (all files included)

The example/ folder holds the two mitogenomes of the linyphiid spider Erigone
atra Blackwall, 1833 from Patrick and Santos (2025): a male, GenBank PP726906
(14,251 bp), and a female, GenBank PP726907 (14,504 bp). Running the program
on these files should reproduce example/output/ exactly, which is a quick way
to check your installation.

Step 1. The inputs

  File                                 Content
  ------------------------------------ -----------------------------------------------------
  GeneOrder_Erigone_atra_male.xlsx     gene order for the male, as in Table 1 of the paper
  GeneOrder_Erigone_atra_female.xlsx   gene order for the female
  Erigone_atra_male.fasta              male genome, from PP726906
  Erigone_atra_female.fasta            female genome, from PP726907

Each spreadsheet has the columns Gene, Begin, End, Strand, Length, Intergenic
nucleotides, Start and Stop. The program ignores the Length and Intergenic
nucleotides columns.

One coordinate in the female spreadsheet was corrected for this example. trnY
began at 14417 in the author's working file, but the published Table 1 and
GenBank both give 14407. That makes trnY 57 bp, the same length as in the
male.

The male gene-order table (the columns the program uses):

  Gene               Begin     End  Strand   Start   Stop
  ---------------- ------- ------- -------- ------- ------
  COX1                   1    1536    +       TTA    TAA
  COX2                1540    2205    +       TTG    TAA
  trnK                2201    2262    +             
  trnD                2246    2300    +             
  ATP8                2294    2449    +       ATT    TAA
  ATP6                2440    3108    +       ATG    TAA
  COX3                3113    3904    +       TTG    TAA
  trnG                3897    3949    +             
  ND3                 3945    4277    +       ATA    TAA
  trnL2               4274    4324    −             
  trnN                4325    4383    +             
  trnA                4428    4480    +             
  trnS1               4477    4527    +             
  trnR                4522    4582    +             
  trnE                4565    4621    +             
  trnF                4605    4658    −             
  ND5                 4658    6294    −       ATA     TA
  trnH                6288    6342    −             
  ND4                 6338    7615    −       TTG    TAA
  ND4L                7616    7874    −       ATT     T
  trnP                7870    7928    −             
  ND6                 7932    8363    +       TTG    TAA
  trnI                8359    8416    +             
  CytB                8414    9541    +       ATA    TAG
  trnS2               9541    9595    +             
  trnT                9596    9652    +             
  ND1                 9641   10558    −       ATT    TAG
  trnL1              10549   10607    −             
  16S                10608   11636    −             
  trnV               11624   11681    −             
  12S                11683   12356    −             
  trnQ               12356   12418    −             
  Control Region     12419   13111   +/−            
  trnM               13112   13171    +             
  ND2                13162   14120    +       ATA     TA
  trnW               14120   14174    +             
  trnY               14154   14210    −             
  trnC               14195       5    −             

TA and T are incomplete stop codons, completed by polyadenylation of the mRNA.
trnC crosses the origin: Begin 14195, End 5.

The FASTA files. The first word of the FASTA header is the sequence ID. It
must match the >Feature line of the feature table, and the program takes it
from the FASTA automatically. It must be one word with no spaces. The rest of
the header can carry table2asn source modifiers:

    >Erigone_atra_male [organism=Erigone atra] [location=mitochondrion] [mgcode=5] [topology=circular] [completeness=complete] ...

Step 2. Run the program

Pick whichever way suits you. All of them produce the same feature table.

A. In IDLE, with the SETTINGS block (main method)

Open GenBank_Feature_Table_Formatter_v2.0.1_2026-11-20.py in IDLE, edit the
SETTINGS block near the top, save, and press F5:

    INPUT_FILE    = r"C:\Users\me\mitogenome-feature-table\example\input\GeneOrder_Erigone_atra_male.xlsx"
    FASTA_FILE    = r"C:\Users\me\mitogenome-feature-table\example\input\Erigone_atra_male.fasta"
    OUTPUT_FILE   = None    # None = same name as INPUT_FILE, ending .tbl
    SEQ_ID        = None    # None = taken from the FASTA header
    GENOME_LENGTH = None    # only needed if there is no FASTA and a gene crosses the origin
    TRANSL_TABLE  = 5       # invertebrate mitochondrial code
    SHEET         = None    # None = first sheet

Keep the r before each opening quote on Windows. On a Mac a path looks like
r"/Users/me/mitogenome-feature-table/example/input/GeneOrder_Erigone_atra_male.xlsx".

B. In IDLE, answering questions

Leave INPUT_FILE = None, open the file in IDLE and press F5. The program asks
for the files in the shell. You can paste a path straight from Windows
Explorer (Shift + right-click → Copy as path); quotes around it are removed:

    GenBank_Feature_Table_Formatter v2.0.1 (2026-11-20)  -- SETTINGS not filled in; answer the questions below.
    Gene-order file (.xlsx/.csv): "C:\Users\me\data\GeneOrder_Erigone_atra_male.xlsx"
    Genome FASTA file (recommended; Enter to skip): "C:\Users\me\data\Erigone_atra_male.fasta"

C. From the IDLE shell prompt

After one F5 run (either way above), run() stays available at the >>> prompt,
so you can do more genomes without editing the file:

    >>> run(r"C:\data\GeneOrder_Erigone_atra_female.xlsx", fasta_file=r"C:\data\Erigone_atra_female.fasta")
    >>> run(r"C:\data\GeneOrder_new.csv", genome_length=14706, seq_id="Minicosa_neptuna_Mt180")

Keyword arguments: fasta_file, output_file, seq_id, genome_length,
transl_table (default 5), sheet, partial_if_invalid_start (default True).

D. From the command line (Windows PowerShell, macOS Terminal, Linux)

From inside the mitogenome-feature-table folder:

    python3 GenBank_Feature_Table_Formatter_v2.0.1_2026-11-20.py example/input/GeneOrder_Erigone_atra_male.xlsx --fasta example/input/Erigone_atra_male.fasta --out example/output/Erigone_atra_male.tbl

On Windows use py -3.13 in place of python3, and \ in place of / if you
prefer. Options:

  Option               Meaning
  -------------------- ---------------------------------------------------------------------------------
  --fasta FILE         genome FASTA (recommended): genome length, sequence ID and the gene checks
  --out FILE           output file (default: the input name ending .tbl)
  --seq-id ID          sequence ID for the >Feature line (default: first word of the FASTA header)
  --length BP          genome length, if no FASTA is given and a gene crosses the origin
  --code N             NCBI genetic code (default 5, invertebrate mitochondrial; also 2, 4, 9, 13, 14)
  --sheet NAME         .xlsx sheet to read (default: the first)
  --no-partial-start   do not mark CDSs with a non-standard start codon as 5′ partial
  --version, -h        version; help

With no arguments the program uses the SETTINGS block, exactly as in IDLE.

E. Many genomes at once (bash, e.g. on a Linux server)

The program needs no scheduler. If each genome has a matching pair of files
(GeneOrder_X.xlsx and X.fasta):

    for f in GeneOrder_*.xlsx; do
        name=${f#GeneOrder_}; name=${name%.xlsx}
        python3 GenBank_Feature_Table_Formatter_v2.0.1_2026-11-20.py "$f" --fasta "$name.fasta" --out "$name.tbl" > "$name.log"
    done

If openpyxl is not installed and you cannot install it
(python3 -m pip install --user openpyxl usually works without admin rights),
save the spreadsheets as CSV and use those instead.

F. From your own Python code

The file name has dots in it, so load it with runpy rather than import:

    import runpy
    fmt = runpy.run_path("GenBank_Feature_Table_Formatter_v2.0.1_2026-11-20.py")
    fmt["run"]("GeneOrder_Erigone_atra_male.xlsx", fasta_file="Erigone_atra_male.fasta")

What the shell shows

    GenBank_Feature_Table_Formatter v2.0.1 (2026-11-20)
      input : example/input/GeneOrder_Erigone_atra_male.xlsx
      FASTA : example/input/Erigone_atra_male.fasta  (14251 bp, ID Erigone_atra_male)
      code  : 5 (invertebrate mitochondrial)
      found : 13 protein-coding, 22 tRNA, 2 rRNA, 1 control region
      Messages (1):
        - row 2 (COX1): TTA is not a start codon in genetic code 5; CDS marked 5' partial (<)
    Done. Feature table written to: .../example/output/Erigone_atra_male.tbl

The one message is expected. TTA is not a start codon in the invertebrate
mitochondrial code, so cox1 is marked 5′ partial, as it is in the GenBank
record. The FASTA checks found no problems in any of the 13 protein-coding
genes of either genome.

Step 3. The output

example/output/Erigone_atra_male.tbl and Erigone_atra_female.tbl are
tab-delimited text. The start of the male table:

    >Feature Erigone_atra_male
    <1  1536    gene
                gene    COX1
    <1  1536    CDS
                product cytochrome c oxidase subunit I
                transl_table    5
                note    Start codon TTA; Stop codon TAA
    1540    2205    gene
                gene    COX2
    ...

A minus-strand protein-coding gene with an incomplete stop codon (larger
coordinate first; the transl_except places the stop on the last two bases):

    6294    4658    gene
                gene    ND5
    6294    4658    CDS
                product NADH dehydrogenase subunit 5
                transl_table    5
                transl_except   (pos:complement(4658..4659),aa:TERM)
                note    Start codon ATA; Stop codon TA; TAA stop codon is completed by the addition of 3' A residues to the mRNA

tRNA-Cys, which crosses the origin on the minus strand. Its 5′ end is at
position 5, so the interval 5→1 comes first, then 14251→14195:

    5   1   tRNA
    14251   14195
                product tRNA-Cys

This is the layout of NCBI's own example of a minus-strand gene that crosses
the origin, in the feature table documentation (NCBI, n.d.):

    500 1   gene
    15000   14200
                gene    nad5

The control region:

    12419   13111   misc_feature
                note    putative control region

The shell messages for each run are saved alongside as
Erigone_atra_male_log.txt and Erigone_atra_female_log.txt.

Step 4. Check against the GenBank records

Both feature tables were run through NCBI's tbl2asn, and the resulting records
were compared feature by feature with PP726906 and PP726907 as they stand in
GenBank (PP726906.2 and PP726907.2). Each record has 51 features. Everything
matched (feature types, products, gene names, notes, transl_table, the
5′-partial cox1, the transl_except qualifiers in the feature tables, and the
translations of all 13 proteins in both genomes) except the following:

  Record              Feature    GenBank                               Program (from the gene table)
  ------------------- ---------- ------------------------------------- -------------------------------------
  PP726906 (male)     tRNA-Pro   complement(7870..7925)                complement(7870..7928)
  PP726907 (female)   tRNA-Ile   8361..8414                            8361..8418
  both                tRNA-Cys   complement(join(1..5,14195..14251))   complement(join(14195..14251,1..5))

-   tRNA-Pro and tRNA-Ile. The program uses the coordinates of the published
    Table 1. The two genomes are identical over both tRNAs, so each tRNA
    should be the same length in both. Only the Table 1 coordinates (59 bp and
    58 bp) make that true. The GenBank values are 3 and 4 bp shorter in one
    specimen only. Both errors are already in the hand-made feature tables
    that were submitted (7925 and 8414), so they are typing errors, not
    changes made by NCBI.
-   tRNA-Cys. GenBank lists the two intervals in the wrong order. Read in that
    order, the tRNA runs from 14251 down to 14195, then jumps to 5 and runs
    down to 1. On the minus strand the gene actually starts at position 5,
    runs down to 1 and continues across the origin from 14251 down to 14195.
    The extracted sequence is the same 62 bases either way, but only the
    program's order gives them in the right sequence: GenBank's order moves
    the first five bases of the tRNA (AGTAA) to its end, which breaks the
    acceptor stem (1–2 of 7 base pairs instead of 5 of 7) and shifts the GCA
    anticodon out of the anticodon loop. The program's order follows NCBI's
    example for a minus-strand gene crossing the origin (Step 3). The reversed
    order also comes from the submitted feature tables.

What NCBI changed in the hand-made submission, and what the program now writes
itself. Comparing the submitted feature tables with the published records
shows the changes GenBank staff made at submission. The program writes each of
them directly:

  Submitted by hand                                                   In GenBank, and written by the program
  ------------------------------------------------------------------- ---------------------------------------------
  no transl_except for the incomplete stops of nad2, nad4l and nad5   transl_except (pos:…,aa:TERM) added
  cox1 at 1..1536 with a TTA start                                    marked 5′ partial: <1..1536
  rRNA products l-rRNA, s-rRNA                                        16S ribosomal RNA, 12S ribosomal RNA
  gene symbol CytB                                                    CYTB
  a note on every tRNA (trnK, trnL2 (CUN) …)                          removed (the product name is enough)
  start and stop codons in separate note lines                        one note: Start codon ATA; Stop codon TA; …
  Control Region in a fourth column of the misc_feature line          removed (the note carries it)

The genome sequences in the submitted FASTA files are identical to the GenBank
sequences, so the coordinates in example/ apply to both.

Apart from the version problem described in the next paragraph, tbl2asn
reported no errors. Its discrepancy report listed only items expected for a
mitochondrial genome: overlapping genes, CDSs overlapping tRNAs, tRNAs and
rRNAs without gene features, upper-case gene symbols (a rule for bacterial
genes), the partial cox1 in a complete sequence, and the assembly comment and
BioProject that are added at submission.

A note on validation software. The check used tbl2asn 25.7, the predecessor of
table2asn and the only NCBI validator that could be run where this program was
tested. That version cannot parse a two-base transl_except (e.g.
(pos:14119..14120,aa:TERM) for a TA stop), so it reports "Missing stop codon"
for nad2 and nad5. The one-base transl_except on nad4l was parsed correctly.
These qualifiers are written exactly as they appear in the accepted GenBank
records, so current table2asn should accept them. Please run table2asn on your
own submission (next section) and read its validation report.

------------------------------------------------------------------------------

Submitting to GenBank

With the FASTA file (.fsa or .fasta), the feature table (.tbl) and a
submission template (template.sbt, made at NCBI's Submission Template page),
run table2asn:

    table2asn -i Erigone_atra_male.fasta -f Erigone_atra_male.tbl -t template.sbt -a s -V vb -Z

NCBI's organelle submission guide lists the checks GenBank staff make on
mitochondrial genomes (gene counts, tRNA sizes and strands, gene features,
items to remove from annotation-software output); it is worth reading before
you submit.

-V vb validates and writes a GenBank flat file (.gbf) to look over. -Z writes
the discrepancy report. Read the .val file and fix every ERROR before you
submit the .sqn file. Source information (organism, [location=mitochondrion],
[mgcode=5], [topology=circular]) can go in the FASTA header as shown in Step
1, or with -j "[organism=...] ...".

Non-standard start codons. Marking a CDS 5′ partial is how the two Erigone
atra records were accepted with a TTA start for cox1, which sits at position 1
of the sequence. If such a gene lies in the middle of your sequence, table2asn
may also warn that the partial start is not at the end of the sequence. Use
--no-partial-start (or PARTIAL_IF_INVALID_START = False) if you would rather
keep the CDS complete and deal with the start-codon warning yourself.

Messages

  Message                                                                         Meaning and what to do
  ------------------------------------------------------------------------------- ------------------------------------------------------------------------------------------------------------
  X is not a start codon in genetic code 5; CDS marked 5' partial (<)             Expected for genes with non-standard starts (see above).
  sequence starts with ATT, table says ATA (or ends with)                         The coordinates or the codon in the table do not match the FASTA. Check Begin/End for an off-by-one error.
  length 919 bp is not a whole number of codons                                   Begin/End is off, or the stop codon should be an incomplete T or TA.
  internal stop codon(s) at codon N                                               Wrong reading frame or coordinates, or the wrong genetic code.
  TAG is not a stop codon in genetic code N                                       Typo in the Stop column, or the wrong genetic code.
  crosses the origin ... genome length is unknown                                 Give the FASTA (or GENOME_LENGTH).
  gene name not recognised; row skipped                                           Use one of the accepted names above.
  COX1 appears more than once                                                     Duplicated row.
  a typical metazoan mitogenome has 13 protein-coding, 22 tRNA and 2 rRNA genes   Shown only when the counts differ, e.g. a gene missing from the table.
  The sequence ID ... contains spaces                                             Use one word, the same as the first word of the FASTA header.

For example, with ND1's Begin mistyped as 9640 and ATP6's stop codon as TAG,
the FASTA check reports:

        - row 7 (ATP6): sequence ends with TAA, table says TAG
        - row 28 (ND1): length 919 bp is not a whole number of codons
        - row 28 (ND1): sequence ends with AGA, table says TAG

The program cannot check tRNA or rRNA boundaries against the sequence, so
check those coordinates against your annotation yourself.

Citing

If you use this program, please cite it as given in CITATION.cff (or with the
Cite this repository button on this page), giving the version used:

  Patrick, L. B. (2026). GenBank Feature Table Formatter: GenBank five-column
  feature tables for mitochondrial genomes (Version 2.0.1) [Computer
  software]. Zenodo. https://doi.org/10.5281/zenodo.24399061

Releases are archived on Zenodo (v2.0.1: 10.5281/zenodo.24399061; v2.0.0:
10.5281/zenodo.24386073); cite the version you used. If you use the worked
example, please also cite Patrick and Santos (2025).

References

Donath, A., Jühling, F., Al-Arab, M., Bernhart, S. H., Reinhardt, F., Stadler,
P. F., Middendorf, M., & Bernt, M. (2019). Improved annotation of
protein-coding genes boundaries in metazoan mitochondrial genomes. Nucleic
Acids Research, 47(20), 10543–10552. https://doi.org/10.1093/nar/gkz833
(MITOS2, whose gene names the program accepts)

National Center for Biotechnology Information. (n.d.). Feature table format
[Five-column feature table used by table2asn and BankIt].
https://www.ncbi.nlm.nih.gov/genbank/feature_table/

National Center for Biotechnology Information. (n.d.). Organelle genome
submissions. https://www.ncbi.nlm.nih.gov/genbank/organelle_submit/

National Center for Biotechnology Information. (n.d.). table2asn.
https://www.ncbi.nlm.nih.gov/genbank/table2asn/

National Center for Biotechnology Information. (n.d.). The genetic codes.
https://www.ncbi.nlm.nih.gov/Taxonomy/Utils/wprintgc.cgi

Ojala, D., Montoya, J., & Attardi, G. (1981). tRNA punctuation model of RNA
processing in human mitochondria. Nature, 290(5806), 470–474.
https://doi.org/10.1038/290470a0 (incomplete stop codons completed by
polyadenylation)

Patrick, L. B., & Santos, S. R. (2025). The complete mitochondrial genome of
Erigone atra Blackwall, 1833 (Araneae, Linyphiidae). Mitochondrial DNA Part B,
10(12), 1243–1248. https://doi.org/10.1080/23802359.2025.2594297 (source of
the worked example)

Python Software Foundation. (2024). Python (Version 3.13) [Computer software].
https://www.python.org/

Provenance and version history

Version 1.0.0 was written by the author with the assistance of Claude
(Anthropic). Version 2.0.0 was drafted with the assistance of Claude Opus 5.5
(Anthropic) on 2026-11-19, from specifications by the author, and reviewed by
the author; version 2.0.1 (2026-11-20) likewise. Version 2.0.0 was tested with
Python 3.13.18 and openpyxl 3.1.5 on the two Erigone atra mitogenomes in
example/. The output was compared feature by feature with GenBank PP726906 and
PP726907 and checked with NCBI tbl2asn 25.7 (Step 4). Origin-crossing features
were also tested on both strands by rotating the male genome so that cox1 (+)
and nad5 (−) cross the origin. All 13 proteins translated by tbl2asn were
identical to those in GenBank.

  Version   Date         Changes
  --------- ------------ ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
  1.0.0     (undated)    Original script (GenBank_Feature_Table_Formatter.py): a six-column GeneOrder Excel file to .tbl; parameters edited at the bottom of the file; needed pandas and openpyxl.
  2.0.0     2026-11-19   Rewrite. Columns found by header name (extra columns ignored; v1.0.0 stopped with an error on a spreadsheet with more than six columns). MITOS2-style gene names accepted. Features crossing the origin handled for every feature type and both strands, with the two intervals in the correct order (v1.0.0 reversed them for minus-strand tRNAs). transl_except added for incomplete stop codons. Footnote letters on codons ignored (v1.0.0 did not recognise TAb as an incomplete stop). Start codons checked against the genetic code. Optional FASTA check of every protein-coding gene. NCBI product names for rRNAs (16S/12S ribosomal RNA). Control region written as a valid misc_feature (v1.0.0 put its label in a fourth column). Warnings printed in the shell rather than written into the .tbl. Sequence ID checked for spaces. SETTINGS block at the top, questions in the shell if it is left empty, command-line options, and genetic codes 2, 4, 5, 9, 13 and 14. pandas no longer needed; CSV/TSV input accepted.
  2.0.1     2026-11-20   Example files only; the program code is unchanged and writes identical feature tables. The example gene-order spreadsheets now hold only the gene-order sheet (a sheet of working notes was removed), without the publication footnote letters (TAb, Tb) or cell highlighting. Zenodo DOI added.

Each version's file names carry its version number and date. Each release on
this page is tagged with its version.

License

Released under the MIT License; see LICENSE.
