# AnalyseIMSCC

A comprehensive Perl script for analysing and restructuring IMS Common Cartridge (`.imscc`) archive files exported from Learning Management Systems (LMS) such as **Schoology**, **Moodle**, **Canvas** and **Edmodo**.

---

## Table of Contents

1. [Overview](#overview)
2. [Features](#features)
3. [Prerequisites](#prerequisites)
4. [Installation](#installation)
5. [Word List Files](#word-list-files)
6. [Cover Files](#cover-files)
7. [Usage](#usage)
8. [Command-Line Options](#command-line-options)
9. [Operating Modes](#operating-modes)
10. [Document Classification (Subtypes)](#document-classification-subtypes)
11. [Output Files and Directories](#output-files-and-directories)
12. [Excel Workbook Columns](#excel-workbook-columns)
13. [Vocabulary Lists](#vocabulary-lists)
14. [Readability Metrics](#readability-metrics)
15. [Save Files (Caching)](#save-files-caching)
16. [Architecture](#architecture)
17. [Known Issues and Limitations](#known-issues-and-limitations)
18. [Reporting Issues](#reporting-issues)
19. [Licence](#licence)

---

## Overview

`AnalyseIMSCC` takes a course backup archive exported from an LMS and:

- Extracts and recreates the original folder hierarchy on the local filesystem
- Classifies each document by subtype (TEXT, TASKS, HANDOUT, ANSWERS, etc.) based on filename keywords
- Performs deep vocabulary analysis against standard EFL word lists (NGSL, NAWL, CEFR A1–C2)
- Calculates readability scores (Flesch, Flesch-Kincaid, Gunning Fog)
- Checks spelling and grammar using LanguageTool (via Python)
- Generates per-document Excel workbooks and an archive-wide inventory workbook
- Merges all PDFs within each subtype into a single booklet (with table of contents and page numbers)
- Optionally colour-highlights vocabulary within `.docx` source files

The name "AnalyseIMSCC" is somewhat historical — the script does far more than simply analyse; it acts as a full content processing pipeline for LMS course materials.

---

## Features

| Feature | Description |
|---|---|
| **Archive extraction** | Unzips `.imscc` files and reconstructs the original Schoology/Moodle folder tree |
| **Directory mode** | Processes a plain directory of files without needing an `.imscc` archive |
| **Single-file mode** | Analyses a single `.docx`, `.pdf`, `.jpg`, `.png` or `.gif` file |
| **Document classification** | Assigns subtypes (TEXT, TASKS, HANDOUT, ANSWERS, TAPESCRIPT, VOCABULARY, WORKSHEET, HOMEWORK, SONG, LESSON PLAN) based on filename |
| **Vocabulary analysis** | Matches words against NGSL (2801 words in 4 bands), NAWL, CEFR A1–C2, and the Pearson Academic Collocations List |
| **Readability** | Flesch Reading Ease, Flesch-Kincaid Grade Level, Gunning Fog Index |
| **N-gram analysis** | Extracts bi-grams, tri-grams and four-grams from TEXT documents |
| **Grammar checking** | Runs LanguageTool (via Python) to identify spelling/grammar issues |
| **PDF generation** | Converts `.docx`/`.pptx`/images to PDF using LibreOffice headless |
| **PDF merging** | Merges all PDFs per subtype into one booklet with table of contents and page numbers |
| **Template application** | Applies a LibreOffice `.ott` template to all `.docx` files |
| **Vocabulary highlighting** | Colour-codes vocabulary in `.docx` files (NGSL=red, NAWL=blue, ACL=purple) |
| **Excel reporting** | Generates a multi-sheet archive inventory workbook with conditional formatting |
| **Duplicate detection** | Identifies duplicate files using MD5 hash comparison |
| **Elective support** | Adds elective course code and name to the front cover |
| **Caching** | Saves generated PDFs and grammar reports to a `SaveFiles/` directory for reuse |

---

## Prerequisites

### Perl Modules

Install all required modules via CPAN:

```bash
cpan install XML::Simple Data::Dumper::Concise File::chdir File::Find::Rule \
  Archive::Zip Getopt::Long File::Basename File::Copy File::Slurp File::Compare \
  Math::Round Image::Size Time::Piece DateTime PDF::API2 PDF::Report PDF::Create \
  Excel::Writer::XLSX Time::HiRes Lingua::EN::Fathom Lingua::EN::Bigram \
  Lingua::StopWords Lingua::EN::Tagger Lingua::EN::Inflexion Lingua::Diversity::VOCD \
  Lingua::Diversity::MTLD File::Path Digest::MD5::File List::MoreUtils \
  Document::OOXML Log::Log4perl MP3::Info
```

### External Tools

| Tool | Purpose | Path expected |
|---|---|---|
| **LibreOffice** | Convert `.docx`/`.pptx` to PDF | `/Applications/LibreOffice.app/Contents/MacOS/soffice` |
| **docx2txt.pl** | Extract plain text from `.docx` | `/usr/local/bin/docx2txt.pl` |
| **Python 3.11** | Run grammar checker and PPTX extractor | `/usr/local/bin/python3.11` |
| **grammar_check.py** | LanguageTool wrapper (in working directory) | `./grammar_check.py` |
| **python_pptx.py** | Extract text from `.pptx` files (in working directory) | `./python_pptx.py` |
| **ConvertImageToJPG.pl** | Convert images to JPEG | `/usr/local/bin/ConvertImageToJPG.pl` |
| **GetXMLDataTest.pl** | Parse Schoology XML resource files | `/usr/local/bin/GetXMLDataTest.pl` |

**Note:** LibreOffice must be installed at the default macOS path. On Linux, adjust the path in `ConvertResourcesToPDF` and `ApplyTemplateToResources`.

### Python Dependencies

`grammar_check.py` requires the `language_tool_python` library:

```bash
pip3 install language-tool-python
```

`python_pptx.py` requires `python-pptx`:

```bash
pip3 install python-pptx
```

---

## Installation

1. Clone or download this repository:
   ```bash
   git clone https://github.com/yourusername/AnalyseIMSCC.git
   cd AnalyseIMSCC
   ```

2. Install all Perl dependencies (see [Prerequisites](#prerequisites) above).

3. Install all Python dependencies.

4. Ensure LibreOffice and `docx2txt.pl` are installed and accessible at the expected paths.

5. Create the `./WordLists/` directory and populate it with the required word list files (see [Word List Files](#word-list-files)).

6. Optionally create a `./Covers/` directory with PDF cover files (see [Cover Files](#cover-files)).

---

## Word List Files

The following files must be present in `./WordLists/`:

| Filename | Contents |
|---|---|
| `NGSL.txt` | New General Service List — one word per line, 2801 entries ranked by frequency |
| `NAWL.txt` | New Academic Word List — one word per line |
| `Supplemental.txt` | Supplemental word list for domain-specific terms |
| `dictionary.txt` | Custom dictionary of correctly-spelled technical terms |
| `AcademicCollocationList.csv` | Pearson Academic Collocations List (CSV with phrase and POS columns) |
| `CEFR-A1.txt` | CEFR A1 vocabulary list — one word/phrase per line |
| `CEFR-A2.txt` | CEFR A2 vocabulary list |
| `CEFR-B1.txt` | CEFR B1 vocabulary list |
| `CEFR-B2.txt` | CEFR B2 vocabulary list |
| `CEFR-C1.txt` | CEFR C1 vocabulary list |
| `CEFR-C2.txt` | CEFR C2 vocabulary list |

CEFR files support multi-word expressions (phrases) as well as single words.

---

## Cover Files

Place PDF cover files in `./Covers/`. The script will automatically match covers to archives based on the course name appearing in the cover filename.

| Convention | Example |
|---|---|
| Front cover | `Writing 3 Front Cover.pdf` |
| Back cover | `Writing Back Cover.pdf` |
| Generic front cover | `Generic Front Cover.pdf` (used if no course-specific cover is found) |
| Generic back cover | `Generic Back Cover.pdf` |
| Subtype-specific covers | `Reading 2 TASKS Front Cover.pdf` (copied only to the TASKS sub-directory) |

If present, front covers are copied with prefix `0000` and back covers with prefix `9999` to ensure they appear at the start and end of merged PDFs.

---

## Usage

### Basic (IMSCC archive)
```bash
perl AnalyseIMSCCv2.22.pl -file=MyCourse.imscc
```

### With PDF conversion
```bash
perl AnalyseIMSCCv2.22.pl -file=MyCourse.imscc -convert
```

### With vocabulary highlighting
```bash
perl AnalyseIMSCCv2.22.pl -file=MyCourse.imscc -convert -highlightvocab=NGSL,NAWL -highlightin=TEXT
```

### Directory mode
```bash
perl AnalyseIMSCCv2.22.pl -directory=/path/to/course/materials
```

### Single file mode
```bash
perl AnalyseIMSCCv2.22.pl -file=MyDocument.docx
```

### Elective course mode
```bash
perl AnalyseIMSCCv2.22.pl -file=Elective.imscc -cc=TAT001
```

---

## Command-Line Options

| Option | Type | Description |
|---|---|---|
| `-file=PATH` | Required* | Path to the `.imscc` archive or a single `.docx`/`.pdf`/image file |
| `-directory=PATH` | Required* | Path to a directory of files to process (alternative to `-file`) |
| `-loglevel=LEVEL` | Optional | Log verbosity: `DEBUG`, `INFO`, `WARNING`, `ERROR`, `CRITICAL` |
| `-highlightvocab=LIST` | Optional | Comma-separated vocab lists to highlight: `NGSL`, `NAWL`, `A1`–`C2`, `ACL` |
| `-highlightin=LIST` | Optional | Comma-separated subtypes to highlight in: `TEXT`, `TASKS`, `ANSWERS`, `HANDOUT` |
| `-addtemplate` | Optional | Apply a LibreOffice `.ott` template to documents. Optionally specify subtypes: `-addtemplate=TEXT,TASKS` |
| `-convert` | Optional | Convert documents to PDF. Optionally specify subtypes: `-convert=TEXT,TASKS` |
| `-cc=CODE` | Optional | Elective course code (e.g. `TAT001`). Triggers elective front-page handling |
| `-quick` | Flag | Skip vocabulary analysis and readability — compile PDFs as fast as possible |
| `-noimages` | Flag | Exclude image files (`.jpg`, `.png`, `.gif`) from the merged PDF |
| `-withanswers` | Flag | Include ANSWERS documents in the merged PDFs (appended at the back) |
| `-forceconvert` | Flag | Convert all documents regardless of the `@exclusions` list |
| `-addheader` | Flag | Add a custom header image (`./WIST.png`) to each page of the merged PDF |

\* Either `-file` or `-directory` must be provided.

---

## Operating Modes

### 1. IMSCC Archive Mode (`-file=course.imscc`)

1. Extracts the archive to `./<basename>/`
2. Parses `imsmanifest.xml` to reconstruct the folder hierarchy
3. Classifies each resource by subtype
4. Runs full analytics (unless `-quick`)
5. Generates the Excel workbook
6. Optionally applies templates and converts to PDF
7. Creates per-subtype merged PDFs with table of contents
8. Creates a master `ALL DOCUMENTS.pdf`

### 2. Directory Mode (`-directory=/path/to/files`)

Processes all files found under the specified directory tree. Folder names are used to determine course/unit context. Unit separator pages are generated for folders matching known course and unit names.

### 3. Single File Mode (`-file=document.docx`)

Analyses a single document. Generates the Excel workbook with vocabulary and readability data for that one file. PDF conversion and merging are available.

---

## Document Classification (Subtypes)

The script assigns a `subtype` to each document based on keywords in the filename. The priority order (top to bottom) is:

| Subtype | Filename must contain |
|---|---|
| `ANSWERS` | `ANSWER`, `TEACHER`, `SOLUTIONS` |
| `HANDOUT` | `HANDOUT` |
| `TAPESCRIPT` | `TAPESCRIPT`, `TRANSCRIPT` |
| `TASKS` | `TASK` |
| `TEXT` | `TEXT` |
| `SONG` | `SONG` |
| `LESSON PLAN` | `LESSON PLAN` |
| `WORKSHEET` | `WORKSHEET`, `ACTIVITY`, `HANDOUT`, `RESOURCE` |
| `VOCABULARY` | `VOCABULARY` |
| `HOMEWORK` | `HOMEWORK`, `VOCABULARY`, `EXAM` |
| `ANSWERS` | Re-checked last to catch e.g. "Homework Solutions" |
| `UNKNOWN` | `.docx`/`.doc`/`.pptx`/`.pdf` not matching any of the above |
| `EXCLUDED` | File > `$MAXIMUM_FILE_SIZE` (100MB), or filename/path contains exclusion keywords |
| `XLSX` | Excel workbooks generated by the script |
| `INCLUDED` | Files in always-include folders (copied to ALL subtype directories) |

**Notes:**
- Image files (`.jpg`, `.png`, `.gif`), audio (`.mp3`) and video (`.mp4`) do not receive a subtype.
- PDF files with ≥ 10 pages are automatically excluded (configurable via `$MAXIMUM_PDF_PAGES`).
- Folders listed in `@exclusions` (e.g. `Archive`, `EXCLUDE`) cause all contained files to be excluded.

---

## Output Files and Directories

After processing `MyCourse.imscc`, the script creates:

```
MyCourse/                           ← Reconstructed course folder hierarchy
MyCourse/ALLDOCUMENTS/              ← All non-excluded documents (with sequential numbering)
MyCourse/ALLDOCUMENTS/0000a Table of Contents.pdf
MyCourse/ALLDOCUMENTS/MyCourse ALL DOCUMENTS.pdf
MyCourse/TEXT/                      ← All TEXT documents
MyCourse/TEXT/MyCourse ALL TEXT.pdf
MyCourse/TASKS/                     ← All TASKS documents
MyCourse/TASKS/MyCourse ALL TASKS.pdf
MyCourse/HANDOUT/
MyCourse/ANSWERS/
MyCourse/TAPESCRIPT/
MyCourse/VOCABULARY/
MyCourse/WORKSHEET/
MyCourse/HOMEWORK/
MyCourse/SONG/
MyCourse/LESSON PLAN/
MyCourse/XLSX/                      ← Per-document analytics workbooks
MyCourse [v2.22].xlsx               ← Archive-wide Excel workbook (Inventory, Vocab, N-Grams, etc.)
MyCourse [v2.22].txt                ← Plain-text archive statistics
MyCourse.log                        ← Log file
SaveFiles/                          ← Cached PDFs and grammar reports (keyed by MD5)
```

Each document also gets:
- `<filename>.pdf` — PDF version (if converted)
- `<filename> [v2.22].xlsx` — Per-document vocabulary workbook (TEXT documents only)
- `<filename>.txt` — Plain text extracted by docx2txt
- `<filename>-PRETTY [v2.22].txt` — Cleaned plain text
- `<filename>-MARKED-UP [v2.22].html` — Colour-coded HTML with vocabulary highlighting
- `<filename>-SPELLING_AND_GRAMMAR.csv` — Grammar checker output
- `<filename>_ORIGINAL.docx` — Backup before template/style application

---

## Excel Workbook Columns

The main `Inventory` sheet contains ~75 columns per document:

| Column range | Contents |
|---|---|
| A–H | Archive name, ID, level, unit, parent folder, location, identifier |
| I–L | Title, web link, file type, subtype |
| M–O | Exists flag, size (KB), MD5 hash |
| P–S | Created by/date, last modified by/date |
| T–U | Revision number, total editing time |
| V–Z | Page count, paragraph count, line count, word count, character count |
| AA–AD | NGSL band counts (A/B/C/D) |
| AE–AH | NGSL total, NAWL total, new words, unknown words |
| AI–AL | NGSL %, NAWL %, new %, unknown % |
| AM–AO | Flesch, Flesch-Kincaid, Gunning Fog |
| AP | Academic collocations count |
| AQ–AU | Source, page size, orientation, page borders, lesson focus |
| AV–AX | Errant newlines, age in days, spelling/grammar error count |
| AY–BD | Spelling error categories (Typography, Punctuation, Typos, Grammar, Style, etc.) |
| BE–BJ | CEFR multiword expression counts per level (A1–C2) |
| BK–BP | CEFR single-word counts per level (A1–C2) |

Conditional formatting highlights:
- **Red**: old `.doc`/`.xls`/`.ppt` files, oversized files, duplicate MD5s, documents outside the expected length/readability range
- **Yellow**: borderline file size, slightly out-of-range length/readability
- **Green**: recently modified files (within 30 days), non-A4 page sizes
- **Blue gradient**: readability scores

---

## Vocabulary Lists

### NGSL (New General Service List)

2801 words divided into four frequency bands:

| Band | Rank range | Description |
|---|---|---|
| Band A | 1–800 | Core high-frequency vocabulary |
| Band B | 801–1600 | Mid-frequency vocabulary |
| Band C | 1601–2400 | Lower-frequency vocabulary |
| Band D | 2401–2801 | Least frequent NGSL words |

### NAWL (New Academic Word List)

Academic vocabulary commonly found in university-level texts.

### CEFR Vocabulary (A1–C2)

European Framework vocabulary lists for each proficiency level. Supports multi-word expressions (e.g. "in spite of").

### Academic Collocations List (ACL)

The Pearson Academic Collocations List: common two- and three-word academic phrases (e.g. "as a result", "play a role").

### Vocabulary Highlighting Colours

| List | Colour | Style |
|---|---|---|
| NAWL | Blue (`0000FF`) | Bold, italic, single underline |
| NGSL (target band) | Red (`FF0000`) | Bold, italic, single underline |
| Academic Collocations | Purple (`8B008B`) | Bold, italic, single underline |

---

## Readability Metrics

| Metric | Description | Tool |
|---|---|---|
| **Flesch Reading Ease** | 0–100, higher = easier. University level ≈ 30 | Lingua::EN::Fathom |
| **Flesch-Kincaid Grade Level** | US school grade equivalent. First-year undergrad ≈ 13.7 | Lingua::EN::Fathom |
| **Gunning Fog Index** | Years of education needed to understand the text | Lingua::EN::Fathom |

Expected Flesch-Kincaid ranges by course level:

| Level | CEFR | FK range |
|---|---|---|
| Reading 1 | A2 | 4–6 |
| Reading 2 | B1 | 6–8 |
| Reading 3 | B1+ | 8–10 |
| Reading 4 | B2 | 10–12 |

Documents outside the expected range are flagged in the Excel workbook.

---

## Save Files (Caching)

The `./SaveFiles/` directory stores generated PDF files and grammar-check CSV reports, named by the MD5 hash of the source `.docx` file.

On subsequent runs, if the source `.docx` has not changed (same MD5), the script reuses the cached PDF/CSV instead of re-running LibreOffice or LanguageTool. This dramatically reduces processing time for large archives.

```
SaveFiles/
  abc123def456...  .pdf    ← Cached PDF of a document
  abc123def456...  .csv    ← Cached grammar report
```

---

## Architecture

The script is structured as a main program block followed by ~54 subroutines:

```
main
 ├── GetCovers             Find front/back cover PDFs
 ├── MakeFolders           Create output directory structure
 ├── LaunchListener        (stub) Start unoconv background listener
 ├── Dive                  Recursively traverse manifest tree
 │    └── PrepareItems     Register each resource in @archivedetails
 ├── FindFileInSaveFiles   Restore cached PDFs/CSVs from SaveFiles/
 ├── ProcessItems          Run analytics on each item
 │    ├── GetOfficeMetadata    Parse .docx XML metadata
 │    ├── GetTagsAndText       Extract plain text
 │    ├── CheckSpellingAndGrammar  Run LanguageTool
 │    ├── GetLexisInformation  Vocabulary analysis
 │    ├── GetReadabilityStats  Flesch/FK/Gunning Fog
 │    ├── GetPDFMetadata       Extract PDF info
 │    └── StyleText            Highlight vocab in .docx
 ├── MakeDirectoryAndCopyFiles  Rebuild folder hierarchy
 ├── FarmResourcesOut           Copy files to subtype folders
 ├── SaveArchiveInventory       Write Excel Inventory sheet
 ├── SaveNGSLAndNAWLInfo        Write vocab worksheet
 ├── SaveCEFRVocabInfo          Write CEFR worksheet
 ├── SaveAcademicCollocationsInfo  Write ACL worksheet
 ├── GetAndSaveNGrams           Write N-Grams worksheet
 ├── SaveSpellingAndGrammarProblems  Write Grammar worksheet
 ├── SaveArchiveStatistics      Write .txt summary
 ├── CheckFilenamesForSpelling  Check filenames for typos
 ├── ApplyTemplateToResources   Apply .ott template via LibreOffice
 ├── ConvertResourcesToPDF      Convert to PDF via LibreOffice
 ├── CreateTableOfContents      Generate ToC file and PDF
 └── MergePDFs                  Merge per-subtype PDFs with page numbers
```

### Key Global Data Structures

- `@archivedetails` — Array of hash references; one entry per resource. Each hash has ~60+ keys including filename variants, metadata, vocabulary counts, readability scores, and spelling error counts.
- `%ngsl`, `%nawl`, `%cefr`, `%AcademicCollocationList` — Archive-wide vocabulary frequency hashes
- `@SpellingAndGrammarProblems` — Array of spelling/grammar error hashes across all documents
- `@tableofcontents`, `@excludedtableofcontents` — For building ToC and statistics report

---

## Known Issues and Limitations

- **macOS only** (as shipped): LibreOffice path, Python path, and several helper scripts reference macOS-specific locations. Linux/Windows users will need to update these paths in `ConvertResourcesToPDF`, `ApplyTemplateToResources`, and `ProcessItems`.

- **`unoconv` support removed**: The unoconv listener thread (`LaunchListener`) is present but the actual `unoconv` system calls are commented out. LibreOffice headless is used instead.

- **Lesson Focus extraction**: Reading lesson focus data from document headers stopped working after a macOS Big Sur upgrade; the relevant code is commented out in `CreateTableOfContents`.

- **`smartmatch` operator**: The script uses the `~~` smartmatch operator (suppressed with `no warnings 'experimental::smartmatch'`). This feature may be removed in future Perl versions.

- **Threads**: The threading model (`use threads`) spawns a single thread for the listener; there is no parallelism of document processing.

- **LanguageTool startup**: Grammar checking requires LanguageTool to be available to the Python script. On first run, it may download language models, causing a delay.

- **Large archives**: The script pauses at two points (after unzipping, and after enumerating items) with a `Press ENTER to continue (or Q to QUIT)` prompt, allowing manual file replacements or edits before processing begins.

---

## Reporting Issues

Issues can be reported at:
- GitHub Issues: https://github.com/ayliffe/AnalyseIMSCC/issues
- Email: github@ayliffe.com or ayliffe.david@gmail.com

---

## Licence

Open source. See individual module licences for their respective terms.

---

*AnalyseIMSCC v2.22 — by David Ayliffe*
