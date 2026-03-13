#!/usr/bin/perl

=head1 NAME

AnalyseIMSCC - Analyse and process IMS Common Cartridge (.imscc) archives

=head1 SYNOPSIS

    perl AnalyseIMSCCv2.22.pl [options] <archive.imscc>

    Options:
      --help        Show this help message
      --version     Show version number
      --directory   Path to directory containing .imscc files
      --file        Path to a single .imscc file

=head1 DESCRIPTION

AnalyseIMSCC unpacks IMS Common Cartridge archives and performs comprehensive
analysis of their contents. It extracts text from documents (DOCX, PDF, PPTX),
runs spelling and grammar checks via LanguageTool, computes readability
statistics (Flesch, Flesch-Kincaid, Gunning Fog), analyses vocabulary against
standard word lists (NGSL, NAWL, CEFR A1-C2 levels), generates Excel
workbooks summarising findings, produces marked-up HTML output, merges
resources into a single PDF, and creates a table of contents.

=head1 DEPENDENCIES

=over 4

=item * XML::Simple - parses imsmanifest.xml

=item * Excel::Writer::XLSX - writes .xlsx workbooks

=item * PDF::API2, PDF::Report, PDF::Create - PDF creation and manipulation

=item * Lingua::EN::Fathom - readability statistics

=item * Digest::MD5::File - MD5 hashing for file caching

=item * Log::Log4perl - logging framework

=item * Archive::Zip - unpacks .imscc archives

=item * File::Find::Rule, File::Basename, File::chdir - file utilities

=item * LibreOffice (headless) - converts DOCX/PPTX to PDF

=item * Python grammar_check.py + LanguageTool - spelling/grammar checking

=back

=head1 WORD LISTS

The script uses several vocabulary reference lists loaded at runtime:

=over 4

=item * NGSL (New General Service List) - ~2800 most frequent English words

=item * NAWL (New Academic Word List) - ~963 academic vocabulary items

=item * CEFR A1-C2 word lists - vocabulary banded by CEFR level

=item * Academic Collocations List - frequent academic phrase combinations

=item * Supplemental word list - project-specific additions

=item * Dictionary - full dictionary for spell-checking purposes

=back

=head1 AUTHOR

David Ayliffe

=head1 VERSION

v2.22

=cut

# https://github.com/unoconv/unoconv/issues/391
# Step3. Check unoconv is working
# unoconv --version
# Step4. Check LibreOffice is working
# /Applications/LibreOffice.app/Contents/MacOS/soffice --version

use utf8;           # allow UTF-8 literals in source
use strict;         # enforce variable declaration, etc.
use v5.10;          # require Perl 5.10+ for 'say', named captures, etc.
use warnings;       # enable runtime warnings
no warnings 'experimental::smartmatch';  # suppress smartmatch (~~) warnings; feature is used intentionally
use diagnostics;    # turn warnings into verbose diagnostic messages

$|++;               # disable STDOUT buffering so progress messages appear immediately

our $version = "v2.22";   # human-readable version string embedded in filenames and PDF metadata
our $VERSION_NO = 1;      # integer version number stamped onto the front cover of merged PDFs

use XML::Simple;
# use Image::Magick;
use Data::Dumper::Concise;
use File::chdir;
use File::Find::Rule;
use Cwd;
use Archive::Zip;
use threads;
use Getopt::Long;
use File::Basename;		#fileparse
use File::Copy;
use File::Slurp;
use File::Compare;
use Math::Round;
use Image::Size;
#use File::stat;	# get the created date of the archive (useful for when we come to add conditional formatting to the manifest)
use Time::localtime;
#use Time::Local;
use locale;
use Time::Piece;
use DateTime;
use PDF::API2;						# http://search.cpan.org/~ssimms/PDF-API2-2.033/lib/PDF/API2.pm
use PDF::Report;					# http://search.cpan.org/~teejay/PDF-Report-1.36/lib/PDF/Report.pm (allows us to see if a PDF is landscape or portrait. useful when adding page numbers)
use PDF::Create;
use POSIX qw( locale_h );
use Excel::Writer::XLSX;			# https://metacpan.org/pod/Excel::Writer::XLSX#write(-$row,-$column,-$token,-$format-)
use Time::HiRes qw( gettimeofday tv_interval );	# use this to show how long the script took to run
use Lingua::EN::Fathom;				# http://search.cpan.org/dist/Lingua-EN-Fathom/lib/Lingua/EN/Fathom.pm#READABILITY
use Lingua::EN::Bigram;				# http://search.cpan.org/dist/Lingua-EN-Bigram/lib/Lingua/EN/Bigram.pm
use Lingua::StopWords qw( getStopWords );
use Lingua::EN::Tagger;				# http://search.cpan.org/dist/Lingua-EN-Tagger/Tagger.pm
use Lingua::EN::Inflexion;			# http://search.cpan.org/~dconway/Lingua-EN-Inflexion-0.001006/lib/Lingua/EN/Inflexion.pm
#use Lingua::EN::CommonMistakes;		# https://metacpan.org/pod/Lingua::EN::CommonMistakes
#use Lingua::EN::Grammarian;
#use Lingua::TreeTagger;			# http://search.cpan.org/dist/Lingua-TreeTagger/lib/Lingua/TreeTagger.pm
use Lingua::Diversity::VOCD;		# http://textinspector.com/help/?page_id=20		http://search.cpan.org/~axanthos/Lingua-Diversity-0.07/lib/Lingua/Diversity/VOCD.pm
use Lingua::Diversity::Utils qw( split_text split_tagged_text );
use Lingua::Diversity::MTLD;		# http://textinspector.com/help/?page_id=20		http://search.cpan.org/~axanthos/Lingua-Diversity-0.06/lib/Lingua/Diversity/MTLD.pm
use File::Path qw( make_path );
use Digest::MD5::File qw( file_md5_hex );
use List::MoreUtils qw( uniq );
use Document::OOXML;
#use Image::Magick;
use Log::Log4perl;					# http://log4perl.sourceforge.net/releases/Log-Log4perl/docs/html/Log/Log4perl.html#e772f=
use MP3::Info;
#use Text::Hunspell;
#use Text::SpellChecker;			# brew install hunspell / brew install aspell	requires Text::Hunspell

my $hasPDFConverter = 1;		# flag: 1 = LibreOffice/soffice is available and can convert DOCX->PDF
my $hasTemplateConverter = 1;	# flag: 1 = LibreOffice template application (unoconv/soffice) is available
my $isUnoconvWorking = 0;		# flag: 1 = unoconv listener is running; 0 = fall back to direct soffice calls

# ─── Constants ───────────────────────────────────────────────────────────────

my $NGRAMS_LIMIT = 20;            # maximum number of n-gram entries written to the NGrams worksheets
my $MAXIMUM_FILE_SIZE = 100485760;# 100 MB hard limit; files larger than this are marked EXCLUDED
                                  # (100 MB = 1024 * 1024 * 100)
my $MAXIMUM_PDF_PAGES = 10;       # PDFs with >= this many pages are auto-excluded from the merged output
my $UPPER_LIMIT_PAGES = 5;        # the "ideal" page-count ceiling used for yellow/red highlighting in the
                                  # inventory spreadsheet (column V)
my $ONE_YEAR_OLD_FILE = 365;      # age threshold in days for red-highlighting stale documents
my $TWO_YEAR_OLD_FILE = 730;      # age threshold in days for very-stale documents
my $ONE_MONTH_OLD_FILE = 30;      # age threshold in days for green-highlighting recently edited documents
my $AVERAGE_READING_SPEED = 200;  # assumed words-per-minute used to estimate reading time

# NGSL band boundaries (rank numbers within the 2,801-item New General Service List)
# Band A = most frequent 800 words (Foundation level)
# Band B = words 801-1600 (Level 1)
# Band C = words 1601-2400 (Level 2)
# Band D = words 2401-2801 (Level 3)
my $NGSL_BAND_A_LOWER 	= 0;
my $NGSL_BAND_A_HIGHER 	= 800;
my $NGSL_BAND_B_LOWER 	= 800;
my $NGSL_BAND_B_HIGHER 	= 1600;
my $NGSL_BAND_C_LOWER 	= 1600;
my $NGSL_BAND_C_HIGHER 	= 2400;
my $NGSL_BAND_D_LOWER 	= 2400;
my $NGSL_BAND_D_HIGHER 	= 2801;

# Ideal word-count windows for each Reading programme level.
# Texts outside these bands are yellow-highlighted (borderline) or
# red-highlighted (significantly out of range) in the inventory spreadsheet.
my $READING_LEVEL_1_LOWER_TEXT_LIMIT 	= 300;		# Reading 1 (A2): target 300-599 words
my $READING_LEVEL_1_HIGHER_TEXT_LIMIT 	= 599;
my $READING_LEVEL_2_LOWER_TEXT_LIMIT 	= 500;		# Reading 2 (B1): target 500-899 words
my $READING_LEVEL_2_HIGHER_TEXT_LIMIT 	= 899;
my $READING_LEVEL_3_LOWER_TEXT_LIMIT 	= 800;		# Reading 3 (B1+): target 800-1199 words
my $READING_LEVEL_3_HIGHER_TEXT_LIMIT 	= 1199;
my $READING_LEVEL_4_LOWER_TEXT_LIMIT 	= 1100;		# Reading 4 (B2): target 1100-1800 words
my $READING_LEVEL_4_HIGHER_TEXT_LIMIT 	= 1800;
my $READING_LEVEL_5_LOWER_TEXT_LIMIT 	= 0;		# Reading 5: placeholder / no constraint currently
my $READING_LEVEL_5_HIGHER_TEXT_LIMIT 	= 99999;	# placeholder upper bound


# Flesch-Kincaid Grade Level target windows per Reading programme level.
# These thresholds drive the wordsdescription field (e.g. "EAAASY", "HAAAARD")
# which is appended to the custom PDF footer so teachers can see at a glance
# whether a text is pitched at the right difficulty.
# Reference scores: IELTS ~12.64; first-year undergrad ~13.66
# (source: http://www.eiken.or.jp/teap/group/pdf/teap_rlspecreview_report.pdf)
my $READING_LEVEL_1_LOWER_FLEISCH_KINCAID_LIMIT 	= 4;   # Reading 1 (A2):  FK 4–6
my $READING_LEVEL_1_HIGHER_FLEISCH_KINCAID_LIMIT 	= 6;
my $READING_LEVEL_2_LOWER_FLEISCH_KINCAID_LIMIT 	= 6;   # Reading 2 (B1):  FK 6–8
my $READING_LEVEL_2_HIGHER_FLEISCH_KINCAID_LIMIT 	= 8;
my $READING_LEVEL_3_LOWER_FLEISCH_KINCAID_LIMIT 	= 8;   # Reading 3 (B1+): FK 8–10
my $READING_LEVEL_3_HIGHER_FLEISCH_KINCAID_LIMIT 	= 10;
my $READING_LEVEL_4_LOWER_FLEISCH_KINCAID_LIMIT 	= 10;  # Reading 4 (B2):  FK 10–12
my $READING_LEVEL_4_HIGHER_FLEISCH_KINCAID_LIMIT 	= 12;


# ─── Command-line argument storage ──────────────────────────────────────────
# These variables are populated by GetOptions() below.

my $hasfilename = "";      # --file: path to a single .imscc or .docx file
my $hasdirectory = "";     # --directory: path to a directory of files
my $loglevel = "";         # --loglevel: Log4perl level (DEBUG|INFO|WARNING|ERROR|CRITICAL)
my $termno = undef;        # --termno: optional term number (not currently used in logic)

my $highlightvocab = undef;   # --highlightvocab: comma-separated list of vocab lists to highlight
my @highlightvocab;            # parsed array of vocab list names (e.g. 'NGSL', 'A2', 'ACL')

my $highlightin = undef;      # --highlightin: comma-separated list of subtypes to apply highlighting to
my @highlightin;               # parsed array of subtypes (e.g. 'TASKS', 'TEXT')

my $addtemplate = undef;      # --addtemplate: comma-separated list of subtypes to apply template to
my @addtemplate;               # parsed array; empty means apply to ALL subtypes

my $convert = undef;          # --convert: comma-separated list of subtypes to convert to PDF
my @convert;                   # parsed array; empty means convert ALL subtypes

my $quick = 0;          # --quick: skip vocabulary, readability, grammar analysis for speed
my $noimages = 0;       # --noimages: exclude jpg/png/gif files from the compiled PDF
my $withanswers = 0;    # --withanswers: also produce a "WITH ANSWERS" merged PDF
my $forceconvert = 0;   # --forceconvert: convert even files that would normally be excluded
my $addheader = 0;      # --addheader: stamp a school logo image into the header of each page

# ---------------------------------------------------------------------------
# Directory and path globals
# ---------------------------------------------------------------------------
# $level tracks the depth of the current node as Dive() recurses through the
# imsmanifest organisation tree.  It starts at 0 and is incremented/decremented
# by Dive() with each level of nesting.
my $level = 0;

# Working output directory; overwritten later once $zipbasename is known.
my $destinationDirectory = './temp/';

# Persistent cache directory.  Files are stored here keyed by their MD5 hash
# so that expensive operations (soffice PDF conversion, grammar checks, etc.)
# are not repeated on re-runs with the same source material.
my $savefileDirectory = './SaveFiles/';

# Absolute path of the directory from which the script was invoked; captured
# via getcwd() after argument parsing so that relative paths remain valid even
# if the working directory changes during processing.
my $HomeWorkingDirectory = "";

# Root output folder name derived from $zipbasename; used as the top-level
# directory created for the archive's output files.
my $RootDirectory = "";

# Stem of the input filename (no extension); used as a prefix for all output
# directories and filenames so that multiple archives can be processed in the
# same working directory without collision.
my $zipbasename = '';	# without extension

# Absolute path to the ALLDOCUMENTS sub-directory inside $destinationDirectory;
# every resource PDF produced during processing is copied here so that
# MergePDFs() can glob them into the final compiled output.
my $AllDocumentsDirectory = '';

# ---------------------------------------------------------------------------
# Template and header file paths
# ---------------------------------------------------------------------------
# OpenDocument Text (.ott) template applied by LibreOffice when converting
# DOCX files so that the output PDF inherits the school's house style.
my $templatefile = '/template_file.ott';

# Separate template used when the ANSWERS subtype is being converted, so that
# answer documents can carry a distinct visual style.
my $answerstemplatefile = '/answers_template_file.ott';

# PNG logo stamped into the page header of every PDF when --addheader is set.
my $header_image_file = './header_image_file.png';

# ---------------------------------------------------------------------------
# Excel workbook handle (Excel::Writer::XLSX)
# ---------------------------------------------------------------------------
# Holds the single .xlsx workbook created by SaveArchiveInventory() and
# populated by all the Save*() subroutines.  Declared here so it is
# accessible throughout the file's package scope.
my $workbook = undef;

# ---------------------------------------------------------------------------
# Accumulated plain-text buffers
# ---------------------------------------------------------------------------
# All readable text extracted from TEXT-subtype documents is concatenated into
# $AllTexts and written to a single .txt file at the end so that corpus-level
# statistics can be recalculated without re-parsing every DOCX.
my $AllTexts = '';
my $AllTasks = '';    # analogous buffer for TASKS-subtype documents
my $AllTextsFilename = '';   # path of the TEXT concatenation output file
my $AllTasksFilename = '';   # path of the TASKS concatenation output file

# ---------------------------------------------------------------------------
# Exclusion and classification arrays
# ---------------------------------------------------------------------------

# Any folder or resource whose name contains one of these strings is silently
# excluded from all processing and PDF compilation.  The check is a substring
# match performed in PrepareItems().  "EXCLUDE" is a catch-all marker that
# content authors can embed in a resource name to force exclusion.
my @exclusions = ( "Archive", "Other Practice Texts (and Tasks)", "EXCLUDE" );

# Exhaustive list of all recognised document subtypes.  The classification
# cascade in PrepareItems() assigns exactly one subtype to each resource; the
# result governs which Save*() analytics worksheets receive a row for the
# resource and which conditional-formatting band is applied.
# "EXCLUDED" and "INCLUDED" are synthetic subtypes injected by the exclusion/
# inclusion rules rather than inferred from the filename.
# "XLSX" is used for spreadsheet resources which require their own handling.
my @subtypes = ( "VOCABULARY", "WORKSHEET", "HOMEWORK", "TEXT", "TASKS", "HANDOUT", "ANSWERS", "TAPESCRIPT", "UNKNOWN", "LESSON PLAN", "EXCLUDED", "INCLUDED", "SONG", "XLSX" );

# Names of the vocabulary lists that can be requested via --highlightvocab.
# These map to the word-list preparation subroutines (PrepareNGSLWordList,
# PrepareNAWLWordList, PrepareCEFRWordList, PrepareAcademicCollocationsList)
# and to the highlighting logic in StyleText().
my @vocablists = ( "NGSL", "NAWL", "A1", "A2", "B1", "B2", "C1", "C2", "ACL" );

# ---------------------------------------------------------------------------
# Course-name arrays used by FindCourseForArchive() and Dive()
# ---------------------------------------------------------------------------
# @toplevels lists every known course (unit) name.  During Dive() traversal,
# when a folder title exactly matches one of these strings the script
# recognises that node as a "top-level" course boundary.  Two things happen:
#   1. $archive_root / $archive_root_code are set (used in all output file
#      names and as a label on cover pages).
#   2. A separator PDF is generated so that the compiled output clearly marks
#      where one course ends and the next begins.
# FindCourseForArchive() also scans this array to infer the course code
# (e.g., "R3", "S2") from the folder title when --cc is not supplied.
my @toplevels = (
"Year 07 - 1 - Clear messaging in digital media",
"Year 07 - 2 - Networks",
"Year 07 - 3 - Gaining support for a cause", 
"Year 07 - 4 - Programming in Scratch", 
"Year 07 - 5 - Spreadsheets",

"Year 08 - 1 - Vector Graphics",
"Year 08 - 2 - Layers of computing systems",
"Year 08 - 3 - Developing for the Web", 
"Year 08 - 4 - Representations (Clay to Silicon)",
"Year 08 - 5 - Mobile app development", 
"Year 08 - 6 - Introduction to Python", 

"Year 09 - 1 - Python Programming", 
"Year 09 - 2 - Media - Animations", 
"Year 09 - 3 - Data science",
"Year 09 - 4 - Representations - Going Audiovisual",  
"Year 09 - 5 - Introduction to Cybersecurity", 
"Year 09 - 6 - Physical Computing",

"Year 10 - 1 - Computer Systems",
"Year 10 - 2 - Algorithms",
"Year 10 - 3 - Data Representations",
"Year 10 - 4 - Programming",

"Year 10 - 9 - IT and the world of work",
"Year 10 - 9 - Media",
"Year 10 - 9 - Online Safety",
"Year 10 - 9 - Physical Computing",
"Year 10 - 9 - Spreadsheets",
"Year 10 - 9 - Using IT in project management",

"Year 11 - 1 - Impacts of Technology",
"Year 11 - 3 - Computer Networks",
"Year 11 - 4 - Databases and SQL",
"Year 11 - 5 - Network Security",
"Year 11 - 6 - Object-Oriented Programming" );

# @secondlevels lists sub-unit folder names that trigger a "unit break"
# separator page inside Dive() without altering the normal inclusion/exclusion
# rules.  This allows a course to be broken into numbered parts (e.g.
# programming topics) each receiving its own divider page in the merged PDF.
# Normal inclusion/exclusion rules still apply — no folder in @secondlevels is
# automatically included or excluded; only the separator page is added.
# NOTE: the previous set of names (Feeding Reading, Test Practice, etc.) is
# preserved in a comment above as a reference for future configuration changes.
# my @secondlevels = ( "Feeding Reading", "In-Class Test Practice Material", "Self-Study Test Practice Material", "Odd Numbered Terms", "Even Numbered Terms", "Odd Terms", "Even Terms" );
my @secondlevels = ( "Part 1 - Sequence", "Part 2 - Selection", "Part 3 - Iteration", "Part 4 - Subroutines", "Part 5 - Strings and lists", "Part 6 - Dictionaries and data files" );

# Any folder whose name appears in @alwaysinclude causes all resources inside
# it to be forcibly included in every PDF, overriding subtype-based exclusion.
# Currently empty — all inclusion/exclusion is driven by subtype and @exclusions.
my @alwaysinclude = ( );

# LanguageTool rule IDs that CheckSpellingAndGrammar() instructs the checker
# to ignore.  Rules listed here generate too many false positives for this
# corpus (e.g., WHITESPACE_RULE fires on double spaces between sentences which
# are common in EFL teaching materials; EN_QUOTES flags smart quotes; DASH_RULE
# flags en-dashes used stylistically).
my @spelling_rules_to_ignore = ( 'WHITESPACE_RULE', 'EN_QUOTES', 'DASH_RULE' );

# ---------------------------------------------------------------------------
# Archive-root tracking variables
# ---------------------------------------------------------------------------
# These three variables are set by FindCourseForArchive() / Dive() when a
# folder matching an entry in @toplevels is encountered, and are used across
# all output file names, cover-page labels, and the merged PDF metadata.

# Human-readable course name extracted from the matching @toplevels entry.
# e.g., "Reading 3" or "Writing - Advanced"
my $archive_root = '';

# Numeric depth at which the archive root was found in the manifest tree.
# -1 means "not yet found".  Used to guard against updating the archive root
# when a deeper nested folder name also matches.
my $archive_root_level = -1;

# Short alphanumeric code derived from the course name.
# e.g., "R3", "S2", "W5", "LA1" — used in output filenames and PDF footers.
my $archive_root_code = '';

# ---------------------------------------------------------------------------
# File-type counters  (used by IncrementFileTypeCount() and SaveArchiveStatistics())
# ---------------------------------------------------------------------------
# Each counter is incremented once per resource of that type encountered during
# Dive() / PrepareItems().  SaveArchiveStatistics() writes them to the Summary
# worksheet and to the log.
my $filecount = 0;           # total resources processed
my $docfiles = 0;            # legacy .doc files
my $docxfiles = 0;           # modern Word documents
my $xlsfiles = 0;            # Excel spreadsheets (.xls / .xlsx)
my $pdffiles = 0;            # pre-existing PDF resources
my $xmlfiles = 0;            # raw XML resources (e.g., Schoology quiz exports)
my $imagefiles = 0;          # images (.jpg, .png, .gif, .jpeg)
my $mp3files = 0;            # audio files
my $mp4files = 0;            # video files
my $pptfiles = 0;            # PowerPoint presentations (.ppt / .pptx)
my $txtfiles = 0;            # plain-text files
my $htmlfiles = 0;           # HTML resources
my $foldercount = 0;         # folders (organisation nodes) in the manifest
my $weblinks = 0;            # web-link resources (no local file)
my $schoologyresources = 0;  # Schoology-native resource types (LTI, etc.)

# Resources that could not be parsed (logged for diagnostics).
my @unabletoparseXML;    # resource IDs / titles that caused XML::Simple to die
my @unabletoparseDOCX;   # DOCX paths that caused Archive::Zip / XML errors

# ---------------------------------------------------------------------------
# Table-of-contents and hierarchy tracking
# ---------------------------------------------------------------------------
# Entries are pushed here by Dive() and later consumed by CreateTableOfContents().
my @tableofcontents;
my $toc_pages = 0;        # page count of the generated ToC PDF (used for page offset in MergePDFs)

# Resources explicitly excluded from the merged PDF still get an entry here
# so the inventory worksheet can show what was omitted.
my @excludedtableofcontents;

# Web-link resources collected during Dive() for the inventory worksheet.
my @weblinks;

# Stack of folder titles maintained by Dive() to track the path from the
# manifest root to the current node.  Used to build breadcrumb strings and
# to determine unit/section context for separator page generation.
my @hierarchy;

my @frontcovers;    # filenames of PDFs suitable for use as front cover pages
my @backcovers;     # filenames of PDFs suitable for use as back cover pages


# ---------------------------------------------------------------------------
# Vocabulary / word-list hashes   (archive-wide accumulators)
# ---------------------------------------------------------------------------
# Each of the following hashes is keyed by a word form and holds a nested
# hash of per-document frequency data.  They are populated incrementally by
# GetLexisInformation() as each document is processed and are written out by
# the Save*Info() subroutines at the end of the run.

# %ngsl  — words drawn from the New General Service List (NGSL, ~2,800 most
# frequent words in English).  Inner hash: { $word => { $docname => $count } }
my %ngsl;

# %nawl  — words drawn from the New Academic Word List (NAWL, ~963 words
# common in academic English but not in the NGSL).
my %nawl;

# %cefr  — words from the Cambridge English CEFR word lists (A1–C2 bands).
# The band (A1, B2, etc.) is stored as an attribute of each entry.
my %cefr;

# %grammarproblems — keyed by rule ID; counts how many times each LanguageTool
# rule fired across all documents.  Used by SaveSpellingAndGrammarProblems().
my %grammarproblems;

# %dictionary — full general-English dictionary loaded from $DICTIONARY_FILE;
# used by GetLexisInformation() to determine whether an unknown word is at
# least a real English word (rather than a proper noun or typo).
my %dictionary;

# %supplemental — domain-specific vocabulary from $SUPPLEMENTAL_FILE; treated
# as "known" words so they are not flagged as off-list during lexis analysis.
my %supplemental;

# %AcademicCollocationList — multi-word academic collocations loaded by
# PrepareAcademicCollocationsList() from the Pearson ACL CSV file.
# See: https://pearsonpte.com/organizations/researchers/academic-collocation-list/
my %AcademicCollocationList;

# ---------------------------------------------------------------------------
# Per-document reusable hashes (reset for each resource in GetLexisInformation)
# ---------------------------------------------------------------------------
# These parallel the archive-wide hashes above but are cleared before each
# document is analysed so they hold only that document's data.  Results are
# then merged up into the archive-wide hashes.
my %NGSL_document;
my %NAWL_document;
my %CEFR_document;
my %AcademicCollocationList_document;

# ---------------------------------------------------------------------------
# Word-level arrays (per-document; populated by GetLexisInformation)
# ---------------------------------------------------------------------------
# Every token extracted from the current document's plain text.
my @AllWords;

# Tokens that did not match any vocabulary list, dictionary, or supplemental
# list — potential misspellings, proper nouns, or genuinely new vocabulary.
my @UnknownWords;

# Tokens that are new to this document compared with all previously processed
# documents in the archive (i.e., appearing in @AllWords but not @OldWords).
my @NewWords;	# new and off-list

# N-gram arrays populated by GetAndSaveNGrams() for the current document.
my @TwoNgrams;
my @ThreeNgrams;
my @FourNgrams;

# Flat array of spelling/grammar problem strings returned by
# CheckSpellingAndGrammar() and later consolidated by
# ReadSpellingAndGrammarFileIntoArray().
my @SpellingAndGrammarProblems;

# ---------------------------------------------------------------------------
# Word-list file paths
# ---------------------------------------------------------------------------
# All word-list data files live under ./WordLists/ relative to the script.
# These constants are passed to the corresponding Prepare*() subroutines.

# Pearson Academic Collocation List — CSV format; each row is a multi-word
# phrase that is academically significant.
my $ACADEMIC_COLLOCATION_LIST_FILE = './WordLists/AcademicCollocationList.csv';

# New General Service List — plain text, one word per line.
my $NGSL_FILE = './WordLists/NGSL.txt';

# New Academic Word List — plain text, one word per line.
my $NAWL_FILE = './WordLists/NAWL.txt';

# Domain-specific supplemental vocabulary treated as "known" during lexis analysis.
my $SUPPLEMENTAL_FILE = './WordLists/Supplemental.txt';

# Comprehensive English dictionary; used to distinguish real (but off-list)
# words from likely misspellings.
my $DICTIONARY_FILE = './WordLists/dictionary.txt';

# ---------------------------------------------------------------------------
# Cover-page and blank-page resources
# ---------------------------------------------------------------------------
# PDF files used for front/back covers and blank writing pages are sourced
# from this directory.
my $COVERS_DIRECTORY = './Covers/';

# Pre-made blank writing page inserted by MergePDFs() at certain points in
# the compiled PDF.  The "4 SIDES" variant is a single sheet folded in half
# (landscape A3 → portrait A4 spreads).
my $BLANK_WRITING_PAGE = 'BLANK WRITING PAGE.pdf';              # 4 sides of paper

# 10-side variant used when more writing space is required.
my $BLANK_WRITING_PAGE_FOR_WRITING = 'BLANK WRITING PAGE FOR WRITING.pdf';  # 10 sides

# Number of blank writing page copies appended by MergePDFs().
my $BLANK_WRITING_PAGE_NO_OF_COPIES = 20;

# ---------------------------------------------------------------------------
# CEFR word-list file paths (one per band)
# ---------------------------------------------------------------------------
my $CEFR_A1_FILE = './WordLists/CEFR-A1.txt';
my $CEFR_A2_FILE = './WordLists/CEFR-A2.txt';
my $CEFR_B1_FILE = './WordLists/CEFR-B1.txt';
my $CEFR_B2_FILE = './WordLists/CEFR-B2.txt';
my $CEFR_C1_FILE = './WordLists/CEFR-C1.txt';
my $CEFR_C2_FILE = './WordLists/CEFR-C2.txt';

# ---------------------------------------------------------------------------
# Vocabulary supplement file handling
# ---------------------------------------------------------------------------
# A course may include one of several pre-made vocabulary PDFs (at Foundation,
# Level 1, 2, or 3).  These are searched for and, if found, appended to the
# compiled PDF by MergePDFs().
my @VOCAB_FILES = (	'Important Vocabulary At Foundation.pdf',
					'Important Vocabulary At Level 1.pdf',
					'Important Vocabulary At Level 2.pdf',
					'Important Vocabulary At Level 3.pdf' );

# Non-zero once a vocab file has been located and copied into ALLDOCUMENTS.
my $addedvocabfile = 0;

# Full ALLDOCUMENTS-relative path of the vocab file that was copied.
my $addedvocabfilename = 0;

# Page count of the vocab file; used so MergePDFs() can account for its pages
# in the running page-number offset.
my $addedvocabpages = 0;

# ---------------------------------------------------------------------------
# Cross-document word accumulator
# ---------------------------------------------------------------------------
# Before GetLexisInformation() analyses a document it copies @AllWords into
# @OldWords so that "new word" detection can compare the current document's
# tokens against all previously seen tokens in the archive.
my @OldWords;

# ---------------------------------------------------------------------------
# Elective course details
# ---------------------------------------------------------------------------
# When --cc is supplied on the command line the script treats the archive as
# an elective course.  GetElectiveName() looks up the course code in the CSV
# to get the full human-readable course name for cover pages.
my $elective_coursecode = '';    # short code, e.g., "ICT3"
my $elective_coursename = '';    # full name, e.g., "ICT - Advanced"
my $elective_filename = 'electives.csv';   # lookup table mapping codes to names
my @elective_details;            # parsed rows from electives.csv

# ---------------------------------------------------------------------------
# Processing-state globals
# ---------------------------------------------------------------------------
# Set by Dive() when the last page of long documents-only mode is reached.
my $lastpageoflongdocssonly = 0;

# Flag set within Dive() when at least one includable item has been found
# inside the current folder; used to suppress empty separator pages.
my $folderhasitems = 0;

# Title of the course (top-level folder) currently being processed.
my $currentcourse = "";

# Title of the unit (second-level folder) currently being processed.
my $currentunit = "";

# Running count of units encountered; drives unit-number labels on separator pages.
my $unitcount = 0;

# Set to 1 once the first unit has been entered, so that unit-break logic is
# not triggered prematurely at the very start of the archive.
my $hasunits = 0;

# Sequential counter incremented for each item pushed onto @archivedetails.
# Used to give each item a stable numeric ID for the inventory worksheet.
my $archiveitemcounter = 0;

# Guards against re-running one-time setup code inside the processing loop.
my $firsttime = 0;

# When non-zero, the next N pages of a document are skipped during page
# counting (used to ignore introductory pages that should not be paginated).
my $ignore_pages = 0;

# XML object returned by XML::Simple for the current manifest or document.
my $xml;

# String literal 'parent' used as a sentinel key when traversing the manifest
# tree to identify the parent node of the current item.
my $parent = 'parent';

# Current resource location (href) extracted from the manifest.
my $location = '';

# ---------------------------------------------------------------------------
# Master item array and mode arrays
# ---------------------------------------------------------------------------
# @archivedetails is the central data structure: an array of hash references,
# one per resource, built by PrepareItems() and Dive(), then read by
# ProcessItems() and all Save*() analytics subroutines.
my @archivedetails;

# File extensions that can be processed when the script is run in single-file
# mode (--file pointing directly at a resource rather than an .imscc archive).
my @individualfilemode = qw ( .doc .docx .jpg .png .gif .jpeg );

# All inflected forms of the verb "to be"; used by GetLexisInformation() to
# exclude copula forms from the lexical density calculation so they do not
# inflate the NGSL word count.
my @tobe = qw ( be is am are were was isn't wasn't weren't aren't );

# PDF metadata author string embedded in every PDF produced by the script.
my $authorinfo = sprintf "AnalyseIMSCC%s.pl with PDF::API2 by David Ayliffe", $version;

# ---------------------------------------------------------------------------
# Log::Log4perl configuration (inline string — no external config file needed)
# ---------------------------------------------------------------------------
# The configuration is passed to Log::Log4perl::init() as a reference to this
# string rather than as a file path, so the script is self-contained.
#
# Category "Foo::Bar" is the logger name used throughout the code via
#   $logger = Log::Log4perl::get_logger("Foo::Bar");
# Setting it to DEBUG means all levels (DEBUG, INFO, WARN, ERROR, FATAL) are
# forwarded to both appenders unless overridden by --loglevel.
#
# Two appenders are configured:
#
#   Logfile — writes to a file whose path is resolved at runtime by calling
#     GetLogFileName(), which returns a timestamped path in the archive's
#     output directory.  The PatternLayout ConversionPattern is:
#       [%d] [%M][%p] %m%n
#     where  %d = ISO date-time, %M = calling method name, %p = priority
#     (DEBUG/INFO/etc.), %m = log message, %n = newline.
#     This gives a full audit trail: timestamp, source subroutine, severity.
#
#   Screen — writes to STDOUT (stderr = 0).  The layout is just %m%n (message
#     plus newline) so the terminal output is clean and uncluttered.  Colour
#     coding is applied per level so important messages stand out:
#       INFO  = bright white   (normal progress)
#       WARN  = bright yellow  (non-fatal issues)
#       DEBUG = bright blue    (verbose detail, often suppressed)
#       ERROR = bright red     (problems that need attention)
my $logger;
my $conf = q(
    log4perl.category.Foo.Bar							= DEBUG, Logfile, Screen
    log4perl.appender.Logfile							= Log::Log4perl::Appender::File
    log4perl.appender.Logfile.filename					= sub { GetLogFileName(); }
    log4perl.appender.Logfile.layout					= Log::Log4perl::Layout::PatternLayout
    log4perl.appender.Logfile.layout.ConversionPattern	= [%d] [%M][%p] %m%n
    log4perl.appender.Screen							= Log::Log4perl::Appender::Screen
    log4perl.appender.Screen.stderr						= 0
    log4perl.appender.Screen.layout						= Log::Log4perl::Layout::PatternLayout
    log4perl.appender.Screen.layout.ConversionPattern	= %m%n
	log4perl.appender.Screen.color.info 				= bright_white
	log4perl.appender.Screen.color.warn 				= bright_yellow
	log4perl.appender.Screen.color.debug 				= bright_blue
	log4perl.appender.Screen.color.error 				= bright_red
);



my $usagestring = "
Usage is:  AnalyseIMSCC.pl -file=XXX.imscc|XXX.docx (-loglevel=XXX) (-addtemplate) (-convert) (-addheader)
						
	-file 			the name (and location) of the file containing the data.  Either .imscc or .docx
	-directory 		the location of the directory containing the data
	-loglevel		(optional)  choices are: 'DEBUG', 'INFO', 'WARNING', 'ERROR', 'CRITICAL'
	-highlightvocab		comma-separated list, default is ???	NGSL|NAWL|A1|A2|B1|B2|C1|C2|ACL
	-highlightin		comma-separated list, default is all	TEXTS|TASKS|ANSWERS|HANDOUTS
	-addtemplate		comma-separated list, default is all	(=ALL (default)|TEXTS|TASKS|ANSWERS|HANDOUTS)
	-convert		(=ALL (default)|TEXTS|TASKS|ANSWERS|HANDOUTS)
	-cc				an optional course code.  if this is specified it is assumed to be an elective and the elective front page is added
	-quick			compiles the documents into a PDF as quick as possible, does no vocabulary work
	-noimages		excludes jpg, png, gif files from the compiled PDF
	-forceconvert	converts all documents to PDF irrespective of the \@exclusions array
	-addheader	adds a custom header to each document
	\n\n";

# Capture the full command line before GetOptions consumes @ARGV; used for
# logging so the exact invocation can be reproduced from the log file.
my $commandline = $0 . " ". ( join " ", @ARGV );

# ---------------------------------------------------------------------------
# Command-line option parsing (Getopt::Long)
# ---------------------------------------------------------------------------
# '=s'  means the option requires a string value.
# ':s'  means the option accepts an optional string value (undef if absent).
# Flag options (no value suffix) are Boolean: present = 1, absent = 0.
#
# After GetOptions() returns, several options undergo post-processing:
#   - $addtemplate, $convert, $highlightin are split on ',' to populate
#     the corresponding arrays (@addtemplate, @convert, @highlightin).
#     An empty array means "apply to ALL subtypes".
#   - $highlightvocab is split to populate @highlightvocab.
#   - Each parsed value is validated against @subtypes / @vocablists;
#     unrecognised values cause logdie().
GetOptions(
    'file=s'			=> \$hasfilename,     # path to the .imscc archive or a single .docx/.jpg/etc.
    'directory=s'		=> \$hasdirectory,    # path to a directory of resources (directory mode)
    'loglevel:s'		=> \$loglevel,        # Log4perl level: DEBUG, INFO, WARN, ERROR, FATAL
    'termno:s'			=> \$termno,          # academic term number; used to filter by odd/even terms
    'highlightvocab:s'	=> \$highlightvocab,  # comma-separated vocab lists to highlight, e.g. "NGSL,ACL,B1"
    'highlightin:s'		=> \$highlightin,     # comma-separated subtypes to apply highlighting to
    'addtemplate:s'		=> \$addtemplate,     # comma-separated subtypes to apply the .ott template to
    'convert:s'			=> \$convert,         # comma-separated subtypes to convert to PDF via soffice
    'cc=s'				=> \$elective_coursecode,  # elective course code (e.g. "ICT3"); triggers elective front page
    'quick'				=> \$quick,           # skip vocabulary / readability / grammar for faster output
    'noimages'			=> \$noimages,        # exclude image files (.jpg, .png, .gif) from the compiled PDF
    'withanswers'		=> \$withanswers,     # also produce a second compiled PDF that includes ANSWERS resources
    'forceconvert'		=> \$forceconvert,    # convert all resources to PDF regardless of @exclusions
    'addheader'			=> \$addheader        # stamp a school logo header image onto every output page
) or die $usagestring;


# Record the wall-clock start time so a total elapsed duration can be logged
# at the end of the run.
my $timeStarted = [gettimeofday];

# Split the input file path into directory, base name, and extension.
# The regex qr/\.[^.]*/ matches the last dot-delimited extension, so
#   "/path/to/Reading 3.imscc" → $filename="Reading 3", $directories="/path/to/",
#   $suffix=".imscc"
# These three variables are used throughout the script to build output paths.
my ( $filename, $directories, $suffix ) = fileparse( $hasfilename, qr/\.[^.]*/ );

# Convert the base filename to Title Case so that directory names and PDF
# titles look polished even when the source filename is all-caps or mixed-case.
# The substitution matches each word token (including words with apostrophes)
# and applies \u (uppercase first char) \L (lowercase rest).
$filename =~ s/([\w']+)/\u\L$1/g;

# ---------------------------------------------------------------------------
# Guard: ensure the input is valid before doing any real work
# ---------------------------------------------------------------------------
# The script has two valid entry points:
#   1. A file that exists (-e) and whose extension is in @individualfilemode or ".imscc"
#   2. A directory path (-d $hasdirectory)
# If neither condition is met, print the usage string and exit implicitly.
unless ( ( -e $hasfilename && ( $suffix ~~ @individualfilemode || $suffix eq "\.imscc" ) ) || -d $hasdirectory )
{
     print $usagestring;
}
else
{
	# -----------------------------------------------------------------------
	# Initialise the logger now that we know the input is valid.
	# The Log4perl config is passed by reference so no config file is needed.
	# The logger name "Foo::Bar" matches the category in $conf.
	# The four banner lines at different levels let the operator verify that
	# colour coding is working correctly in their terminal.
	Log::Log4perl::init( \$conf );
	$logger = Log::Log4perl::get_logger( "Foo::Bar" );
	$logger->info ( "*" x 80 );
	$logger->info ( "*" x 80 );
	$logger->info ( "*" x 80 );

	$logger->debug ( "*" x 80 );    # blue  — visible only at DEBUG level
	$logger->info  ( "*" x 80 );    # white
	$logger->warn  ( "*" x 80 );    # yellow
	$logger->error ( "*" x 80 );    # red

	$logger->info ( sprintf "Command line was [%s]\n", $commandline );

	# -----------------------------------------------------------------------
	# Post-process --addtemplate
	# -----------------------------------------------------------------------
	# Normalise the alias "TEXTS" → "TEXT" (both spellings are accepted on the
	# command line for convenience).  Split on comma to get an array of subtypes.
	# Set $addtemplate = 1 (truthy "enabled") and validate each element
	# against the master @subtypes array to catch typos early.
	if ( defined $addtemplate )
	{
		$addtemplate =~ s/TEXTS/TEXT/g;    # normalise plural alias
		@addtemplate = split ( ',', $addtemplate );
		$addtemplate = 1;

		# if the user has specified subtypes to be converted, check that these are valid subtypes
		if ( scalar @addtemplate > 0 )
		{
			foreach my $totemplate ( @addtemplate )
			{
				$logger->logdie ( "Unknown subtype specified as template option: [$totemplate]\n" ) unless ( $totemplate ~~ @subtypes );
			}
		}
	}
	else
	{
		$addtemplate = 0;    # --addtemplate was not supplied; template application is disabled
	}

	# -----------------------------------------------------------------------
	# Post-process --convert
	# -----------------------------------------------------------------------
	# Same pattern as --addtemplate: normalise, split, validate.
	# $convert = 1 means "conversion is active for the subtypes in @convert".
	# An empty @convert means convert ALL subtypes.
	if ( defined $convert )
	{
		$convert =~ s/TEXTS/TEXT/g;
		@convert = split ( ',', $convert );
		$convert = 1;

		# if the user has specified subtypes to be converted, check that these are valid subtypes
		if ( scalar @convert > 0 )
		{
			foreach my $converter ( @convert )
			{
				$logger->logdie ( "Unknown subtype specified as convert option: [$converter]\n" ) unless ( $converter ~~ @subtypes );
			}
		}
	}
	else
	{
		$convert = 0;    # --convert was not supplied; conversion is disabled
	}

	# -----------------------------------------------------------------------
	# Post-process --highlightin  (e.g. TASKS, TEXT, HANDOUTS)
	# -----------------------------------------------------------------------
	# Determines which document subtypes receive vocabulary highlighting via
	# StyleText().  An empty @highlightin means highlight in ALL subtypes.
	if ( defined $highlightin )
	{
		$highlightin =~ s/TEXTS/TEXT/g;    # normalise plural alias
		@highlightin = split ( ',', $highlightin );
		$highlightin = 1;

		# if the user has specified subtypes to be highlighted, check that these are valid subtypes
		if ( scalar @highlightin > 0 )
		{
			foreach my $highlighter ( @highlightin )
			{
				$logger->logdie ( "Unknown subtype specified as highlighter option: [$highlighter]\n" ) unless ( $highlighter ~~ @subtypes );
			}
		}
	}
	else
	{
		$highlightin = 0;    # --highlightin was not supplied; highlighting is disabled
	}

	# -----------------------------------------------------------------------
	# Post-process --highlightvocab  (e.g. NGSL, NAWL, ACL)
	# -----------------------------------------------------------------------
	# Determines which vocabulary lists are used for in-document highlighting.
	# Validated against @vocablists (not @subtypes) since these are word lists,
	# not document subtypes.
	if ( defined $highlightvocab )
	{
		@highlightvocab = split ( ',', $highlightvocab );
		$highlightvocab = 1;

		# if the user has specified subtypes to be highlighted, check that these are valid subtypes
		if ( scalar @highlightvocab > 0 )
		{
			foreach my $highlighter ( @highlightvocab )
			{
				$logger->logdie ( "Unknown vocab list specified as highlighter option: [$highlighter]\n" ) unless ( $highlighter ~~ @vocablists );
			}
		}
	}
	else
	{
		$highlightvocab = 0;    # --highlightvocab was not supplied
	}

	# If --forceconvert was set, ensure $convert is truthy even if --convert
	# was not explicitly supplied, so that all resources are converted.
	$convert = 1 if $forceconvert == 1;
	
	# Log today's date (from DateTime::today) so the log file is self-timestamped
	# even when archived or copied to a different system.
	my $dt = DateTime->today;
	$logger->info ( sprintf "Today is: %s\n", $dt->date );

	# Log all runtime option values so any invocation can be exactly reproduced
	# from the log file.  Array-valued options are listed one-per-line indented.
	$logger->info ( sprintf "Version Number is 		[%s]\n", $VERSION_NO );
	$logger->info ( sprintf "File is 		[%s]\n", $hasfilename );
	$logger->info ( sprintf "Directory is 		[%s]\n", $hasdirectory );
	$logger->info ( sprintf "Course code is 		[%s]\n", $elective_coursecode );
	$logger->info ( sprintf "LogLevel is 		[%s]\n", $loglevel );
	$logger->info ( sprintf "Quick is 		[%s]\n", $quick );
	$logger->info ( sprintf "NoImages is 		[%s]\n", $noimages );
	$logger->info ( sprintf "WithAnswers is 		[%s]\n", $withanswers );

	$logger->info ( sprintf "AddTemplate is		[%s]\n", $addtemplate );
	# If specific subtypes were given, list them one per line for clarity
	$logger->info ( sprintf "	[%s]\n", join( "\n\t", @addtemplate ) ) if ( scalar( @addtemplate ) > 0 );

	$logger->info ( sprintf "HighlightVocab is	[%s]\n", $highlightvocab );  # vocab lists e.g. ACL, A2, NGSL
	$logger->info ( sprintf "	[%s]\n", join( "\n\t", @highlightvocab ) ) if ( scalar( @highlightvocab ) > 0 );

	$logger->info ( sprintf "HighlightIn is	[%s]\n", $highlightin );         # subtypes e.g. TASKS, HANDOUTS
	$logger->info ( sprintf "	[%s]\n", join( "\n\t", @highlightin ) ) if ( scalar( @highlightin ) > 0 );

	$logger->info ( sprintf "Convert is	 	[%s]\n", $convert );
	$logger->info ( sprintf "	[%s]\n", join( "\n\t", @convert ) ) if ( scalar( @convert ) > 0 );

	$logger->info ( sprintf "Electives code is	 	[%s]\n", $elective_coursecode );
	$logger->info ( sprintf "Quick is	 	[%s]\n", $quick );
	$logger->info ( sprintf "No Images is	 	[%s]\n", $noimages );
	$logger->info ( sprintf "With Answers is	 	[%s]\n", $withanswers );
	$logger->info ( sprintf "Force Convert is	 	[%s]\n", $forceconvert );

	$logger->info ( sprintf "Directory is [%s]; Filename is [%s]; Suffix is [%s]\n", $directories, $filename, $suffix );

	# -----------------------------------------------------------------------
	# Derive the output stem ($zipbasename) and set dependent path variables.
	# -----------------------------------------------------------------------
	# In archive mode $filename holds the .imscc stem; in directory mode it
	# is empty so we fall back to "temp".  The second assignment then overrides
	# with the directory name if --directory was given.
	$zipbasename = ( $filename ne '' ) ? $filename : 'temp';      # archive mode
	$zipbasename = ( $hasdirectory ne '' ) ? $hasdirectory . " OUTPUT" : 'temp';  # directory mode

	$logger->info ( sprintf "zipbasename is [%s]\n", $zipbasename );

	# All output files live under ./$zipbasename/ to keep a run's outputs grouped.
	$destinationDirectory = './'.$zipbasename.'/';

	# Remember the invoking directory so we can build absolute paths later.
	$HomeWorkingDirectory = getcwd;

	# Top-level folder name used by MakeFolders() to create the directory tree.
	$RootDirectory = $zipbasename.' Home Folder';

	# All individual resource PDFs are copied into ALLDOCUMENTS so MergePDFs()
	# can glob them in a single pass without having to recurse subtype folders.
	$AllDocumentsDirectory = $destinationDirectory.'ALLDOCUMENTS/';

	# Plain-text concatenation files used for corpus-level readability analysis.
	# Note: the directory is TEXT (singular), not TEXTS.
	$AllTextsFilename = $zipbasename.'/TEXT/ALL TEXTS.txt';
	$AllTasksFilename = $zipbasename.'/TASKS/ALL TASKS.txt';

	# Excel workbook — named with the archive stem and version so multiple runs
	# produce distinctly named files (e.g., "Reading 3 [v2.22].xlsx").
	my $ArchiveWorkbook = './'.$zipbasename." [".$version."].xlsx";

	# Plain-text statistics export file (tab-separated summary for pasting).
	my $exportTXTfile = './'.$zipbasename." [".$version."].txt";

	$logger->info ( sprintf "Archive Workbook is [%s]\n", $ArchiveWorkbook );
	$logger->info ( sprintf "Export Text file is [%s]\n", $exportTXTfile );
	
	if ( -e $hasfilename && $suffix eq "\.imscc" )
	{
		#printf "\nArchive creation time: %s\n", ctime( stat($hasfilename)->ctime );
		
		$logger->info ( sprintf "Unzipping archive %s...", $hasfilename );
		my $zip = Archive::Zip->new( $hasfilename );
		
		foreach my $member ($zip->members)
		{
   			next if $member->isDirectory;
    		#(my $extractName = $member->fileName) =~ s{.*/}{};
			my $extractName = $member->fileName;
			$member->extractToFileNamed( "$destinationDirectory/$extractName" );
		}

		if ( $elective_coursecode eq '' )
		{
			GetCovers( $filename );
		}
		else
		{
			GetElectiveName();
			GetCovers( 'Electives' );
		}

		MakeFolders();
		
		# create this listener nice and early - we might do conversions in-line
		# if we have unoconv we can do this:
		if ( ( $hasTemplateConverter == 1 || $hasPDFConverter == 1 ) && ( $addtemplate == 1 || $convert == 1 ) )
		{
			# give the worker thread a chance to start the unoconv process
			$logger->info ( "Starting the unoconv worker thread...\n" );
			my $thread = threads->create( \&LaunchListener );
			if ( defined $thread ) 
			{
				$logger->info ( "Thread creation successful!\n" );
			}
			else
			{
				$logger->info ( "Thread creation failed\n" );
			}
		}
		
		# we might want to replace a file with a newly-modified version of that file.
		# this saves having to upload onto Schoology, make a new archive and download the new archive 
		print "Press ENTER to continue (or Q to QUIT):\n";
		my $input = <STDIN>;
		chomp $input;
		die if ( $input eq 'q' or $input eq 'Q');
		
		# create object
		$xml = new XML::Simple;

		# read XML file
		$logger->info ( "Reading imsmanifest.xml..." );
		my $data = $xml->XMLin( $destinationDirectory."imsmanifest.xml", forcearray => 1 );

		Dive( $data->{organizations}[0]{organization}[0]{item} );	# this calls PrepareItems

		$logger->info ( sprintf "Finished enumerating items.  Archive has [%i] items\n", scalar @archivedetails );
		
		# this will try to find the spelling and grammar file we will shortly generate and copy it FROM the savefiles dir.
		# the generation of this file is time consuming so if the MD5s haven't changed, we can avoid having to do this
		# the file will be copied from the savefiles dir to our local dir
		FindFileInSaveFiles( "csv" );
		
		print "Press ENTER to continue (or Q to QUIT):\n";
		$input = <STDIN>;
		chomp $input;
		die if ( $input eq 'q' or $input eq 'Q');
		
		ProcessItems();

		MakeDirectoryAndCopyFiles();
		
		FarmResourcesOut();
				
		unlink $ArchiveWorkbook;	# delete this if it exists
		$workbook = Excel::Writer::XLSX->new( $ArchiveWorkbook );

		$logger->info ( sprintf "\n\nFinished processing items.  Saving details...\n" );
		
		SaveArchiveInventory( $ArchiveWorkbook, 'Inventory' );
		
		SaveAcademicCollocationsInfo ( undef, 'Academic Collocations' );
		
		SaveNGSLAndNAWLInfo( undef, 'NGSL+NAWL Vocab Info' );
		
		SaveCEFRVocabInfo( undef, 'CEFR Vocab' );
		
		GetAndSaveNGrams( undef );
		
		SaveSpellingAndGrammarProblems( undef, 'Grammar Problems' );
		
		SaveStrings();
		
		SaveArchiveStatistics( $exportTXTfile );
		
		CheckFilenamesForSpelling();
		
		$logger->info ( sprintf "Closing workbook %s...", $ArchiveWorkbook ); 
		$workbook->close() if ( defined $workbook );
		$logger->info ( "ok!\n" );
		
		# if we have unoconv we can do this:
		if ( $hasTemplateConverter == 1 && $addtemplate == 1 )
		{
			ApplyTemplateToResources ( "docx" );
				
			MakeDirectoryAndCopyFiles();
		
			FarmResourcesOut();
		}
			
		# if we have unoconv we can do this:
		if ( $hasPDFConverter == 1 && $convert == 1 )
		{
			FindFileInSaveFiles( "pdf" );
			
			ConvertResourcesToPDF ( "pdf" );
			
			MakeDirectoryAndCopyFiles();
		
			FarmResourcesOut();
		
			$toc_pages = CreateTableOfContents ( $AllDocumentsDirectory, $AllDocumentsDirectory, "*.pdf" );
		}

		# if we have unoconv we can do this:
		if ( $hasPDFConverter == 1 && $convert == 1 )
		{
			foreach my $sub ( @subtypes )
			{
				next if ( $sub eq 'EXCLUDED' || $sub eq 'INCLUDED' || $sub eq 'UNKNOWN' );
				next unless ( DoesArchiveHave( $sub ) == 1 );
				
				$toc_pages = CreateTableOfContents ( $sub, $destinationDirectory.$sub.'/', "*.pdf" );
				
				MergePDFs( 	$sub, 															# this subtype
							$destinationDirectory.$sub.'/', 									# the directory of the subtype	
							$destinationDirectory.$sub.'/'.$zipbasename.' ALL '.$sub.'.pdf',	# the destination filename
							$ignore_pages + $toc_pages,											# ??
							1,																	# ??
							'',																	# subtype to append to the end of the file
							'' );																# ignore this file when merging

				MergePDFs( 	$sub,
							$destinationDirectory.$sub.'/', 
							$destinationDirectory.$sub.'/'.$zipbasename.' ALL '.$sub.' WITH ANSWERS.pdf', 
							$ignore_pages + $toc_pages, 
							1,
							'ANSWERS',
							$destinationDirectory.$sub.'/'.$zipbasename.' ALL '.$sub.'.pdf' ) if ( $sub ne 'ANSWERS' && $withanswers == 1 );
				
				MakeFileListFile ( $sub );
			}

			# finally
			# create one big PDF for each folder
			# 			subtype			 directory, 		 	 outputfile, 											   ignore pages, 			start counting pages at
			MergePDFs( 	'ALL DOCUMENTS', $AllDocumentsDirectory, $AllDocumentsDirectory.$zipbasename.' ALL DOCUMENTS.pdf', $ignore_pages + $toc_pages, 1, '', '' );

			MergePDFs( 	'ALL DOCUMENTS', $AllDocumentsDirectory, $AllDocumentsDirectory.$zipbasename.' ALL DOCUMENTS WITH ANSWERS.pdf', $ignore_pages + $toc_pages, 1, 'ANSWERS', $AllDocumentsDirectory.$zipbasename.' ALL DOCUMENTS.pdf' ) if ( $withanswers == 1 );						
		}
	}
	elsif ( -e $hasfilename && $suffix ~~ @individualfilemode )
	{
		$logger->info ( "Running in single file mode\n" );

		FindCourseForArchive( $filename );
		
		PrepareItems ( $destinationDirectory, $hasfilename, undef );
		
		$logger->info ( sprintf "Finished enumerating items.  Archive has [%i] items\n", scalar @archivedetails );
		
		ProcessItems();
		
		ApplyTemplateToResources ( "docx" ) if ( $hasTemplateConverter == 1 && $addtemplate == 1 );
		
		ConvertResourcesToPDF ( "pdf" ) if ( $hasPDFConverter == 1 && $convert == 1 );
		
		unlink $ArchiveWorkbook;	# delete this if it exists
		$workbook = Excel::Writer::XLSX->new( $ArchiveWorkbook );
		
		$logger->info ( sprintf "\n\nFinished processing items.  Saving details...\n" );
		
		SaveArchiveInventory( $ArchiveWorkbook, 'Inventory' );
		
		# SaveArchiveStatistics( $exportTXTfile );	# if we are running in single-file mode, we don't need to do this

		SaveAcademicCollocationsInfo ( undef, 'Academic Collocations' );
		
		SaveNGSLAndNAWLInfo( undef, 'NGSL+NAWL Vocab Info' );
		
		SaveCEFRVocabInfo( undef, 'CEFR Vocab' );
		
		GetAndSaveNGrams( undef );
		
		SaveStrings();
		
		SaveSpellingAndGrammarProblems( undef, 'Grammar Problems' );
		
		CheckFilenamesForSpelling();
		
		$logger->info ( sprintf "Closing workbook %s...", $ArchiveWorkbook ); 
		$workbook->close() if ( defined $workbook );
		$logger->info ( "ok!\n" );
	}
	elsif ( -e $hasdirectory && -d $hasdirectory )
	{
		my @inputfiles = File::Find::Rule->in( $hasdirectory );

    	$logger->info ( sprintf "Running with directory switch: [%s].  Found [%i] files.\n", $hasdirectory, scalar @inputfiles );
				
		GetCovers( $hasdirectory );
		
		# create this listener nice and early - we might do conversions in-line
		# if we have unoconv we can do this:
		if ( ( $hasTemplateConverter == 1 || $hasPDFConverter == 1 ) && ( $addtemplate == 1 || $convert == 1 ) )
		{
			# give the worker thread a chance to start the unoconv process
			$logger->info ( "Starting the unoconv worker thread...\n" );
			my $thread = threads->create( \&LaunchListener );
		}

		MakeFolders();
		
		foreach my $file ( sort @inputfiles )
		{
			unless ( -d $file )		# unless it's a directory do this.  i.e. if it's a file do this:
			{
				my( $filename, $directories, $suffix ) = fileparse( $file, qr/\.[^.]*/ );
				$logger->info ( sprintf "\nWorking with file [%s] in [%s]...\n", $filename.$suffix, $directories );

				IncrementFileTypeCount ( $filename.$suffix );
				
				# location is reported in the XLSX report (not directory)
				$location = $directories;
				
				PrepareItems ( $directories, $file, undef );
				$archiveitemcounter++;
			}
			else
			{
				$logger->info ( sprintf "\n\nNew directory is [%s]...\n", $file );
				
				my @dirs = split /\//, $file;
				$level = () = $file =~ /\//g;  # count the number of / in the directory string.  e.g. [TestMaterial/General Interest Monologue/Grass Roofs] = 2
				
				# save this 'unit change page' to the 'All Documents' directory
				my $filename = sprintf( '%s%04db%i %s.pdf', $AllDocumentsDirectory, $archiveitemcounter, $level, $dirs[-1] );
				if ( $dirs[-1] ~~ @secondlevels || $dirs[-1] ~~ @toplevels )
				{
					$logger->info ( sprintf "Making new page marker called [%s]", $filename );
					SaveTextAsPDF( $filename, $dirs[-1], 0, 0 ) 
				}
				
				# save this 'unit change page' to all subtypes directory
				foreach my $sub ( @subtypes )
				{
					my $filename = sprintf( '%s%04db%i %s.pdf', $destinationDirectory.$sub.'/', $archiveitemcounter, $level, $dirs[-1] );
					if ( $dirs[-1] ~~ @secondlevels || $dirs[-1] ~~ @toplevels )
					{
						$logger->info ( sprintf "Making new page marker called [%s]", $filename );
						SaveTextAsPDF( $filename, $dirs[-1], 0, 0 )
					}
				}
			}
		}
		
		# this will try to find the spelling and grammar file we will shortly generate and copy it FROM the savefiles dir.
		# the generation of this file is time consuming so if the MD5s haven't changed, we can avoid having to do this
		# the file will be copied from the savefiles dir to our local dir
		FindFileInSaveFiles( "csv" );
		
		ProcessItems();
		
		unlink $ArchiveWorkbook;	# delete this if it exists
		$workbook = Excel::Writer::XLSX->new( $ArchiveWorkbook );

		$logger->info ( sprintf "\n\nFinished processing items.  Saving details...\n" );
		
		SaveArchiveInventory( $ArchiveWorkbook, 'Inventory' );	# does not currently work with '-directory' switch

		SaveArchiveStatistics( $exportTXTfile );
		
		SaveAcademicCollocationsInfo ( undef, 'Academic Collocations' );
		
		SaveNGSLAndNAWLInfo( undef, 'NGSL+NAWL Vocab Info' );
		
		SaveCEFRVocabInfo( undef, 'CEFR Vocab' );
		
		GetAndSaveNGrams( undef );
		
		SaveStrings();
		
		SaveSpellingAndGrammarProblems( undef, 'Grammar Problems' );
		
		CheckFilenamesForSpelling();
		
		$logger->info ( sprintf "Closing workbook %s...", $ArchiveWorkbook ); 
		$workbook->close() if ( defined $workbook );
		$logger->info ( "ok!\n" );
		
		ApplyTemplateToResources ( "docx" ) if ( $hasTemplateConverter == 1 && $addtemplate == 1 );
		
		# if we have unoconv we can do this:
		if ( $hasPDFConverter == 1 && $convert == 1 )
		{
			FindFileInSaveFiles( "pdf" );

			ConvertResourcesToPDF ( "pdf" );

			FarmResourcesOut();
		
			$toc_pages = CreateTableOfContents ( $AllDocumentsDirectory, $AllDocumentsDirectory, "*.pdf" );
		}
		
		foreach my $sub ( @subtypes )
		{
			next if ( $sub eq 'EXCLUDED' || $sub eq 'INCLUDED' || $sub eq 'UNKNOWN' );
			next unless ( DoesArchiveHave( $sub ) == 1 );
			
			$toc_pages = CreateTableOfContents ( $sub, $destinationDirectory.$sub.'/', "*.pdf" );
			
			MergePDFs( 	$sub, 																# this subtype
						$destinationDirectory.$sub.'/', 									# the directory of the subtype	
						$destinationDirectory.$sub.'/'.$zipbasename.' ALL '.$sub.'.pdf',	# the destination filename
						$ignore_pages + $toc_pages,											# ??
						1,																	# ??
						'',																	# subtype to append to the end of the file
						'' );																# ignore this file when merging

			MergePDFs( 	$sub,
						$destinationDirectory.$sub.'/', 
						$destinationDirectory.$sub.'/'.$zipbasename.' ALL '.$sub.' WITH ANSWERS.pdf', 
						$ignore_pages + $toc_pages, 
						1,
						'ANSWERS',
						$destinationDirectory.$sub.'/'.$zipbasename.' ALL '.$sub.'.pdf' ) if ( $sub ne 'ANSWERS' && $withanswers == 1 );
			
			MakeFileListFile ( $sub );
		}
		
		# create one big PDF for each folder
		# 			subtype			 directory, 		 	               outputfile, 											   ignore pages, 			   start counting pages at
		MergePDFs( 'ALL DOCUMENTS', $destinationDirectory.'ALLDOCUMENTS/', $destinationDirectory.'ALLDOCUMENTS/ALL DOCUMENTS.pdf', $ignore_pages + $toc_pages, 1, '', '' );
	}
	else
	{
		$logger->info ( sprintf "File or directory (%s or %s) does not exist.  Stopping.\n", $hasfilename, $hasdirectory );
	}
	
	my $rawseconds = tv_interval ( $timeStarted );
	
	#my $DD = int( $rawseconds / ( 24 * 60 * 60 ) );		# 24 * 60 * 60 = 86400
	my $HH = ( $rawseconds / ( 60 * 60 ) ) % 24;
	my $MM = ( $rawseconds / 60 ) % 60;
	my $SS = $rawseconds % 60;
	
	# make some noise to signal that we've finished.
	print "\a\a\a\a\a\a\a\a";
	
	$logger->info ( sprintf "\n\nTook %i raw seconds to complete.\n", $rawseconds );
	$logger->info ( sprintf "Took %02ih:%02im:%02is to complete.\n", $HH, $MM, $SS );
	$logger->info ( sprintf "Finished at %s\n", scalar localtime );	# display current date and time
}



# MakeFileListFile
# Creates a plain-text file listing all resources of a given subtype found in
# the archive. The list is written to the subtype subdirectory of the
# destination directory and is used for downstream processing.
sub MakeFileListFile
{
	my $subtype = shift;
	my $output_filename = $destinationDirectory.$subtype.'/'.'filelist-'.$subtype.'.txt'; 
	my $thisunit = "";
			
	open( my $fh, '>', $output_filename );
	$logger->info ( sprintf "Making filelist file in %s.\n", $output_filename );

	foreach my $item ( @archivedetails )
	{
		unless ( $item->{unit} eq $thisunit )
		{
			print $fh $item->{unit}."\n";
			$thisunit = $item->{unit};
		}
		
		if ( $item->{subtype} eq $subtype || $item->{subtype} eq 'INCLUDED' )
		{
			$item->{location} =~ s/-->|<--/\//g;
			my ( $filename, $directories, $suffix ) = fileparse( $item->{filename}, qr/\.[^.]*/ );
			my $string = $item->{location}.'/'.$filename.$suffix;
			$string =~ s/$zipbasename//g;
			
			print $fh "\t".$string."\n";
		}
	}
	
	close $fh;
	$logger->info ( "Done." );
	return;
}


# SaveSpellingAndGrammarProblems
# Writes spelling and grammar errors (held in @SpellingAndGrammarProblems) to
# a named worksheet in the output Excel workbook. Each row contains the
# resource filename, the flagged text, the suggested correction, and the
# rule/category that triggered the error.
sub SaveSpellingAndGrammarProblems
{
	my $keyhash = shift;
	my $sheetname = shift;	
	my $i = 0;
	my $record_count = scalar @SpellingAndGrammarProblems + 1;		# add one to this for the header row

	my $worksheet;
	my $bold = $workbook->add_format( bold => 1 );
	my $centre = $workbook->add_format( align => 'center' );
	
	$logger->info ( "Saving Spelling And Grammar Problems..." );
	
	if ( scalar @SpellingAndGrammarProblems > 0 )
	{
		$worksheet = $workbook->add_worksheet( $sheetname );
		$worksheet->set_tab_color( 'red' );
		$worksheet->freeze_panes( 1, 0 );    # Freeze the first row
		$worksheet->set_zoom( 150 );
		$worksheet->write_row( 0, 0, ['#', 'Filename', 'Rule ID', 'Message', 'Suggested Replacements', 'Context', 'Offset', 'Error Length', 'Category', 'Rule Issue Type'], $bold );

		$worksheet->set_column( 'A:A',  2 );
		$worksheet->set_column( 'B:B', 60 );
		$worksheet->set_column( 'C:C', 25 );
		$worksheet->set_column( 'D:D', 60 );
		$worksheet->set_column( 'E:E', 30 );
		$worksheet->set_column( 'F:F', 60 );
		$worksheet->set_column( 'G:G', 10 );
		$worksheet->set_column( 'H:H', 10 );
		$worksheet->set_column( 'I:I', 15 );
		$worksheet->set_column( 'J:J', 15 );
		
		my $red_format =    $workbook->add_format( bg_color => '#FFC7CE', color => '#9C0006' );
		my $yellow_format = $workbook->add_format( bg_color => '#FFEB9C', color => '#9C6500' );
		my $green_format =  $workbook->add_format( bg_color => '#C6EFCE', color => '#006100' );
		my $orange_format = $workbook->add_format( bg_color => '#fce5cd', color => '#080806' );
		my $pink_format =  	$workbook->add_format( bg_color => '#ead1dc', color => '#080806' );
		my $blue_format =  	$workbook->add_format( bg_color => '#cfe2f3', color => '#080806' );
	
		$worksheet->conditional_formatting( 'J2:J'.$record_count, { type => 'cell', criteria => 'equal to', value => 'typographical', format => $red_format } );
		$worksheet->conditional_formatting( 'J2:J'.$record_count, { type => 'cell', criteria => 'equal to', value => 'style', format => $yellow_format } );
		$worksheet->conditional_formatting( 'J2:J'.$record_count, { type => 'cell', criteria => 'equal to', value => 'duplication', format => $green_format } );
		$worksheet->conditional_formatting( 'J2:J'.$record_count, { type => 'cell', criteria => 'equal to', value => 'uncategorized', format => $orange_format } );
		$worksheet->conditional_formatting( 'J2:J'.$record_count, { type => 'cell', criteria => 'equal to', value => 'misspelling', format => $pink_format } );
		$worksheet->conditional_formatting( 'J2:J'.$record_count, { type => 'cell', criteria => 'equal to', value => 'grammar', format => $blue_format } );
		
		foreach my $row ( @SpellingAndGrammarProblems )
		{
			$i++;
			$worksheet->write_number( $i, 0, $i, $centre );		# row, column, token
			$worksheet->write_string( $i, 1, $row->{original_filename} );
			$worksheet->write_string( $i, 2, $row->{ruleId} );
			$worksheet->write_string( $i, 3, $row->{message} );
			$worksheet->write_string( $i, 4, $row->{replacements} );
			$worksheet->write_string( $i, 5, $row->{context} );
			$worksheet->write_number( $i, 6, $row->{offset}, $centre );
			$worksheet->write_number( $i, 7, $row->{errorLength}, $centre );
			$worksheet->write_string( $i, 8, $row->{category} );
			$worksheet->write_string( $i, 9, $row->{ruleIssueType} );
		}
	}
	
	return;
}



# GetLogFileName
# Returns the path to the Log4perl log file for this run, constructed from the
# destination directory path and the script version string.
sub GetLogFileName
{
	return ( $zipbasename eq "" ) ? "Logfile.log" : $zipbasename . ".log";
}



# MakeFolders
# Creates the full output directory structure beneath the destination directory.
# Subdirectories are created for each resource subtype (pdf, docx, pptx, etc.)
# as well as for merged output, covers, and temporary working files.
sub MakeFolders
{	
	my $cover = "";
	
	mkdir './' . $zipbasename if ( -e $hasdirectory && -d $hasdirectory );	# if we're running in directory mode, make a temporary home for our files
	
	mkdir $AllDocumentsDirectory or $logger->logdie( "Cannot make master directory $AllDocumentsDirectory: $!" );
	$logger->info ( sprintf "Created 'all documents' subdirectory %s ok!\n", $AllDocumentsDirectory );
	
	if ( -e $BLANK_WRITING_PAGE_FOR_WRITING && $archive_root_code =~ /W1|W2|W3|W4/ )
	{
		my $destination_filename = $AllDocumentsDirectory.'9998 '.$BLANK_WRITING_PAGE_FOR_WRITING; 
		copy $BLANK_WRITING_PAGE_FOR_WRITING, $destination_filename;
		$logger->info ( sprintf "	Copied the blank writing page (for Writing) to %s ok!\n", $destination_filename );
	}
	elsif ( -e $BLANK_WRITING_PAGE && $archive_root_code !~ /W1|W2|W3|W4/ )
	{
		my $destination_filename = $AllDocumentsDirectory.'9998 '.$BLANK_WRITING_PAGE; 
		copy $BLANK_WRITING_PAGE, $destination_filename;
		$logger->info ( sprintf "	Copied the blank writing page to %s ok!\n", $destination_filename );
	}
	else
	{
		$logger->info ( "	Could not copy the blank writing page.\n" );
	}
	
	#
	# AllDocumentsDirectory
	#
	
	if ( scalar @frontcovers > 0 )
	{
		foreach $cover ( @frontcovers )
		{
			my ( $cover_filename, $cover_directories, $cover_suffix ) = fileparse( $cover, qr/\.[^.]*/ );
			my $dont_copy = 0;
			
			foreach my $sub (@subtypes)
			{
				$dont_copy++ if ( index( lc($cover_filename), lc($sub) ) != -1 );
			}
			
			if ( $dont_copy > 0 )
			{
				my $destination_filename = $AllDocumentsDirectory.'0000 '.$cover_filename.$cover_suffix; 
				copy $cover, $destination_filename;
		
				# find out how many pages there are in this cover document.  there might be more than one page
				my $pdf = PDF::API2->open( $cover );
				$ignore_pages = $pdf->pages();
				$logger->info ( sprintf "	Copied front cover to [%s] ok!\n", $destination_filename );				
			}
			else
			{
				$logger->info ( sprintf "	Front cover [%s] belongs to a subtype.  Skipping\n", $cover_filename );								
			}
		}
	}
	
	#
	#
	
	if ( scalar @backcovers > 0 )
	{
		foreach $cover ( @backcovers )
		{
			my ( $cover_filename, $cover_directories, $cover_suffix ) = fileparse( $cover, qr/\.[^.]*/ );		# take the last cover we found.  hit and miss strategy.
			my $destination_filename = $AllDocumentsDirectory.'9999 '.$cover_filename.$cover_suffix; 
			copy $cover, $destination_filename;
			$logger->info ( printf "	Copied back cover to %s ok!\n", $destination_filename );
		}	
	}

	#
	# set up the subtype subdirectories
	#
	
	foreach my $sub ( @subtypes )
	{
		my $subdirectory = $destinationDirectory.$sub.'/';
		mkdir $subdirectory or warn "Cannot make directory $subdirectory: $!";
		$logger->info ( sprintf "\nCreated subtype subdirectory %s ok!\n", $subdirectory );	
		
		if ( -e $BLANK_WRITING_PAGE )
		{
			my $destination_filename = $subdirectory.'9998 '.$BLANK_WRITING_PAGE; 
			copy $BLANK_WRITING_PAGE, $destination_filename;
			$logger->info ( sprintf "	Copied the blank writing page to %s ok!\n", $destination_filename );
		}

		#
		#
		
		if ( scalar @frontcovers > 0 )
		{
			my $copied_cover = 0;
			foreach $cover ( @frontcovers )
			{
				my ( $cover_filename, $cover_directories, $cover_suffix ) = fileparse( $cover, qr/\.[^.]*/ );			
				# we found a match
				if ( index( lc($cover_filename), lc($sub) ) != -1 )		# string, substring, position - case INSENSITIVE
				{
					my $destination_filename = $subdirectory.'0000 '.$cover_filename.$cover_suffix; 
					copy $cover, $destination_filename;
					$logger->info ( sprintf "	Copied subtype specific front cover [%s] to [%s] ok!\n", $cover, $destination_filename );	
					$copied_cover++;					
				}
			}
			
			# if we didn't find a TASKS cover or a HANDOUT cover, but we know we've got a cover, find one without a subtype in the name
			if ( $copied_cover == 0 )
			{
				foreach $cover ( @frontcovers )
				{
					my ( $cover_filename, $cover_directories, $cover_suffix ) = fileparse( $cover, qr/\.[^.]*/ );
					
					if ( index( $cover_filename, "Generic Front Cover" ) != -1 )
					{
						my $destination_filename = $subdirectory.'0000 '.$cover_filename.$cover_suffix;  
						copy $cover, $destination_filename;
		
						# find out how many pages there are in this cover document.  there might be more than one page
						my $pdf = PDF::API2->open( $cover );
						$ignore_pages = $pdf->pages();
						$logger->info ( sprintf "	No subtype specific front cover found!  Copied front cover [%s] to [%s] ok!\n", $cover, $destination_filename );
					}
					else
					{
						$logger->info ( sprintf "	Front cover %s belongs to a subtype.  Skipping\n", $cover_filename );								
					}
				}
			}
		}
		
		#
		#
		
		if ( scalar @backcovers > 0 )
		{
			my $copied_cover = 0;
			foreach $cover ( @backcovers )
			{
				my ( $cover_filename, $cover_directories, $cover_suffix ) = fileparse( $cover, qr/\.[^.]*/ );			
				if ( index( $cover_filename, $sub ) != -1 )		# string, substring, position
				{
					my $destination_filename = $subdirectory.'9999 '.$cover_filename.$cover_suffix;  
					copy $cover, $destination_filename;
					$logger->info ( sprintf "	Copied subtype specific back cover [%s] to [%s] ok!\n", $cover, $destination_filename );
					$copied_cover++;					
				}
			}
			
			# if we didn't find a TASKS cover or a HANDOUT cover, but we know we've got a cover, just copy the first one to all folders
			if ( $copied_cover == 0 )
			{
				my ( $cover_filename, $cover_directories, $cover_suffix ) = fileparse( $backcovers[0], qr/\.[^.]*/ );	
				my $destination_filename = $subdirectory.'9999 '.$cover_filename.$cover_suffix;  
				copy $backcovers[0], $destination_filename;
				$logger->info ( sprintf "	No subtype specific back cover found!  Copied back cover [%s] to [%s] ok!\n", $backcovers[0], $destination_filename );
			}
		}
	}
	
	return;
}


# FindCourseForArchive
# Searches the course data loaded from the manifest to locate the course record
# that corresponds to the archive currently being processed, returning a
# reference to that course hash.
sub FindCourseForArchive
{
	my $archivename = shift;
	
	$logger->info ( sprintf "Trying to find a course for [%s]...\n", $archivename );

	foreach my $toplevel ( @toplevels )
	{
		if ( index( lc $archivename, lc $toplevel ) != -1 )			# string, substring, position
		{
			$archive_root = $toplevel;								# set this global, we may need it later
			$archive_root_level = substr $archive_root, -1;			# get the final character (this should be the level) 
			$archive_root_code = $toplevel;
			$archive_root_code =~ s/[a-z ]//g;						# get the code, from the first characters of each word.  Writing 1 -> W1
			$logger->info ( sprintf "Archive [%s] looks like [%s]; level [%i]; code [%s]\n", $archivename, $archive_root, $archive_root_level, $archive_root_code );
		}
		#else
		#{
		#	printf "Archive [%s] doesn't look like [%s]\n", $archivename, $toplevel;
		#}
	}
	
	return $archive_root;
}




# GetCovers
# Locates and copies cover-page PDF files for each unit into the working
# directory. Cover pages are matched to units by comparing unit titles against
# the filenames of PDFs found in the covers source directory.
sub GetCovers
{
	my $archivename = shift;
	$archivename =~ s/-/ /g;
	$archivename =~ s/  / /g;
	$archive_root = 'an unknown archive';
	
	if ( -e $COVERS_DIRECTORY && -d $COVERS_DIRECTORY )
	{
		my $toplevel = FindCourseForArchive( $archivename );
		
		# now have a look in the covers directory
		my $files = $COVERS_DIRECTORY."*.pdf";
		my @coverfiles = glob $files;
		$logger->info ( sprintf "Found %i covers in [%s]: \n\t%s\n", scalar @coverfiles, $COVERS_DIRECTORY, join("\n\t", @coverfiles) );
		
		# find a front cover
		foreach my $cover ( @coverfiles )
		{
			my ( $cover_filename, $cover_directories, $cover_suffix ) = fileparse( $cover, qr/\.[^.]*/ );
			#printf "[%s]\n", $cover_filename;
			
			if ( index( $cover_filename, $toplevel ) != -1 || index( $cover_filename, "Generic Front Cover" ) != -1 )		# string, substring, position
			{
				next if $cover =~ m/Back/i;
				$logger->info ( sprintf "Cover [%s] looks like [%s]\n", $cover, $toplevel );
				push @frontcovers, $cover;
			}
		}
		
		$logger->info ( sprintf "Found %i potential front covers\n", scalar @frontcovers );
		$logger->info ( sprintf "\t%s\n", join( "\n\t", @frontcovers ) ) if scalar @frontcovers > 0;
		
		# find a back cover
		my $truncated_toplevel = $toplevel;
		$truncated_toplevel =~ s/ [0-9]//g;	#remove numbers and spaces from the top level. I.e. "Writing 1" --> "Writing"
		$truncated_toplevel .= ' Back';
		
		foreach my $cover ( @coverfiles )
		{
			my ( $cover_filename, $cover_directories, $cover_suffix ) = fileparse( $cover, qr/\.[^.]*/ );
			#printf "[%s]\n", $cover_filename;
			
			if ( ( index( $cover_filename, $truncated_toplevel ) != -1 || index( $cover_filename, "Generic Back Cover" ) != -1 ) && not $cover ~~ @frontcovers )		# string, substring, position
			{
				$logger->info ( sprintf "Cover [%s] looks like [%s]\n", $cover, $truncated_toplevel );
				push @backcovers, $cover;
			}
			#else
			#{
			#	printf "Cover [%s] does not look like [%s]\n", $cover, $toplevel;
			#}
		}
				
		$logger->info ( sprintf "Found %i potential back covers", scalar @backcovers );
		$logger->info ( sprintf "\t%s\n", join("\n\t", @backcovers) ) if scalar @backcovers > 0;
	}

	$logger->info ( sprintf "Didn't find any front covers :-( for %s\n", $archive_root )	if ( scalar @frontcovers == 0 );	
	$logger->info ( sprintf "Didn't find any back covers :-( for %s\n", $archive_root )		if ( scalar @backcovers == 0 );	
	return;
}


# SaveStrings
# Persists key string values (e.g. course title, unit names, resource counts)
# to a plain-text file in the destination directory so they can be read back
# by other tools or scripts in the pipeline.
sub SaveStrings
{
	$logger->info ( sprintf "Saving all text to file [%s]...", $AllTextsFilename );
	
	# save the text string to file
	if ( length $AllTexts > 0 )
	{
		if ( open my $fh, '>', $AllTextsFilename )
		{
			print $fh $AllTexts;
			close $fh;
		}
		else
		{
			$logger->info ( sprintf "Cannot make file %s: $!", $AllTextsFilename );
		}
	}

	# save the tasks string to file
	if ( length $AllTasks > 0 )
	{
		if ( open my $fh, '>', $AllTasksFilename )
		{
			print $fh $AllTasks;
			close $fh;
		}
		else
		{
			$logger->info ( sprintf "Cannot make file %s: $!", $AllTasksFilename );
		}
	}
	
	return;
}




# MakeDirectoryAndCopyFiles
# Creates a subdirectory inside the destination directory (if it does not
# already exist) and copies a set of files into it. Used to stage resources
# for conversion or further processing.
sub MakeDirectoryAndCopyFiles
{
	$logger->info ( "Building directories and copying resources...\n" );
	
	foreach my $item ( @archivedetails )
	{
		$item->{location} =~ s/-->|<--/\//g;
		make_path( $item->{location} ) unless ( -e $item->{location} && -d $item->{location} );
		
		if( -e $item->{filename} )	# this should always exist
		{
			my ( $filename, $directories, $suffix ) = fileparse( $item->{filename}, qr/\.[^.]*/ );
			my $destination = $item->{location}.'/'.$filename.$suffix;
			
			# don't do this, it causes problems when we come to farm out documents for the nth time
			# we modify this document (e.g. add a template) so this check is problematic
			#unless ( -e $destination )
			if ( not -e $destination )
			{
				$logger->info ( sprintf "	Copying [%s] to [%s]... \n", $item->{filename}, $destination );
				copy( $item->{filename}, $destination ) or warn "Cannot copy from [$item->{filename}] to [$destination]: $!";
			}
			elsif ( -e $destination && compare ( $item->{filename}, $destination ) != 0 )
			{
				$logger->info ( sprintf "	Found updated [%s].  Updating folder... \n", $item->{filename} );
				copy( $item->{filename}, $destination ) or warn "Cannot copy from [$item->{filename}] to [$destination]: $!";
			}
			else
			{
				#printf "	File [%s] has not changed, so not copying!\n", $item->{filename};
			}
		}
				
		if( -e $item->{TXTfilename} )	# this will only exist if it's a TEXT
		{
			my($filename, $directories, $suffix) = fileparse( $item->{TXTfilename} );
			my $destination = $item->{location}.'/'.$filename.$suffix;
			
			unless ( -e $destination )
			{
				$logger->info ( sprintf "	Copying to [%s]\n", $destination );
				copy( $item->{TXTfilename}, $destination ) or warn "Cannot copy from [$item->{TXTfilename}] to [$destination]: $!";
			}
		}
		
		if( -e $item->{PrettyTXTfilename} )	# this will only exist if it's a TEXT
		{
			my($filename, $directories, $suffix) = fileparse( $item->{PrettyTXTfilename}, qr/\.[^.]*/ );
			my $destination = $item->{location}.'/'.$filename.$suffix;
			
			unless ( -e $destination )
			{
				$logger->info ( sprintf "	Copying to [%s]\n", $destination );
				copy( $item->{PrettyTXTfilename}, $destination ) or warn "Cannot copy from [$item->{PrettyTXTfilename}] to $destination: $!";
			}
		}

		if( -e $item->{Workbook} )	# this will only exist if it's a TEXT/TAPESCRIPT
		{
			# do not beautify this
			my($filename, $directories, $suffix) = fileparse( $item->{Workbook} );
			my $destination = $item->{location}.'/'.$filename.$suffix;
			
			unless ( -e $destination )
			{
				$logger->info ( sprintf "	Copying to [%s]\n", $destination );
				copy( $item->{Workbook}, $destination ) or warn "Cannot copy from [$item->{Workbook}] to [$destination]: $!";
			}
		}
		
		if( -e $item->{PDFfilename} )	# this **may** exist if it's a .doc(x)/image/ppt(x) file
		{
			my($filename, $directories, $suffix) = fileparse( $item->{PDFfilename}, qr/\.[^.]*/ );			
			my $destination = $item->{location}.'/'.$filename.$suffix;

			unless ( -e $destination )
			{
				$logger->info ( sprintf "	Copying [%s] to [%s]... \n", $item->{PDFfilename}, $destination );
				copy( $item->{PDFfilename}, $destination ) or warn "Cannot copy from [$item->{PDFfilename}] to [$destination]: $!";
			}
		}
				
		if( -e $item->{UntouchedFilename} )	# this will only exist if it's a TEXT
		{
			my($filename, $directories, $suffix) = fileparse( $item->{UntouchedFilename}, qr/\.[^.]*/ );
			my $destination = $item->{location}.'/'.$filename.$suffix;
			
			unless ( -e $destination )
			{
				$logger->info ( sprintf "	Copying [%s] to [%s]... \n", $item->{UntouchedFilename}, $destination );
				copy( $item->{UntouchedFilename}, $destination ) or warn "Cannot copy from [$item->{UntouchedFilename}] to $destination: $!";
			}
		}
	}
	
	return;
}




# FarmResourcesOut
# Dispatches each resource file to the appropriate processing pipeline
# depending on its type (DOCX, PDF, PPTX, etc.). Coordinates conversion to
# PDF via LibreOffice, extraction of text content, and population of the
# per-resource hash with metadata and analysis results.
sub FarmResourcesOut
{
	$logger->info ( "Copying resources to their dedicated folders...\n" );
	my $id = 1;
	
	foreach my $resource ( @archivedetails )
	{
		$logger->info ( sprintf "Resource is [%i], [%s]\n", $id, $resource->{IdAndFilenameNoPath} );
	
		if ( ( $resource->{type} eq "doc" || $resource->{type} eq "docx" || $resource->{type} eq "pptx" || $resource->{type} eq "pdf" ) && $resource->{FilenameNoPath} !~ /^~/ )
		{
			# copy item to the subtype specific resource folder.  i.e. copy all texts to the TEXT folder etc.
			if ( $resource->{subtype} ne '' && $resource->{subtype} ne 'INCLUDED' )
			{
				my $destination = sprintf( "%s/%s", $destinationDirectory.$resource->{subtype}, $resource->{IdAndFilenameNoPath} );
			
				# it's possible that we copied resources, applied the template and now we're attempting to copy the new file back.
				# check that the files are different before we copy the new one over
				#unless ( -e $destinationfilename )
				if ( not -e $destination )
				{
					$logger->info ( sprintf "	Copying [%s] to '%s' as [%s]... \n", $resource->{filename}, $resource->{subtype}, $destination );
					copy( $resource->{filename}, $destination ) or warn "Cannot copy from [$resource->{filename}] to [$destination]: $!";
				}
				elsif ( -e $destination && compare ( $resource->{filename}, $destination ) != 0 )
				{
					$logger->info ( sprintf "	Found updated [%s].  Updating %s folder... \n", $resource->{filename}, $resource->{subtype} );
					copy( $resource->{filename}, $destination ) or warn "Cannot copy from [$resource->{filename}] to [$destination]: $!";
				}
				else
				{
					#printf "	File [%s] has not changed, so not copying!\n", $$resource->{filename};
				}
				
				# check to see if there is a PDF file
				# this may exist if we've done some conversions
				if( -e $resource->{PDFfilename} )
				{
					my $destination = sprintf( "%s/%s", $destinationDirectory.$resource->{subtype}, $resource->{IdAndFilenameNoPathAsPDF} );

					$logger->info ( sprintf "	Copying [%s] to '%s' as [%s]...\n", $resource->{PDFfilename}, $resource->{subtype}, $destination );
					copy( $resource->{PDFfilename}, $destination ) or warn "Cannot copy [$resource->{PDFfilename}] to [$destination]: $!\n";
				}
				
				# check to see if there is a XLSX file
				# this may exist if we've made it, copy these to their own 'XLSX' folder
				if( -e $resource->{Workbook} )
				{
					my $destination = sprintf( "%s/%s", $destinationDirectory.'XLSX', $resource->{BareWorkbookFilename} );

					$logger->info ( sprintf "	Copying [%s] to '%s' as [%s]...\n", $resource->{Workbook}, 'XLSX', $destination );
					copy( $resource->{Workbook}, $destination ) or warn "Cannot copy [$resource->{Workbook}] to [$destination]: $!\n";
				}
			}
			elsif ( $resource->{subtype} eq 'INCLUDED' )	# && $resource->{parent} ~~ @alwaysinclude
			{
				foreach my $sub ( @subtypes )
				{
					next if ( $sub eq 'XLSX' );

					my $destination = sprintf( "%s/%s", $destinationDirectory.$sub, $resource->{IdAndFilenameNoPath} );
			
					# it's possible that we copied resources, applied the template and now we're attempting to copy the new file back.
					# check that the files are different before we copy the new one over
					#unless ( -e $destinationfilename )
					if ( not -e $destination )
					{
						$logger->info ( sprintf "	Copying [%s] to '%s' as [%s]... \n", $resource->{filename}, $sub, $destination );
						copy( $resource->{filename}, $destination ) or warn "Cannot copy from [$resource->{filename}] to [$destination]: $!";
					}
					elsif ( -e $destination && compare ( $resource->{filename}, $destination ) != 0 )
					{
						$logger->info ( sprintf "	Found updated [%s].  Updating %s folder... \n", $resource->{filename}, $sub );
						copy( $resource->{filename}, $destination ) or warn "Cannot copy from [$resource->{filename}] to [$destination]: $!";
					}
					else
					{
						#printf "	File [%s] has not changed, so not copying!\n", $$resource->{filename};
					}
					
					# check to see if there is a PDF file
					# this may exist if we've done some conversions
					if( -e $resource->{PDFfilename} )
					{
						my $destination = sprintf( "%s/%s", $destinationDirectory.$sub, $resource->{IdAndFilenameNoPathAsPDF} );

						$logger->info ( sprintf "	Copying [%s] to '%s' as [%s]...\n", $resource->{PDFfilename}, $sub, $destination );
						copy( $resource->{PDFfilename}, $destination ) or warn "Cannot copy [$resource->{PDFfilename}] to [$destination]: $!\n";
					}					
				}
			}
		
			# ALL DOCUMENTS (exclude these)
			# we exclude based on three criterion: 
			# 	file size > 10mb,
			# 	resource is in a folder marked "assessment" or "archive"
			# 	resource is a pdf with more than 10 pages.
			# OR ANSWERS
			if ( $resource->{subtype} ne 'EXCLUDED' && ( $resource->{subtype} ne 'ANSWERS' || $withanswers == 1 ) )
			{
				my $tempstring = sprintf "%s<><>%s", $resource->{title}, $resource->{MetaData}->{pages}; 
				push @tableofcontents, $tempstring;
			
				my $destination = sprintf("%s%04d %s", $AllDocumentsDirectory, $id, $resource->{FilenameNoPath});

				unless ( -e $destination )
				{
					$logger->info ( sprintf "	Copying [%s] to [%s]...\n", $resource->{filename}, $destination );
					copy( $resource->{filename} , $destination ) or warn "Cannot copy [$resource->{filename}] to [$destination]: $!\n";
				}
				
				# check to see if there is a PDF file
				# this may exist if we've done some conversions
				if( -e $resource->{PDFfilename} )
				{
					my $destination = sprintf( "%s/%s", $AllDocumentsDirectory, $resource->{IdAndFilenameNoPathAsPDF} );
					
					# if the files are different
					if ( compare ( $resource->{PDFfilename}, $destination ) != 0 )
					{
						$logger->info ( sprintf "	Copying [%s] to [%s]...\n", $resource->{PDFfilename}, $destination );
						copy( $resource->{PDFfilename}, $destination ) or warn "Cannot copy [$resource->{PDFfilename}] to [$destination]: $!\n";						
					}
				}
			}
			else
			{
				push @excludedtableofcontents, $resource->{FilenameNoPath};
				#printf "	Not copying [%s] to [%s].  Excluded (size, pages, words in title, etc).\n", $resource->{FilenameNoPath}, $AllDocumentsDirectory;
			}
		}
		else
		{
			# check to see if there is a PDF file
			# this will only exist if it's a .doc(x) file ** or an image or a PPT(x) **
			if( -e $resource->{PDFfilename} )
			{
				my $tempstring = sprintf "%s<><>%s", $resource->{title}, $resource->{MetaData}->{pages}; 
				push @tableofcontents, $tempstring;
			
				my $destination = sprintf( "%s%s", $AllDocumentsDirectory, $resource->{IdAndFilenameNoPathAsPDF} );

				# if the files are different
				if ( compare ( $resource->{PDFfilename}, $destination ) != 0 )
				{
					$logger->info ( sprintf "	Copying [%s] to [%s]... \n", $resource->{PDFfilename}, $destination );
					copy( $resource->{PDFfilename}, $destination ) or warn "Cannot copy [$resource->{PDFfilename}] to [$destination]: $!\n";
				}
			}
		}

		$id++
	}
	
	return;
}



# SaveArchiveInventory
# Writes a full inventory of all resources found in the archive to the
# 'Inventory' worksheet of the output Excel workbook. Columns include
# resource type, filename, unit, file size, page count, word count,
# readability scores, vocabulary statistics, and nonstandard phrases.
sub SaveArchiveInventory
{
	my $exportfile = shift;
	my $sheetname = shift;
	my $i = 0;
	my $record_count = scalar @archivedetails + 1;		# add one to this for the header row
	
	$logger->info ( sprintf "Saving archive Inventory details to archive summary file %s...", $exportfile );
	
	if ( defined $workbook )
	{
		my $worksheet = $workbook->add_worksheet( $sheetname );
		$worksheet->set_tab_color( 'green' );
		$worksheet->freeze_panes( 1, 0 );    # Freeze the first row
		$worksheet->set_zoom( 150 );
	
		my $bold = $workbook->add_format( bold => 1 );
		my $centre = $workbook->add_format( align => 'center' );
		my $date_format = $workbook->add_format( num_format => 'yyyy-mm-dd' );
		my $url_format = $workbook->add_format( color => 'blue', underline => 1 );

		my $red_format =    $workbook->add_format( bg_color => '#FFC7CE', color => '#9C0006' );
		my $yellow_format = $workbook->add_format( bg_color => '#FFEB9C', color => '#9C6500' );
		my $green_format =  $workbook->add_format( bg_color => '#C6EFCE', color => '#006100' );
	
		$worksheet->write_row( 0, 0, ["ARCHIVE", "ID", "LEVEL", "UNIT", "PARENT FOLDER", "LOCATION", "IDENTIFIER", "IDENTIFIERREF", "TITLE", "WEBLINK", "TYPE", "SUBTYPE", "EXISTS", "SIZE (KB)", "MD5", "CREATED BY", "CREATED DATE", "LAST MODIFIED BY", "LAST MODIFIED DATE", "REVISION NUMBER", "TOTAL EDITING TIME (MINS)", "#PAGES", "#PARAGRAPHS", "#LINES", "#WORDS", "#CHARACTERS", "NGSL_TO_800", "NGSL_TO_1600", "NGSL_TO_2400", "NGSL_TO_2601", "NGSL_TOTAL", "NAWL_TOTAL", "NEW_WORDS_TOTAL", "UNKNOWN_WORDS_TOTAL", "NGSL_PERCENT", "NAWL_PERCENT", "NEW_PERCENT", "UNKNOWN_PERCENT", "FLESCH", "FLESCH-KINCAID", "GUNNING FOG", "#ACADEMIC COLLOCATIONS", "SOURCE", "PAGE SIZE", "PAGE ORIENTATION", "PAGE BORDERS", "LESSON FOCUS", "ERRANT NEWLINES", "AGE IN DAYS SINCE LAST EDIT", '# OF SPELLING AND GRAMMAR ERRORS', 'TYPOGRAPHY', 'PUNCTUATION', 'MISC', 'TYPOS', 'CASING', 'CONFUSEDWORDS', 'GRAMMAR', 'STYLE', 'REDUNDANCY', 'SEMANTICS', 'NONSTANDARD PHRASE', 'COLLOCATIONS', 'KEY VOCABULARY', 'A1 MULTIWORD', 'A2 MULTIWORD', 'B1 MULTIWORD', 'B2 MULTIWORD', 'C1 MULTIWORD', 'C2 MULTIWORD', 'A1 COUNT', 'A2 COUNT', 'B1 COUNT', 'B2 COUNT', 'C1 COUNT', 'C2 COUNT'], $bold );
	
		# set the widths
		$worksheet->set_column( 'A:A', 15, undef, 1 );	# hide this column
		$worksheet->set_column( 'B:B',  4 );
		$worksheet->set_column( 'C:C',  5 );
		$worksheet->set_column( 'D:D', 30 );

		$worksheet->set_column( 'E:E', 15 );
		$worksheet->set_column( 'F:F', 40 );
		$worksheet->set_column( 'G:G', 11, undef, 1 );	# hide this column
		$worksheet->set_column( 'H:H', 11, undef, 1 );	# hide this column

		$worksheet->set_column( 'I:I', 30 );
		$worksheet->set_column( 'J:J',  7 );
		$worksheet->set_column( 'K:K',  6 );
		$worksheet->set_column( 'L:L',  9 );
	
		$worksheet->set_column( 'M:N',  6 );
		$worksheet->set_column( 'N:N',  6 );

		$worksheet->set_column( 'O:O', 33 );
		$worksheet->set_column( 'P:P', 12 );
		$worksheet->set_column( 'Q:Q', 12 );
		$worksheet->set_column( 'R:R', 12 );
		$worksheet->set_column( 'S:S', 12 );

		$worksheet->set_column( 'AP:AP',  22 );
		$worksheet->set_column( 'AQ:AQ',  90 );

		$worksheet->set_column( 'AR:AR',  9 );
		$worksheet->set_column( 'AS:AS',  15 );
		$worksheet->set_column( 'AT:AT',  27 );		# page borders
		$worksheet->set_column( 'AU:AU',  40 );		# lesson focus

		$worksheet->set_column( 'AV:AV',  10 );		# errant new lines
		$worksheet->set_column( 'AW:AW',  10 );		# age in days since last edit
		$worksheet->set_column( 'AX:AX',  10 );		# number of spelling and grammar errors

		$worksheet->set_column( 'AY:AY',  10 );		# A1 multiword
		$worksheet->set_column( 'AZ:AZ',  10 );
		$worksheet->set_column( 'BA:BA',  10 );
		$worksheet->set_column( 'BB:BB',  10 );
		$worksheet->set_column( 'BC:BC',  10 );
		$worksheet->set_column( 'BD:BD',  10 );

		$worksheet->set_column( 'BE:BE',  10 );		# A1
		$worksheet->set_column( 'BF:BF',  10 );
		$worksheet->set_column( 'BG:BG',  10 );
		$worksheet->set_column( 'BH:BH',  10 );
		$worksheet->set_column( 'BI:BI',  10 );
		$worksheet->set_column( 'BJ:BJ',  10 );
	
		#
		# apply our conditional formatting rules
		#
	
		# highlight any old office (.doc .xls .ppt) files
		$worksheet->conditional_formatting( 'K2:K'.$record_count, { type => 'cell', criteria => 'equal to', value => 'doc', format => $red_format } );
		$worksheet->conditional_formatting( 'K2:K'.$record_count, { type => 'cell', criteria => 'equal to', value => 'xls', format => $red_format } );
		$worksheet->conditional_formatting( 'K2:K'.$record_count, { type => 'cell', criteria => 'equal to', value => 'ppt', format => $red_format } );
	
		# highlight any big files		
		$worksheet->conditional_formatting( 'N2:N'.$record_count, { type => 'cell', criteria => 'greater than', value => ($MAXIMUM_FILE_SIZE/1024) * 2, format => $red_format } );
		$worksheet->conditional_formatting( 'N2:N'.$record_count, { type => 'cell', criteria => 'greater than', value => ($MAXIMUM_FILE_SIZE/1024), format => $yellow_format } );
	
		# highlight any duplicated MD5s		
		$worksheet->conditional_formatting( 'O2:O'.$record_count, { type => 'duplicate', format => $red_format } );
	
		# highlight any old/new documents
		my $today = localtime;		
		my $newly_changed = $today - ( 86400 * $ONE_MONTH_OLD_FILE );
		my $one_year_old =  $today - ( 86400 * $ONE_YEAR_OLD_FILE );
		my $two_years_old = $today - ( 86400 * $TWO_YEAR_OLD_FILE );

		my $today_string = $today->strftime( '%Y-%m-%dT' );						# e.g. '2011-01-01T'
		my $newly_changed_string = $newly_changed->strftime( '%Y-%m-%dT' );		# e.g. '2011-01-01T'
		my $one_year_old_string = $one_year_old->strftime( '%Y-%m-%dT' );		# e.g. '2011-01-01T'
		my $two_years_old_string = $two_years_old->strftime( '%Y-%m-%dT' );		# e.g. '2011-01-01T'
	
		$worksheet->conditional_formatting( 'Q2:Q'.$record_count, { type => 'date', criteria => 'between', minimum  => $newly_changed_string, maximum  => $today_string, format => $green_format } );
		$worksheet->conditional_formatting( 'S2:S'.$record_count, { type => 'date', criteria => 'between', minimum  => $newly_changed_string, maximum  => $today_string, format => $green_format } );
	
		$worksheet->conditional_formatting( 'S2:S'.$record_count, { type => 'date', criteria => 'less than', value => $two_years_old_string, format => $red_format } );
		$worksheet->conditional_formatting( 'S2:S'.$record_count, { type => 'date', criteria => 'less than', value => $one_year_old_string, format => $yellow_format } );


		# highlight any long documents
		$worksheet->conditional_formatting( 'V2:V'.$record_count, { type => 'cell', criteria => 'greater than', value => $UPPER_LIMIT_PAGES * 2, format => $red_format } );
		$worksheet->conditional_formatting( 'V2:V'.$record_count, { type => 'cell', criteria => 'greater than', value => $UPPER_LIMIT_PAGES, format => $yellow_format } );
	
		# highlight norm-referenced text length
		$worksheet->conditional_formatting( 'Y2:Y'.$record_count, { type  => 'data_bar', bar_color => "#FF0080", min_value  => 1 } );	# ignore zero-length items
	
		if ( $archive_root eq 'Reading 1' )
		{
			$worksheet->conditional_formatting( 'Y2:Y'.$record_count, { type => 'cell', criteria => 'less than', value => $READING_LEVEL_1_LOWER_TEXT_LIMIT-100, format => $red_format } );
			$worksheet->conditional_formatting( 'Y2:Y'.$record_count, { type => 'cell', criteria => 'less than', value => $READING_LEVEL_1_LOWER_TEXT_LIMIT, format => $yellow_format } );

			$worksheet->conditional_formatting( 'Y2:Y'.$record_count, { type => 'cell', criteria => 'greater than', value => $READING_LEVEL_1_HIGHER_TEXT_LIMIT+100, format => $red_format } );
			$worksheet->conditional_formatting( 'Y2:Y'.$record_count, { type => 'cell', criteria => 'greater than', value => $READING_LEVEL_1_HIGHER_TEXT_LIMIT, format => $yellow_format } );
		}
		elsif ( $archive_root eq 'Reading 2' )
		{
			$worksheet->conditional_formatting( 'Y2:Y'.$record_count, { type => 'cell', criteria => 'less than', value => $READING_LEVEL_2_LOWER_TEXT_LIMIT-100, format => $red_format } );
			$worksheet->conditional_formatting( 'Y2:Y'.$record_count, { type => 'cell', criteria => 'less than', value => $READING_LEVEL_2_LOWER_TEXT_LIMIT, format => $yellow_format } );

			$worksheet->conditional_formatting( 'Y2:Y'.$record_count, { type => 'cell', criteria => 'greater than', value => $READING_LEVEL_2_HIGHER_TEXT_LIMIT+100, format => $red_format } );
			$worksheet->conditional_formatting( 'Y2:Y'.$record_count, { type => 'cell', criteria => 'greater than', value => $READING_LEVEL_2_HIGHER_TEXT_LIMIT, format => $yellow_format } );
		} 
		elsif ( $archive_root eq 'Reading 3' )
		{
			$worksheet->conditional_formatting( 'Y2:Y'.$record_count, { type => 'cell', criteria => 'less than', value => $READING_LEVEL_3_LOWER_TEXT_LIMIT-100, format => $red_format } );
			$worksheet->conditional_formatting( 'Y2:Y'.$record_count, { type => 'cell', criteria => 'less than', value => $READING_LEVEL_3_LOWER_TEXT_LIMIT, format => $yellow_format } );

			$worksheet->conditional_formatting( 'Y2:Y'.$record_count, { type => 'cell', criteria => 'greater than', value => $READING_LEVEL_3_HIGHER_TEXT_LIMIT+100, format => $red_format } );
			$worksheet->conditional_formatting( 'Y2:Y'.$record_count, { type => 'cell', criteria => 'greater than', value => $READING_LEVEL_3_HIGHER_TEXT_LIMIT, format => $yellow_format } );	
		} 
		elsif ( $archive_root eq 'Reading 4' )
		{
			$worksheet->conditional_formatting( 'Y2:Y'.$record_count, { type => 'cell', criteria => 'less than', value => $READING_LEVEL_4_LOWER_TEXT_LIMIT-100, format => $red_format } );
			$worksheet->conditional_formatting( 'Y2:Y'.$record_count, { type => 'cell', criteria => 'less than', value => $READING_LEVEL_4_LOWER_TEXT_LIMIT, format => $yellow_format } );

			$worksheet->conditional_formatting( 'Y2:Y'.$record_count, { type => 'cell', criteria => 'greater than', value => $READING_LEVEL_4_HIGHER_TEXT_LIMIT+100, format => $red_format } );
			$worksheet->conditional_formatting( 'Y2:Y'.$record_count, { type => 'cell', criteria => 'greater than', value => $READING_LEVEL_4_HIGHER_TEXT_LIMIT, format => $yellow_format } );	
		} 
		elsif ( $archive_root eq 'Reading 5' )
		{
			$worksheet->conditional_formatting( 'Y2:Y'.$record_count, { type => 'cell', criteria => 'less than', value => $READING_LEVEL_5_LOWER_TEXT_LIMIT-100, format => $red_format } );
			$worksheet->conditional_formatting( 'Y2:Y'.$record_count, { type => 'cell', criteria => 'less than', value => $READING_LEVEL_5_LOWER_TEXT_LIMIT, format => $yellow_format } );

			$worksheet->conditional_formatting( 'Y2:Y'.$record_count, { type => 'cell', criteria => 'greater than', value => $READING_LEVEL_5_HIGHER_TEXT_LIMIT+100, format => $red_format } );
			$worksheet->conditional_formatting( 'Y2:Y'.$record_count, { type => 'cell', criteria => 'greater than', value => $READING_LEVEL_5_HIGHER_TEXT_LIMIT, format => $yellow_format } );		
		} 
	
		# highlight our readability indicies
		$worksheet->conditional_formatting( 'AM2:AM'.$record_count, { type  => '2_color_scale', min_color => "#C5D9F1", max_color => "#538ED5", min_value  => 100, max_value  => 0 } );
		$worksheet->conditional_formatting( 'AN2:AN'.$record_count, { type  => '2_color_scale', min_color => "#C5D9F1", max_color => "#538ED5", min_value  => 2, max_value  => 18 } );
		$worksheet->conditional_formatting( 'AO2:AO'.$record_count, { type  => '2_color_scale', min_color => "#C5D9F1", max_color => "#538ED5", min_value  => 4, max_value  => 20 } );
	
		# highlight any non-A4 page sizes
		$worksheet->conditional_formatting( 'AR2:AR'.$record_count, { type => 'cell', criteria => 'not equal to', value => 'A4', format => $red_format } );

		#
		# save resource details
		#
	
		foreach my $resource ( @archivedetails )
		{		
			$worksheet->write_string( $i+1, 0, $zipbasename );		# row, column, token
			$worksheet->write_number( $i+1, 1, $i, $centre );
			$worksheet->write_number( $i+1, 2, $resource->{level}, $centre );
			$worksheet->write_string( $i+1, 3, $resource->{unit} );
			$worksheet->write_string( $i+1, 4, $resource->{parent} );

			$worksheet->write_string( $i+1, 5, $resource->{location} );
			$worksheet->write( $i+1, 6, $resource->{identifier} );
			$worksheet->write( $i+1, 7, $resource->{identifierref} );
			$worksheet->write_string( $i+1, 8, $resource->{title} );
		
			if ( defined $resource->{XML}->{URL} && $resource->{XML}->{URL} ne '' )
			{
				$worksheet->write_url( $i+1, 9, $resource->{XML}->{URL}, $url_format );
			}
			else
			{
				$worksheet->write_string( $i+1, 9, '' );
			}
		
			$worksheet->write_string( $i+1, 10, $resource->{type}, $centre );
			$worksheet->write_string( $i+1, 11, $resource->{subtype}, $centre );
			$worksheet->write_string( $i+1, 12, $resource->{exists}, $centre );
		
			$worksheet->write_number( $i+1, 13, int($resource->{size}/1024), $centre );
			$worksheet->write( $i+1, 14, $resource->{MD5} );
			$worksheet->write( $i+1, 15, $resource->{MetaData}->{createdby} );
		
			if ( defined $resource->{MetaData}->{createddate} && $resource->{MetaData}->{createddate} ne '' )
			{
				$worksheet->write_date_time( $i+1, 16, substr( $resource->{MetaData}->{createddate}, 0, 11 ), $date_format );
			}
			else
			{
				$worksheet->write_string( $i+1, 16, '' );
			}
		
			$worksheet->write( $i+1, 17, $resource->{MetaData}->{lastmodifiedby} );

			if ( defined $resource->{MetaData}->{lastmodifieddate} && $resource->{MetaData}->{lastmodifieddate} ne '' )
			{
				$worksheet->write_date_time( $i+1, 18, substr( $resource->{MetaData}->{lastmodifieddate}, 0, 11 ), $date_format );
			}
			else
			{
				$worksheet->write_string( $i+1, 18, '' );
			}

			$worksheet->write( $i+1, 19, $resource->{MetaData}->{revisionnumber}, $centre );
			$worksheet->write( $i+1, 20, $resource->{MetaData}->{totaleditingtime}, $centre );
			$worksheet->write( $i+1, 21, $resource->{MetaData}->{pages}, $centre );
		
			$worksheet->write( $i+1, 22, $resource->{WordMetaData}->{paragraphs}, $centre );
			$worksheet->write( $i+1, 23, $resource->{WordMetaData}->{lines}, $centre );
			$worksheet->write( $i+1, 24, $resource->{WordMetaData}->{words}, $centre );
			$worksheet->write( $i+1, 25, $resource->{WordMetaData}->{characters}, $centre );
		
			$worksheet->write_number( $i+1, 26, $resource->{WordListMetaData}->{NGSLBandA}, $centre );
			$worksheet->write_number( $i+1, 27, $resource->{WordListMetaData}->{NGSLBandB}, $centre );
			$worksheet->write_number( $i+1, 28, $resource->{WordListMetaData}->{NGSLBandC}, $centre );
			$worksheet->write_number( $i+1, 29, $resource->{WordListMetaData}->{NGSLBandD}, $centre );
			$worksheet->write_number( $i+1, 30, $resource->{WordListMetaData}->{NGSLTotalCount}, $centre );
			$worksheet->write_number( $i+1, 31, $resource->{WordListMetaData}->{NAWLTotalCount}, $centre );
			$worksheet->write_number( $i+1, 32, $resource->{WordListMetaData}->{NewWords}, $centre );
			$worksheet->write_number( $i+1, 33, $resource->{WordListMetaData}->{TyposCount}, $centre );

			$worksheet->write( $i+1, 34, $resource->{WordListMetaData}->{NGSLPercent}, $centre );
			$worksheet->write( $i+1, 35, $resource->{WordListMetaData}->{NAWLPercent}, $centre );
			$worksheet->write( $i+1, 36, $resource->{WordListMetaData}->{NewPercent}, $centre );
			$worksheet->write( $i+1, 37, $resource->{WordListMetaData}->{UnknownPercent}, $centre );

			$worksheet->write( $i+1, 38, $resource->{Readability}->{Flesch}, $centre );
			$worksheet->write( $i+1, 39, $resource->{Readability}->{FleschKincaid}, $centre );
			$worksheet->write( $i+1, 40, $resource->{Readability}->{GunningFog}, $centre );

			$worksheet->write( $i+1, 41, $resource->{WordListMetaData}->{AcademicCollocationsCount}, $centre );

			$worksheet->write_string( $i+1, 42, $resource->{WordMetaData}->{source} );

			$worksheet->write_string( $i+1, 43, $resource->{WordMetaData}->{PageSize}, $centre );
			$worksheet->write_string( $i+1, 44, $resource->{WordMetaData}->{PageOrientation} );
			$worksheet->write_string( $i+1, 45, $resource->{WordMetaData}->{PageBorders} );
			$worksheet->write_string( $i+1, 46, "" );		# $resource->{MetaData}->{LessonFocus}.  this is an array
			$worksheet->write_string( $i+1, 47, $resource->{ErrantCRLFs} );		

			$worksheet->write_number( $i+1, 48, $resource->{MetaData}->{AgeInDaysSinceLastEdit}, $centre );
			
			$worksheet->write_number( $i+1, 49, $resource->{NumberOfSpellingAndGrammarErrors}, $centre );

			$worksheet->write_number( $i+1, 50, $resource->{SpellingMistakes}->{Typography}, $centre );
			$worksheet->write_number( $i+1, 51, $resource->{SpellingMistakes}->{Punctuation}, $centre );
			$worksheet->write_number( $i+1, 52, $resource->{SpellingMistakes}->{Miscellaneous}, $centre );
			$worksheet->write_number( $i+1, 53, $resource->{SpellingMistakes}->{Typos}, $centre );
			$worksheet->write_number( $i+1, 54, $resource->{SpellingMistakes}->{Casing}, $centre );
			$worksheet->write_number( $i+1, 55, $resource->{SpellingMistakes}->{ConfusedWords}, $centre );
			$worksheet->write_number( $i+1, 56, $resource->{SpellingMistakes}->{Grammar}, $centre );
			$worksheet->write_number( $i+1, 57, $resource->{SpellingMistakes}->{Style}, $centre );
			$worksheet->write_number( $i+1, 58, $resource->{SpellingMistakes}->{Redundancy}, $centre );
			$worksheet->write_number( $i+1, 59, $resource->{SpellingMistakes}->{Semantics}, $centre );
			$worksheet->write_number( $i+1, 60, $resource->{SpellingMistakes}->{NonstandardPhrases}, $centre );
			$worksheet->write_number( $i+1, 61, $resource->{SpellingMistakes}->{Collocations}, $centre );
			$worksheet->write_string( $i+1, 62, $resource->{WordMetaData}->{KeyVocabulary} );

			$worksheet->write_string( $i+1, 63, $resource->{WordMetaData}->{CEFR_A1_MultiWord_Count} );
			$worksheet->write_string( $i+1, 64, $resource->{WordMetaData}->{CEFR_A2_MultiWord_Count} );
			$worksheet->write_string( $i+1, 65, $resource->{WordMetaData}->{CEFR_B1_MultiWord_Count} );
			$worksheet->write_string( $i+1, 66, $resource->{WordMetaData}->{CEFR_B2_MultiWord_Count} );
			$worksheet->write_string( $i+1, 67, $resource->{WordMetaData}->{CEFR_C1_MultiWord_Count} );
			$worksheet->write_string( $i+1, 68, $resource->{WordMetaData}->{CEFR_C2_MultiWord_Count} );

			$worksheet->write_string( $i+1, 69, $resource->{WordMetaData}->{CEFR_A1_Count} );
			$worksheet->write_string( $i+1, 70, $resource->{WordMetaData}->{CEFR_A2_Count} );
			$worksheet->write_string( $i+1, 71, $resource->{WordMetaData}->{CEFR_B1_Count} );
			$worksheet->write_string( $i+1, 72, $resource->{WordMetaData}->{CEFR_B2_Count} );
			$worksheet->write_string( $i+1, 73, $resource->{WordMetaData}->{CEFR_C1_Count} );
			$worksheet->write_string( $i+1, 74, $resource->{WordMetaData}->{CEFR_C2_Count} );

			$i++;
		}		
	}
	else
	{
		$logger->info ( sprintf "	Workbook object is undefined!" );
	}
	
	return;
}



# GetElectiveName
# Reads a small configuration file to determine the elective/course name
# associated with the archive being processed. Returns the name string,
# or a default value if the file cannot be found or parsed.
sub GetElectiveName
{
	$logger->info ( sprintf "Getting elective names from %s...", $elective_filename );
	
	if ( open( my $fh, '<', $elective_filename ) )
	{
		while ( my $line = <$fh> )		# for each line in the csv file...
		{
			chomp($line);
			my ( $coursecode, $coursename ) = split "\t", $line, 2;
		
			$logger->info ( sprintf "[%s]	-->	[%s]", $coursecode, $coursename );
		
			my %key = ( course_code	=> $coursecode, course_name	=> $coursename );

			push @elective_details, \%key;		# *probably* unnecessary since we only ever deal with one elective per archive
		
			if ( $elective_coursecode eq $coursecode )
			{
				$elective_coursename = $coursename;
				$logger->info ( sprintf "Found!  Elective name is [%s]...", $elective_coursename );	
			}
		}
		
		close $fh;
	}
	else
	{
		$logger->warn ( "Cannot open $elective_filename: $!" ) if ( $elective_coursename eq '' );
	}

	$logger->info ( sprintf ":-<  Cound not find an elective name for [%s] in [%s]!", $elective_coursecode, $elective_filename ) if ( $elective_coursename eq '');
	return;
}



# SaveArchiveStatistics
# Writes summary statistics for the entire archive (file type counts, total
# word counts, aggregate readability scores, vocabulary coverage percentages,
# etc.) to the 'Statistics' worksheet of the output Excel workbook.
sub SaveArchiveStatistics
{
	my $exportfile = shift;

	my $i = 1;
	my $AllDocuments_cumulative_page_count = 0;

	$logger->info ( sprintf "Saving archive statistics to TXT file %s...", $exportfile );

	my $string = '';
	$string .= "***** ARCHIVE DETAILS *****\n";; 
	
	$string .= sprintf "Archive name is %s\n", $hasfilename 	if ( $hasfilename ne '' && $hasfilename =~ m/\.imscc$/ ); 
	$string .= sprintf "File name is %s\n", $hasfilename 		if ( $hasfilename ne '' && $hasfilename =~ m/\.docx$/ ); 
	$string .= sprintf "Directory name is %s\n", $hasdirectory 	if ( $hasdirectory ne '' ); 
	$string .= sprintf "Found: %i doc files\n", $docfiles;
	$string .= sprintf "Found: %i docx files\n", $docxfiles;
	$string .= sprintf "	Unable to parse %i of them\n", scalar(@unabletoparseDOCX) if scalar(@unabletoparseDOCX) > 0;
	$string .= sprintf "		%s\n", join("\n\t", @unabletoparseDOCX) if scalar(@unabletoparseDOCX) > 0;
	$string .= sprintf "Found: %i pdf files\n", $pdffiles;
	$string .= sprintf "Found: %i image files\n", $imagefiles;
	$string .= sprintf "Found: %i powerpoint files\n", $pptfiles;
	$string .= sprintf "Found: %i audio files\n", $mp3files;
	$string .= sprintf "Found: %i video files\n", $mp4files;
	$string .= sprintf "Found: %i text files\n", $txtfiles;
	$string .= sprintf "Found: %i spreadsheet files\n", $xlsfiles;
	$string .= sprintf "Found: %i xml/html files.  Of which...\n", $xmlfiles+$htmlfiles;
	
	$i = 1;	# reset this
	$string .= sprintf "	Found: %i weblinks\n", $weblinks;
	foreach my $link ( @weblinks )
	{
		$string .= sprintf "		%03i) %s\n", $i, $link;
		$i++;
	}
	
	$string .= sprintf "	Found: %i Schoology Resources (quizzes, discussions, etc)\n", $schoologyresources;
	$string .= sprintf "	Found: %i Schoology Assignments files\n", $htmlfiles;
	$string .= sprintf "	Unable to parse %i of them\n", scalar(@unabletoparseXML) if scalar(@unabletoparseXML) > 0;;
	$string .= sprintf "	%s\n", join("\n\t", @unabletoparseXML) if scalar(@unabletoparseXML) > 0;
	
	$string .= sprintf "Found: %i files in total\n", $filecount;
	$string .= sprintf "Found: %i folders\n", $foldercount;
	
	$string .= "\n\n***** MATERIALS DETAILS *****\n";
	
	$i = 1;	# reset this
	$string .= sprintf "\nFiles included in 'All Documents' directory are:\n";
	foreach my $item ( @tableofcontents )
	{
		my ( $file, $page_count ) = split ( '<><>', $item );

		$string .= sprintf "%03i)	%s\n", $i, $file;
		$i++;
	}
	
	$i = 1;	# reset this
	$string .= sprintf "\nFiles NOT included in 'All Documents' directory are:\n" if scalar @excludedtableofcontents > 0;
	foreach my $item (@excludedtableofcontents)
	{
		$string .= sprintf "%03i)	%s\n", $i, $item;
		$i++;
	}
	
	$string .= "\n\n***** VOCABULARY DETAILS *****\n";
	
	# save vocab information
	if ( scalar keys %ngsl > 0 )
	{
		my $band_a_count = 0;
		my $band_b_count = 0;
		my $band_c_count = 0;
		my $band_d_count = 0;
		
		my $found_band_a_count = 0;
		my $found_band_b_count = 0;
		my $found_band_c_count = 0;
		my $found_band_d_count = 0;
		
		foreach my $outerkey ( keys %ngsl )
		{ 
			my $innerkey = $ngsl{$outerkey};
			
			# we know how many items there are in each band, but count them anyway
			$band_a_count++ if $innerkey->{level} eq "BandA";
			$band_b_count++ if $innerkey->{level} eq "BandB";
			$band_c_count++ if $innerkey->{level} eq "BandC";
			$band_d_count++ if $innerkey->{level} eq "BandD";
			
			$found_band_a_count += $innerkey->{found_count} if ( $innerkey->{level} eq "BandA" && $innerkey->{found_count} > 0 );
			$found_band_b_count += $innerkey->{found_count} if ( $innerkey->{level} eq "BandB" && $innerkey->{found_count} > 0 );
			$found_band_c_count += $innerkey->{found_count} if ( $innerkey->{level} eq "BandC" && $innerkey->{found_count} > 0 );
			$found_band_d_count += $innerkey->{found_count} if ( $innerkey->{level} eq "BandD" && $innerkey->{found_count} > 0 );
		}
		
		$string .= sprintf "\nNGSL Vocab Information:\n";
		$string .= sprintf "	Found %-4i Band A words from %4i total (%.2f percent coverage)\n", $found_band_a_count, $band_a_count, ($found_band_a_count/$band_a_count)*100; 
		$string .= sprintf "	Found %-4i Band B words from %4i total (%.2f percent coverage)\n", $found_band_b_count, $band_b_count, ($found_band_b_count/$band_b_count)*100; 
		$string .= sprintf "	Found %-4i Band C words from %4i total (%.2f percent coverage)\n", $found_band_c_count, $band_c_count, ($found_band_c_count/$band_c_count)*100; 
		$string .= sprintf "	Found %-4i Band D words from %4i total (%.2f percent coverage)\n", $found_band_d_count, $band_d_count, ($found_band_d_count/$band_d_count)*100; 
	}
	
	
	# save NAWL vocab information
	if ( scalar keys %nawl > 0 )
	{
		my $found_nawl_count = 0;
		
		foreach my $outerkey ( keys %nawl )
		{ 
			my $innerkey = $nawl{$outerkey};
			$found_nawl_count += $innerkey->{found_count} if ( $innerkey->{found_count} > 0 );
		}
		
		$string .= sprintf "\nNAWL Vocab Information:\n";
		$string .= sprintf "	Found %-4i academic words from %4i total (%.2f percent coverage)\n", $found_nawl_count, scalar keys %nawl, ( $found_nawl_count / scalar keys %nawl ) * 100;
	}
	
	
	# save Academic Collocation vocab information
	if ( scalar keys %AcademicCollocationList > 0 )
	{
		my $found_colls_count = 0;
		
		foreach my $outerkey ( keys %AcademicCollocationList )
		{ 
			my $innerkey = $AcademicCollocationList{$outerkey};		
			$found_colls_count += $innerkey->{found_count} if ( $innerkey->{found_count} > 0 );
		}
		
		$string .= sprintf "\nAcademic Collocations Vocab Information:\n";
		$string .= sprintf "	Found %-4i academic collocations from %4i total (%.2f percent coverage)\n", $found_colls_count, scalar keys %AcademicCollocationList, ( $found_colls_count / scalar keys %AcademicCollocationList ) * 100; 
	}
	
	$logger->info ( sprintf "%s", $string );
	
	# save details of 'All Documents' manifest to file
	if ( open my $fh, '>', $exportfile )
	{
		print $fh $string;
		close $fh;
	}
	else
	{
		warn "Cannot open $exportfile: $!\n";
	}
	
	return;
}




# SaveTextAsPDF
# Converts a plain-text string to a PDF file using PDF::Create. The text is
# wrapped to fit the page width, paginated automatically, and saved to the
# path specified in the resource hash.
sub SaveTextAsPDF
{
	my $filename = shift;
	my $string = shift;
	my $blankpagesbefore = shift;
	my $blankpagesafter = shift;
	
	my $pdf = new PDF::Report( PageSize => "A4", PageOrientation => "portrait", undef => undef );

	for ( my $i = 0; $i < $blankpagesbefore; $i++ )
	{
		$pdf->newpage();
	}

	$pdf->newpage(1);
	$pdf->setFont( 'Helvetica-bold' );
	$pdf->setSize(100);
	my ( $width, $height ) = $pdf->getPageDimensions();

	$string =~ s/  / /g;	# replace two spaces with one
	
	# does the string look like this:  Unit 1 of 7: Forming Theories.  The : may or may not exist
	# or TAT 001: English for Travel and Tourism
	# https://regexr.com/
	if ( $string =~ m/^Unit \d of \d: \w.+/i || $string =~ m/^[A-Z]{2,4} 001: \w.+/g )
	{
		my ( $part1, $part2 ) = split /:/, $string;
		
		$part1 =~ s/^\s+|\s+$//g;	# trim leading and trailing spaces	
		$part2 =~ s/^\s+|\s+$//g;	# trim leading and trailing spaces	
		
		$pdf->centerString( 20, $width-20, $height/4 + $height/3, $part1 );
		$pdf->centerString( 20, $width-20, $height/2.5, $part2 );
	}
	else
	{
		$pdf->centerString( 20, $width-20, $height/2, $string );
	}

	for ( my $i = 0; $i < $blankpagesafter; $i++ )
	{
		$pdf->newpage();
	}
	
	open( PDF, "> $filename" ) or warn "Error opening $filename: $!\n";
	print PDF $pdf->Finish();
	close PDF;
	
	$logger->info ( sprintf "Saved text [%s] as PDF file [%s]... ok.\n", $string, $filename );
	return;
}


# GetTagsAndText
# Extracts the plain text content and any inline tags (e.g. heading levels,
# list markers) from a resource. For DOCX files this involves unzipping and
# parsing the word/document.xml; for PDFs it calls pdftotext. Returns the
# extracted text and a parallel tagged version.
sub GetTagsAndText
{
	my $keyhash = shift;
	my $count = 0;
	my $number = 0;
	my $i = 0;
		
	$logger->info ( sprintf "\tGetting tags and text for file [%s]...", $keyhash->{TXTfilename} );

	# read text into one long string
	# we have more flexibility doing it this way than with a traditional file handle
	my $text = read_file( $keyhash->{TXTfilename} );
	
	# can we find a line which reads like this:  "Key vocabulary" anywhere in the text?  This exists in the lesson plans documents
	# or if the line is just a URL to a website
	my $start = "Key vocabulary";
	my $end  = "Preparation";
	
	if ( $text =~ m/$start/ )
	{
		$logger->info ( sprintf "\t\tFinding %s in file\n", $start );

		if (my ($vocab) = $text =~ /$start(.*?)$end/s)
		{
			$vocab =~ s/\n//g;	# remove newline characters
			
			$logger->info ( sprintf "\t\t\t[%s] %s [%s]\n", $start, $vocab, $end );
			$keyhash->{WordMetaData}->{KeyVocabulary} = $vocab;
		}
		else
		{
			$logger->info ( sprintf "\t\t\tNothing found\n" );					
		}
	}
	else
	{
		printf "\t\tNo '%s' string found in this file :-(\n", $start;
	}
	
	#printf( "Text of document is [%s]", $text );
	
	# if it's a vocab file, it will have this text, let's extract the items into our CSV file
	my $definition = "Definition and example";
	
	if ( $text =~ m/$definition/ )
	{
		my $temp_text = $text;
		$temp_text =~ s/\n/ /g;	# replace newline characters with SPACE
		
		$logger->info ( sprintf "\t\tFinding %s in file\n", $definition );

		if ( my ($vocab) = $temp_text =~ /Definition and example \(if appropriate\)(.*)/msg )		# test with https://regex101.com/
		{
			$vocab =~ s/ +/ /;	# replace many spaces with one
			
			$logger->info ( sprintf "\t\t\t[%s] %s\n", $definition, $vocab );
			$keyhash->{WordMetaData}->{KeyVocabulary} = $vocab;
		}
		else
		{
			$logger->info ( sprintf "\t\t\tNothing found\n" );					
		}
	}
	else
	{
		printf "\t\tNo '%s' string found in this file :-(\n", $definition;
	}
	
	# count how many \n's we've removed from the end of the file
	do
	{
		$i = chomp $text;	# will be 1 until there are no more to chomp.  annoyingly, chomp only removes one at a time
		$number += $i;
		#printf "\ni:%i	number: %s	text is [%s]", $i, $number, $text;
	}
	while ( $i == 1 );
	
	$logger->info ( sprintf "\t\tFile has [%i] errant newlines at the EOF\n", $number );
	$keyhash->{ErrantCRLFs} = $number;
	
	my @temp_sentences = split( '\n', $text );	
	
	# keep a running string of texts and tasks
	if ( $keyhash->{subtype} eq 'TASKS' )
	{
		$text =~ s/\n\n/\n/g;			# replace two new lines with one (saves space)
		
		my $location = $keyhash->{location};
		$location =~ s/-->|<--/\//g;
		my($filename, $directories, $suffix) = fileparse( $keyhash->{TXTfilename} );	
		
		$AllTasks .= "----> ". $location.'/'.$filename."\n";
		$AllTasks .= $text;
		$AllTasks .= "<---- ". $location.'/'.$filename."\n\n"; 
	}
	elsif ( $keyhash->{subtype} eq 'TEXT' || $keyhash->{subtype} eq 'TAPESCRIPT' )
	{
		$text =~ s/\n\n/\n/g;		# replace two new lines with one (saves space)
		
		$AllTexts .= "----> ". $keyhash->{TXTfilename}."\n";
		$AllTexts .= $text;
		$AllTexts .= "<---- ". $keyhash->{TXTfilename}."\n\n"; 	
	}
	
	#
	#
	#
	
	foreach my $sentence ( @temp_sentences )
	{
		$sentence =~ s/^\s+|\s+$//g;	# remove leading and trailing spaces
		$sentence =~ s/  / /g;			# replace two spaces with one
		$sentence =~ s/\t//g;			# get rid of tabs

		if ( length $sentence > 1 )
		{	
			# can we find a line which reads like this:  "Taken from: Inside Reading 2" anywhere in the text?
			# or if the line is just a URL to a website
			if ( $sentence =~ m/^(Taken from:|Adapted from|Adapted by|Retrieved from|Source)/ || $sentence =~ m/^(http:\/\/www\.|https:\/\/www\.|http:\/\/|https:\/\/)?[a-z0-9]+([\-\.]{1}[a-z0-9]+)*\.[a-z]{2,5}(:[0-9]{1,5})?(\/.*)?$/ )
			{
				$keyhash->{WordMetaData}->{source} = $sentence;
				next;		# don't store this link in the 'textofdocument' member (and don't add a full stop at the end either)						
			}	
			
			# if the sentence doesn't end with one of these, add a full stop
			unless( substr($sentence, -1) eq "." || 
					substr($sentence, -1) eq "'" || 
					substr($sentence, -1) eq "\"" || 
					substr($sentence, -1) eq "!" || 
					substr($sentence, -1) eq ";" )
			{
				$sentence .= ".";
			}
			
			#printf "[%s]\n", $sentence;
			$keyhash->{TextofDocument} .= $sentence . "\n";
			$count++;
		}
	}
		
	my $contents = $keyhash->{TextofDocument};
	$contents =~ s/\[HYPERLINK: \S+\]//g;
	$contents =~ tr/0-9a-z A-Z.'\012\015-//dc;

	$keyhash->{PrettyTextofDocument} = lc $contents;
	chomp $keyhash->{PrettyTextofDocument};
	
	
	# We don't currently use this output so this is commented out for now.
	# Add part of speech tags to a text
	# Tags are explained here http://cpansearch.perl.org/src/ACOBURN/Lingua-EN-Tagger-0.23/README
	# Create a parser object
	#my $tagger = new Lingua::EN::Tagger;

	#$keyhash->{TaggedTextofDocument} = $tagger->add_tags($keyhash->{PrettyTextofDocument});
	
	#if ( open( my $fh, '>', $keyhash->{TaggedTXTfilename} ) )
	#{
	#	print $fh $keyhash->{TaggedTextofDocument};
	#	close $fh;		
	#}
	
	# save this polished text to a 'pretty' file
	#if ( open( my $fh, '>', $keyhash->{PrettyTXTfilename} ) )
	#{
	#	print $fh $keyhash->{PrettyTextofDocument};
	#	close $fh;
	#}
	
	return;	
}



# ReadSpellingAndGrammarFileIntoArray
# Reads the CSV output produced by the external grammar_check.py / LanguageTool
# process back into the @SpellingAndGrammarProblems array for later writing
# to the Excel workbook.
sub ReadSpellingAndGrammarFileIntoArray
{
	my $keyhash = shift;
	my $i = 0;

	open my $handle, '<', $keyhash->{SpellingAndGrammarErrorsfilename};

	while( <$handle> )
	{	
		$i++;
		my $line = $_;
		chomp $line;
		$line =~ s/\t//g;		# remove tab characters
		$line =~ s/\s+$//;		# remove trailing spaces

		my ( $ruleId, $message, $replacements, $context, $offset, $errorLength, $category, $ruleIssueType ) = split( "!!", $line );	# CSV file

		my %ErrorEntry = (
			original_filename	=> $keyhash->{filename},
			filename			=> $keyhash->{TXTfilename},
			ruleId				=> $ruleId,
			message  			=> $message,
			replacements 		=> $replacements,
			context				=> $context,
			offset				=> $offset,
			errorLength			=> $errorLength,
			category  			=> $category,
			ruleIssueType		=> $ruleIssueType );
	
		push @SpellingAndGrammarProblems, \%ErrorEntry if ( $i > 1 );	# the first row will be a header row.
		
		$keyhash->{SpellingMistakes}->{Punctuation}++ 			if $category eq "PUNCTUATION";
		$keyhash->{SpellingMistakes}->{Typography}++ 			if $category eq "TYPOGRAPHY";
		$keyhash->{SpellingMistakes}->{Typos}++ 				if $category eq "TYPOS";
		$keyhash->{SpellingMistakes}->{Casing}++ 				if $category eq "CASING";
		$keyhash->{SpellingMistakes}->{Grammar}++ 				if $category eq "GRAMMAR";
		$keyhash->{SpellingMistakes}->{Redundancy}++ 			if $category eq "REDUNDANCY";
		$keyhash->{SpellingMistakes}->{Style}++ 				if $category eq "STYLE";
		$keyhash->{SpellingMistakes}->{ConfusedWords}++ 		if $category eq "CONFUSED_WORDS";
		$keyhash->{SpellingMistakes}->{Semantics}++ 			if $category eq "SEMANTICS";
		$keyhash->{SpellingMistakes}->{NonstandardPhrases}++ 	if $category eq "NONSTANDARD_PHRASES";
		$keyhash->{SpellingMistakes}->{Collocations}++ 			if $category eq "COLLOCATIONS";
		$keyhash->{SpellingMistakes}->{Miscellaneous}++ 		if $category eq "MISC"

#			$logger->info ( "RULE           | MESSAGE              | CONTEXT" ) if ( $i == 1 );
#			my $string = sprintf "%-20s	| %-40s	| %-40s\n", $ruleId, $message, $context;
#			$logger->info ( $string );
	}
	
	close $handle;
	
	$logger->info ( sprintf "\t\tRead file ok.  File has [%i errors: PU:%02i TY:%02i TO:%02i CA:%02i GR:%02i RE:%02i ST:%02i CW:%02i SE:%02i NP:%02i CO:%02i MI:%02i]\n", $i, $keyhash->{SpellingMistakes}->{Punctuation}, $keyhash->{SpellingMistakes}->{Typography}, $keyhash->{SpellingMistakes}->{Typos}, $keyhash->{SpellingMistakes}->{Casing}, $keyhash->{SpellingMistakes}->{Grammar}, $keyhash->{SpellingMistakes}->{Redundancy}, $keyhash->{SpellingMistakes}->{Style}, $keyhash->{SpellingMistakes}->{ConfusedWords}, $keyhash->{SpellingMistakes}->{Semantics}, $keyhash->{SpellingMistakes}->{NonstandardPhrases}, $keyhash->{SpellingMistakes}->{Collocations}, $keyhash->{SpellingMistakes}->{Miscellaneous} );
	
	CopyFileToSaveDirectory( 0, $keyhash->{SpellingAndGrammarErrorsfilename}, $keyhash->{MD5}, ".csv" );

	return $i;
}


# CheckSpellingAndGrammar
# Invokes the external Python grammar_check.py script (backed by LanguageTool)
# on the plain-text version of a resource. Waits for the process to complete,
# then calls ReadSpellingAndGrammarFileIntoArray to load the results.
sub CheckSpellingAndGrammar
{
	my $keyhash = shift;
	my $i = 0;
	
	if ( not -e $keyhash->{SpellingAndGrammarErrorsfilename} )	# couldn't find it in savefiles
	{
		$logger->info ( sprintf "\t\tFile [%s] does not exist in our savefiles (or MD5 changed).  Creating...\n", $keyhash->{SpellingAndGrammarErrorsfilename} );

		my $Started = [gettimeofday];		# start the timer!
		$logger->info ( sprintf "\tRunning grammar_check.py [%s] -> [%s]\n", $keyhash->{TXTfilename}, $keyhash->{SpellingAndGrammarErrorsfilename} );
		
		if( system ( "/opt/homebrew/bin/python3", "grammar_check.py", "-i", $keyhash->{TXTfilename}, "-o", $keyhash->{SpellingAndGrammarErrorsfilename} ) == 0 )	# success!
		{
			if ( -e $keyhash->{SpellingAndGrammarErrorsfilename} )
			{
				# open the file we've just had created, and read contents into @SpellingAndGrammarProblems
				$i = ReadSpellingAndGrammarFileIntoArray( $keyhash );

				# found the spreadsheet for this document in the savefiles directory
				$logger->info ( sprintf "\t\tGenerated file [%s] with [%i] errors.\n", $keyhash->{SpellingAndGrammarErrorsfilename}, $i );
			}
			else
			{
				$logger->info ( sprintf "\tExpected to find [%s] but couldn't.\n", $keyhash->{SpellingAndGrammarErrorsfilename} );
			}
		
			$keyhash->{NumberOfSpellingAndGrammarErrors} = $i;
			$logger->info ( sprintf "\t\tDone. Found %i possible errors.  Took %is\n", $i, tv_interval ( $Started ) ) if ( -e $keyhash->{SpellingAndGrammarErrorsfilename} );
		}
	}
	else
	{
		# open the file provided for us, and read contents into @SpellingAndGrammarProblems
		$i = ReadSpellingAndGrammarFileIntoArray( $keyhash );
		
		# found the spreadsheet for this document in the savefiles directory
		$logger->info ( sprintf "\t\tFound file [%s] with [%i] errors already generated for us :-).  Moving on...\n", $keyhash->{SpellingAndGrammarErrorsfilename}, $i );
	}

	return;
}



# GetReadabilityStats
# Uses Lingua::EN::Fathom to compute readability metrics (Flesch Reading Ease,
# Flesch-Kincaid Grade Level, Gunning Fog Index) and surface-level text
# statistics (sentence count, word count, syllable count) for the plain-text
# content of a resource. Results are stored in the resource hash.
sub GetReadabilityStats
{
	my $keyhash = shift;
	$logger->info ( sprintf "Getting readability info for file [%s]...", $keyhash->{TXTfilename} );

	my $text = Lingua::EN::Fathom->new();
    $text->analyse_file( $keyhash->{TXTfilename}, 0 );
    #$text->analyse_block( $keyhash->{TextofDocument}, 0 );

	$keyhash->{Readability}{NumChars} 				= $text->num_chars;
	$keyhash->{Readability}{NumWords} 				= $text->num_words;
	$keyhash->{Readability}{PercentComplexWords} 	= sprintf "%.02f", $text->percent_complex_words;
	$keyhash->{Readability}{NumSentences} 			= $text->num_sentences;
	$keyhash->{Readability}{NumTextLines} 			= $text->num_text_lines;
	$keyhash->{Readability}{NumBlankLines} 			= $text->num_blank_lines;
	$keyhash->{Readability}{NumParagraphs} 			= $text->num_paragraphs;
	$keyhash->{Readability}{AverageSyllablesPerWord} = sprintf "%.02f", $text->syllables_per_word;
	$keyhash->{Readability}{AverageWordsPerSentence} = sprintf "%.02f", $text->words_per_sentence;

	# Save these
	$keyhash->{Readability}->{EstimatedReadingTime} = sprintf "%.02f", $keyhash->{WordListMetaData}->{WordCount}/$AVERAGE_READING_SPEED; 
	$keyhash->{Readability}->{Flesch} 		 = sprintf "%.02f", $text->flesch;
	$keyhash->{Readability}->{FleschKincaid} = sprintf "%.02f", $text->kincaid;
	$keyhash->{Readability}->{GunningFog} 	 = sprintf "%.02f", $text->fog;

	if ( $archive_root eq 'Reading 1' )
	{
		if ( $keyhash->{Readability}->{FleschKincaid} < $READING_LEVEL_1_LOWER_FLEISCH_KINCAID_LIMIT )
		{
			$keyhash->{Readability}->{FleschKincaidDescription} = 'EA'.'A' x ( $READING_LEVEL_1_LOWER_FLEISCH_KINCAID_LIMIT - $keyhash->{Readability}->{FleschKincaid} ).'SY';
		}
		elsif ( $keyhash->{Readability}->{FleschKincaid} > $READING_LEVEL_1_HIGHER_FLEISCH_KINCAID_LIMIT )
		{
			$keyhash->{Readability}->{FleschKincaidDescription} = 'HA'.'A' x ( $keyhash->{Readability}->{FleschKincaid} - $READING_LEVEL_1_HIGHER_FLEISCH_KINCAID_LIMIT ).'RD';			
		}
	}
	elsif ( $archive_root eq 'Reading 2' )
	{
		if ( $keyhash->{Readability}->{FleschKincaid} < $READING_LEVEL_2_LOWER_FLEISCH_KINCAID_LIMIT )
		{
			$keyhash->{Readability}->{FleschKincaidDescription} = 'EA'.'A' x ( $READING_LEVEL_2_LOWER_FLEISCH_KINCAID_LIMIT - $keyhash->{Readability}->{FleschKincaid}).'SY';
		}
		elsif ( $keyhash->{Readability}->{FleschKincaid} > $READING_LEVEL_2_HIGHER_FLEISCH_KINCAID_LIMIT )
		{
			$keyhash->{Readability}->{FleschKincaidDescription} = 'HA'.'A' x ( $keyhash->{Readability}->{FleschKincaid} - $READING_LEVEL_2_HIGHER_FLEISCH_KINCAID_LIMIT ).'RD';			
		}
	}
	elsif ( $archive_root eq 'Reading 3' )
	{
		if ( $keyhash->{Readability}->{FleschKincaid} < $READING_LEVEL_3_LOWER_FLEISCH_KINCAID_LIMIT )
		{
			$keyhash->{Readability}->{FleschKincaidDescription} = 'EA'.'A' x ( $READING_LEVEL_3_LOWER_FLEISCH_KINCAID_LIMIT - $keyhash->{Readability}->{FleschKincaid} ).'SY';
		}
		elsif ( $keyhash->{Readability}->{FleschKincaid} > $READING_LEVEL_3_HIGHER_FLEISCH_KINCAID_LIMIT )
		{
			$keyhash->{Readability}->{FleschKincaidDescription} = 'HA'.'A' x ( $keyhash->{Readability}->{FleschKincaid} - $READING_LEVEL_3_HIGHER_FLEISCH_KINCAID_LIMIT ).'RD';			
		}
	}
	elsif ( $archive_root eq 'Reading 4' )
	{
		if ( $keyhash->{Readability}->{FleschKincaid} < $READING_LEVEL_4_LOWER_FLEISCH_KINCAID_LIMIT )
		{
			$keyhash->{Readability}->{FleschKincaidDescription} = 'EA'.'A' x ( $READING_LEVEL_4_LOWER_FLEISCH_KINCAID_LIMIT - $keyhash->{Readability}->{FleschKincaid} ).'SY';
		}
		elsif ( $keyhash->{Readability}->{FleschKincaid} > $READING_LEVEL_4_HIGHER_FLEISCH_KINCAID_LIMIT )
		{
			$keyhash->{Readability}->{FleschKincaidDescription} = 'HA'.'A' x ( $keyhash->{Readability}->{FleschKincaid} - $READING_LEVEL_4_HIGHER_FLEISCH_KINCAID_LIMIT ).'RD';			
		}
	}
	
	
	# A hash of all unique words and the number of times they occur is generated.
	my %words = $text->unique_words;
	foreach my $word ( sort keys %words )
	{
		$keyhash->{Readability}->{ReadabilityReport} .= sprintf "%4i\t%s\n", $words{$word}, $word;
	}
	
	if ( open my $fh, '>', $keyhash->{ReadabilityFilename} )
	{	
		print $fh $keyhash->{Readability}->{ReadabilityReport};
		close $fh;
	}

	return;
}


# GetPDFMetadata
# Opens a PDF file with PDF::API2 and extracts document metadata (title,
# author, subject, creator, creation date) along with the page count.
# Stores the values in the resource hash under the MetaData key.
sub GetPDFMetadata
{
	my $keyhash = shift;
	
	my $pdf = PDF::API2->open( $keyhash->{filename} );
	my %pdfinfo = $pdf->info;
	my $string;

	$keyhash->{MetaData}->{createdby} = $pdfinfo{Author};
	if ( defined $pdfinfo{CreationDate} && $pdfinfo{CreationDate} ne '' )
	{
		$keyhash->{MetaData}->{createddate} = sprintf "%s-%s-%sT", substr( $pdfinfo{CreationDate}, 2, 4 ), substr( $pdfinfo{CreationDate}, 6, 2 ), substr( $pdfinfo{CreationDate}, 8, 2 );		# sprintf into strings otherwise we lose the leading 0s
	}

	$keyhash->{MetaData}->{lastmodifiedby} = $pdfinfo{Author};
	if ( defined $pdfinfo{ModDate} && $pdfinfo{ModDate} ne '' )
	{
		$keyhash->{MetaData}->{lastmodifieddate} = sprintf "%s-%s-%sT", substr( $pdfinfo{ModDate}, 2, 4 ), substr( $pdfinfo{ModDate}, 6, 2 ), substr( $pdfinfo{ModDate}, 8, 2 ); 		# sprintf into strings otherwise we lose the leading 0s
	}
	
	my $now = localtime;
	#my $tp = Time::Piece->strptime( $keyhash->{MetaData}->{lastmodifieddate},"%Y-%m-%dT" ) if defined ( $keyhash->{MetaData}->{lastmodifieddate} );
	#my $diff = $now - $tp;
	
	# we don't have last modified dates for PDFs, set this to zero
	$keyhash->{MetaData}->{AgeInDaysSinceLastEdit} = 0;	#	int $diff->days;		# round this
	
	$keyhash->{MetaData}->{title} = $pdfinfo{Title};
	$keyhash->{MetaData}->{pages} = $pdf->pages();

	# Irrespective of the subtype, add these

	$keyhash->{SummaryText} .= sprintf "		** PDF Metadata **\n";
	$keyhash->{SummaryText} .= sprintf "		Created By:	%s\n", 	  $keyhash->{MetaData}->{createdby} 		if ( defined $keyhash->{MetaData}->{createdby} && $keyhash->{MetaData}->{createdby} ne '' );	
	$keyhash->{SummaryText} .= sprintf "		Created Date:	%s\n", $keyhash->{MetaData}->{createddate} 		if ( defined $keyhash->{MetaData}->{createddate} && $keyhash->{MetaData}->{createddate} ne '' );	
	$keyhash->{SummaryText} .= sprintf "		Last Modified By:	%s\n", $keyhash->{MetaData}->{lastmodifiedby} 	if ( defined $keyhash->{MetaData}->{lastmodifiedby} && $keyhash->{MetaData}->{lastmodifiedby} ne '' );
	$keyhash->{SummaryText} .= sprintf "		Last Modified Date:	%s (%i days old)\n", $keyhash->{MetaData}->{lastmodifieddate}, $keyhash->{MetaData}->{AgeInDaysSinceLastEdit} 	if ( defined $keyhash->{MetaData}->{lastmodifieddate} && $keyhash->{MetaData}->{lastmodifieddate} ne '' );
	$keyhash->{SummaryText} .= sprintf "\n";
	$keyhash->{SummaryText} .= sprintf "		Document has:\n";	
	$keyhash->{SummaryText} .= sprintf "			%i pages.\n", 	$keyhash->{MetaData}->{pages} if ( defined $keyhash->{MetaData}->{pages} && $keyhash->{MetaData}->{pages} ne '' );

	$keyhash->{subtype} = 'EXCLUDED' if ( $keyhash->{MetaData}->{pages} >= $MAXIMUM_PDF_PAGES && $keyhash->{filename} !~ /Important Vocabulary At/i )
}



# GetOfficeMetadata
# Unzips the docProps/core.xml and docProps/app.xml files from a DOCX or PPTX
# archive and parses them with XML::Simple to extract document metadata
# (title, author, last modified by, revision, page/slide count). Stores the
# values in the resource hash under the MetaData key.
sub GetOfficeMetadata
{
	my $keyhash = shift;
	my $filename = '';
	my $data;
	my @focus;
	my $text = "";
	
	$logger->info ( sprintf "\tGetting office metadata file [%s]...", $keyhash->{TXTfilename} );

	if( $keyhash->{exists} && -e $keyhash->{filename} && -s $keyhash->{filename} > 0 )
	{
		my $zip = Archive::Zip->new($keyhash->{filename});
		
		foreach my $member ($zip->members)
		{
	    	next if $member->isDirectory;
	    	#(my $extractName = $member->fileName) =~ s{.*/}{};
	    	my $extractName = $member->fileName;
			$member->extractToFileNamed($keyhash->{directory}."/".$extractName);
		}
		
		# create object
		my $xml = new XML::Simple;

		# read XML file.  this exists for .docx, .xlsx, .pptx
		$filename = $keyhash->{directory}."/docProps/core.xml";
		
		if ( -e $filename )
		{
			$data = $xml->XMLin( $filename );
		
			$keyhash->{MetaData}->{createdby} = $data->{'dc:creator'} if ( defined $data->{'dc:creator'} );
			$keyhash->{MetaData}->{createddate} = $data->{'dcterms:created'}->{'content'} if ( defined $data->{'dcterms:created'}->{'content'} );
		
			$keyhash->{MetaData}->{lastmodifiedby} = $data->{'cp:lastModifiedBy'} if ( defined $data->{'cp:lastModifiedBy'} );
			$keyhash->{MetaData}->{lastmodifieddate} = $data->{'dcterms:modified'}->{'content'}  if ( defined $data->{'dcterms:modified'}->{'content'} );			
		
			my $now = localtime;
			my $tp = Time::Piece->strptime( $keyhash->{MetaData}->{lastmodifieddate},"%Y-%m-%dT%H:%M:%SZ" );
			
			my $diff = $now - $tp;
			$logger->info ( sprintf "\t\tFile [%s] is unchanged in [%i] days", $keyhash->{filename}, $diff->days );
			
			$keyhash->{MetaData}->{AgeInDaysSinceLastEdit} = int $diff->days;
		}
		
		# read XML file.  this exists for .docx, .xlsx, .pptx
		$filename = $keyhash->{directory}."/docProps/app.xml";
		
		if ( -e $filename && ( $keyhash->{type} eq 'docx' || $keyhash->{type} eq 'pptx' ) )
		{
			$data = $xml->XMLin( $filename );

			# these only exists for .docx OR .pptx files (and even then, not always, some may exist and some may not)		
			$keyhash->{MetaData}->{totaleditingtime} = $data->{'TotalTime'} if ( defined $data->{'TotalTime'} );
			$keyhash->{MetaData}->{revisionnumber} = $data->{'cp:revision'} if ( defined $data->{'cp:revision'} );
			$keyhash->{MetaData}->{pages} = ( defined $data->{'Pages'} ) ? $data->{'Pages'} : '';		
			$keyhash->{WordMetaData}->{paragraphs} = ( defined $data->{'Paragraphs'} ) ? $data->{'Paragraphs'} : '';
			$keyhash->{WordMetaData}->{lines} = ( defined $data->{'Lines'} ) ? $data->{'Lines'} : '';
			$keyhash->{WordMetaData}->{words} = ( defined $data->{'Words'} ) ? $data->{'Words'} : 0;
			$keyhash->{WordMetaData}->{characters} = ( defined $data->{'Characters'} ) ? $data->{'Characters'} : '';
			
			# get a count of the number of slides for presentations and store this in the 'pages' property 
			if ( $keyhash->{type} eq 'pptx' )
			{
				# if it's a presentation get the number of slides
				my $slidepath = "$keyhash->{directory}ppt/slides/slide*.xml";
				# $logger->info ( sprintf "slidepath is [%s]", $slidepath );
				
				my @count_of_slides = glob "'$slidepath'";		# don't fuck with the quotes here:  https://stackoverflow.com/questions/32260485/using-perl-glob-with-spaces-in-the-pattern
				$keyhash->{MetaData}->{pages} = scalar @count_of_slides;
							
				$logger->info ( sprintf "\t\tPresentation [%s] has [%i] slides", $keyhash->{filename}, scalar @count_of_slides );
				
				#foreach my $slide ( sort @count_of_slides )
				#{
				#	$logger->info ( sprintf "\t\t[%s]", $slide );
				#}
			}		
		}
		
				
		# a document might have many headers
		foreach my $numheader ( 1..10 )
		{
			# read XML file.  this exists for .docx files
			$filename = sprintf "%s/word/header%i.xml", $keyhash->{directory}, $numheader;
		
			if ( -e $filename )
			{
				$data = $xml->XMLin( $filename, forcearray => 1 ); 
				
				for my $level1 ( @{ $data->{'w:p'} } )
				{
					$text = "";		# reset this

					for my $level2 ( @{ $level1->{'w:r'} } )
					{
						for my $level3 ( @{ $level2->{'w:t'} } )
						{
							if ( ref( $level3 ) eq 'HASH' && defined $level3->{'content'} )
							{
								$text .= $level3->{'content'};
							}
							else
							{
								$text .= $level3;
							}
						}
					}
										
					if ( $text =~ /^Focus/i || $text =~ /^Lesson/i || $text =~ /^Skill/i )	# case insensitive
					{
						# beautify it, make the Lesson Focus Sentence Case
						$text =~ s/([\w']+)/\u\L$1/g;
						$logger->info ( sprintf "Lesson focus is: [%s]\n", $text );
						push @focus, $text;
					}
				}
			}
		}
		
		$keyhash->{MetaData}->{LessonFocus} = \@focus;		# add a reference to this array to our hash
						
		#
		#
		#
		
		if ( $archive_root eq 'Reading 1' )
		{
			if ( $keyhash->{WordMetaData}->{words} < $READING_LEVEL_1_LOWER_TEXT_LIMIT )
			{
				$keyhash->{Readability}->{wordsdescription} = 'SHO'.'O' x ( ( $READING_LEVEL_1_LOWER_TEXT_LIMIT-$keyhash->{WordMetaData}->{words}) / 50 ).'RT';
			}
			elsif ( $keyhash->{WordMetaData}->{words} > $READING_LEVEL_1_HIGHER_TEXT_LIMIT )
			{
				$keyhash->{Readability}->{wordsdescription} = 'LO'.'O' x ( ( $keyhash->{WordMetaData}->{words}-$READING_LEVEL_1_HIGHER_TEXT_LIMIT) / 50 ).'NG';
			}
		}
		elsif ( $archive_root eq 'Reading 2' )
		{
			if ( $keyhash->{WordMetaData}->{words} < $READING_LEVEL_2_LOWER_TEXT_LIMIT )
			{
				$keyhash->{Readability}->{wordsdescription} = 'SHO'.'O' x ( ( $READING_LEVEL_2_LOWER_TEXT_LIMIT-$keyhash->{WordMetaData}->{words}) / 50 ).'RT';
			}
			elsif ( $keyhash->{WordMetaData}->{words} > $READING_LEVEL_2_HIGHER_TEXT_LIMIT )
			{
				$keyhash->{Readability}->{wordsdescription} = 'LO'.'O' x ( ( $keyhash->{WordMetaData}->{words}-$READING_LEVEL_2_HIGHER_TEXT_LIMIT) / 50 ).'NG';
			}
		}
		elsif ( $archive_root eq 'Reading 3' )
		{
			if ( $keyhash->{WordMetaData}->{words} < $READING_LEVEL_3_LOWER_TEXT_LIMIT )
			{
				$keyhash->{Readability}->{wordsdescription} = 'SHO'.'O' x ( ( $READING_LEVEL_3_LOWER_TEXT_LIMIT-$keyhash->{WordMetaData}->{words}) / 50 ).'RT';
			}
			elsif ( $keyhash->{WordMetaData}->{words} > $READING_LEVEL_3_HIGHER_TEXT_LIMIT )
			{
				$keyhash->{Readability}->{wordsdescription} = 'LO'.'O' x ( ( $keyhash->{WordMetaData}->{words}-$READING_LEVEL_3_HIGHER_TEXT_LIMIT) / 50 ).'NG';
			}
		}
		elsif ( $archive_root eq 'Reading 4' )
		{
			if ( $keyhash->{WordMetaData}->{words} < $READING_LEVEL_4_LOWER_TEXT_LIMIT )
			{
				$keyhash->{Readability}->{wordsdescription} = 'SHO'.'O' x ( ( $READING_LEVEL_4_LOWER_TEXT_LIMIT-$keyhash->{WordMetaData}->{words}) / 50 ).'RT';
			}
			elsif ( $keyhash->{WordMetaData}->{words} > $READING_LEVEL_4_HIGHER_TEXT_LIMIT )
			{
				$keyhash->{Readability}->{wordsdescription} = 'LO'.'O' x ( ( $keyhash->{WordMetaData}->{words}-$READING_LEVEL_4_HIGHER_TEXT_LIMIT) / 50 ).'NG';
			}			
		}
		
		# read XML file.  this exists for .docx ONLY
		$filename = $keyhash->{directory}."/word/document.xml";
		
		if ( -e $filename )
		{
			$data = $xml->XMLin( $filename );
			#print Dumper( $data );
			
			# for example: <w:pgSz w:w="16820" w:h="11900" w:orient="landscape"/>
			if ( defined $data->{'w:body'}->{'w:sectPr'}->{'w:pgSz'}->{'w:h'} && defined $data->{'w:body'}->{'w:sectPr'}->{'w:pgSz'}->{'w:w'} )
			{
				my $height = $data->{'w:body'}->{'w:sectPr'}->{'w:pgSz'}->{'w:h'};
				my $width  = $data->{'w:body'}->{'w:sectPr'}->{'w:pgSz'}->{'w:w'};
				
				# it appears that different versions of word have different A4 sizes, also LibreOffice.  Officially, A4 = 16838 x 11906
				if ( ( $height == 16840 && $width == 11900 ) || ( $height == 16820 && $width == 11900 ) || ( $height == 16838 && $width == 11906 ) )		# https://startbigthinksmall.wordpress.com/2010/01/04/points-inches-and-emus-measuring-units-in-office-open-xml/
				{
					$keyhash->{WordMetaData}->{PageSize} = 'A4';
				}
				elsif ( ( $height == 11900 && $width == 16840 ) || ( $height == 11906 && $width == 16838 ) )
				{
					$keyhash->{WordMetaData}->{PageSize} = 'A4'; 	# A4 landscape
				}
				elsif ( $height == 15840 && $width == 12240 )
				{
					$keyhash->{WordMetaData}->{PageSize} = 'US Letter';
				}
				elsif ( $height == 12240 && $width == 15840 )
				{
					$keyhash->{WordMetaData}->{PageSize} = 'US Letter';	 # US Letter landscape
				}
				elsif ( $height == 23800 && $width == 16820 )
				{
					$keyhash->{WordMetaData}->{PageSize} = 'A3';
				}
				elsif ( $height == 16820 && $width == 23800 )	
				{
					$keyhash->{WordMetaData}->{PageSize} = 'A3';	# A3 landscape
				}
				elsif ( $height == 11900 && $width == 8380 )
				{
					$keyhash->{WordMetaData}->{PageSize} = 'A5';
				}
				elsif ( $height == 8380 && $width == 11900 )
				{
					$keyhash->{WordMetaData}->{PageSize} = 'A5';	# A5 landscape
				}
				elsif ( $height == 540 && $width == 960 )
				{
					$keyhash->{WordMetaData}->{PageSize} = 'Widescreen PowerPoint Slide (16:9 aspect ratio)';	# Widescreen PowerPoint Slide
				}
				elsif ( $height == 540 && $width == 720 )
				{
					$keyhash->{WordMetaData}->{PageSize} = 'Standard-width PowerPoint Slide (4:3 aspect ratio)';	# Widescreen PowerPoint Slide
				}
				elsif ( $height == 648 && $width == 432 )
				{
					$keyhash->{WordMetaData}->{PageSize} = 'US 1 Catalog envelope';	# US 1 Catalog envelope
				}
				else
				{
					$keyhash->{WordMetaData}->{PageSize} = sprintf "??? height: %i; width: %i", $height, $width;
				}
			}
		
			# by default this is 'portrait'
			$keyhash->{WordMetaData}->{PageOrientation} = 'portrait';
			$keyhash->{WordMetaData}->{PageOrientation} = $data->{'w:body'}->{'w:sectPr'}->{'w:pgSz'}->{'w:orient'} if ( defined $data->{'w:body'}->{'w:sectPr'}->{'w:pgSz'}->{'w:orient'} );
			
			#w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" w:header="708" w:footer="708" w:gutter="0"
			if ( defined $data->{'w:body'}->{'w:sectPr'}->{'w:pgMar'}->{'w:top'} && defined $data->{'w:body'}->{'w:sectPr'}->{'w:pgMar'}->{'w:right'} && defined $data->{'w:body'}->{'w:sectPr'}->{'w:pgMar'}->{'w:bottom'} && defined $data->{'w:body'}->{'w:sectPr'}->{'w:pgMar'}->{'w:left'} )
			{
				my $top = $data->{'w:body'}->{'w:sectPr'}->{'w:pgMar'}->{'w:top'};
				my $right  = $data->{'w:body'}->{'w:sectPr'}->{'w:pgMar'}->{'w:right'};
				my $bottom = $data->{'w:body'}->{'w:sectPr'}->{'w:pgMar'}->{'w:bottom'};
				my $left = $data->{'w:body'}->{'w:sectPr'}->{'w:pgMar'}->{'w:left'};
				
				# narrow borders
				if ( $top == 720 && $right == 720 && $bottom == 720 && $left == 720 )
				{
					$keyhash->{WordMetaData}->{PageBorders} = 'Narrow';					
				}
				elsif ( $top == 1440 && $right == 1440 && $bottom == 1440 && $left == 1440 )
				{
					$keyhash->{WordMetaData}->{PageBorders} = 'Normal';
				}
				elsif ( $top == 1440 && $right == 1080 && $bottom == 1440 && $left == 1080 )
				{
					$keyhash->{WordMetaData}->{PageBorders} = 'Moderate';
				}
				elsif ( $top == 1440 && $right == 2880 && $bottom == 1440 && $left == 2880 )
				{
					$keyhash->{WordMetaData}->{PageBorders} = 'Wide';
				}
				else
				{
					$keyhash->{WordMetaData}->{PageBorders} = sprintf "Custom: %i, %i, %i, %i", $top, $right, $bottom, $left;
				}
			}		
		}
		
		#
		# Save the info to the summary property
		#
		
		$keyhash->{SummaryText} .= sprintf "		** Office Metadata **\n";
		$keyhash->{SummaryText} .= sprintf "		Created By:	%s\n", 	  $keyhash->{MetaData}->{createdby} 		if ( defined $keyhash->{MetaData}->{createdby} && $keyhash->{MetaData}->{createdby} ne '' );	
		$keyhash->{SummaryText} .= sprintf "		Created Date:	%s\n", $keyhash->{MetaData}->{createddate} 		if ( defined $keyhash->{MetaData}->{createddate} && $keyhash->{MetaData}->{createddate} ne '' );	
		$keyhash->{SummaryText} .= sprintf "		Last Modified By:	%s\n", $keyhash->{MetaData}->{lastmodifiedby} 	if ( defined $keyhash->{MetaData}->{lastmodifiedby} && $keyhash->{MetaData}->{lastmodifiedby} ne '' );
		$keyhash->{SummaryText} .= sprintf "		Last Modified Date:	%s	(%i days ago)\n", $keyhash->{MetaData}->{lastmodifieddate}, $keyhash->{MetaData}->{AgeInDaysSinceLastEdit} if ( defined $keyhash->{MetaData}->{lastmodifieddate} && $keyhash->{MetaData}->{lastmodifieddate} ne '' );
		$keyhash->{SummaryText} .= sprintf "\n";
		$keyhash->{SummaryText} .= sprintf "		** Document Statistics **\n";
		$keyhash->{SummaryText} .= sprintf "		Document is:\n";	
		$keyhash->{SummaryText} .= sprintf "			%s page size.\n", 	$keyhash->{WordMetaData}->{PageSize}	if ( defined $keyhash->{WordMetaData}->{PageSize} && $keyhash->{WordMetaData}->{PageSize} ne '' );
		$keyhash->{SummaryText} .= sprintf "			%s page orientation.\n",$keyhash->{WordMetaData}->{PageOrientation}	if ( defined $keyhash->{WordMetaData}->{PageOrientation} && $keyhash->{WordMetaData}->{PageOrientation} ne '' );
		$keyhash->{SummaryText} .= sprintf "			%s page borders.\n",$keyhash->{WordMetaData}->{PageBorders}	if ( defined $keyhash->{WordMetaData}->{PageBorders} && $keyhash->{WordMetaData}->{PageBorders} ne '' );
		$keyhash->{SummaryText} .= sprintf "\n";		
		$keyhash->{SummaryText} .= sprintf "		Document has:\n";	
		$keyhash->{SummaryText} .= sprintf "			%i pages.\n", 	$keyhash->{MetaData}->{pages}				if ( defined $keyhash->{MetaData}->{pages} && $keyhash->{MetaData}->{pages} ne '' );
		$keyhash->{SummaryText} .= sprintf "			%i paragraphs.\n",$keyhash->{WordMetaData}->{paragraphs}	if ( defined $keyhash->{WordMetaData}->{paragraphs} && $keyhash->{WordMetaData}->{paragraphs} ne '' );
		$keyhash->{SummaryText} .= sprintf "			%i lines.\n", 	$keyhash->{WordMetaData}->{lines}			if ( defined $keyhash->{WordMetaData}->{lines} && $keyhash->{WordMetaData}->{lines} ne '' );
		$keyhash->{SummaryText} .= sprintf "			%i words.\n", 	$keyhash->{WordMetaData}->{words}			if ( defined $keyhash->{WordMetaData}->{words} && $keyhash->{WordMetaData}->{words} ne '' );
		$keyhash->{SummaryText} .= sprintf "			%i characters.\n",$keyhash->{WordMetaData}->{characters}	if ( defined $keyhash->{WordMetaData}->{characters} && $keyhash->{WordMetaData}->{characters} ne '' );
		$keyhash->{SummaryText} .= "\n";

		my @focusarray = @{ $keyhash->{MetaData}->{LessonFocus} };
		$keyhash->{SummaryText} .= sprintf "		Focus:\n", 									if ( scalar( @focusarray ) > 0 );
		$keyhash->{SummaryText} .= sprintf "			%s\n", join( "\n\t\t\t", @focusarray ) 	if ( scalar( @focusarray ) > 0 );
		$keyhash->{SummaryText} .= "\n"															if ( scalar( @focusarray ) > 0 );
		
		$keyhash->{SummaryText} .= sprintf "		Total Editing Time:	%i.\n", $keyhash->{WordMetaData}->{totaleditingtime}			if ( defined $keyhash->{WordMetaData}->{totaleditingtime} && $keyhash->{WordMetaData}->{totaleditingtime} ne '' );
		$keyhash->{SummaryText} .= sprintf "		Revision number:	%i characters.\n", $keyhash->{WordMetaData}->{revisionnumber} 	if ( defined $keyhash->{WordMetaData}->{revisionnumber} && $keyhash->{WordMetaData}->{revisionnumber} ne '' );
	}
	else
	{
		$logger->info ( sprintf "File %s does not exist\n", $keyhash->{filename} );
		push @unabletoparseDOCX, $keyhash->{filename};
	}
	
	return;
}


# PrepareCEFRWordList
# Reads the CEFR-banded vocabulary files (A1 through C2) from disk and
# populates the %CEFRwords hash, keyed by CEFR level, so that each word can
# later be looked up to determine its CEFR band during lexical analysis.
sub PrepareCEFRWordList
{
	my $keyhash = shift;
	my $levellist = shift;
	my $level = shift;
	my $member;
	my $index = 0;
	
	$logger->info ( sprintf "Preparing CEFR Wordlists at %s level", $level );

	if ( -e $levellist && open my $handle, '<', $levellist )
	{
		while( <$handle> )
		{ 
			my $word = '';
			my $type = '';
			my $meaning = '';
			my $ismultiword = 0;
			my $level = '';
			my $colour = '';
			my @strings;
		
			my $item = '';
		
			$member = $_;
			chomp $member;
			$member =~ s/\t//g;		# remove tab characters
			$member =~ s/\s+$//;	# remove trailing spaces
		
			my @wordinfo = split(' ', $member);
		
			#
			# THREE WORDS:
			#
			# nice and simple
			# walk verb A1
			# wall noun A1
			# want verb A1
			#
			# FOUR WORDS:
			#
			# what pronoun QUESTION A1
			# window noun GLASS A1
			# see you later A1
			# the Internet noun A1
			# How are you? A1
			# holiday noun VISIT A1
			#
			# FIVE WORDS:
			#
			# but conjunction DIFFERENT STATEMENT A1
			# can modal verb ABILITY A1
			# can modal verb REQUEST A1
			# can modal verb OFFER A1
			# can modal verb POSSIBILITY A1
		
			foreach $item ( @wordinfo )
			{
				$item =~ tr/0-9 a-z A-Z\/\';!-().//dc;		# remove most special characters

				# if the word is written in UPPERCASE (and isn't one of: A1-C2), it shows us the meaning/category of the word
				if ( $item eq uc $item && $item !~ '^[A-C][1-2]$' )	# && length $item >= 1 		# the next thing *I* knew
				{
					$meaning .= $item . ' ';
				}
				elsif ( $item ~~ ['auxiliary', 'exclamation', 'pronoun', 'noun', 'adverb', 'modal', 'verb', 'adjective', 'preposition', 'determiner', 'conjunction'] )
				{
					# maybe of more than one type:
					# above adverb, preposition TOO IMPORTANT C2
					# all right adjective, adverb WITHOUT PROBLEMS A1
	
					$type .= $item . ' ';
				}
				elsif ( $item =~ '^[A-C][1-2]$' )
				{
					$level = $item;
				}
				else
				{
					$item = 'sth' if $item eq 'something';		# make the wording consistent
					next if ( $item eq 'be' || $item eq '(be)' ) && $word eq '';	# if the verb to 'be' is at the start of the phrase, remove it
				
					$ismultiword = 1 if ( length $word > 0 );
					$word .= $item . ' ';
				}
			}		
			
			$word =~ s/\s+$//;	# remove trailing spaces
			$type =~ s/\s+$//;	# remove trailing spaces
			$meaning =~ s/\s+$//;	# remove trailing spaces
			#$logger->debug( sprintf "%4i: [%s] -> [%s] -> [%s] -> [%i] -> [%s]\n", $index, $word, $type, $meaning, $ismultiword, $level ) if ( $ismultiword == 1 );
		
			push @strings, $word;
			push @strings, verb( $word )->singular( 1 )	if ( $type eq 'verb' && not verb( $word )->singular( 1 ) ~~ @strings );	# first person singular
			push @strings, verb( $word )->singular( 3 )	if ( $type eq 'verb' && not verb( $word )->singular( 3 ) ~~ @strings );	# third person singular
			push @strings, verb( $word )->past 			if ( $type eq 'verb' && not verb( $word )->past  ~~ @strings );
			push @strings, verb( $word )->past_part 	if ( $type eq 'verb' && not verb( $word )->past_part  ~~ @strings );
			push @strings, verb( $word )->pres_part		if ( $type eq 'verb' && not verb( $word )->pres_part  ~~ @strings );
			push @strings, noun( $word )->singular   	if ( $type eq 'noun' && not noun( $word )->singular  ~~ @strings );
			push @strings, noun( $word )->plural	   	if ( $type eq 'noun' && not noun( $word )->plural ~~ @strings );
		
			# if we already have an item with this vocab item, don't add a second
			# this means that where a word has different precisions of meaning at different levels,
			# we always get the lowest level
			unless ( exists( $cefr{lc $word} ) || length $word == 0 )
			{
				my %ArchiveKey = (
					key				=> lc $word . '-'. $level,
					index  			=> $index,
				    member 			=> lc $word,
				    type			=> $type,
					meaning			=> $meaning,
					ismultiword		=> $ismultiword,
					level  			=> $level,
					found_count  	=> 0,
					strings			=> \@strings );
				
				my %DocumentArchiveKey = (
					key				=> lc $word . '-'. $level,
					index  			=> $index,
					member 			=> lc $word,
					type			=> $type,
					meaning			=> $meaning,
					ismultiword		=> $ismultiword,
					level  			=> $level,
					found_count  	=> 0,
					strings			=> \@strings );
		
				$cefr{lc $word} = \%ArchiveKey;
				$CEFR_document{lc $word} = \%DocumentArchiveKey;
			}
			else
			{
				#$logger->info ( sprintf "		Key [%s] already exists at index [%i] and level [%s]\n", lc $word, $cefr{lc $word}->index, $cefr{lc $word}->level );					
			}
		
			$index++;
		}
	
		$logger->info ( sprintf "		First time.  Prepared archive-wide CEFR wordlist %s.     	Dictionary has %i tokens\n", $levellist, scalar keys %cefr );
		$logger->info ( sprintf "		First time.  Prepared document-specific CEFR wordlist %s.	Dictionary has %i tokens\n", $levellist, scalar keys %CEFR_document );

		$keyhash->{Dictionaries}->{CEFR} = \%CEFR_document;
		$firsttime = 1;
		close $handle;
	}
	else
	{
		warn "Cannot open $levellist: $!\n";
	}
	
	return;
}


# PrepareAcademicCollocationsList
# Reads the Academic Collocations List data file from disk and populates the
# @AcademicCollocations array with phrase entries (multi-word expressions)
# for use during vocabulary analysis.
sub PrepareAcademicCollocationsList
{
	my $keyhash = shift;	
	my $member;
	undef %AcademicCollocationList;		# empty this
	
	$logger->info ( "Preparing Academic Collocations Wordlists" );

	if ( -e $ACADEMIC_COLLOCATION_LIST_FILE )
	{
		# this came from https://pearsonpte.com/organizations/researchers/academic-collocation-list/
		# we have modified it slightly 
		open my $handle, '<', $ACADEMIC_COLLOCATION_LIST_FILE;
	
		while( <$handle> )
		{
			my @one;
			my @two;
			my @strings;
			my @component1;
			my @component2;
			
			my @addition1;
			my @addition2;
	
			# PROBLEMATIC COLLOCATIONS MAY LOOK LIKE THIS
			#
			#	ADDITION	POS			COMPONENT		POS		COMPONENT		ADDITION
			#				adj			high/er			n		frequency
			#				v			pose			n		(a) threat		(to)
			#	(be) 		adv			specifically	vpp		designed		(to, for)
			#				v			take on			n		(the) role		(of, as)
			#	(be)		adv			best			vpp		described		(as, in terms of)
			#	(on/upon)	adj			closer			n		inspection
			#				adj			geographic(al)	n		area
			#				v			give (sb)		n		(an) impression
			#				adj			high/er			n		frequency
			#	(the only)	adj			plausible		n		explanation
				
			# remember: adverb + adjective collocations might be hyphenated when modifying a noun (e.g. a radically-different idea)
			# 			a verb + noun collocation may need the verb inflecting for person or tense (e.g. achieve a goal --> she achieves her goal OR she achieved her goal)

			$member = $_;
			chomp $member;

			next if ( $member =~ /^#/ );	# ignore any lines starting with a #

			my @info = split( ',', $member );

			my $index 		= $info[0];
			$index =~ s/^\s+|\s+$//g;		# remove leading and trailing spaces from this
			
			
			#
			#	left side
			#
			
			my $addition1 	= $info[1];
			$addition1 =~ s/^\s+|\s+$//g;	# remove leading and trailing spaces from this

			my $POS1 		= $info[2];
			$POS1 =~ s/^\s+|\s+$//g;		# remove leading and trailing spaces from this

			my $component1 	= $info[3];
			$component1 =~ s/^\s+|\s+$//g;	# remove leading and trailing spaces from this
			
			#
			#	right side
			#
			
			my $POS2 		= $info[4];
			$POS2 =~ s/^\s+|\s+$//g;		# remove leading and trailing spaces from this

			my $component2 	= $info[5];
			$component2 =~ s/^\s+|\s+$//g;	# remove leading and trailing spaces from this

			my $addition2 	= $info[6];
			$addition2 =~ s/^\s+|\s+$//g;	# remove leading and trailing spaces from this
			
			$addition1 =~ s/^\(|\)$//g;		# remove opening and closing brackets
			$addition2 =~ s/^\(|\)$//g;		# remove opening and closing brackets
			
			#$addition1 =~ tr/a-z,A-Z//dc if ( $addition1 ne "" );	# remove all characters except lowercase and uppercase a-z and ,
			#$addition2 =~ tr/a-z,A-Z//dc if ( $addition2 ne "" );	# remove all characters except lowercase and uppercase a-z and ,
			
			@addition1 = split /[,\/]/, $addition1;	# Use a character class in the regex delimiter to match on a set of possible delimiters.  Split on either , or /
			@addition2 = split /[,\/]/, $addition2;	# Use a character class in the regex delimiter to match on a set of possible delimiters.
			
			my $string = sprintf "[%-4s]	| [%-5s]	| [%-10s]	| [%-20s]	| [%-5s]	| [%-20s]	| [%-10s]\n", 
																							$index,
																							 
																							$POS1, 
																							join( "|", @addition1 ),
																							$component1, 

																							$POS2, 
																							$component2, 
																							join( "|", @addition2 );

			if ( $component1 =~ m/\(/ )
			{
				my $component1_with_optional_text = $component1;
				my $component1_without_optional_text = $component1;
	
				$component1_with_optional_text =~ tr/a-z //dc;	# remove all characters except a-z
				$component1_without_optional_text =~ s/\((.*?)\)//gsi;
				$component1_without_optional_text =~ s/^\s//g;
		
				push @component1, $component1_with_optional_text 	if ( $component1_with_optional_text ne $component1 );
				push @component1, $component1_without_optional_text	if ( $component1_without_optional_text ne $component1 );
			}
			else
			{
				push @component1, $component1;
			}


			if ( $component2 =~ m/\(/ )
			{
				my $component2_with_optional_text = $component2;
				my $component2_without_optional_text = $component2;
	
				$component2_with_optional_text =~ tr/a-z //dc;	# remove all characters except a-z
				$component2_without_optional_text =~ s/\((.*?)\)//gsi;
				$component2_without_optional_text =~ s/^\s//g;
		
				push @component2, $component2_with_optional_text 	if ( $component2_with_optional_text ne $component2 );
				push @component2, $component2_without_optional_text	if ( $component2_without_optional_text ne $component2 );
			}
			else
			{
				push @component2, $component2;
			}
		
			if ( $addition1 eq 'be' )
			{
				foreach my $tobe ( @tobe )
				{
					foreach my $firstcomponent ( @component1 )
					{
						push @one, $firstcomponent  						unless ( $firstcomponent ~~ @one );
						push @one, $tobe . ' ' . $firstcomponent			unless ( $tobe . ' ' . $firstcomponent ~~ @one );
	
						push @one, 'to ' . $firstcomponent 					if ( $POS1 eq 'v' && not 'to ' . $firstcomponent ~~ @one );	# first person singular
						push @one, verb( $firstcomponent )->singular( 1 )	if ( $POS1 eq 'v' && not verb( $firstcomponent )->singular( 1 ) ~~ @one );	# first person singular
						push @one, verb( $firstcomponent )->singular( 3 )	if ( $POS1 eq 'v' && not verb( $firstcomponent )->singular( 3 ) ~~ @one );	# third person singular
						push @one, verb( $firstcomponent )->past 			if ( $POS1 eq 'v' && not verb( $firstcomponent )->past  ~~ @one);
						push @one, verb( $firstcomponent )->past_part 		if ( $POS1 eq 'v' && not verb( $firstcomponent )->past_part  ~~ @one);
						push @one, verb( $firstcomponent )->pres_part		if ( $POS1 eq 'v' && not verb( $firstcomponent )->pres_part  ~~ @one);
						push @one, noun( $firstcomponent )->singular   		if ( $POS1 eq 'n' && not noun( $firstcomponent )->singular  ~~ @one);
						push @one, noun( $firstcomponent )->plural	   		if ( $POS1 eq 'n' && not noun( $firstcomponent )->plural ~~ @one );
					}
				}
			}
			else
			{
				foreach my $firstcomponent ( @component1 )
				{
					push @one, $firstcomponent  						unless ( $firstcomponent ~~ @one );
					#push @one, $addition1 . ' ' . $firstcomponent 		unless ( $addition1 eq '' || $addition1 . ' ' . $firstcomponent ~~ @one );

					push @one, 'to ' . $firstcomponent 					if ( $POS1 eq 'v' && not 'to ' . $firstcomponent ~~ @one );	# first person singular
					push @one, verb( $firstcomponent )->singular( 1 )	if ( $POS1 eq 'v' && not verb( $firstcomponent )->singular( 1 ) ~~ @one );	# first person singular
					push @one, verb( $firstcomponent )->singular( 3 )	if ( $POS1 eq 'v' && not verb( $firstcomponent )->singular( 3 ) ~~ @one );	# third person singular
					push @one, verb( $firstcomponent )->past 			if ( $POS1 eq 'v' && not verb( $firstcomponent )->past  ~~ @one);
					push @one, verb( $firstcomponent )->past_part 		if ( $POS1 eq 'v' && not verb( $firstcomponent )->past_part  ~~ @one);
					push @one, verb( $firstcomponent )->pres_part		if ( $POS1 eq 'v' && not verb( $firstcomponent )->pres_part  ~~ @one);
					push @one, noun( $firstcomponent )->singular   		if ( $POS1 eq 'n' && not noun( $firstcomponent )->singular  ~~ @one);
					push @one, noun( $firstcomponent )->plural	   		if ( $POS1 eq 'n' && not noun( $firstcomponent )->plural ~~ @one );
				}
			}

			foreach my $secondcomponent ( @component2 )
			{
				push @two, $secondcomponent 						unless ( $secondcomponent ~~ @two );
				#push @two, $secondcomponent . ' ' . $addition2 		unless ( $addition2 eq '' || $secondcomponent . ' ' . $addition2 ~~ @two );

				push @two, verb( $secondcomponent )->singular( 1 )	if ( $POS2 eq 'v' && not verb( $secondcomponent )->singular( 1 ) ~~ @two );	# first person singular
				push @two, verb( $secondcomponent )->singular( 3 )	if ( $POS2 eq 'v' && not verb( $secondcomponent )->singular( 3 ) ~~ @two);	# third person singular
				push @two, verb( $secondcomponent )->past 			if ( $POS2 eq 'v' && not verb( $secondcomponent )->past ~~ @two);
				push @two, verb( $secondcomponent )->past_part 		if ( $POS2 eq 'v' && not verb( $secondcomponent )->past_part ~~ @two);
				push @two, verb( $secondcomponent )->pres_part 		if ( $POS2 eq 'v' && not verb( $secondcomponent )->pres_part ~~ @two);
				push @two, noun( $secondcomponent )->singular   	if ( $POS2 eq 'n' && not noun( $secondcomponent )->singular ~~ @two);
				push @two, noun( $secondcomponent )->plural			if ( $POS2 eq 'n' && not noun( $secondcomponent )->plural ~~ @two);
			}
			
			push @addition1, '';	# this ensures that we iterate through the array at least once
			push @addition2, '';	# this ensures that we iterate through the array at least once
			
			foreach my $one ( @one )
			{
				foreach my $two ( @two )
				{
					foreach my $front ( @addition1 )
					{
						next if $front eq 'be';		# we've already dealth with this
						
						foreach my $back ( @addition2 )
						{
							my $string = $front . ' ' . $one . ' ' . $two . ' ' . $back;
							$string =~ s/^\s+|\s+$//g;	# trim leading and trailing spaces
							push @strings, $string unless ( $string ~~ @strings );
						}
					}
				}
			}
			
			my @sortedbylength = sort { length $b <=> length $a } @strings;
			
			# for debugging
			#printf "		Left-side is 	%s\n", join( "\n\t\t\t\t", @one );
			#printf "		Right-side is	%s\n", join( "\n\t\t\t\t", @two );	
			#print $string;
			#printf "		Strings are	[%s]\n", join( "\n\t\t\t\t", @sortedbylength );
	
			my %ACLKey = (
			    index  			=> $index,
			    addition_1		=> $addition1,
				component_1		=> $component1,
				POS_1			=> $POS1,
				component_2 	=> $component2,
			    POS_2  			=> $POS2,
				addition_2		=> $addition2,
				found_count		=> 0,
				strings			=> \@sortedbylength );
				
			my %ACLKey_Document = (
				index  			=> $index,
		    	addition_1		=> $addition1,
				component_1		=> $component1,
				POS_1			=> $POS1,
				component_2 	=> $component2,
		    	POS_2  			=> $POS2,
				addition_2		=> $addition2,
				found_count		=> 0,
				strings			=> \@sortedbylength );
		
			my $key = $ACLKey{component_1}.' '.$ACLKey{component_2};
			
			unless( exists $AcademicCollocationList{ $key } )
			{
				$AcademicCollocationList{ $key } = \%ACLKey;
				$AcademicCollocationList_document{ $key } = \%ACLKey_Document;
			}
			else
			{
				$logger->info ( sprintf "		Key [%s] already exists [%s]\n", lc $key, '' );
			}
		}
	
		$keyhash->{Dictionaries}->{AcademicCollocations} = \%AcademicCollocationList_document;
		
		close $handle;
		$logger->info ( sprintf "		First time.  Prepared ACL wordlist.  Dictionary has %i items\n", scalar keys %AcademicCollocationList );
		$firsttime = 1;
	}
	else
	{
		$logger->info ( "		Could not prepare ACL wordlist.  File not found\n" );
	}

	return;
}



# PrepareNGSLWordList
# Reads the New General Service List (NGSL) data file and populates the
# %NGSLwords hash. Each entry maps a headword to its frequency rank, enabling
# NGSL coverage calculations during lexical analysis.
sub PrepareNGSLWordList
{
	my $keyhash = shift;	
	my $member;
	my $index = 0;
	my $newroot = 0;
	my $root = '';
	undef %ngsl;	# empty this
	
	$logger->info ( "Preparing NGSL Wordlists" );
	
	if ( -e $NGSL_FILE )
	{
		open my $handle, '<', $NGSL_FILE;	# this came from http://www.newgeneralservicelist.org & https://github.com/d5ve/passgen/blob/master/ngsl.txt
	
		while( <$handle> )
		{ 
			$member = $_;
			chomp $member;
			$member =~ s/\t//g;		# remove tab characters
			$member =~ s/\s+$//;	# remove trailing spaces
		
			next if ( $member =~ /^#/ );	# ignore any lines starting with a #
		
			if ( $member =~ /^[0-9]+$/ )  # if is a number
			{
				$index = $member;
				$newroot = 1;
			}
			elsif ( length $member > 0 )
			{
				my $level = '';
				$root = $member if ( $newroot == 1 );
				
				$level = "BandA" if ( $index > $NGSL_BAND_A_LOWER && $index <= $NGSL_BAND_A_HIGHER );
				$level = "BandB" if ( $index > $NGSL_BAND_B_LOWER && $index <= $NGSL_BAND_B_HIGHER );
				$level = "BandC" if ( $index > $NGSL_BAND_C_LOWER && $index <= $NGSL_BAND_C_HIGHER );
				$level = "BandD" if ( $index > $NGSL_BAND_D_LOWER && $index <= $NGSL_BAND_D_HIGHER );
			
				my %ArchiveKey = (
				    index  			=> $index,
				    root			=> $root,
					isroot			=> $newroot,
					root_found_count=> 0,
					member 			=> lc $member,
				    found_count  	=> 0,
					level			=> $level );
		
				my %DocumentArchiveKey = (
					index  			=> $index,
				    root			=> $root,
					isroot			=> $newroot,
					root_found_count=> 0,
					member 			=> lc $member,
					found_count  	=> 0,
					level			=> $level,
					isnew			=> 0 );	# is this new vocab for our students?  this doesn't exist in the archivekey because we base this on whether or not the archivekey found_count > 0
				
				# do this down here
				$newroot = 0 if ( $newroot == 1 );
			
				unless( exists $ngsl{ lc $member } )
				{
					$ngsl{ lc $member } = \%ArchiveKey;
					$NGSL_document{ lc $member } = \%DocumentArchiveKey;
				}
				else
				{
					$logger->info ( sprintf "		Key [%s] already exists [%s]\n", lc $member, '' );
				}
			}
			else
			{
				# do nothing
				$logger->info ( "		Length of member is zero, not adding to vocab hash\n" );
			}
		} 		

		$logger->info ( sprintf "		First time.  Prepared NGSL wordlist.  Dictionary has %i items and %i tokens\n", $index, scalar keys %ngsl );
		$keyhash->{Dictionaries}->{NGSL} = \%NGSL_document;
		$firsttime = 1;
		close $handle;
	}
	else
	{
		$logger->info ( "		Could not prepare NGSL wordlist.  File not found\n" );		
	}

	return;
}



# PrepareNAWLWordList
# Reads the New Academic Word List (NAWL) data file and populates the
# %NAWLwords hash. Each entry maps a headword to its frequency rank, enabling
# NAWL coverage calculations during lexical analysis.
sub PrepareNAWLWordList
{
	my $keyhash = shift;
	my $member;
	my $index = 0;
	my $root = '';
	my $newroot = 0;
	undef %nawl;	# empty this
	
	$logger->info ( "Preparing NAWL Wordlists" );
	
	if ( -e $NAWL_FILE )
	{
		open my $handle, '<', $NAWL_FILE;
	
		while( <$handle> )
		{ 
			$index++;
			$member = $_;
			chomp $member;
			$member =~ s/\s+$//;	# remove trailing spaces
		
			next if ( $member =~ /^#/ );	# ignore any lines starting with a #
		
			if ( length $member > 0 )
			{
				my $level = "NAWL";
			
				# NAWL looks like this:
				# the roots are accumulate and accumulation.
				#accumulate
				#	accumulates
				#	accumulated
				#	accumulating
				#	accumulatings
				#accumulation
				#	accumulations
				unless ( $member =~ /^\t/ )
				{
					$root = $member;
					$newroot = 1;
				}
				else
				{
					$newroot = 0;
				}
			
				# remove tab characters
				# do this here
				$member =~ s/\t//g;
			
				my %ArchiveKey = (
				    index  			=> $index,
				    root			=> $root,
					isroot			=> $newroot,
					root_found_count=> 0,
					member 			=> lc $member,
				    found_count  	=> 0,
					level			=> $level );
			
				my %DocumentArchiveKey = (
					index  			=> $index,
				    root			=> $root,
					isroot			=> $newroot,
					root_found_count=> 0,
					member 			=> lc $member,
					found_count  	=> 0,
					level			=> $level,
					isnew			=> 0 );
		
				unless( exists $nawl{lc $member} )
				{
					$nawl{lc $member} = \%ArchiveKey;
					$NAWL_document{lc $member} = \%DocumentArchiveKey;
				}
				else
				{
					$logger->info ( sprintf "		Key already exists [%s]\n", lc $member, '' );
				}
			}
			else
			{
				$logger->info ( "		Length of member is zero, not adding to vocab hash\n" );
			}
		} 

		$keyhash->{Dictionaries}->{NAWL} = \%NAWL_document;
		close $handle;
		$logger->info ( sprintf "		First time.  Prepared NAWL wordlist.  Dictionary has %i tokens\n", scalar keys %nawl );
		$firsttime = 1;		
	}
	else
	{
		$logger->info ( "		Could not prepare NAWL wordlist.  File not found\n" );		
	}

	return;
}


# PrepareSupplementalWordList
# Reads a project-specific supplemental word list from disk and populates the
# %SupplementalWords hash. Words in this list are excluded from 'unknown word'
# counts so that domain-specific terminology is not flagged erroneously.
sub PrepareSupplementalWordList
{
	my $member;
	
	$logger->info ( "Preparing Supplemental Wordlists" );
	
	if ( -e $SUPPLEMENTAL_FILE )
	{
		open my $handle, '<', $SUPPLEMENTAL_FILE;
	
		while( <$handle> )
		{ 
			$member = $_;
			chomp $member;
			$member =~ s/\t//g;		# remove tab characters
			$member =~ s/\s+$//;	# remove trailing spaces
		
			$supplemental{ lc $member } = 1;
		} 

		$logger->info ( sprintf "		First time.  Prepared Supplemental wordlist.  Dictionary has %i tokens\n", scalar keys %supplemental );
		$firsttime = 1;		

		close $handle;
	}
	else
	{
		$logger->info ( "		Could not prepare supplemental wordlist.  File not found\n" );		
	}

	return;
}



# PrepareDictionary
# Reads the full dictionary word list from disk into the %Dictionary hash.
# This hash is used during vocabulary analysis to distinguish known English
# words from potential spelling errors or non-words.
sub PrepareDictionary
{
	my $member;
	
	$logger->info ( "Preparing Dictionary" );

	if ( -e $DICTIONARY_FILE )
	{
		open my $handle, '<', $DICTIONARY_FILE;	# this came from https://raw.githubusercontent.com/eneko/data-repository/master/data/words.txt
	
		while( <$handle> )
		{ 
			$member = $_;
			chomp $member;
			$member =~ s/\t//g;		# remove tab characters
			$member =~ s/\s+$//;	# remove trailing spaces
	    
			$dictionary{lc $member } = 1;
		} 

		$logger->info ( sprintf "		First time.  Prepared dictionary.     Dictionary has %i tokens\n", scalar keys %dictionary );
		
		$firsttime = 1;
		close $handle;	
	}
	else
	{
		$logger->info ( "		Could not prepare dictionary.  File not found\n" );		
	}

	return;
}


# GetLexisInformation
# Performs full lexical analysis of a resource's text content. Tokenises the
# text, lemmatises each token, then classifies every word against the NGSL,
# NAWL, CEFR, Academic Collocations, supplemental, and dictionary lists.
# Populates the resource hash with vocabulary coverage statistics and arrays
# of words at each CEFR level.
sub GetLexisInformation
{
	my $keyhash = shift;
		
	my %AcademicCollocationsInDocument;
	my %NGSLInDocument;
	my %NAWLInDocument;
	my %CEFRInDocument;
	
	my @all_words_in_document;			# just a simple array of all words in the document
	my @new_words_in_document;			# just a simple array of NEW words which exist in this text but not previously seen
	my @all_unknown_words_in_document;	# just a simple array of typos/unknown words in the document
			
	$logger->info ( "Getting Lexis Information" );

	if( $firsttime == 0 )
	{
		PrepareNGSLWordList( $keyhash );
		PrepareNAWLWordList( $keyhash );
		PrepareAcademicCollocationsList( $keyhash );
		PrepareSupplementalWordList();
		PrepareDictionary();
		PrepareCEFRWordList( $keyhash, $CEFR_A1_FILE, "A1" );
		PrepareCEFRWordList( $keyhash, $CEFR_A2_FILE, "A2" );
		PrepareCEFRWordList( $keyhash, $CEFR_B1_FILE, "B1" );
		PrepareCEFRWordList( $keyhash, $CEFR_B2_FILE, "B2" );
		PrepareCEFRWordList( $keyhash, $CEFR_C1_FILE, "C1" );
		PrepareCEFRWordList( $keyhash, $CEFR_C2_FILE, "C2" );
	}
	else
	{
		#
		# reset these found counts and re-use the **SAME** DOCUMENT hash
		# BEWARE THAT ALL POINTERS/REFERENCES WILL POINT TO THE SAME DOCUMENT HASH IRRESPECTIVE OF THE ARCHIVE ITEM (FIX THIS)
		#
			
		foreach my $outerkey ( keys %NGSL_document )
		{ 
			my $innerkey = $NGSL_document{$outerkey};
			$innerkey->{root_found_count} = 0;
			$innerkey->{found_count} = 0;
		}
	
		foreach my $outerkey ( keys %NAWL_document )
		{ 
			my $innerkey = $NAWL_document{$outerkey};
			$innerkey->{root_found_count} = 0;
			$innerkey->{found_count} = 0;
		}
	
		foreach my $outerkey ( keys %CEFR_document )
		{
			my $innerkey = $CEFR_document{$outerkey};
			$innerkey->{found_count} = 0;
		}
		
		foreach my $outerkey ( keys %AcademicCollocationList_document )
		{ 
			my $innerkey = $AcademicCollocationList_document{$outerkey};
			$innerkey->{found_count} = 0;
		}
		
		# we should take a copy of these here before we use them
		$keyhash->{Dictionaries}->{AcademicCollocations} = \%AcademicCollocationList_document;
		$keyhash->{Dictionaries}->{NGSL} = \%NGSL_document;
		$keyhash->{Dictionaries}->{NAWL} = \%NAWL_document;
		$keyhash->{Dictionaries}->{CEFR} = \%CEFR_document;
	}
	
	#
	# WE NEED TO USE THESE IN THE CODE BELOW THEY ARE THE DOCUMENT SPECIFIC HASHES
	#
	
	%AcademicCollocationsInDocument = %{ $keyhash->{Dictionaries}->{AcademicCollocations} };
	%NGSLInDocument = %{ $keyhash->{Dictionaries}->{NGSL} };
	%NAWLInDocument = %{ $keyhash->{Dictionaries}->{NAWL} };
	%CEFRInDocument = %{ $keyhash->{Dictionaries}->{CEFR} };
					
	#
	# can we find any academic collocations?
	#
	$logger->info ( "Finding academic collocations...\n" );
	
	foreach my $outerkey ( keys %AcademicCollocationsInDocument )
	{
		my $innerkey = $AcademicCollocationsInDocument{$outerkey};	
		my @strings = @{ $innerkey->{strings} };
		
		foreach my $string ( @strings )
		{
			my @matches = $keyhash->{PrettyTextofDocument} =~ m/\b$string\b/g; 	# use PrettyTextofDocument as it's all lowercase search with the regex 'i' switch is slow as hell!
			$logger->info ( sprintf "		Found %i academic collocation for: [%s]\n", scalar @matches, $string ) if ( scalar @matches > 0 );

			$keyhash->{WordListMetaData}->{AcademicCollocationsCount} += scalar @matches;
			$AcademicCollocationList{$outerkey}{found_count} += scalar @matches;
			$AcademicCollocationsInDocument{$outerkey}{found_count} += scalar @matches;
		}
	}
	
	#
	# for each CEFR word/multiword in the text, check if it's in the text
	#
	$logger->info ( "Finding CEFR vocab...\n" );
	
	foreach my $outerkey ( keys %CEFRInDocument )
	{ 
		my $innerkey = $CEFRInDocument{$outerkey};
		my @strings = @{ $innerkey->{strings} };
		
		#
		# if member is a verb, search for inflections, likewise nouns
		#
		foreach my $string ( @strings )
		{
			# use PrettyTextofDocument as it's all lowercase search with the regex 'i' switch is slow as hell!
			my $count = 0;
			
			if ( $innerkey->{ismultiword} == 1 )
			{
				my @matches = $keyhash->{PrettyTextofDocument} =~ m/\b$string\b/g;
				$count = scalar @matches;
			
				$keyhash->{WordListMetaData}->{CEFR_A1_MultiWord_Count} += $count if ( $innerkey->{level} eq 'A1' );
				$keyhash->{WordListMetaData}->{CEFR_A2_MultiWord_Count} += $count if ( $innerkey->{level} eq 'A2' );
				$keyhash->{WordListMetaData}->{CEFR_B1_MultiWord_Count} += $count if ( $innerkey->{level} eq 'B1' );
				$keyhash->{WordListMetaData}->{CEFR_B2_MultiWord_Count} += $count if ( $innerkey->{level} eq 'B2' );
				$keyhash->{WordListMetaData}->{CEFR_C1_MultiWord_Count} += $count if ( $innerkey->{level} eq 'C1' );
				$keyhash->{WordListMetaData}->{CEFR_C2_MultiWord_Count} += $count if ( $innerkey->{level} eq 'C2' );
				$innerkey->{found_count} += $count;
				
				$logger->info ( sprintf "	 Matched multiword (%s)[%s] @ level %s; found count is now %i\n", $string, $innerkey->{member}, $innerkey->{level}, $innerkey->{found_count} ) if ( $innerkey->{found_count} > 0 );				
			}
			else
			{
				my @matches = $keyhash->{PrettyTextofDocument} =~ m/\b$string\b/g;
				$count = scalar @matches;
			
				$keyhash->{WordListMetaData}->{CEFR_A1_Count} += $count if ( $innerkey->{level} eq 'A1' );
				$keyhash->{WordListMetaData}->{CEFR_A2_Count} += $count if ( $innerkey->{level} eq 'A2' );
				$keyhash->{WordListMetaData}->{CEFR_B1_Count} += $count if ( $innerkey->{level} eq 'B1' );
				$keyhash->{WordListMetaData}->{CEFR_B2_Count} += $count if ( $innerkey->{level} eq 'B2' );
				$keyhash->{WordListMetaData}->{CEFR_C1_Count} += $count if ( $innerkey->{level} eq 'C1' );
				$keyhash->{WordListMetaData}->{CEFR_C2_Count} += $count if ( $innerkey->{level} eq 'C2' );
				$innerkey->{found_count} += $count;					
			}
			
			$keyhash->{WordListMetaData}->{CEFRTotalCount} += $count;
			
			$logger->info ( sprintf "	Found [%02i] match(es) CEFR item at level [%s] for: [%s] in [%s]\t\t CEFR Total Count is [%s] \n", $count, $innerkey->{level}, $string, $innerkey->{member}, $keyhash->{WordListMetaData}->{CEFRTotalCount} ) if ( $count > 0 );
		}
	}
	
	$logger->info ( sprintf "MULTIWORD\tA1: %02i\tA2: %02i\tB1: %02i\tB2: %02i\tC1: %02i\tC2: %02i\n", $keyhash->{WordListMetaData}->{CEFR_A1_MultiWord_Count}, $keyhash->{WordListMetaData}->{CEFR_A2_MultiWord_Count}, $keyhash->{WordListMetaData}->{CEFR_B1_MultiWord_Count}, $keyhash->{WordListMetaData}->{CEFR_B2_MultiWord_Count}, $keyhash->{WordListMetaData}->{CEFR_C1_MultiWord_Count}, $keyhash->{WordListMetaData}->{CEFR_C2_MultiWord_Count} );
	$logger->info ( sprintf "WORD\tA1: %02i\tA2: %02i\tB1: %02i\tB2: %02i\tC1: %02i\tC2: %02i\n", $keyhash->{WordListMetaData}->{CEFR_A1_Count}, $keyhash->{WordListMetaData}->{CEFR_A2_Count}, $keyhash->{WordListMetaData}->{CEFR_B1_Count}, $keyhash->{WordListMetaData}->{CEFR_B2_Count}, $keyhash->{WordListMetaData}->{CEFR_C1_Count}, $keyhash->{WordListMetaData}->{CEFR_C2_Count} );

	#
	# check to see if words from our wordlists ( (i) CEFR (ii) NGSL (iii) NAWL ) exist in the document 
	#
	
	$logger->info ( "Finding NGSL vocab...\n" );
	
	foreach my $outerkey ( keys %NGSLInDocument )
	{
		my $innerkey = $NGSLInDocument{$outerkey};
		
		while ( $keyhash->{PrettyTextofDocument} =~ m/\b$innerkey->{member}\b/g )	# use PrettyTextofDocument as it's all lowercase search with the regex 'i' switch is slow as hell!
		{
			#$logger->info ( sprintf "		Found match NGSL item for: [%s]\n", $innerkey->{member} );
			$innerkey->{isnew} = 1 if ( $ngsl{ $innerkey->{member} }->{found_count} == 0 );	# do this before we increment the found_count

			$ngsl{ $innerkey->{member} }->{found_count}++;
			$innerkey->{found_count}++;
			
			$keyhash->{WordListMetaData}->{NGSLBandA}++ if ( $ngsl{ $innerkey->{member} }{level} eq "BandA" );
			$keyhash->{WordListMetaData}->{NGSLBandB}++ if ( $ngsl{ $innerkey->{member} }{level} eq "BandB" );
			$keyhash->{WordListMetaData}->{NGSLBandC}++ if ( $ngsl{ $innerkey->{member} }{level} eq "BandC" );
			$keyhash->{WordListMetaData}->{NGSLBandD}++ if ( $ngsl{ $innerkey->{member} }{level} eq "BandD" );
		
			$keyhash->{WordListMetaData}->{NGSLTotalCount}++;
			
			if ( $innerkey->{isroot} == 1 )
			{
				# increment the root_found_count in our archive-wide and document-wide dictionaries
				$ngsl{ $innerkey->{member} }->{root_found_count}++;
				$innerkey->{root_found_count}++;
			}
			else
			{
				# find the root and increment the root_found_count in our archive-wide dictionary
				$ngsl{ $innerkey->{root} }->{root_found_count}++;
				$ngsl{ $innerkey->{member} }{root_found_count} = $ngsl{ $ngsl{ $innerkey->{member} }->{root} }->{root_found_count};
				
				# find the root and increment the root_found_count in our document-wide dictionary
				$NGSLInDocument{ $NGSLInDocument{ $innerkey->{member} }->{root} }->{root_found_count}++;
				$NGSLInDocument{ $innerkey->{member} }->{root_found_count} = $NGSLInDocument{ $NGSLInDocument{ $innerkey->{member} }->{root} }->{root_found_count};
			}
		}
	}
	
	$logger->info ( "Finding NAWL vocab...\n" );
	
	foreach my $outerkey ( keys %NAWLInDocument )
	{
		my $innerkey = $NAWLInDocument{$outerkey};
		
		while ( $keyhash->{PrettyTextofDocument} =~ m/\b$innerkey->{member}\b/g )	# use PrettyTextofDocument as it's all lowercase search with the regex 'i' switch is slow as hell!
		{
			#$logger->info ( sprintf "		Found match NAWL item for: [%s]\n", $innerkey->{member} );
			$NAWLInDocument{ $innerkey->{member} }->{isnew} = 1 if ( $nawl{ $innerkey->{member} }->{found_count} == 0 );	# do this before we increment the found_count
			
			$nawl{ $innerkey->{member} }->{found_count}++;
			$NAWLInDocument{ $innerkey->{member} }->{found_count}++;
			
			$keyhash->{WordListMetaData}->{NAWLTotalCount}++;
		
			if ( $nawl{ $innerkey->{member} }->{isroot} == 1 )
			{
				# increment the root_found_count in our archive-wide and document-wide dictionaries
				$nawl{ $innerkey->{member} }->{root_found_count}++;
				$NAWLInDocument{ $innerkey->{member} }->{root_found_count}++;
			}
			else
			{
				# find the root and increment the root_found_count in our archive-wide dictionary
				$nawl{ $nawl{ $innerkey->{member} }{root} }{root_found_count}++;
				$nawl{ $innerkey->{member} }{root_found_count} = $nawl{ $nawl{ $innerkey->{member} }->{root} }->{root_found_count};

				# find the root and increment the root_found_count in our document-wide dictionary
				$NAWLInDocument{ $NAWLInDocument{ $innerkey->{member} }->{root} }->{root_found_count}++;
				$NAWLInDocument{ $innerkey->{member} }->{root_found_count} = $NAWLInDocument{ $NAWLInDocument{ $innerkey->{member} }->{root} }->{root_found_count};
			}
		}	
	}
	
	#
	#
	#
	
	$logger->info ( "	Got vocab. Sorting and storing...\n" );
	
	
	# get an array of words in the document	
	my @words = split(' ', $keyhash->{PrettyTextofDocument});
	push @all_words_in_document, @words;		# add to the complete word list
	
	# this might be different from what Word tells us.
	$keyhash->{WordListMetaData}->{WordCount} = scalar( @all_words_in_document );
	
	@AllWords = uniq @all_words_in_document;
	@AllWords = sort @AllWords;
	$keyhash->{Dictionaries}->{AllWords} = \@AllWords;
	
	#
	# get new words in this document
	#
	
	$logger->info ( "	Finding new words...\n" );
	
	foreach my $uniqueword ( @AllWords )			# for each unique word in the document
	{
		#
		# get "typos" in this document
		#
		if ( length $uniqueword > 3 && $uniqueword =~ /^[a-zA-Z]+$/ )		# only check if the word is actually a word (no numbers) and longer than 3 characters
		{
			#
			# does the word exist in the dictionary (or supplementary dictionary which contains months, days, numbers)??
			#
			if ( exists $dictionary{$uniqueword} || exists $supplemental{$uniqueword} ) 	# yes!
			{
				#
				# get new words in this document which are NOT in the NGSL or NAWL
				#
				unless( $uniqueword ~~ @OldWords || exists $ngsl{$uniqueword} || exists $ngsl{$uniqueword} )
				{
					#printf "					We've NOT seen this word (%s) before.  New Word.\n", $word;
					push @new_words_in_document, $uniqueword;
					$keyhash->{WordListMetaData}->{NewWords}++;
				}
			}
			else
			{
				push @all_unknown_words_in_document, $uniqueword;	
				$keyhash->{WordListMetaData}->{TyposCount}++;
			}
		}
	}
	
	push @OldWords, sort @new_words_in_document;		# at the end, add these words to our big global array ready for next time
	
	
	@UnknownWords = uniq @all_unknown_words_in_document;
	@UnknownWords = sort @UnknownWords;
	$keyhash->{Dictionaries}->{UnknownWords} = \@UnknownWords;
	
	@NewWords = uniq @new_words_in_document;
	@NewWords = sort @NewWords;
	$keyhash->{Dictionaries}->{NewWords} = \@NewWords;
	
	$keyhash->{WordListMetaData}->{NGSLPercent} = sprintf "%.02f", $keyhash->{WordListMetaData}->{NGSLTotalCount}/$keyhash->{WordListMetaData}->{WordCount}*100 if ( $keyhash->{WordListMetaData}->{WordCount} > 0 );
	$keyhash->{WordListMetaData}->{NAWLPercent} = sprintf "%.02f", $keyhash->{WordListMetaData}->{NAWLTotalCount}/$keyhash->{WordListMetaData}->{WordCount}*100 if ( $keyhash->{WordListMetaData}->{WordCount} > 0 );
	$keyhash->{WordListMetaData}->{NewPercent} =  sprintf "%.02f", $keyhash->{WordListMetaData}->{NewWords}/$keyhash->{WordListMetaData}->{WordCount}*100		if ( $keyhash->{WordListMetaData}->{WordCount} > 0 );
	$keyhash->{WordListMetaData}->{UnknownPercent} = sprintf "%.02f", $keyhash->{WordListMetaData}->{TyposCount}/$keyhash->{WordListMetaData}->{WordCount}*100	if ( $keyhash->{WordListMetaData}->{WordCount} > 0 );
	
	#
	# display and save all this document-level information
	#

	$logger->info ( "	Creating statistics information for this document...\n" );
	
	$keyhash->{SummaryText} .= sprintf "		Source:	%s\n", $keyhash->{WordMetaData}->{source};
	$keyhash->{SummaryText} .= "\n";
	$keyhash->{SummaryText} .= sprintf "		** Vocabulary Information **\n";
	$keyhash->{SummaryText} .= sprintf "		Found %i CEFR words (and phrases) in document\n", $keyhash->{WordListMetaData}->{CEFRTotalCount};
	$keyhash->{SummaryText} .= sprintf "			A1:	%3i	%3i\n", $keyhash->{WordListMetaData}->{CEFR_A1_Count}, $keyhash->{WordListMetaData}->{CEFR_A1_MultiWord_Count};
	$keyhash->{SummaryText} .= sprintf "			A2:	%3i	%3i\n", $keyhash->{WordListMetaData}->{CEFR_A2_Count}, $keyhash->{WordListMetaData}->{CEFR_A2_MultiWord_Count};
	$keyhash->{SummaryText} .= sprintf "			B1:	%3i	%3i\n", $keyhash->{WordListMetaData}->{CEFR_B1_Count}, $keyhash->{WordListMetaData}->{CEFR_B1_MultiWord_Count};
	$keyhash->{SummaryText} .= sprintf "			B2:	%3i	%3i\n", $keyhash->{WordListMetaData}->{CEFR_B2_Count}, $keyhash->{WordListMetaData}->{CEFR_B2_MultiWord_Count};
	$keyhash->{SummaryText} .= sprintf "			C1:	%3i	%3i\n", $keyhash->{WordListMetaData}->{CEFR_C1_Count}, $keyhash->{WordListMetaData}->{CEFR_C1_MultiWord_Count};
	$keyhash->{SummaryText} .= sprintf "			C2:	%3i	%3i\n", $keyhash->{WordListMetaData}->{CEFR_C2_Count}, $keyhash->{WordListMetaData}->{CEFR_C2_MultiWord_Count};
	$keyhash->{SummaryText} .= "\n";
	$keyhash->{SummaryText} .= sprintf "		Found %i NGSL words in document (%.02f%%)\n", $keyhash->{WordListMetaData}->{NGSLTotalCount}, $keyhash->{WordListMetaData}->{NGSLPercent};
	$keyhash->{SummaryText} .= sprintf "			0000 --> 0800:	%i\n", $keyhash->{WordListMetaData}->{NGSLBandA};
	$keyhash->{SummaryText} .= sprintf "			0801 --> 1600:	%i\n", $keyhash->{WordListMetaData}->{NGSLBandB};
	$keyhash->{SummaryText} .= sprintf "			1601 --> 2400:	%i\n", $keyhash->{WordListMetaData}->{NGSLBandC};
	$keyhash->{SummaryText} .= sprintf "			2401 --> 2801:	%i\n", $keyhash->{WordListMetaData}->{NGSLBandD};
	$keyhash->{SummaryText} .= "\n";
	$keyhash->{SummaryText} .= sprintf "		Found %i NAWL words in document (%.02f%%)\n", $keyhash->{WordListMetaData}->{NAWLTotalCount}, , $keyhash->{WordListMetaData}->{NAWLPercent};	
	$keyhash->{SummaryText} .= sprintf "		Found %i new words in document (%.02f%%)\n", $keyhash->{WordListMetaData}->{NewWords}, $keyhash->{WordListMetaData}->{NewPercent};	
	$keyhash->{SummaryText} .= sprintf "		Found %i unknown words in document (%.02f%%)\n", $keyhash->{WordListMetaData}->{TyposCount}, $keyhash->{WordListMetaData}->{UnknownPercent};	
	$keyhash->{SummaryText} .= sprintf "		Found %i Academic Collocations in document\n", $keyhash->{WordListMetaData}->{AcademicCollocationsCount};
	$keyhash->{SummaryText} .= "\n";
	
	return;
}




# SaveReadabilityInfo
# Writes the readability statistics and detailed per-sentence analysis for a
# resource to its dedicated readability worksheet in the output Excel workbook.
# Includes Flesch, Flesch-Kincaid and Gunning Fog scores with interpretation
# guidance.
sub SaveReadabilityInfo
{
	my $keyhash = shift;
	my $sheetname = shift;
	my $start_at = 0;
	my $rownum = 0;
	
	$logger->info ( "Saving Readability Information" );

	my $worksheet = $workbook->add_worksheet( $sheetname );
	$worksheet->set_tab_color( 'green' );
	$worksheet->set_zoom( 150 );

	my $bold = $workbook->add_format( bold => 1 );
	
	my $highlight_format =  $workbook->add_format( bg_color => '#d3d9c3', color => '#080806' );
	#my $green_format =	$workbook->add_format( bg_color => '#d9ead3', color => '#080806' );
	#my $yellow_format =	$workbook->add_format( bg_color => '#fff2cc', color => '#080806' );
	#my $orange_format = $workbook->add_format( bg_color => '#fce5cd', color => '#080806' );
	#my $pink_format =  	$workbook->add_format( bg_color => '#ead1dc', color => '#080806' );
	#my $blue_format =  	$workbook->add_format( bg_color => '#cfe2f3', color => '#080806' );


	$worksheet->set_column( 'A:A',  15 );
	$worksheet->set_column( 'B:B', 30 );
	$worksheet->set_column( 'C:C', 50 );
	$worksheet->set_column( 'D:D', 18 );
	
	#$worksheet->write_number( 1, 1,  );
	#$worksheet->write_number( 2, 1, $keyhash->{Readability}->{FleschKincaid} );
	#$worksheet->write_number( 3, 1, $keyhash->{Readability}->{GunningFog} );

	#
	
	$worksheet->write( $start_at+0, 0, '# characters' );
	$worksheet->write_number( $start_at+0, 1, $keyhash->{Readability}{NumChars} );

	$worksheet->write( $start_at+1, 0, '# words' );
	$worksheet->write_number( $start_at+1, 1, $keyhash->{Readability}{NumWords} );

	$worksheet->write( $start_at+2, 0, '% complex words' );
	$worksheet->write_number( $start_at+2, 1, $keyhash->{Readability}{PercentComplexWords} );

	$worksheet->write( $start_at+3, 0, '# sentences' );
	$worksheet->write_number( $start_at+3, 1, $keyhash->{Readability}{NumSentences} );

	$worksheet->write( $start_at+4, 0, '# text lines' );
	$worksheet->write_number( $start_at+4, 1, $keyhash->{Readability}{NumTextLines} );

	$worksheet->write( $start_at+5, 0, '# blank lines' );
	$worksheet->write_number( $start_at+5, 1, $keyhash->{Readability}{NumBlankLines} );
	
	$worksheet->write( $start_at+6, 0, '# paragraphs' );
	$worksheet->write_number( $start_at+6, 1, $keyhash->{Readability}{NumParagraphs} );
	
	$worksheet->write( $start_at+7, 0, 'Average # syllables per word' );
	$worksheet->write_number( $start_at+7, 1, $keyhash->{Readability}{AverageSyllablesPerWord} );
	
	$worksheet->write( $start_at+8, 0, 'Average # words per sentence' );
	$worksheet->write_number( $start_at+8, 1, $keyhash->{Readability}{AverageWordsPerSentence} );
	
	$worksheet->write( $start_at+9, 0, 'Estimated reading time' );
	$worksheet->write( $start_at+9, 1, $keyhash->{Readability}->{EstimatedReadingTime} . ' mins' );
	
	#
		
	$start_at = 11;
		
	$worksheet->write( $start_at+0, 0, 'About Gunning Fog', $bold );
	$worksheet->write( $start_at+1, 0, 'The Fog index, developed by Robert Gunning, is a well known and simple formula for measuring readability.' );
	$worksheet->write( $start_at+2, 0, 'The index indicates the number of years of formal education a reader of average intelligence would need to read the text once and understand that piece of writing with its word sentence workload.' );
	$worksheet->write( $start_at+3, 0, 'It is calculated by: ( average number of words per sentence + percentage of complex words ) * 0.4  (Note: A complex word consists of three or more syllables)' );

	$start_at++;
	
	$worksheet->write( $start_at+4, 0, '6 - 7.99' );
	$worksheet->write( $start_at+5, 0, '8 - 9.99' );
	$worksheet->write( $start_at+6, 0, '10 - 11.99' );
	$worksheet->write( $start_at+7, 0, '12 - 13.99' );
	$worksheet->write( $start_at+8, 0, '14 - 17.99' );
	$worksheet->write( $start_at+9, 0, '18+' );

	$worksheet->write( $start_at+4, 1, 'Very childish' );
	$worksheet->write( $start_at+5, 1, 'Childish' );
	$worksheet->write( $start_at+6, 1, 'Acceptable' );
	$worksheet->write( $start_at+7, 1, 'Ideal' );
	$worksheet->write( $start_at+8, 1, 'Difficult' );
	$worksheet->write( $start_at+9, 1, 'Unreadable' );
	
	$rownum = 4 if ( $keyhash->{Readability}->{GunningFog} >=  6 && $keyhash->{Readability}->{GunningFog} < 8 );
	$rownum = 5 if ( $keyhash->{Readability}->{GunningFog} >=  8 && $keyhash->{Readability}->{GunningFog} < 10 );
	$rownum = 6 if ( $keyhash->{Readability}->{GunningFog} >= 10 && $keyhash->{Readability}->{GunningFog} < 12 );
	$rownum = 7 if ( $keyhash->{Readability}->{GunningFog} >= 12 && $keyhash->{Readability}->{GunningFog} < 14 );
	$rownum = 8 if ( $keyhash->{Readability}->{GunningFog} >= 14 && $keyhash->{Readability}->{GunningFog} < 18 );
	$rownum = 9 if ( $keyhash->{Readability}->{GunningFog} >= 18 );
	
	$worksheet->set_row( $start_at+$rownum, undef, $highlight_format );
	$worksheet->write( $start_at+$rownum, 2, 'This text is: ' . $keyhash->{Readability}->{GunningFog} );
	
	#
	
	$worksheet->write( $start_at+11, 0, 'About Flesch', $bold );
	$worksheet->write( $start_at+12, 0, 'This score rates text on a 100 point scale.' );
	$worksheet->write( $start_at+13, 0, 'The higher the score, the easier it is to understand the text.' );
	$worksheet->write( $start_at+14, 0, 'A score of 60 to 70 is considered to be optimal.' );
	$worksheet->write( $start_at+15, 0, 'It is calculated by: 206.835 - ( 1.015 * average number of words per sentence ) - ( 84.6 * average number of syllables per word )' );
	
	$start_at++;
	
	$worksheet->write( $start_at+16, 0, '90 - 100' );
	$worksheet->write( $start_at+17, 0, '80 - 89' );
	$worksheet->write( $start_at+18, 0, '70 - 79' );
	$worksheet->write( $start_at+19, 0, '60 - 69' );
	$worksheet->write( $start_at+20, 0, '50 - 59' );
	$worksheet->write( $start_at+21, 0, '30 - 49' );
	$worksheet->write( $start_at+22, 0, '0 - 29' );

	$worksheet->write( $start_at+16, 1, 'Very Easy' );
	$worksheet->write( $start_at+17, 1, 'Easy' );
	$worksheet->write( $start_at+18, 1, 'Fairly Easy' );
	$worksheet->write( $start_at+19, 1, 'Standard' );
	$worksheet->write( $start_at+20, 1, 'Fairly Difficult' );
	$worksheet->write( $start_at+21, 1, 'Difficult' );
	$worksheet->write( $start_at+22, 1, 'Very Confusing' );
	
	$rownum = 16 if ( $keyhash->{Readability}->{Flesch} >= 90 );
	$rownum = 17 if ( $keyhash->{Readability}->{Flesch} >= 80 && $keyhash->{Readability}->{Flesch} < 90 );
	$rownum = 18 if ( $keyhash->{Readability}->{Flesch} >= 70 && $keyhash->{Readability}->{Flesch} < 80 );
	$rownum = 19 if ( $keyhash->{Readability}->{Flesch} >= 60 && $keyhash->{Readability}->{Flesch} < 70 );
	$rownum = 20 if ( $keyhash->{Readability}->{Flesch} >= 50 && $keyhash->{Readability}->{Flesch} < 60 );
	$rownum = 21 if ( $keyhash->{Readability}->{Flesch} >= 30 && $keyhash->{Readability}->{Flesch} < 50 );
	$rownum = 22 if ( $keyhash->{Readability}->{Flesch} >=  0 && $keyhash->{Readability}->{Flesch} < 30 );
	
	$worksheet->set_row( $start_at+$rownum, undef, $highlight_format );
	$worksheet->write( $start_at+$rownum, 2, 'This text is: ' . $keyhash->{Readability}->{Flesch} );
	#
	
	$worksheet->write( $start_at+24, 0, 'About Flesch-Kincaid', $bold );
	$worksheet->write( $start_at+25, 0, 'This score rates text on U.S. grade school level.' );
	$worksheet->write( $start_at+26, 0, 'So a score of 8.0 means that the document can be understood by an native-speaking eighth grader.' );
	$worksheet->write( $start_at+27, 0, 'The US Government Department of Defense uses Flesch-Kincaid Grade Level formula as a standard test.' );
	$worksheet->write( $start_at+28, 0, 'It is calculated by: (11.8 * average number of syllables per word) + (0.39 * average number of words per sentence) - 15.59' );

	$start_at++;

	$worksheet->write( $start_at+29, 0, 'Readability Index', $bold );
	$worksheet->write( $start_at+30, 0, '<3' );
	$worksheet->write( $start_at+31, 0, '3 - 5' );
	$worksheet->write( $start_at+32, 0, '5 - 8' );
	$worksheet->write( $start_at+33, 0, '8 - 10' );
	$worksheet->write( $start_at+34, 0, '10 - 12' );
	$worksheet->write( $start_at+35, 0, '12 - 16' );
	$worksheet->write( $start_at+36, 0, '16+' );

	$worksheet->write( $start_at+29, 1, 'Level', $bold );
	$worksheet->write( $start_at+30, 1, 'Emergent and Early Readers' );
	$worksheet->write( $start_at+31, 1, 'Childrens' );
	$worksheet->write( $start_at+32, 1, 'Young Adult' );
	$worksheet->write( $start_at+33, 1, 'General Adult' );
	$worksheet->write( $start_at+34, 1, 'General Adult' );
	$worksheet->write( $start_at+35, 1, 'Undergraduate' );
	$worksheet->write( $start_at+36, 1, 'Graduate, Post-Graduate, Professional' );

	$worksheet->write( $start_at+29, 2, 'Examples', $bold );
	$worksheet->write( $start_at+30, 2, 'Picture and early reader books.' );
	$worksheet->write( $start_at+31, 2, 'Chapter Books.' );
	$worksheet->write( $start_at+32, 2, 'Advertising copy, Young Adult literature, some news articles.' );
	$worksheet->write( $start_at+33, 2, 'Novels, news articles, blog posts, political speeches.' );
	$worksheet->write( $start_at+34, 2, 'Novels, news articles, blog posts, political speeches.' );
	$worksheet->write( $start_at+35, 2, 'College textbooks.' );
	$worksheet->write( $start_at+36, 2, 'Scholarly Journal and Technical articles.' );

	$worksheet->write( $start_at+29, 3, 'AGU Reading Programme', $bold );
	$worksheet->write( $start_at+30, 3, '' );
	$worksheet->write( $start_at+31, 3, 'READING 1 (4 - 6)' );
	$worksheet->write( $start_at+32, 3, 'READING 2 (6 - 8)' );
	$worksheet->write( $start_at+33, 3, 'READING 3 (8 - 10)' );
	$worksheet->write( $start_at+34, 3, 'READING 4 (10 - 12)' );
	$worksheet->write( $start_at+35, 3, '' );
	$worksheet->write( $start_at+36, 3, '' );

	$rownum = 30 if ( $keyhash->{Readability}->{FleschKincaid} <  3 );
	$rownum = 31 if ( $keyhash->{Readability}->{FleschKincaid} >=  3 && $keyhash->{Readability}->{FleschKincaid} < 5 );
	$rownum = 32 if ( $keyhash->{Readability}->{FleschKincaid} >=  5 && $keyhash->{Readability}->{FleschKincaid} < 8 );
	$rownum = 33 if ( $keyhash->{Readability}->{FleschKincaid} >=  8 && $keyhash->{Readability}->{FleschKincaid} < 10 );
	$rownum = 34 if ( $keyhash->{Readability}->{FleschKincaid} >= 10 && $keyhash->{Readability}->{FleschKincaid} < 12 );
	$rownum = 35 if ( $keyhash->{Readability}->{FleschKincaid} >= 12 && $keyhash->{Readability}->{FleschKincaid} < 16 );
	$rownum = 36 if ( $keyhash->{Readability}->{FleschKincaid} >= 16 );

	$worksheet->set_row( $start_at+$rownum, undef, $highlight_format );
	$worksheet->write( $start_at+$rownum, 4, 'This text is: ' . $keyhash->{Readability}->{FleschKincaid} );

	$worksheet->write( $start_at+38, 0, 'Lexical Density Results', $bold );
	
	$worksheet->write( $start_at+39, 0, 'VOCD Lexical Diversity:' );
	$worksheet->write( $start_at+40, 0, 'VOCD Variance:' );

	$worksheet->write_number( $start_at+39, 1, sprintf "%.2f", $keyhash->{Readability}->{VOCD}->{LexicalDiversity} );
	$worksheet->write_number( $start_at+40, 1, sprintf "%.2f", $keyhash->{Readability}->{VOCD}->{Variance} );

	$worksheet->write( $start_at+42, 0, 'MTLD Lexical Diversity:' );
	$worksheet->write( $start_at+43, 0, 'MTLD Variance:' );

	$worksheet->write_number( $start_at+42, 1, sprintf "%.2f", $keyhash->{Readability}->{MTLD}->{LexicalDiversity} );
	$worksheet->write_number( $start_at+43, 1, sprintf "%.2f", $keyhash->{Readability}->{MTLD}->{Variance} );

	return;
}



# SaveSummaryDetails
# Writes a one-row summary of a resource (filename, word count, readability
# scores, NGSL/NAWL/CEFR coverage percentages) to the 'Summary' worksheet of
# the output Excel workbook.
sub SaveSummaryDetails
{
	my $keyhash = shift;
	my $sheetname = shift;
	my $x = 0;
	my $y = 0;
	
	$logger->info ( "Saving Summary Details" );

	my $worksheet = $workbook->add_worksheet( $sheetname );
	$worksheet->set_tab_color( 'green' );
	$worksheet->freeze_panes( 1, 0 );    # Freeze the first row
	$worksheet->set_zoom( 150 );
	my $bold = $workbook->add_format( bold => 1 );
	my $header = sprintf "Summary of: %s", $keyhash->{filename};
	$worksheet->write_row( 0, 0, [$header], $bold );

	$worksheet->set_column( 'A:A',  14 );
	#$worksheet->set_column( 'B:B', 20 );
	#$worksheet->set_column( 'C:C',  7 );
	#$worksheet->set_column( 'D:D',  7 );
	#$worksheet->set_column( 'E:E', 80 );
	
	my @summary_info = split ( /\n/, $keyhash->{SummaryText} );
	foreach my $item ( @summary_info )
	{
		$x = 0;
		$item =~ s/^\t+//g;		# remove *leading* tabs
		
		my @columns = split( /\t/, $item );
		foreach my $value ( @columns )
		{
			if ( $value =~ /^[0-9,.]+$/ )		# filter to allow only numbers e.g. 1.1	66	7,777.55	NOT -1
			{
				$worksheet->write_number( $y+1, $x, $value );
			}
			else
			{
				$worksheet->write( $y+1, $x, $value );
			}
			$x++;	# horizontal
		}
		$y++;	# vertical
	}
	
	return;
}



# produces a report which looks like this
	#	TOKEN		INSTANCE	LOCATION	IN CONTEXT
	#1	definition	1			37			...=situation...
	#2	despite		1			114			Despite this, there are many different opinions ab...
	#3	experts		1			809			One thing experts do agree on is that any kind of exercise i...
	#4	promote		1			933			...ong with exercise, having a healthy diet can help promote good health.
	#5	stress		1			1591		...n today's modern world, we all have some level of stress in our life.
	#6	stress		2			1634		Different things cause stress for different people.
	#7	stress		3			1734		...and relationships with other people can all cause stress.
	#8	remove		1			1893		...hing to remember is that you can never completely remove stress from your life.
	#9	stress		4			1900		...remember is that you can never completely remove stress from your life.

# SaveVocabArray
# Writes the full vocabulary frequency list for a resource to its dedicated
# vocabulary worksheet. Each row contains a word, its frequency, its CEFR
# level, and flags indicating NGSL/NAWL membership.
sub SaveVocabArray
{
	my $keyhash = shift;
	my $sheetname = shift;	
	my @array;
	my $i = 0;
	
	$logger->info ( sprintf "		Saving %s vocab info to %s...", $sheetname, $keyhash->{Workbook} );	

	my $worksheet = $workbook->add_worksheet( $sheetname );
	$worksheet->set_tab_color( 'purple' );
	$worksheet->freeze_panes( 1, 0 );    # Freeze the first row
	$worksheet->set_zoom( 150 );
	my $bold = $workbook->add_format( bold => 1 );
	$worksheet->write_row( 0, 0, ['ID', 'Word', 'Instance', 'Location', 'In Context'], $bold );

	$worksheet->set_column( 'A:A',  4 );
	$worksheet->set_column( 'B:B', 20 );
	$worksheet->set_column( 'C:C',  7 );
	$worksheet->set_column( 'D:D',  7 );
	$worksheet->set_column( 'E:E', 80 );
	
	@array = @{ $keyhash->{Dictionaries}->{NewWords} } 		if ( $sheetname eq 'New Words' );
	@array = @{ $keyhash->{Dictionaries}->{UnknownWords} } 	if ( $sheetname eq 'Unknown Words' );	
	
	foreach my $item ( @array )
	{
		$worksheet->write( $i+1, 0, $i );		# row, column, token
		$worksheet->write( $i+1, 1, $item );
		$worksheet->write( $i+1, 2, '' );
		$worksheet->write( $i+1, 3, '' );
		$worksheet->write( $i+1, 4, '0' );

		$i++;
	}
	
	return;
}



# SaveAcademicCollocationsInfo
# Writes the academic collocations found in a resource to its dedicated
# collocations worksheet. Each row shows the collocation, the number of
# occurrences, and the surrounding sentence context.
sub SaveAcademicCollocationsInfo
{
	my $keyhash = shift;
	my $sheetname = shift;		
	my $i = 0;
	my $hascollocations = 0;

	$logger->info ( "Saving Academic Collocations Info" );
	
	# before we do anything, find out if we have collocations to export...
	if ( defined $keyhash )
	{
		my %AcademicCollocationsInDocument = %{ $keyhash->{Dictionaries}->{AcademicCollocations} };	# dereference this
		foreach my $outerkey ( sort keys %AcademicCollocationsInDocument )
		{ 
			my $innerkey = $AcademicCollocationsInDocument{$outerkey};
			next if ( $innerkey->{found_count} == 0 );
			$hascollocations++;
		}
	}
	else
	{
		foreach my $outerkey ( sort keys %AcademicCollocationList )
		{ 
			my $innerkey = $AcademicCollocationList{$outerkey};
			next if ( $innerkey->{found_count} == 0 );
			$hascollocations++;
		}
	}
	
	#
	#
	#
	
	if ( $hascollocations > 0 )
	{
		$hascollocations++; # add one to this to ensure the conditional formatting works all the way down to the bottom
		my $worksheet = $workbook->add_worksheet( $sheetname );
		$worksheet->set_tab_color( 'pink' );
		$worksheet->freeze_panes( 1, 0 );    # Freeze the first row
		$worksheet->set_zoom( 150 );
		
		my $bold = $workbook->add_format( bold => 1 );
		my $centre = $workbook->add_format( align => 'center' );		
		my $olive_format =  $workbook->add_format( bg_color => '#d3d9c3', color => '#080806' );
		my $green_format =	$workbook->add_format( bg_color => '#d9ead3', color => '#080806' );
		my $yellow_format =	$workbook->add_format( bg_color => '#fff2cc', color => '#080806' );
		my $orange_format = $workbook->add_format( bg_color => '#fce5cd', color => '#080806' );
		my $pink_format =  	$workbook->add_format( bg_color => '#ead1dc', color => '#080806' );
		my $blue_format =  	$workbook->add_format( bg_color => '#cfe2f3', color => '#080806' );
	
		$worksheet->write_row( 0, 0, ['#', 'Addition', 'Part Of Speech 1', 'Component 1', 'Part Of Speech 2', 'Component 2', 'Addition', 'Found Count'], $bold );
	
		# highlight different parts of speech of words POS_1
		$worksheet->conditional_formatting( 'C2:C'.$hascollocations, { type => 'text', criteria => 'containing', value => 'adj', format => $blue_format } );
		$worksheet->conditional_formatting( 'C2:C'.$hascollocations, { type => 'text', criteria => 'containing', value => 'adv', format => $pink_format } );
		$worksheet->conditional_formatting( 'C2:C'.$hascollocations, { type => 'text', criteria => 'containing', value => 'n', format => $orange_format } );
		$worksheet->conditional_formatting( 'C2:C'.$hascollocations, { type => 'text', criteria => 'containing', value => 'vpp', format => $yellow_format } );
		$worksheet->conditional_formatting( 'C2:C'.$hascollocations, { type => 'text', criteria => 'containing', value => 'v' , format => $green_format } );

		# highlight different parts of speech of words POS_2
		$worksheet->conditional_formatting( 'E2:E'.$hascollocations, { type => 'text', criteria => 'containing', value => 'adj', format => $blue_format } );
		$worksheet->conditional_formatting( 'E2:E'.$hascollocations, { type => 'text', criteria => 'containing', value => 'adv', format => $pink_format } );
		$worksheet->conditional_formatting( 'E2:E'.$hascollocations, { type => 'text', criteria => 'containing', value => 'n', format => $orange_format } );
		$worksheet->conditional_formatting( 'E2:E'.$hascollocations, { type => 'text', criteria => 'containing', value => 'vpp', format => $yellow_format } );
		$worksheet->conditional_formatting( 'E2:E'.$hascollocations, { type => 'text', criteria => 'containing', value => 'v' , format => $green_format } );

		$worksheet->set_column( 'A:A',  5 );
		$worksheet->set_column( 'B:B',  8, undef, 1 );	# hide this column
		$worksheet->set_column( 'C:C', 12, undef, 1 );	# hide this column
		$worksheet->set_column( 'D:D', 12 );
		$worksheet->set_column( 'E:E', 12, undef, 1 );	# hide this column
		$worksheet->set_column( 'F:F', 12 );
		$worksheet->set_column( 'G:G',  8, undef, 1 );	# hide this column
		$worksheet->set_column( 'H:H', 11 );
	
		if ( defined $keyhash )
		{
			my %AcademicCollocationsInDocument = %{ $keyhash->{Dictionaries}->{AcademicCollocations} };	# dereference this
			$logger->info ( sprintf "		Saving %s vocab info to %s...", $sheetname, $keyhash->{Workbook} );
		
			foreach my $outerkey ( sort { $b->{found_count} <=> $a->{found_count} } values %AcademicCollocationsInDocument )
			{ 
				#my $innerkey = $AcademicCollocationsInDocument{$outerkey};
				next if ( $outerkey->{found_count} == 0 );		# only export details of collocations we've encountered
			
				$worksheet->write_number( $i+1, 0, $outerkey->{index}, $centre );
				$worksheet->write_string( $i+1, 1, $outerkey->{addition_1} );
				$worksheet->write_string( $i+1, 2, $outerkey->{POS_1}, $centre );
				$worksheet->write_string( $i+1, 3, $outerkey->{component_1} );
				$worksheet->write_string( $i+1, 4, $outerkey->{POS_2}, $centre );
				$worksheet->write_string( $i+1, 5, $outerkey->{component_2} );
				$worksheet->write_string( $i+1, 6, $outerkey->{addition_2} );
				$worksheet->write_number( $i+1, 7, $outerkey->{found_count}, $centre );
			
				$i++;
			}
		}
		else
		{
			$logger->info ( sprintf "Saving %s vocab info to archive summary file...", $sheetname );
		
			foreach my $outerkey ( sort { $b->{found_count} <=> $a->{found_count} } values %AcademicCollocationList )
			#foreach my $outerkey ( sort keys %AcademicCollocationList )
			{ 
				#my $innerkey = $AcademicCollocationList{$outerkey};
				next if ( $outerkey->{found_count} == 0 );		# only export details of collocations we've encountered
			
				$worksheet->write_number( $i+1, 0, $outerkey->{index}, $centre );
				$worksheet->write_string( $i+1, 1, $outerkey->{addition_1} );
				$worksheet->write_string( $i+1, 2, $outerkey->{POS_1}, $centre );
				$worksheet->write_string( $i+1, 3, $outerkey->{component_1} );
				$worksheet->write_string( $i+1, 4, $outerkey->{POS_2}, $centre );
				$worksheet->write_string( $i+1, 5, $outerkey->{component_2} );
				$worksheet->write_string( $i+1, 6, $outerkey->{addition_2} );
				$worksheet->write_number( $i+1, 7, $outerkey->{found_count}, $centre );
			
				$i++;
			}
		}
		
		$worksheet->write_string( $i+2, 0, "The complete list of collocations can be found at: https://pearsonpte.com/organizations/researchers/academic-collocation-list/", $bold );
	}

	return;
}



# SaveNGSLAndNAWLInfo
# Writes NGSL and NAWL vocabulary analysis results to the dedicated NGSL/NAWL
# worksheet for a resource. Includes coverage percentages, word-level
# breakdown tables, and lists of off-list words.
sub SaveNGSLAndNAWLInfo
{
	my $keyhash = shift;
	my $sheetname = shift;		
	my $i = 0;
	my $record_count = scalar %ngsl;
		
	$logger->info ( "Saving NGSL and NAWL Info" );

	my $worksheet = $workbook->add_worksheet( $sheetname );
	$worksheet->set_tab_color( 'blue' );
	$worksheet->freeze_panes( 1, 0 );    # Freeze the first row
	$worksheet->set_zoom( 150 );
	my $bold = $workbook->add_format( bold => 1 );
	my $centre = $workbook->add_format( align => 'center' );

	my $olive_format =  $workbook->add_format( bg_color => '#d3d9c3', color => '#080806' );
	my $green_format =	$workbook->add_format( bg_color => '#d9ead3', color => '#080806' );
	my $yellow_format =	$workbook->add_format( bg_color => '#fff2cc', color => '#080806' );
	my $orange_format = $workbook->add_format( bg_color => '#fce5cd', color => '#080806' );
	my $pink_format =  	$workbook->add_format( bg_color => '#ead1dc', color => '#080806' );
	my $blue_format =  	$workbook->add_format( bg_color => '#cfe2f3', color => '#080806' );
	
	$worksheet->write_row( 0, 0, ['#', 'Index', 'Level', 'Root', 'Member', 'Is Root?', 'All Inflections Found Count', 'Member Found Count', 'Is New?'], $bold );
	
	# highlight different bands of words
	$worksheet->conditional_formatting( 'C2:C'.$record_count, { type => 'text', criteria => 'containing', value => 'BandA', format => $blue_format } );
	$worksheet->conditional_formatting( 'C2:C'.$record_count, { type => 'text', criteria => 'containing', value => 'BandB', format => $pink_format } );
	$worksheet->conditional_formatting( 'C2:C'.$record_count, { type => 'text', criteria => 'containing', value => 'BandC', format => $orange_format } );
	$worksheet->conditional_formatting( 'C2:C'.$record_count, { type => 'text', criteria => 'containing', value => 'BandD', format => $yellow_format } );
	$worksheet->conditional_formatting( 'C2:C'.$record_count, { type => 'text', criteria => 'containing', value => 'NAWL' , format => $green_format } );
	
	$worksheet->set_column( 'A:A',  4 );
	$worksheet->set_column( 'B:B',  6, undef, 1 );	# hide this column
	$worksheet->set_column( 'C:C',  6 );
	$worksheet->set_column( 'D:D', 12 );
	$worksheet->set_column( 'E:E', 12 );
	$worksheet->set_column( 'F:F',  7 );
	$worksheet->set_column( 'G:G', 21.5 );
	$worksheet->set_column( 'H:H', 18 );
	$worksheet->set_column( 'I:I',  6 );
	
	if ( defined $keyhash )
	{
		my %NGSLInDocument = %{ $keyhash->{Dictionaries}->{NGSL} };	# dereference this
		my %NAWLInDocument = %{ $keyhash->{Dictionaries}->{NAWL} };	# dereference this
		
		$logger->info ( sprintf "		Saving %s vocab info to %s...", $sheetname, $keyhash->{Workbook} );

		my %vocab_hash_document = ( %NGSLInDocument, %NAWLInDocument );		
		foreach my $outerkey ( sort { $b->{found_count} <=> $a->{found_count} } values %vocab_hash_document )
		{ 
			next if ( $outerkey->{found_count} == 0 && $outerkey->{root_found_count} == 0 );		# export only keys which we've found?
			my $rootstring = ( $outerkey->{isroot} == 1 ) ? 'YES' : '';
		
			$worksheet->write_number( $i+1, 0, $i, $centre );		# row, column, token
			$worksheet->write_number( $i+1, 1, $outerkey->{index}, $centre );
			$worksheet->write_string( $i+1, 2, $outerkey->{level}, $centre );
			$worksheet->write_string( $i+1, 3, $outerkey->{root} );
			$worksheet->write_string( $i+1, 4, $outerkey->{member} );
			$worksheet->write_string( $i+1, 5, $rootstring, $centre );
			$worksheet->write( $i+1, 6, ( $outerkey->{isroot} == 1 ) ? $outerkey->{root_found_count} : '' );
			$worksheet->write_number( $i+1, 7, $outerkey->{found_count} );
			$worksheet->write( $i+1, 8, ( $outerkey->{isnew} == 1 ) ? 'YES' : '' );
			$i++;
		}			
	}
	else
	{
		$logger->info ( sprintf "Saving %s vocab info to archive summary file...", $sheetname );

		my %vocab_hash = ( %ngsl, %nawl );
		foreach my $outerkey ( sort { $b->{found_count} <=> $a->{found_count} } values %vocab_hash )
		{ 
			next if ( $outerkey->{found_count} == 0 && $outerkey->{root_found_count} == 0 );		# export only keys which we've found?
			my $rootstring = ( $outerkey->{isroot} == 1 ) ? 'YES' : '';
		
			$worksheet->write_number( $i+1, 0, $i, $centre );		# row, column, token
			$worksheet->write_number( $i+1, 1, $outerkey->{index}, $centre );
			$worksheet->write_string( $i+1, 2, $outerkey->{level}, $centre );
			$worksheet->write_string( $i+1, 3, $outerkey->{root} );
			$worksheet->write_string( $i+1, 4, $outerkey->{member} );
			$worksheet->write_string( $i+1, 5, $rootstring, $centre );
			$worksheet->write( $i+1, 6, ( $outerkey->{isroot} == 1 ) ? $outerkey->{root_found_count} : '' );
			$worksheet->write_number( $i+1, 7, $outerkey->{found_count} );
			$i++;
		}		
	}
	
	#
	# Add a 'key' worksheet
	#
	
	$worksheet = $workbook->add_worksheet( 'NGSL+NAWL Key' );
	$worksheet->set_tab_color( 'blue' );
	$worksheet->freeze_panes( 1, 0 );    # Freeze the first row
	$worksheet->set_zoom( 150 );
	
	$worksheet->write_row( 0, 0, ['Level', 'Meaning'], $bold );

	# highlight different bands of words
	$worksheet->conditional_formatting( 'A2:A'.$record_count, { type => 'text', criteria => 'containing', value => 'BandA', format => $blue_format } );
	$worksheet->conditional_formatting( 'A2:A'.$record_count, { type => 'text', criteria => 'containing', value => 'BandB', format => $pink_format } );
	$worksheet->conditional_formatting( 'A2:A'.$record_count, { type => 'text', criteria => 'containing', value => 'BandC', format => $orange_format } );
	$worksheet->conditional_formatting( 'A2:A'.$record_count, { type => 'text', criteria => 'containing', value => 'BandD', format => $yellow_format } );
	$worksheet->conditional_formatting( 'A2:A'.$record_count, { type => 'text', criteria => 'containing', value => 'NAWL' , format => $green_format } );

	$worksheet->set_column( 'A:A', 7 );
	$worksheet->set_column( 'B:B', 90 );
	
	$worksheet->write_string( 1, 0, 'BandA', $centre );
	$worksheet->write_string( 2, 0, 'BandB', $centre  );
	$worksheet->write_string( 3, 0, 'BandC', $centre );
	$worksheet->write_string( 4, 0, 'BandD', $centre  );
	$worksheet->write_string( 5, 0, 'NAWL', $centre );

	my $BandAString = sprintf "Words %i to %i from the NGSL (New General Service List).  Foundation students should know all of these words.", $NGSL_BAND_A_LOWER, $NGSL_BAND_A_HIGHER; 
	my $BandBString = sprintf "Words %i to %i from the NGSL (New General Service List).  Level 1 students should know all of these words.", $NGSL_BAND_B_LOWER, $NGSL_BAND_B_HIGHER;
	my $BandCString = sprintf "Words %i to %i from the NGSL (New General Service List).  Level 2 students should know all of these words.", $NGSL_BAND_C_LOWER, $NGSL_BAND_C_HIGHER;
	my $BandDString = sprintf "Words %i to %i from the NGSL (New General Service List).  Level 3 students should know all of these words.", $NGSL_BAND_D_LOWER, $NGSL_BAND_D_HIGHER;
	my $NAWLString = sprintf "All word from the NAWL (New Academic Word List).  Level 4 students should know all of these words.";
	
	$worksheet->write_string( 1, 1, $BandAString );
	$worksheet->write_string( 2, 1, $BandBString );
	$worksheet->write_string( 3, 1, $BandCString );
	$worksheet->write_string( 4, 1, $BandDString );
	$worksheet->write_string( 5, 1, $NAWLString );

	return;
}


# SaveCEFRVocabInfo
# Writes CEFR vocabulary analysis results to the dedicated CEFR worksheet for
# a resource. Includes per-band coverage percentages and word lists for each
# CEFR level (A1 through C2) plus an off-list category.
sub SaveCEFRVocabInfo
{		
	my $keyhash = shift;	
	my $sheetname = shift;	

	my $i = 0;
	my $nth = 1;
	my $record_count = scalar %cefr;
		
	$logger->info ( "Saving CEFR Vocab Info" );
	
	if ( defined $keyhash )
	{
		my %CEFRInDocument = %{ $keyhash->{Dictionaries}->{CEFR} };	# dereference this
		$logger->info ( sprintf "		Saving %s vocab info to %s... ", $sheetname, $keyhash->{Workbook} );	

		my $worksheet = $workbook->add_worksheet( $sheetname );
		$worksheet->set_tab_color( 'purple' );
		$worksheet->freeze_panes( 1, 0 );    # Freeze the first row
		$worksheet->set_zoom( 150 );
		my $bold = $workbook->add_format( bold   => 1 );
		my $centre = $workbook->add_format( align => 'center' );

		my $olive_format =  $workbook->add_format( bg_color => '#d3d9c3', color => '#080806' );
		my $green_format =	$workbook->add_format( bg_color => '#d9ead3', color => '#080806' );
		my $yellow_format =	$workbook->add_format( bg_color => '#fff2cc', color => '#080806' );
		my $orange_format = $workbook->add_format( bg_color => '#fce5cd', color => '#080806' );
		my $pink_format =  	$workbook->add_format( bg_color => '#ead1dc', color => '#080806' );
		my $blue_format =  	$workbook->add_format( bg_color => '#cfe2f3', color => '#080806' );
	
		# highlight different bands of words
		$worksheet->conditional_formatting( 'B2:B'.$record_count, { type => 'text', criteria => 'containing', value => 'A1', format => $blue_format } );
		$worksheet->conditional_formatting( 'B2:B'.$record_count, { type => 'text', criteria => 'containing', value => 'A2', format => $pink_format } );
		$worksheet->conditional_formatting( 'B2:B'.$record_count, { type => 'text', criteria => 'containing', value => 'B1', format => $orange_format } );
		$worksheet->conditional_formatting( 'B2:B'.$record_count, { type => 'text', criteria => 'containing', value => 'B2', format => $yellow_format } );
		$worksheet->conditional_formatting( 'B2:B'.$record_count, { type => 'text', criteria => 'containing', value => 'C1' , format => $green_format } );
		$worksheet->conditional_formatting( 'B2:B'.$record_count, { type => 'text', criteria => 'containing', value => 'C2' , format => $olive_format } );
	
		$worksheet->write_row( 0, 0, ['#', 'Level', 'Member', 'Found Count'], $bold );

		$worksheet->set_column( 'A:A', 4.5 );
		$worksheet->set_column( 'B:B', 5 );
		$worksheet->set_column( 'C:C', 40 );
		$worksheet->set_column( 'D:D', 11 );

		foreach my $outerkey ( sort { $b->{found_count} <=> $a->{found_count} or
									  $b->{level} cmp $a->{level} or
									  $a->{member} cmp $b->{member}} values %CEFRInDocument )
		{ 
			next if ( $outerkey->{found_count} == 0 );		# export only keys which we've found?
		
			$i++;		
			$worksheet->write( $i, 0, $i, $centre );		# row, column, token
			$worksheet->write( $i, 1, $outerkey->{level}, $centre );
			$worksheet->write( $i, 2, $outerkey->{member} );
			$worksheet->write( $i, 3, $outerkey->{found_count}, $centre );
			#$worksheet->write( $i, 4, '' );
			#$worksheet->write( $i, 5, '' );
			#$worksheet->write( $i, 6, '' );
		}
	}
	else
	{
		$logger->info ( sprintf "Saving %s vocab info to archive summary file... ", $sheetname );	

		my $worksheet = $workbook->add_worksheet( $sheetname );
		$worksheet->set_tab_color( 'purple' );
		$worksheet->freeze_panes( 1, 0 );    # Freeze the first row
		$worksheet->set_zoom( 150 );
		my $bold = $workbook->add_format( bold   => 1 );
		my $centre = $workbook->add_format( align => 'center' );

		my $olive_format =  $workbook->add_format( bg_color => '#d3d9c3', color => '#080806' );
		my $green_format =	$workbook->add_format( bg_color => '#d9ead3', color => '#080806' );
		my $yellow_format =	$workbook->add_format( bg_color => '#fff2cc', color => '#080806' );
		my $orange_format = $workbook->add_format( bg_color => '#fce5cd', color => '#080806' );
		my $pink_format =  	$workbook->add_format( bg_color => '#ead1dc', color => '#080806' );
		my $blue_format =  	$workbook->add_format( bg_color => '#cfe2f3', color => '#080806' );
	
		# highlight different bands of words
		$worksheet->conditional_formatting( 'B2:B'.$record_count, { type => 'text', criteria => 'containing', value => 'A1', format => $blue_format } );
		$worksheet->conditional_formatting( 'B2:B'.$record_count, { type => 'text', criteria => 'containing', value => 'A2', format => $pink_format } );
		$worksheet->conditional_formatting( 'B2:B'.$record_count, { type => 'text', criteria => 'containing', value => 'B1', format => $orange_format } );
		$worksheet->conditional_formatting( 'B2:B'.$record_count, { type => 'text', criteria => 'containing', value => 'B2', format => $yellow_format } );
		$worksheet->conditional_formatting( 'B2:B'.$record_count, { type => 'text', criteria => 'containing', value => 'C1' , format => $green_format } );
		$worksheet->conditional_formatting( 'B2:B'.$record_count, { type => 'text', criteria => 'containing', value => 'C2' , format => $olive_format } );
	
		$worksheet->write_row( 0, 0, ['#', 'Level', 'Member', 'Found Count'], $bold );

		$worksheet->set_column( 'A:A', 4.5 );
		$worksheet->set_column( 'B:B', 5 );
		$worksheet->set_column( 'C:C', 40 );
		$worksheet->set_column( 'D:D', 11 );

		foreach my $outerkey ( sort { $b->{found_count} <=> $a->{found_count} or
									  $b->{level} cmp $a->{level} or
									  $a->{member} cmp $b->{member}} values %cefr )
		{ 
			# next if ( $outerkey->{found_count} == 0 );		# export only keys which we've found?
		
			$i++;		
			$worksheet->write( $i, 0, $i, $centre );		# row, column, token
			$worksheet->write( $i, 1, $outerkey->{level}, $centre );
			$worksheet->write( $i, 2, $outerkey->{member} );
			$worksheet->write( $i, 3, $outerkey->{found_count}, $centre );
			#$worksheet->write( $i, 4, '' );
			#$worksheet->write( $i, 5, '' );
			#$worksheet->write( $i, 6, '' );
		}
	}
	
	#
	# Add a 'key' worksheet
	#
	
	my $worksheet = $workbook->add_worksheet( 'CEFR Key' );
	$worksheet->set_tab_color( 'purple' );
	$worksheet->freeze_panes( 1, 0 );    # Freeze the first row
	$worksheet->set_zoom( 150 );
	
	my $bold = $workbook->add_format( bold   => 1 );
	my $centre = $workbook->add_format( align => 'center' );
	my $olive_format =  $workbook->add_format( bg_color => '#d3d9c3', color => '#080806' );
	my $green_format =	$workbook->add_format( bg_color => '#d9ead3', color => '#080806' );
	my $yellow_format =	$workbook->add_format( bg_color => '#fff2cc', color => '#080806' );
	my $orange_format = $workbook->add_format( bg_color => '#fce5cd', color => '#080806' );
	my $pink_format =  	$workbook->add_format( bg_color => '#ead1dc', color => '#080806' );
	my $blue_format =  	$workbook->add_format( bg_color => '#cfe2f3', color => '#080806' );
	
	$worksheet->write_row( 0, 0, ['Level', 'Category', 'Meaning'], $bold );

	# highlight different bands of words
	$worksheet->conditional_formatting( 'A2:A7', { type => 'text', criteria => 'containing', value => 'A1', format => $blue_format } );
	$worksheet->conditional_formatting( 'A2:A7', { type => 'text', criteria => 'containing', value => 'A2', format => $pink_format } );
	$worksheet->conditional_formatting( 'A2:A7', { type => 'text', criteria => 'containing', value => 'B1', format => $orange_format } );
	$worksheet->conditional_formatting( 'A2:A7', { type => 'text', criteria => 'containing', value => 'B2', format => $yellow_format } );
	$worksheet->conditional_formatting( 'A2:A7', { type => 'text', criteria => 'containing', value => 'C1' , format => $green_format } );
	$worksheet->conditional_formatting( 'A2:A7', { type => 'text', criteria => 'containing', value => 'C2' , format => $olive_format } );

	$worksheet->set_column( 'A:A', 4.5 );
	$worksheet->set_column( 'B:B', 20 );
	$worksheet->set_column( 'C:C', 20 );
	
	$worksheet->write_row( 1, 0, ['C2', 'Proficient User', 'Mastery'] );
	$worksheet->write_row( 2, 0, ['C1', 'Proficient User', 'Effective Operational Proficiency'] );
	$worksheet->write_row( 3, 0, ['B2', 'Independent User', 'Vantage'] );
	$worksheet->write_row( 4, 0, ['B1', 'Independent User', 'Threshold'] );
	$worksheet->write_row( 5, 0, ['A2', 'Basic User', 'Waystage'] );
	$worksheet->write_row( 6, 0, ['A1', 'Basic User', 'Breakthrough'] );
	
	return;
}




# GenerateHTMLFile
# Generates a marked-up HTML file for a resource in which words are
# colour-coded by CEFR level. The HTML file is written to the destination
# directory and its path recorded in the resource hash.
sub GenerateHTMLFile
{
	my $keyhash = shift;	
	$logger->info ( "Generating HTML file" );
	
	my %NGSLInDocument = %{ $keyhash->{Dictionaries}->{NGSL} } if defined ( $keyhash->{Dictionaries}->{NGSL} );
	my %NAWLInDocument = %{ $keyhash->{Dictionaries}->{NAWL} } if defined ( $keyhash->{Dictionaries}->{NAWL} );
	my %CEFRInDocument = %{ $keyhash->{Dictionaries}->{CEFR} } if defined ( $keyhash->{Dictionaries}->{CEFR} );
	
	if ( open my $HTML, '>', $keyhash->{HTMLfilename} )
	{
		printf $HTML "
<!DOCTYPE html>
<html lang=\"en\">
<style>
		h1 {color:black;  font-family: sans-serif; font-size: 250%%;}
		h2 {color:black;  font-family: sans-serif; font-size: 200%%;}
		p  {			  font-family: sans-serif; font-size: 125%%;}
		.A1 { color: #B22222; }
		.A2 { color: #C71585; }
		.B1 { color: #FF8C00; }
		.B2 { color: #008000; }
		.C1 { color: #00008B; }
		.C2 { color: #663399; }
		.BandA  { color: #52BE80; }
		.BandB { color: #3498DB; }
		.BandC  { color: #9B59B6; }
		.BandD { color: #E67E22; }
		.BandE  { color: #F39C12; }
		.BandF  { color: #D35400; }
		.NAWL { color: #663399; }
</style>
<head>
		<title>%s</title>
</head>
<h1 align=center>%s</h1>
<body>", $keyhash->{TitleofDocument}, $keyhash->{TitleofDocument};

		#
		# CEFR Dictionary
		#

		my $htmltext = $keyhash->{TextofDocument};

		print $HTML "\n\n<br><h2>Highlighted: CEFR</h2>\n";

		print $HTML "<p><br><span class=\"A1\">A1 Words look like this.</span>\n";
		print $HTML "<br><span class=\"A2\">A2 Words look like this.</span>\n";
		print $HTML "<br><span class=\"B1\">B1 Words look like this.</span>\n";
		print $HTML "<br><span class=\"B2\">B2 Words look like this.</span>\n";
		print $HTML "<br><span class=\"C1\">C1 Words look like this.</span>\n";
		print $HTML "<br><span class=\"C2\">C2 Words look like this.</span>\n";
		print $HTML "<br><br><br>";
		
		foreach my $outerkey ( keys %CEFRInDocument )
		{ 
			my $innerkey = $CEFRInDocument{$outerkey};
			if ( $innerkey->{found_count} > 0 )
			{
				if ( index ( "span class", $innerkey->{member} ) == -1 )
				{
					my $from = $innerkey->{member};
					my $to = sprintf "<span class=\"%s\">%s</span>", $innerkey->{level}, $innerkey->{member}; 

					$htmltext =~ s/\b$from\b/$to/g;		# \b matches a word boundary otherwise "or" is matched in "for"
				
					$from = ucfirst $from;				# catch any sentences which start with a CEFR word
					$to = sprintf "<span class=\"%s\">%s</span>", $innerkey->{level}, ucfirst $innerkey->{member}; 
		
					$htmltext =~ s/\b$from\b/$to/g;		# \b matches a word boundary otherwise "or" is matched in "for"
				}
			}
		}
	
		$htmltext =~ s/\n/<br><br>/g;
		print $HTML $htmltext;
		print $HTML "\n</p><br><br><br>\n\n\n";
	
		#
		# New General Service List Dictionary.  2801 items.
		#
	
		$htmltext = $keyhash->{TextofDocument};

		print $HTML "\n\n<br><h2>Highlighted: New General Service List </h2>\n";
	
		print $HTML "<p><br><span class=\"BandA\">(0000-0800) Foundation words look like this.</span>\n";
		print $HTML "<br><span class=\"BandB\">(0801-1600) Level 1 words look like this.</span>\n";
		print $HTML "<br><span class=\"BandC\">(1601-2400) Level 2 words look like this.</span>\n";
		print $HTML "<br><span class=\"BandD\">(2401-2801) Level 3 words look like this.</span>\n";
		#print $HTML "<br><span class=\"Fifth500\">2001-2500 Words look like this.</span>\n";
		#print $HTML "<br><span class=\"Sixth500\">2501-2801 Words look like this.</span>\n";
		print $HTML "<br><br><br>";
	
		foreach my $outerkey ( keys %NGSLInDocument )
		{ 
			my $innerkey = $NGSLInDocument{$outerkey};
			if ( $innerkey->{found_count} > 0 )
			{
				if ( index ( "span class", $innerkey->{member} ) == -1 )
				{
					my $from = $innerkey->{member};
					my $to = sprintf "<span class=\"%s\">%s</span>", $innerkey->{level}, $innerkey->{member}; 

					$htmltext =~ s/\b$from\b/$to/g;		# \b matches a word boundary otherwise "or" is matched in "for"
				
					$from = ucfirst $from;				# catch any sentences which start with a CEFR word
					$to = sprintf "<span class=\"%s\">%s</span>", $innerkey->{level}, ucfirst $innerkey->{member}; 
		
					$htmltext =~ s/\b$from\b/$to/g;		# \b matches a word boundary otherwise "or" is matched in "for"
				}
			}
		}
	
		$htmltext =~ s/\n/<br><br>/g;
		print $HTML $htmltext;
		print $HTML "\n</p><br><br><br>\n\n\n";
	
		#
		# New Academic Word List Dictionary.
		#
	
		$htmltext = $keyhash->{TextofDocument};

		print $HTML "\n\n<br><h2>Highlighted: New Academic Word </h2>\n";
	
		print $HTML "<p><br><span class=\"NAWL\">NAWL Words look like this.</span>\n";
		print $HTML "<br><br><br>";
	
		foreach my $outerkey ( keys %NAWLInDocument )
		{ 
			my $innerkey = $NAWLInDocument{$outerkey};
			if ( $innerkey->{found_count} > 0 )
			{
				if ( index ( "span class", $innerkey->{member} ) == -1 )
				{
					my $from = $innerkey->{member};
					my $to = sprintf "<span class=\"%s\"><strong>%s</strong></span>", $innerkey->{level}, $innerkey->{member}; 

					$htmltext =~ s/\b$from\b/$to/g;		# \b matches a word boundary otherwise "or" is matched in "for"
				
					$from = ucfirst $from;				# catch any sentences which start with a CEFR word
					$to = sprintf "<span class=\"%s\">%s</span>", $innerkey->{level}, ucfirst $innerkey->{member}; 

					$htmltext =~ s/\b$from\b/$to/g;		# \b matches a word boundary otherwise "or" is matched in "for"
				}
			}
		}
	
		$htmltext =~ s/\n/<br><br>/g;
		print $HTML $htmltext;
		print $HTML "\n</p><br><br><br>\n\n\n";
	
		#
		#
		#

		print $HTML '
</body>
</html>';	

		close $HTML;
	
		$logger->info ( sprintf "		Made html file [%s] OK :-)\n", $keyhash->{HTMLfilename} );

	}
	else
	{
		warn "Cannot open $HTML: $!\n";
	}

	return;
} 



# GetReadabilityVOCD
# Calculates the vocd-D lexical diversity measure for the text of a resource
# by repeated random sampling of the token stream at varying sample sizes and
# fitting a curve to the resulting type-token ratios.
sub GetReadabilityVOCD
{
	my $keyhash = shift;
	my $text = $keyhash->{TextofDocument};

    # Create a Diversity object...
	my $diversity = Lingua::Diversity::VOCD->new(
		'length_range'      => [ 35..50 ],
		'num_subsamples'    => 100,
		'min_value'         => 1,
		'max_value'         => 200,
		'precision'         => 0.01,
		'num_trials'        => 3
	);

	# Given some text, get a reference to an array of words...
	my $word_array_ref = split_text(
		'text'          => \$text,
		'unit_regexp'   => qr{[^a-zA-Z]+},
	);

	# Measure lexical diversity...
	my $result = $diversity->measure( $word_array_ref );
    
	# Save results
	$keyhash->{Readability}->{VOCD}->{LexicalDiversity} = $result->get_diversity();
	$keyhash->{Readability}->{VOCD}->{Variance} = $result->get_variance();
	
    # Tag text using Lingua::TreeTagger...
	#my $tagger = Lingua::TreeTagger->new(
	#	'language' => 'english-utf8',
	#	'options'  => [ qw( -token -lemma -no-unknown ) ],
	#);
    
	#my $tagged_text = $tagger->tag_text( \$text );

	# Get references to an array of wordforms and an array of lemmas...
	#my ( $wordform_array_ref, $lemma_array_ref ) = split_tagged_text(
	#	'tagged_text'   => $tagged_text,
	#	'unit'          => 'original',
	#	'category'      => 'lemma',
	#);

    # Measure morphological diversity...
	#$result = $diversity->measure_per_category( $wordform_array_ref, $lemma_array_ref );

	# Save results...
	#$keyhash->{Readability}->{VOCD}->{MorphologicalDiversity} = $result->get_diversity();
	#$keyhash->{Readability}->{VOCD}->{MorphologicalVariance} = $result->get_variance();
	
	return;
}




# GetReadabilityMTLD
# Calculates the Measure of Textual Lexical Diversity (MTLD) for the text of a
# resource by scanning the token stream and measuring how far it travels before
# the type-token ratio drops below a threshold, averaging over forward and
# backward passes.
sub GetReadabilityMTLD
{
	my $keyhash = shift;
	my $text = $keyhash->{TextofDocument};

	# Create a Diversity object...
	my $diversity = Lingua::Diversity::MTLD->new(
		'threshold'         => 0.71,
		'weighting_mode'    => 'within_and_between',
	);

	# Given some text, get a reference to an array of words...
	my $word_array_ref = split_text(
		'text'          => \$text,
		'unit_regexp'   => qr{[^a-zA-Z]+},
	);

	# Measure lexical diversity...
	my $result = $diversity->measure( $word_array_ref );
    
	# Save results
	$keyhash->{Readability}->{MTLD}->{LexicalDiversity} = $result->get_diversity();
	$keyhash->{Readability}->{MTLD}->{Variance} = $result->get_variance();
	
	# Tag text using Lingua::TreeTagger...
	#my $tagger = Lingua::TreeTagger->new(
	#	'language' => 'english-utf8',
	#	'options'  => [ qw( -token -lemma -no-unknown ) ],
	#);

	#my $tagged_text = $tagger->tag_text( \$text );

	# Get references to an array of wordforms and an array of lemmas...
	#my ( $wordform_array_ref, $lemma_array_ref ) = split_tagged_text(
	#	'tagged_text'   => $tagged_text,
	#	'unit'          => 'original',
	#	'category'      => 'lemma',
	#);

	# Measure morphological diversity...
	#$result = $diversity->measure_per_category(
	#	$wordform_array_ref,
	#	$lemma_array_ref,
	#);

	# Save results...
	#$keyhash->{Readability}->{MTLD}->{MorphologicalDiversity} = $result->get_diversity();
	#$keyhash->{Readability}->{MTLD}->{MorphologicalVariance} = $result->get_variance();
	
	return;
}





# BeautifyFilename
# Transforms a raw filename (as found in the manifest or on disk) into a
# human-readable display name by removing path components, stripping the file
# extension, replacing underscores and hyphens with spaces, and applying
# title-case capitalisation rules.
sub BeautifyFilename
{
	my $path_and_filename = shift;
	my $old_path_and_filename = $path_and_filename;
	
	my($filename, $directories, $suffix) = fileparse( $path_and_filename, qr/\.[^.]*/ );	
	$directories = "" if $directories eq "./";
	my $oldfilename = $filename;	# save the original
	
	#printf "BeautifyFilename: Directory [%s]; Filename [%s]; Suffix [%s]\n", $directories, $filename, $suffix;
	
	$filename =~ s/TEXT-/TEXT /g;			# fix this
	$filename =~ s/TEXT-/TEXT /g;			# fix this
	$filename =~ s/TASKS1/TASKS/g;			# fix this
	$filename =~ s/\bTEXT/TEXT/g;			# fix this (filenames such as: 0300 The New FamilyTEXT-.docx)
	$filename =~ s/\b- /\b /g;				# fix this (filenames such as: 0300 The New FamilyTEXT-.docx)
	$filename =~ s/_/ /g;					# replace underlines with spaces
	$filename =~ s/  / /g;					# replace two spaces with one		
	$filename =~ s/--/-/g;					# replace -- with -	
	$filename =~ s/\b \.docx/\.docx/g;		# replace answers .docx with answers.docx
	$filename =~ s/\s+\./\./g;				# replace space dot with dot		
	$filename =~ s/^\s+|\s+$//g;			# remove leading and trailing spaces2
	$filename =~ s/ - / /g;					# fix this
	
	my @WordsInFilename = split( /\s+/, $filename);
	
	foreach my $word ( @WordsInFilename )
	{
		$word = ucfirst $word;
		$word = uc $word if ( lc $word eq 'text' || lc $word eq 'tasks' || lc $word eq 'answers' );
		
		$word = 'TEXT.docx' if $word =~ /^Text.docx/;
		$word = 'HANDOUT.docx' if $word =~ /^Handout.docx/;
		$word = 'TASKS.docx' if $word =~ /^Tasks.docx/;
		$word = 'TASKS.docx' if $word =~ /^Task.docx/;
		$word = 'ANSWERS.docx' if $word =~ /^Answers.docx/;
	}
	
	# put the filename back together
	$filename = join ( ' ', @WordsInFilename );
	
	$logger->info ( sprintf "\tBeautifyFilename was [%s] is now [%s]\n", $old_path_and_filename, $directories.$filename.$suffix ) if ( $old_path_and_filename ne $directories.$filename.$suffix );
	return $directories.$filename.$suffix;
}





# GetXMLData
# Parses the imsmanifest.xml file inside an unpacked archive using XML::Simple
# and returns the resulting data structure. Also extracts the top-level
# organisation and resource elements for use during the Dive phase.
sub GetXMLData
{
	my $keyhash = shift;

	$logger->info ( sprintf "Getting XML data for file [%s]", $keyhash->{filename} );
	if ( -e $keyhash->{filename} )
	{
		# read XML file
		my $data = $xml->XMLin( $keyhash->{filename} ) or warn "Could not read $keyhash->{filename}: $!";
	
		#print "Printing contents of data\n";
		#print Dumper($data);
	
		if ( defined $data->{'url'}->{'href'} && length($data->{'url'}->{'href'}) != 0 )
		{
			$keyhash->{title} = $data->{'title'};
			$keyhash->{XML}->{URL} = $data->{'url'}->{'href'};
			$keyhash->{type} = 'weblink';
			$weblinks++;
		
			push @weblinks, $keyhash->{XML}->{URL};
			#printf "	XML Title = %s.  URL = %s.  Type = %s\n", $keyhash->{title}, $keyhash->{XML}->{URL}, $keyhash->{type};	
		}
		else
		{
			#print "Printing contents of data\n";
			#print Dumper($data);
	
			#$keyhash->{title} = (length($data->{'title'}) > 0)  ? $data->{'title'} : $keyhash->{title};
			$keyhash->{title} = (defined $data->{'title'})  ? $data->{'title'} : $keyhash->{title};
			$keyhash->{type} = 'Schoology Resource';		# only a quiz or discussion????
			$schoologyresources++;
	
			#printf "	XML Title = %s.  Type = %s\n", $keyhash->{title}, $keyhash->{type};	
		}		
	}
	else
	{
		$logger->info ( sprintf "	XML File [%s] does not exist.\n", $keyhash->{filename} );			
	}
	
	return;
}


# GetAndSaveNGrams
# Extracts bigram and trigram frequencies from the text of a resource and
# writes them to the NGrams worksheet of the output Excel workbook. Common
# function-word n-grams are filtered out so that content-bearing phrases are
# highlighted.
sub GetAndSaveNGrams
{
	my $keyhash = shift;
	my $index = 0;
	my $ngrams = Lingua::EN::Bigram->new;
	my $stopwords = &getStopWords( 'en' );
	my $worksheet;
	my $text;
	my $bold = $workbook->add_format( bold => 1 );
	my $centre = $workbook->add_format( align => 'center' );
	
	$logger->info ( "Getting and saving NGrams" );
	
	if ( defined $keyhash && length $keyhash->{PrettyTextofDocument} > 0 )
	{
		# remove hyperlinks from text.  e.g.: The storm formed over the Bahamas [HYPERLINK: http://simple.wikipedia.org/wiki/Bahamas] on August 23...
		$text = $keyhash->{PrettyTextofDocument};
		$text =~ s/\[HYPERLINK: \S+\]//g;
		
		$ngrams->text( $text ); 
		$logger->info ( sprintf "Saving N-Grams data to %s (text length is %i characters)...", $keyhash->{Workbook}, length $text );	
	}
	else
	{
		$text = $AllTexts;
		if ( length $text > 0 )
		{
			$text =~ s/\[HYPERLINK: \S+\]//g;
			$text =~ s/(<-+|-+>).*//g;
			$text =~ tr/0-9a-z A-Z.'-//dc;
		
			$ngrams->text( $text );
			$logger->info ( sprintf "Saving N-Grams data to archive workbook (text length is %i characters)...", length $text );			
		}
	}
	
	if ( length $text > 0 )
	{
		# list grams according to frequency
		my @bigrams = $ngrams->ngram( 2 );
		my $count = $ngrams->ngram_count( \@bigrams );
	
		if ( $count > 0 )
		{
			$worksheet = $workbook->add_worksheet( '2 NGrams');
			$worksheet->set_tab_color( 'red' );
			$worksheet->freeze_panes( 1, 0 );    # Freeze the first row
			$worksheet->set_zoom( 150 );
			$worksheet->write_row( 0, 0, ['#', 'First Token', 'Second Token', 'Frequency'], $bold );

			$worksheet->set_column( 'A:A',  4 );
			$worksheet->set_column( 'B:B', 15 );
			$worksheet->set_column( 'C:C', 15 );
			$worksheet->set_column( 'D:D',  8 );
	
			foreach my $bigram ( sort { $$count{ $b } <=> $$count{ $a } } keys %$count )
			{
				# get the tokens of the bigram
				my ( $first_token, $second_token ) = split / /, $bigram;
	
				# skip stopwords and punctuation, and counts less than 1
				next if ( $$stopwords{ $first_token } );
				next if ( $first_token =~ /[,.?!:;()\-]/ );
				next if ( $$stopwords{ $second_token } );
				next if ( $second_token =~ /[,.?!:;()\-]/ );
				next if ( $$count{ $bigram } <= 1 );
		
				# increment
				$index++;
				last if ( $index > $NGRAMS_LIMIT );

				# output
				$logger->info ( sprintf "2 NGRAMS: #%i\t%i\t[%s %s]\t\n", $index, $$count{ $bigram }, $first_token, $second_token );
		
				$worksheet->write( $index, 0, $index, $centre );		# row, column, token
				$worksheet->write( $index, 1, $first_token );
				$worksheet->write( $index, 2, $second_token );
				$worksheet->write( $index, 3, $$count{ $bigram }, $centre );
			}
		}
	
		$index = 0;		# reset this
	
		# list trigrams according to frequency
		my @trigrams = $ngrams->ngram( 3 );
		$count = $ngrams->ngram_count( \@trigrams );
	
		if ( $count > 0 )
		{
			$worksheet = $workbook->add_worksheet( '3 NGrams');
			$worksheet->set_tab_color( 'red' );
			$worksheet->freeze_panes( 1, 0 );    # Freeze the first row
			$worksheet->set_zoom( 150 );
			$worksheet->write_row( 0, 0, ['#', 'First Token', 'Second Token', 'Third Token', 'Frequency'], $bold );
	
			$worksheet->set_column( 'A:A',  4 );
			$worksheet->set_column( 'B:B', 15 );
			$worksheet->set_column( 'C:C', 15 );
			$worksheet->set_column( 'D:D', 15 );
			$worksheet->set_column( 'E:E',  8 );
	  
			foreach my $trigram ( sort { $$count{ $b } <=> $$count{ $a } } keys %$count )
			{
				# get the tokens of the bigram
				my ( $first_token, $second_token, $third_token ) = split / /, $trigram;
	
				# skip stopwords and punctuation
				next if ( $$stopwords{ $first_token } );
				next if ( $first_token =~ /[,.?!:;()\-]/ );
				next if ( $$stopwords{ $second_token } );
				next if ( $second_token =~ /[,.?!:;()\-]/ );
				next if ( $$stopwords{ $third_token } );
				next if ( $third_token =~ /[,.?!:;()\-]/ );
				next if ( $$count{ $trigram } <= 1 );
	
				# increment
				$index++;
				last if ( $index > $NGRAMS_LIMIT );

				# output
				$logger->info ( sprintf "3 NGRAMS: #%i\t%i\t[%s %s %s]\n", $index, $$count{ $trigram }, $first_token, $second_token, $third_token );
		
				$worksheet->write( $index, 0, $index, $centre );		# row, column, token
				$worksheet->write( $index, 1, $first_token );
				$worksheet->write( $index, 2, $second_token );	
				$worksheet->write( $index, 3, $third_token );	
				$worksheet->write( $index, 4, $$count{ $trigram }, $centre );
			}
		}
	
		$index = 0;		# reset this
	
		# list quadgrams according to frequency
		my @quadgrams = $ngrams->ngram( 4 );
		$count = $ngrams->ngram_count( \@quadgrams );
	
		if ( $count > 0 )
		{
			$worksheet = $workbook->add_worksheet( '4 NGrams');
			$worksheet->set_tab_color( 'red' );
			$worksheet->freeze_panes( 1, 0 );    # Freeze the first row
			$worksheet->set_zoom( 150 );
			$worksheet->write_row( 0, 0, ['#', 'First Token', 'Second Token', 'Third Token', 'Fourth Token', 'Frequency'], $bold );

			$worksheet->set_column( 'A:A',  4 );
			$worksheet->set_column( 'B:B', 15 );
			$worksheet->set_column( 'C:C', 15 );
			$worksheet->set_column( 'D:D', 15 );
			$worksheet->set_column( 'E:E', 15 );
			$worksheet->set_column( 'F:F',  8 );
	  
			foreach my $quadgram ( sort { $$count{ $b } <=> $$count{ $a } } keys %$count )
			{
				# get the tokens of the bigram
				my ( $first_token, $second_token, $third_token, $fourth_token ) = split / /, $quadgram;
	
				# skip stopwords and punctuation
				next if ( $$stopwords{ $first_token } );
				next if ( $first_token =~ /[,.?!:;()\-]/ );
				next if ( $$stopwords{ $second_token } );
				next if ( $second_token =~ /[,.?!:;()\-]/ );
				next if ( $$stopwords{ $third_token } );
				next if ( $third_token =~ /[,.?!:;()\-]/ );
				next if ( $$stopwords{ $fourth_token } );
				next if ( $fourth_token =~ /[,.?!:;()\-]/ );
				next if ( $$count{ $quadgram } <= 1 );
	
				# increment
				$index++;
				last if ( $index > $NGRAMS_LIMIT );

				# output
				$logger->info ( sprintf "4 NGRAMS: %i\t%i\t[%s %s %s %s]\n", $index, $$count{ $quadgram }, $first_token, $second_token, $third_token, $fourth_token );
		
				$worksheet->write( $index, 0, $index, $centre );		# row, column, token
				$worksheet->write( $index, 1, $first_token );
				$worksheet->write( $index, 2, $second_token );	
				$worksheet->write( $index, 3, $third_token );
				$worksheet->write( $index, 4, $fourth_token );
				$worksheet->write( $index, 5, $$count{ $quadgram }, $centre );
			}	
		}		
	}
	else
	{
		$logger->info ( sprintf "Cannot save N-Grams data to archive workbook (text length is %i characters)...", length $text );
	}
	
	return;
}


# PrepareItems
# Prepares a single resource item for processing. Populates the resource hash
# (%key) with derived filenames, initialises counters and flags, determines
# the resource subtype, and calls the appropriate extraction and analysis
# routines (GetTagsAndText, GetReadabilityStats, GetLexisInformation, etc.).
sub PrepareItems
{
	my $directory = shift;
	my $filename = shift;
	my $item = shift;

	my $exists = 'FALSE';
	my $size = 0;
	my $md5 = 0;
	my $PrettyFilename = '';
	my $OriginalFilename = $filename;
	my $fileondisk = '';
	my %key;
		
	$logger->info ( sprintf "    Preparing item #%04i: %s...\n", $archiveitemcounter, $filename );
	
	if ( -e $filename )
	{
		$exists = 'TRUE';
		$size = -s $filename;
		$md5 = file_md5_hex( $filename );
	}
	else
	{
		$logger->info ( sprintf "    Item [%s] does not exist!  Doing nothing.\n", $filename );
	}
	
	# before we do anything with this, tidy up the filename.
	# $filename = BeautifyFilename( $filename );
	
	if ( $OriginalFilename ne $filename )
	{
		move ( $OriginalFilename, $filename );
		#printf "	* Renamed file from [%s] to [%s]\n", $OriginalFilename, $filename;
	}
		
	my ( $FilenameNoExt, $Directories, $Ext ) = fileparse( $filename, qr/\.[^.]*/ );
	$Ext =~ s/\.//g;	# remove the dot  (before:	.docx	.pdf	.xml	after:	docx	pdf		xml)
	$Ext = lc $Ext;
	
	$key{type} = $Ext;
	$key{level} = $level;
	$key{parent} = $parent;
	$key{location} = $location;
	
	if ( defined ($item) )
	{
		$key{identifier} = $item->{identifier};
		$key{identifierref} = $item->{identifierref};
	}
	
	# make this filename Sentence Case
	# $FilenameNoExt =~ s/([\w']+)/\u\L$1/g unless ( $FilenameNoExt =~ /^[A-Z]{2,4} 001/ );
	
	my @dirs = split /\//, $location;
	my $last_directory = $dirs[-1];
	
	$key{title} = $FilenameNoExt;
	$key{directory} = $directory;
	$key{LastDirectory} = $last_directory;
	$key{filename} = $filename;
	$key{FilenameNoPathNoExt} = $FilenameNoExt;
	$key{FilenameNoPath} = $FilenameNoExt.'.'.$Ext;
	$key{IdAndFilenameNoPath} = sprintf "%04i %s", $archiveitemcounter + 1, $FilenameNoExt.'.'.$Ext;		# add one so we don't get two files called 0000 SOMETHING.pdf
	$key{IdAndFilenameNoPathAsPDF} = sprintf "%04i %s", $archiveitemcounter + 1, $FilenameNoExt.'.pdf';

	$key{FilenameWithoutExtension} = $Directories.$FilenameNoExt;
	
	$key{course} = $currentcourse;
	$key{unit} = $currentunit;
	
	$key{subtype}  = '';
	$key{exists} = $exists;
	$key{size} = $size;		# store this as the real size in bytes
	$key{MD5} = $md5;	
	$key{Workbook} = $key{FilenameWithoutExtension}." [".$version."].xlsx";
	$key{BareWorkbookFilename} = $key{FilenameNoPathNoExt}." [".$version."].xlsx";
	$key{PDFfilename} = $key{FilenameWithoutExtension}.".pdf";		# these don't have version numbers because we are not responsible for the output
	$key{TXTfilename} = $key{FilenameWithoutExtension}.".txt";		# these don't have version numbers because we are not responsible for the output
	$key{SpellingAndGrammarErrorsfilename} = $key{FilenameWithoutExtension}."-SPELLING_AND_GRAMMAR.csv";		# these don't have version numbers because we are not responsible for the output
	$key{UntouchedFilename} = $key{FilenameWithoutExtension}."_ORIGINAL.docx";
	$key{PrettyTXTfilename} = $key{FilenameWithoutExtension}."-PRETTY [".$version."].txt";
	$key{TaggedTXTfilename} = $key{FilenameWithoutExtension}."-TAGGED [".$version."].txt";
	$key{HTMLfilename} = $key{FilenameWithoutExtension}."-MARKED-UP [".$version."].html";
	$key{ReadabilityFilename} = $key{FilenameWithoutExtension}."-READABILITY [".$version."].txt";
	
	$key{AddedWhiteRectangle} = 0;
				
	# 'zero' these to avoid 'uninitialized value' warnings
	$key{TextofDocument} = '';
	$key{TitleofDocument} = '';
	$key{PrettyTextofDocument} = '';		# same as above but with all punctuation and other superfluous characters removed
	$key{TaggedTextofDocument} = '';
	$key{XML}->{URL} = '';
	
	$key{MetaData}->{createdby} = '';
	$key{MetaData}->{createddate} = '';
	$key{MetaData}->{lastmodifiedby} = '';
	$key{MetaData}->{lastmodifieddate} = '';
	$key{MetaData}->{totaleditingtime} = '';
	$key{MetaData}->{revisionnumber} = '';
	$key{MetaData}->{pages} = '';	# this is populated with the information from Word, unless it's a PDF file in which case we get this information from the PDF
	$key{MetaData}->{AgeInDaysSinceLastEdit} = 0;
	
	$key{WordMetaData}->{paragraphs} = '';
	$key{WordMetaData}->{lines} = '';
	$key{WordMetaData}->{words} = 0;
	$key{Readability}->{wordsdescription} = '';	# this does not come from Word but we determine this based on our word limits per level
	$key{WordMetaData}->{characters} = '';
	$key{WordMetaData}->{source} = '';
	$key{WordMetaData}->{PageSize} = 'A4';					# default
	$key{WordMetaData}->{PageOrientation} = 'portrait';		# default
	$key{WordMetaData}->{PageBorders} = '';	
	$key{MetaData}->{LessonFocus} = undef;		# this is an array.  some documents might have this information in their header
	$key{WordMetaData}->{KeyVocabulary} = '';

	# if the file is an image, we'll have values for these.
	$key{Image}->{OriginalX} = 0;
	$key{Image}->{OriginalY} = 0;
	$key{Image}->{NewX} = 0;
	$key{Image}->{NewY} = 0;
	$key{Image}->{ScaleFactor} = 1;
	
	# our document-specific dictionaries
	#$key{Dictionaries}->{NGSL};
	#$key{Dictionaries}->{NAWL};
	#$key{Dictionaries}->{CEFR};
	#$key{Dictionaries}->{AcademicCollocations};
	
	$key{WordListMetaData}->{WordCount} = '';
	
	$key{WordListMetaData}->{NGSLBandA} = 0;
	$key{WordListMetaData}->{NGSLBandB} = 0;
	$key{WordListMetaData}->{NGSLBandC} = 0;
	$key{WordListMetaData}->{NGSLBandD} = 0;
	$key{WordListMetaData}->{NGSLTotalCount} = 0;
	
	$key{WordListMetaData}->{NAWLTotalCount} = 0;
	$key{WordListMetaData}->{NewWords} = 0;
	$key{WordListMetaData}->{TyposCount} = 0;	
	$key{WordListMetaData}->{NGSLPercent} = 0;
	$key{WordListMetaData}->{NAWLPercent} = 0;
	$key{WordListMetaData}->{NewPercent} = 0;
	$key{WordListMetaData}->{UnknownPercent} = 0;
		
	$key{WordListMetaData}->{AcademicCollocationsCount} = 0;
	
	$key{WordListMetaData}->{CEFR_A1_Count} = 0;
	$key{WordListMetaData}->{CEFR_A2_Count} = 0;
	$key{WordListMetaData}->{CEFR_B1_Count} = 0;
	$key{WordListMetaData}->{CEFR_B2_Count} = 0;
	$key{WordListMetaData}->{CEFR_C1_Count} = 0;
	$key{WordListMetaData}->{CEFR_C2_Count} = 0;

	$key{WordListMetaData}->{CEFR_A1_MultiWord_Count} = 0;
	$key{WordListMetaData}->{CEFR_A2_MultiWord_Count} = 0;
	$key{WordListMetaData}->{CEFR_B1_MultiWord_Count} = 0;
	$key{WordListMetaData}->{CEFR_B2_MultiWord_Count} = 0;
	$key{WordListMetaData}->{CEFR_C1_MultiWord_Count} = 0;
	$key{WordListMetaData}->{CEFR_C2_MultiWord_Count} = 0;
	
	$key{WordListMetaData}->{CEFRTotalCount} = 0;
			
	$key{Readability}->{Flesch} = 0;			# don't zero this, it'll distort the colour scales we apply in the spreadsheet 
	$key{Readability}->{FleschKincaid} = 0;		# don't zero this, it'll distort the colour scales we apply in the spreadsheet
	$key{Readability}->{GunningFog} = 0;		# don't zero this, it'll distort the colour scales we apply in the spreadsheet
	
	$key{Readability}->{FleschDescription} = '';		# text which details if the text is too easy/too difficult
	$key{Readability}->{FleschKincaidDescription} = '';
	$key{Readability}->{GunningFogDescription} = '';
	
	$key{Readability}->{EstimatedReadingTime} = '';
	$key{Readability}->{ReadabilityReport} = '';
	$key{SummaryText} = '';
	
	$key{NumberOfSpellingAndGrammarErrors} = 0;

	$key{SpellingMistakes}->{Punctuation} = 0;
	$key{SpellingMistakes}->{Typography} = 0;
	$key{SpellingMistakes}->{Typos} = 0;
	$key{SpellingMistakes}->{Miscellaneous} = 0;
	$key{SpellingMistakes}->{Casing} = 0;
	$key{SpellingMistakes}->{ConfusedWords} = 0;
	$key{SpellingMistakes}->{Grammar} = 0;
	$key{SpellingMistakes}->{Style} = 0;
	$key{SpellingMistakes}->{Redundancy} = 0;
	$key{SpellingMistakes}->{Semantics} = 0;
	$key{SpellingMistakes}->{NonstandardPhrases} = 0;
	$key{SpellingMistakes}->{Collocations} = 0;
	
	
	$logger->info ( sprintf "	key{PDFfilename} is [%s]\n", $key{PDFfilename} );
	
	#
	#
	#
	
	# these are in a particular order, do not re-order unless you know what you're doing!
	$key{subtype} = 'ANSWERS' 		if ( ( $key{type} eq 'docx' || $key{type} eq 'doc' || $key{type} eq 'pptx' || $key{type} eq 'pdf' ) && $key{FilenameNoPath} =~ /\bANSWER|TEACHER/i );	# E.G. "... TEACHER'S NOTES"
	$key{subtype} = 'HANDOUT' 		if ( ( $key{type} eq 'docx' || $key{type} eq 'doc' || $key{type} eq 'pptx' || $key{type} eq 'pdf' ) && $key{FilenameNoPath} =~ /\bHANDOUT\b/i );	# what about: 'A New Sports Centre Handout Transcript.docx'.  This should be labelled a tapescript. 
	$key{subtype} = 'TAPESCRIPT' 	if ( ( $key{type} eq 'docx' || $key{type} eq 'doc' || $key{type} eq 'pptx' || $key{type} eq 'pdf' ) && $key{FilenameNoPath} =~ /\bTAPESCRIPT\b|\bTRANSCRIPT\b/i );		# if the filename is "Seminar On Rock Art Tapescript With ANSWERS", then the subtype ANSWERS will be applied 
	$key{subtype} = 'TASKS' 		if ( ( $key{type} eq 'docx' || $key{type} eq 'doc' || $key{type} eq 'pptx' || $key{type} eq 'pdf' ) && $key{FilenameNoPath} =~ /\bTASK\b/i );	# have this before the answers to catch things like: "La 3 Reporting Verb Patterns Tasks Ak" 
	$key{subtype} = 'TEXT' 			if ( ( $key{type} eq 'docx' || $key{type} eq 'doc' || $key{type} eq 'pptx' || $key{type} eq 'pdf' ) && $key{FilenameNoPath} =~ /\bTEXT\b/i );
	$key{subtype} = 'SONG' 			if ( ( $key{type} eq 'docx' || $key{type} eq 'doc' || $key{type} eq 'pptx' || $key{type} eq 'pdf' ) && $key{FilenameNoPath} =~ /\bSONG\b/i );
	$key{subtype} = 'LESSON PLAN'   if ( ( $key{type} eq 'docx' || $key{type} eq 'doc' || $key{type} eq 'pptx' || $key{type} eq 'pdf' ) && $key{FilenameNoPath} =~ /\bLESSON PLAN\b/i );
	$key{subtype} = 'UNKNOWN' 		if ( ( $key{type} eq 'docx' || $key{type} eq 'doc' || $key{type} eq 'pptx' || $key{type} eq 'pdf' ) && $key{subtype} eq '' );
	
	$key{subtype} = 'WORKSHEET' 	if ( ( $key{type} eq 'docx' || $key{type} eq 'doc' || $key{type} eq 'pptx' || $key{type} eq 'pdf' ) && $key{FilenameNoPath} =~ /\bWORKSHEET|ACTIVITY|HANDOUT|RESOURCE\b/i );
	$key{subtype} = 'VOCABULARY' 	if ( ( $key{type} eq 'docx' || $key{type} eq 'doc' || $key{type} eq 'pptx' || $key{type} eq 'pdf' ) && $key{FilenameNoPath} =~ /\bVOCABULARY\b/i );
	$key{subtype} = 'HOMEWORK' 		if ( ( $key{type} eq 'docx' || $key{type} eq 'doc' || $key{type} eq 'pptx' || $key{type} eq 'pdf' ) && $key{FilenameNoPath} =~ /\bHOMEWORK|VOCABULARY|EXAM\b/i );
	
	# do this last in case we have any 'homework solutions' in the filename
	$key{subtype} = 'ANSWERS' 		if ( ( $key{type} eq 'docx' || $key{type} eq 'doc' || $key{type} eq 'pptx' || $key{type} eq 'pdf' ) && $key{FilenameNoPath} =~ /\bANSWER\b|TEACHER|\bANSWERS\b|\bSOLUTIONS\b/i );		# /i = case insensitive; ignore speAKing. OLD: | AK|_AK|ANSWER KEY|ANSWER_KEY|-Ak|A\.K| A\.K\.|_A\.K| ATD

	# if the file is too big, exclude it
	$key{subtype} = 'EXCLUDED' if ( $key{size} > $MAXIMUM_FILE_SIZE );
	
	# if the title has any of these words in it, exclude it
	$key{subtype} = 'EXCLUDED' if ( $key{title} =~ /(talk about the topic|booklet|EXCLUDE)/i );
	$key{subtype} = 'EXCLUDED' if ( $key{title} =~ /(lecture note|lecture_note)/i );
	
	# does the folder name/structure contain any words which mean we need to exclude these resources?
	if ( $forceconvert == 0 )
	{
		foreach my $exclude ( @exclusions )
		{
			if ( index ( $key{location}, $exclude ) > -1 )	# STR, SUBSTR, POSITION
			{
				$key{subtype} = 'EXCLUDED';
				$logger->info ( sprintf "	Subtype is now [%s] due to maching exclusions\n", $key{subtype} );
			}
		}
	}
	
	# does the foler name/structure contain anything in @alwaysinclude, if so, we want it in all subtype directories
	foreach my $tier ( @alwaysinclude )
	{
		# $key{subtype} = 'INCLUDED' if ( index ( $key{location}, $tier ) != -1 );
	}
	
	#
	# now that we have determined the subtype, set these up
	#
	
	$key{CommonFilename} 	= sprintf( '%s%04d %s', $destinationDirectory.$key{subtype}.'/', $archiveitemcounter, $key{FilenameNoPath} );
	$key{CommonFilenamePDF} = sprintf( '%s%04d %s', $destinationDirectory.$key{subtype}.'/', $archiveitemcounter, $key{FilenameNoPathNoExt}.'.pdf' );
	
	#$logger->info ( sprintf "CommonFilename     is [%s]\n", $key{CommonFilename} );
	#$logger->info ( sprintf "CommonFilenamePDF  is [%s]\n", $key{CommonFilenamePDF} );	
	
	return if $key{type}     eq "ds_store";
	return if $key{filename} eq "Thumbs.db";
	return if $key{filename} =~ /^\~/;			# starts with ~ 	i.e. temporary file
	
	
	# show this again so we know what the final state of these is
	$logger->info ( sprintf "	* Item location is [%s], parent is [%s], subtype is [%s]\n", $key{location}, $key{parent}, $key{subtype} );
	push @archivedetails, \%key;
	
	return;
}


# ProcessItems
# Iterates over the list of resource items passed in from Dive, calling
# PrepareItems for each one in turn. Handles both individual files and
# directory-based resources (where the actual file must be located by scanning
# the directory).
sub ProcessItems
{
	$logger->info ( sprintf "Processing %i items...\n", scalar @archivedetails );
	
	$workbook = undef;
	my $i = 0;
	
	#
	# the stuff in 'prepareitems' is what we must set up for the convert to work, all of this below is the analytics gathering code
	#
	
	foreach my $keyhash ( @archivedetails )
	{
		$i++;
		$logger->info ( sprintf "\n\n%03i of %03i: Item location is [%s], parent is [%s], subtype is [%s], MD5 is [%s]\n", $i, scalar @archivedetails, $keyhash->{location}, $keyhash->{parent}, $keyhash->{subtype}, $keyhash->{MD5} );

		if ( $quick == 0 )
		{
			if ( $keyhash->{type} eq 'doc' || $keyhash->{type} eq 'docx' || $keyhash->{type} eq 'pdf' || $keyhash->{type} eq 'xlsx' || $keyhash->{type} eq 'pptx' )
			{
				if ( $keyhash->{FilenameNoPath} !~ /^\~/ && ( $keyhash->{type} eq 'docx' || $keyhash->{type} eq 'pptx' ) )	# ignore temporary files such as "~$stival Handout.docx" which may exist when run with the -directory switch
				{
					$logger->info ( sprintf "\tGetting word metadata for %s...\n", $keyhash->{filename} );
					GetOfficeMetadata( $keyhash );
		
					if( $keyhash->{type} eq 'docx' && system ( "/usr/bin/perl", "/usr/local/bin/docx2txt.pl", $keyhash->{filename} ) == 0 )	# success!
					{
						print "\tMade txt file from docx :-)\n";
						
						$logger->info ( sprintf "\tGetting tags for %s\n", $keyhash->{filename} );
						GetTagsAndText( $keyhash );	 # this populates key.TextOfDocument
					
						$logger->info ( sprintf "\tChecking document for grammar and spelling errors...\n" );
						CheckSpellingAndGrammar( $keyhash );
	
						# do this if it's a .docx file
						$logger->info ( sprintf "\tGetting document information for %s...\n", $keyhash->{filename} );
						GetLexisInformation( $keyhash );
					
						if ( $keyhash->{subtype} eq 'TEXT' || $keyhash->{subtype} eq 'TAPESCRIPT' || $keyhash->{subtype} eq 'SONG' )
						{
							# create the stats workbook
							$workbook  = Excel::Writer::XLSX->new( $keyhash->{Workbook} );
					
							# set properties
							$workbook->set_properties(
							    title    => 'This is spreadsheet showing the analytics for '.$keyhash->{title},
							    author   => 'David Ayliffe',
							    comments => 'Created with Perl and Excel::Writer::XLSX',
							);
										
							if ( $keyhash->{WordMetaData}->{words} >= 50 && length( $keyhash->{TextofDocument}) > 0 )
							{	
								$logger->info ( sprintf "\tGetting Readability stats for %s...\n", $keyhash->{filename} );
								GetReadabilityStats( $keyhash );
							
								#$logger->info ( sprintf "\tMeasuring the Textual Lexical Diversity (MTLD) for %s...\n", $keyhash->{filename} );
								#GetReadabilityMTLD( $keyhash );
							
								#$logger->info ( sprintf "\tCalculating the Vocabulary Diversity (VOCD) for %s...\n", $keyhash->{filename} );
								#GetReadabilityVOCD( $keyhash );
							
								$logger->info ( sprintf "	* Saving document and readability information for %s...\n", $keyhash->{filename} );
								SaveSummaryDetails( $keyhash, 'Document' );
								SaveReadabilityInfo( $keyhash, 'Readability' );
							
								$logger->info ( sprintf "\tGetting N-Grams for %s...\n", $keyhash->{filename} );
								GetAndSaveNGrams( $keyhash );
												
								$logger->info ( sprintf "\tMaking HTML file for %s...\n", $keyhash->{filename} );
								GenerateHTMLFile( $keyhash );
						
								$logger->info ( sprintf "\tSaving vocabulary information for %s\n", $keyhash->{filename} );
								SaveAcademicCollocationsInfo( $keyhash, 'Academic Collocations' ) 	if $keyhash->{WordListMetaData}->{AcademicCollocationsCount} > 0;
								SaveNGSLAndNAWLInfo( $keyhash, 'NGSL+NAWL Vocab' ) 					if $keyhash->{WordListMetaData}->{NGSLTotalCount} > 0;
								SaveCEFRVocabInfo( $keyhash, 'CEFR Vocab' ) 						if $keyhash->{WordListMetaData}->{CEFRTotalCount} > 0;
								SaveVocabArray( $keyhash, 'New Words' ) 							if $keyhash->{WordListMetaData}->{NewWords} > 0;
								SaveVocabArray( $keyhash, 'Unknown Words' ) 						if $keyhash->{WordListMetaData}->{TyposCount} > 0;
							}
						} 
						elsif ( $keyhash->{subtype} eq 'TASKS' )
						{
							$logger->info ( sprintf "	* Not getting AWL or Readability data - is a TASK\n" );
						}
				
						if ( $highlightin == 1 && ( scalar @highlightin == 0 || ( scalar @highlightin > 0 && $keyhash->{filename} ~~ /@convert/ ) ) )
						{
							$logger->info ( sprintf "\tApplying style to document %s...\n", $keyhash->{filename} );
							StyleText( $keyhash );
						}	
					}
					elsif( $keyhash->{type} eq 'pptx' )	# success!
					{
						#if ( system ( "/usr/bin/perl", "/usr/local/bin/pptx2txt.pl", $keyhash->{filename}, $keyhash->{TXTfilename} ) == 0 )
						if ( system ( "/opt/homebrew/bin/python3", "python_pptx.py", $keyhash->{filename}, $keyhash->{TXTfilename} ) == 0 )
						{
							print "\tMade txt file from pptx :-)\n";	
							
							$logger->info ( sprintf "\tChecking presentation for grammar and spelling errors...\n" );
							CheckSpellingAndGrammar( $keyhash );
						}
						else
						{
							print "Failed to make txt file from pptx: $_\n";
						}
					}
					elsif ( $keyhash->{type} ne 'docx' && $keyhash->{type} ne 'pptx'  )
					{
						$logger->info ( "It's not a docx or pptx file, no need to convert to txt.\n" );
					}
					else
					{
						$logger->info ( "Failed to convert .docx/.pptx to .txt :-(\n" );
					}	
				}
				elsif ( $keyhash->{type} eq 'pdf' )
				{
					$logger->info ( sprintf "\tGetting PDF metadata for %s...\n", $keyhash->{filename} );
					GetPDFMetadata( $keyhash );
				}
				elsif ( $keyhash->{FilenameNoPath} =~ /^\~/ )	# this could be a temporary docx file or pptx
				{
					$logger->info ( sprintf "\t%s looks like a temporary file.  Ignoring.\n", $keyhash->{filename} );
					
				}
		
				$logger->info ( sprintf "	* Document summary for %s is:\n", $keyhash->{filename} ) if ( length $keyhash->{SummaryText} > 0 );
				$logger->info ( $keyhash->{SummaryText} );
			}
			elsif ( $keyhash->{type} eq 'xml' )
			{
				# this sets the title, the XML URL and the updates the type
				#printf "Getting XML information for %s (type is %s)\n", $key{filename}, $key{type};
		
				if( system ( "/usr/bin/perl", "/usr/local/bin/GetXMLDataTest.pl", $keyhash->{filename}, ">/dev/null" ) == 0 )	# success!
				{
					$logger->info ( sprintf "\tGetting XML data for %s...\n", $keyhash->{filename} );
					GetXMLData( $keyhash );
				}
				else
				{
					push @unabletoparseXML, $keyhash->{filename};
					$logger->info ( sprintf "Failed to parse %s\n", $keyhash->{filename} );
				}
			}
			elsif( $keyhash->{type} eq 'html' )
			{
				$keyhash->{type} = 'Schoology Assignment (or page)';
			}
			elsif ( $keyhash->{type} eq 'mp3' )
			{
				# get mp3 audio length here
				my $mp3 = get_mp3info( $keyhash->{filename} );
				$keyhash->{AudioLength} = $mp3->{MM}.'m:'.$mp3->{SS}.'s';

				$logger->info ( sprintf "	* Audio is %s in length\n", $keyhash->{AudioLength} );
			}	
		}
		else
		{
			$logger->info ( "	Running in quick mode, skipping lots of stuff\n" );	
		
			# if we want the answers, we need to do this, even if we're running in quick mode
			if( $withanswers == 1 && $keyhash->{type} eq 'docx' )
			{
				if( system ( "/usr/bin/perl", "/usr/local/bin/docx2txt.pl", $keyhash->{filename} ) == 0 )	# success!
				{
					$logger->info ( sprintf "	* Getting tags for %s\n", $keyhash->{filename} );
					GetTagsAndText( $keyhash );	 # this populates key.TextOfDocument
				}			
			}
		
			# we still want to do this because it might be a big PDF which we want to exclude
			GetPDFMetadata( $keyhash ) if ( $keyhash->{type} eq 'pdf' )
		}
	
		#
		#
		#

		if ( defined $workbook )
		{
			$logger->info ( sprintf "Closing workbook %s...", $keyhash->{Workbook} ); 
			$workbook->close();
			$logger->info ( "ok!\n" );
		}
	}
	
	return;
}




# StyleText
# Applies Excel cell formatting (font, size, colour, bold, italic, alignment,
# borders, number format) to a range of cells in the output workbook. Acts as
# a thin wrapper around the Excel::Writer::XLSX format API to centralise
# style definitions.
sub StyleText
{
	my $keyhash = shift;
	my $BandToFind = "";
	my $colour_code;
	my $i = 0;
	
	# always take a backup!!!
	copy( $keyhash->{filename}, $keyhash->{UntouchedFilename} ) or warn "Cannot copy from [$keyhash->{filename}] to [$keyhash->{UntouchedFilename}]: $!";  # save the original
	
	if ( -e $keyhash->{UntouchedFilename} )
	{
		my $doc = Document::OOXML->read_document( $keyhash->{UntouchedFilename} );

		# Ensure all "matching" text is merged into single runs, so find/replace can actually find all the words.
		$doc->merge_runs();

		#
		$i = 0;
		#
		
		# NAWL (or blank which means ALL lists)
		if ( $highlightvocab == 1 && ( scalar @highlightvocab == 0 || ( scalar @highlightvocab > 0 && "NAWL" ~~ @highlightvocab ) ) )
		{
			$logger->info ( sprintf "		Now highlighting NAWLs...\n" );
			
			my %NAWLInDocument = %{ $keyhash->{Dictionaries}->{NAWL} };
			$colour_code = '0000FF';		# bright blue

			foreach my $outerkey ( keys %NAWLInDocument )
			{ 
				$i++;
				my $innerkey = $NAWLInDocument{$outerkey};
				next if ( $innerkey->{found_count} == 0 );
		
				my $word = $innerkey->{member};
				my $words_highlighted = $doc->style_text( qr/\b$word\b/, bold => 1, italic => 1, underline_style => 'single', color => $colour_code );	# italic => 1, qr/\Qthe\E/		Allowed underline styles are: dash, dashDotDotHeavy, dashDotHeavy, dashedHeavy, dashLong, dashLongHeavy, dotDash, dotDotDash, dotted, dottedHeavy, double, none, single, thick, wave, wavyDouble, wavyHeavy, words
				$logger->info ( sprintf "		%02i) Highlighted NAWL word [%s] which appears [%i] time(s) in text.\n", $i, $innerkey->{member}, $innerkey->{found_count} ) if ( $words_highlighted > 0 );
			}
			
			$logger->info ( sprintf "		Didn't find any NAWLs in the text.\n" ) if ( $i == 0 );
		}
		else
		{
			$logger->info ( sprintf "		Not tasked to highlight NAWLs.  Skipping.\n" );
		}
		
		
		#
		$i = 0;
		#
		
		# NAWL (or blank which means ALL lists)
		if ( $highlightvocab == 1 && ( scalar @highlightvocab == 0 || ( scalar @highlightvocab > 0 && "NGSL" ~~ @highlightvocab ) ) )
		{
			$logger->info ( sprintf "		Now highlighting NGSLs...\n" );
			
			$BandToFind = "BandA" if ( $archive_root_level == 0 );	# Foundation
			$BandToFind = "BandB" if ( $archive_root_level == 1 );	# Level 1
			$BandToFind = "BandC" if ( $archive_root_level == 2 );	# Level 2
			$BandToFind = "BandD" if ( $archive_root_level == 3 );	# Level 3
		
			if ( $BandToFind ne "" )
			{
				$logger->info ( sprintf "	Highlighting words in %s\n", $BandToFind );
				my %NGSLInDocument = %{ $keyhash->{Dictionaries}->{NGSL} };
				$colour_code = 'FF0000';		# red
	
				foreach my $outerkey ( keys %NGSLInDocument )
				{ 
					my $innerkey = $NGSLInDocument{$outerkey};
					next if ( $innerkey->{found_count} == 0 );
			
					if ( $innerkey->{level} eq $BandToFind )
					{
						$i++;
						my $word = $innerkey->{member};
						my $words_highlighted = $doc->style_text( qr/\b$word\b/, bold => 1, italic => 1, underline_style => 'single', color => $colour_code );	# italic => 1, qr/\Qthe\E/		Allowed underline styles are: dash, dashDotDotHeavy, dashDotHeavy, dashedHeavy, dashLong, dashLongHeavy, dotDash, dotDotDash, dotted, dottedHeavy, double, none, single, thick, wave, wavyDouble, wavyHeavy, words
						$logger->info ( sprintf "		%02i) Highlighted word [%s] which appears [%i] time(s) in text.\n", $i, $innerkey->{member}, $innerkey->{found_count} ) if ( $words_highlighted > 0 );
					}
				}
				
				$logger->info ( sprintf "		Didn't find any NGSLs in the text.\n" ) if ( $i == 0 );
			}
			else
			{
				$logger->info ( sprintf "		Could not determine the band of vocab to highlight.  Nothing done.\n" );
			}
		}
		else
		{
			$logger->info ( sprintf "		Not tasked to highlight NGSLs.  Skipping.\n" );
		}
		
		#
		$i = 0;
		#

		# ACL (or blank which means ALL lists)
		if ( $highlightvocab == 1 && ( scalar @highlightvocab == 0 || ( scalar @highlightvocab > 0 && "ACL" ~~ @highlightvocab ) ) )
		{
			$logger->info ( sprintf "		Now highlighting ACLs...\n" );
			
			my %AcademicCollocationsInDocument = %{ $keyhash->{Dictionaries}->{AcademicCollocations} };
			$colour_code = '7D3C98';		# a purple
		
			foreach my $outerkey ( keys %AcademicCollocationsInDocument )
			{
				my $innerkey = $AcademicCollocationsInDocument{$outerkey};
				next if ( $innerkey->{found_count} == 0 );
		
				#$logger->info ( sprintf "		this has a found count of %i\n", $innerkey->{found_count} );
			
				my @strings = @{ $innerkey->{strings} };

				foreach my $string ( @strings )
				{
					$i++;
					my $words_highlighted = $doc->style_text( qr/\b$string\b/i, bold => 1, italic => 1, underline_style => 'single', color => $colour_code );	# qr/\Qthe\E/
					$logger->info ( sprintf "		%02i) Highlighted ACL [%s] which appears [%i] time(s) in text.\n", $i, $string, $words_highlighted )  if ( $words_highlighted > 0 );
				}
			}
			
			$logger->info ( sprintf "		Didn't find any ACLs in the text.\n" ) if ( $i == 0 );
		}
		else
		{
			$logger->info ( sprintf "		Not tasked to highlight ACLs.  Skipping.\n" );
		}


		$doc->save_to_file( $keyhash->{filename} );
	}
	
	return;
}


# IncrementFileTypeCount
# Increments the global file-type counters ($filecount, $docfiles, $pdffiles,
# etc.) based on the extension of the filename passed in. Called once per
# resource after its type has been determined.
sub IncrementFileTypeCount
{
	my $filename = shift;
	
	# increment these
	$filecount++;
	$docfiles++ 	if ( $filename =~ /\.doc$/i );
	$docxfiles++ 	if ( $filename =~ /\.docx$/i );
	$pdffiles++ 	if ( $filename =~ /\.pdf$/i );
	$xmlfiles++ 	if ( $filename =~ /\.xml$/i );
	$imagefiles++ 	if ( $filename =~ /\.(?:jpg|jpe|jpeg|png|bmp|gif|tif)$/i );
	$pptfiles++ 	if ( $filename =~ /\.(?:ppt|pptx)$/i );
	$mp3files++ 	if ( $filename =~ /\.(?:mp3|m4a|wma)$/i );
	$mp4files++ 	if ( $filename =~ /\.(?:mp4|avi)$/i );
	$txtfiles++ 	if ( $filename =~ /\.txt$/i );
	$htmlfiles++ 	if ( $filename =~ /\.(?:html|htm)$/i );
	$xlsfiles++ 	if ( $filename =~ /\.(?:xls|xlsx)$/i );
	
	return;
}


# Dive
# Recursively traverses the organisation tree parsed from the imsmanifest.xml.
# For each node it determines whether the node represents a resource (leaf) or
# a folder (branch). Leaf nodes are passed to ProcessItems; branch nodes push
# their title onto the @hierarchy stack and recurse into their children.
sub Dive
{
	my ($ref) = @_;
		
	$logger->info ( "Dive! Dive! Dive!" );

	$location = join( "-->", @hierarchy );
	printf "%s\n", $location;
	my $folderitemcounter = 0;
	
	for my $item ( @$ref )
	{
		my $filename = '';
		my $fileondisk = '';
		my $directory = '';
		
		my $currentitemrref = $item->{identifierref};
		
		if ( defined $currentitemrref )	# if NOT a folder
		{
			# don't count folders
			$archiveitemcounter++;
			$folderitemcounter++;
		
			#  the filename may not be the same as the title
			#  on upload the uploader can give the file a different title to the filename
			#  to be on the safe side, browse to the directory and get the filename proper.
			
			my $directory = $destinationDirectory.$item->{identifierref};
			my $filename = $directory.'/'.$item->{title}[0];
			
			# if the the title provided in the manifest doesn't match a filename which exists, look for it
			unless ( -e $filename )
			{
				# try to find it
				if( opendir( DIR, $directory ) )
				{
					while ( my $file = readdir( DIR ) )
					{
						# Use a regular expression to ignore files beginning with a period
						next if ($file =~ m/^\./);

						$fileondisk = $file;
					}
					closedir( DIR );
				
					$filename = $directory.'/'.$fileondisk;	
				}	
			}
			
			IncrementFileTypeCount ( $filename );
			
			PrepareItems ( $directory, $filename, $item );
		}
		else	# it's a folder
		{
			my $oldunit = $parent;
			my $donotmakePDF = 0;
			$parent = $item->{title}[0];

			# beautify the unit name, make the Unit Change Page Sentence Case
			# unless it starts with an uppercase word followed by a space followed by 001
			$parent =~ s/([\w']+)/\u\L$1/g unless ( $parent =~ /^[A-Z]{2,4} 001/ );
			
			$parent = $RootDirectory if ( length( $parent ) == 0 );
			$logger->info ( sprintf "Parent is [%s]\n", $parent );
			
			# does the folder name/structure contain any words which mean we need to exclude these resources?
			foreach my $exclude ( @exclusions )
			{
				$location = join( "<--", @hierarchy );
				$logger->info ( sprintf "Checking [%s] for exclusion: [%s]\n", $location, $exclude );
			
				# if the unit title contains an exclusion word then don't bother making a PDF for it 
				$donotmakePDF++ if ( index ( $location, $exclude ) > -1 );		# STR, SUBSTR, POSITION
			}
			
			$logger->info ( sprintf "Found a new folder.  Was [%s] is now [%s].  DoNotMakePDF is [%i]\n", $oldunit, $parent, $donotmakePDF );
			
			# only add the vocab files for LA
			if ( $oldunit eq 'Course Documents' && $parent ne 'Assessments' && $archive_root_code =~ /XXXXXXX/ )	# /LA1|LA2|LA3|LA4/
			{
				foreach my $vocabfile ( @VOCAB_FILES )
				{
					if ( index ( $vocabfile, $archive_root_level ) != -1 && $addedvocabfile == 0 )
					{					
					    my $pdf = PDF::API2->open( $vocabfile );
						$addedvocabpages = $pdf->pages();					
						$logger->info ( sprintf "Found vocabulary file [%s] for this level; it has [%i] pages; copying to:\n", $vocabfile, $addedvocabpages );						
									
						my $destination_filename = sprintf( '%s%04da %s', $AllDocumentsDirectory, $archiveitemcounter, $vocabfile );
						$logger->info ( sprintf "	[%s]...\n", $destination_filename );
						copy $vocabfile, $destination_filename;
						$addedvocabfilename = $destination_filename;		# set this, we need it when we come to make the ToC
										
						foreach my $sub ( @subtypes )
						{
							$destination_filename = sprintf( '%s%04da %s', $destinationDirectory.$sub.'/', $archiveitemcounter, $vocabfile );
							$logger->info ( sprintf "	[%s]...\n", $destination_filename );						
							copy $vocabfile, $destination_filename;
						}
						
						$addedvocabfile++;
						last;	
					}
				}
			}
			elsif ( $oldunit eq 'Course Documents' && $parent ne 'Assessments' && $archive_root_code !~ /LA1|LA2|LA3|LA4/ )
			{
				$logger->info ( sprintf "Archive is [%s] not LA.  Not adding vocab file\n", $archive_root_code );	
			}
			
			#
			#
			#
			
			if ( ( $parent =~ /^Lesson\b/i || $parent ~~ @toplevels || $parent ~~ @alwaysinclude || $parent ~~ @secondlevels ) && $donotmakePDF == 0 )
			{
				if ( $parent =~ /^Lesson\b/i || $parent ~~ @alwaysinclude || $parent ~~ @secondlevels )
				{
					$logger->info ( sprintf "Starting a new Lesson.  Was [%s] is now [%s].  Level is now [%i]\n", $oldunit, $parent, $level );
					$logger->info ( sprintf "Archive has lessons.  This is the [%i] unit found\n", $hasunits );

					$currentunit = $parent;
					$hasunits++;
					$unitcount++;					
				}
				elsif ( $parent ~~ @toplevels )	# new course
				{
					$logger->info ( sprintf "Starting a new Course.  Was [%s] is now [%s].  Level is now [%i]\n", $currentcourse, $parent, $level );
					
					$currentcourse = $parent;
					$currentunit = '';	# reset this
					$unitcount++;	# we should still increment this here
				}
				
				# save this 'unit change page' to the 'All Documents' directory
				my $filename = sprintf( '%s%04db%i %s.pdf', $AllDocumentsDirectory, $archiveitemcounter, $level, $parent );
				SaveTextAsPDF( $filename, $parent, 0, 0 );
				
				# save this 'unit change page' to all subtypes directory
				foreach my $sub ( @subtypes )
				{
					my $filename = sprintf( '%s%04db%i %s.pdf', $destinationDirectory.$sub.'/', $archiveitemcounter, $level, $parent );
					SaveTextAsPDF( $filename, $parent, 0, 0 );
				}
				$foldercount++;
			}
			else
			{
				$logger->info ( sprintf "New Unit [%s] does not match criterion.  No action.\n", $parent );				
			}
		}
		
		if ( defined $item->{item} )
		{
			push @hierarchy, $parent;
			
			$level++;
			Dive( $item->{item} );
			$level--;
		}
	}
	
	$parent = $hierarchy[-1];	# the last one.
	pop @hierarchy;
	$location = join( "<--", @hierarchy );
	$logger->info ( sprintf "Left [%s].  Parent is now [%s].  Level is now [%i].\n", $location, $parent, $level );
	return;
}


# IsUniqueMD5
# Computes the MD5 hash of a file and checks it against the %seenMD5s cache.
# Returns true if this is the first time the file content has been encountered,
# false if a duplicate has already been processed. Used to skip redundant
# analysis of identical files.
sub IsUniqueMD5
{
	my $MD5 = shift;
	my $subtype = shift;
	
	my $foundcount = 0;
	
	foreach my $item ( @archivedetails )
	{
		# we'll find it once, if we find it twice, we have duplicates
		# must match the same subtype e.g. TASKS == TASKS
		$foundcount++ if ( ( defined $item->{MD5} && $item->{MD5} eq $MD5 ) && $item->{subtype} eq $subtype );
	}
	
	return $foundcount-1;
}


# LaunchListener
# Starts a background thread that listens for inter-process messages (via a
# socket or named pipe) from the external grammar_check.py processes. Used to
# collect results asynchronously as multiple grammar checks run in parallel.
sub LaunchListener
{
	my $returnvalue = -1;
	
	#$returnvalue = system ( 'unoconv', '--listener' );
	
	if ( $returnvalue == 0 ) # success!
	{
		$logger->info ( "Started unoconv listener successfully." );
	}
	else
	{
		$logger->info ( "System command failed: $!  Returning $returnvalue\n" );
	}		
	
	return $returnvalue;
}


# CopyFileToSaveDirectory
# Copies a processed output file (e.g. a generated PDF or Excel workbook) into
# the designated save/archive directory, creating the directory if necessary.
# Records the destination path in the global @SavedFiles array.
sub CopyFileToSaveDirectory
{
	my $index = shift;
	my $filename = shift;
	my $md5 = shift;
	my $ext = shift;
	
	my $destinationfilename = $savefileDirectory.$md5.$ext;
	
	mkdir $savefileDirectory unless -d $savefileDirectory;
	
	$logger->info ( sprintf "\tCopyFileToSaveDirectory: Destination file is [%s]  Does it exist? [%s]\n", $destinationfilename, ( -e $destinationfilename ) ? "Yes :-)" : "No :-(" );
	
	# if we don't already have it in our savefiles directory, copy it to our save directory for possible use later
	if ( not -e $destinationfilename )
	{
		copy( $filename, $destinationfilename ) or warn "Cannot copy from [$filename] to [$destinationfilename]: $!";
		$logger->info ( sprintf "\tCopied [%s] to [%s] for possible use again.\n", $filename, $destinationfilename );
	}
	
	return;
}



# FindFileInSaveFiles
# Searches the @SavedFiles array for a file matching a given name or pattern.
# Returns the full path of the first match found, or undef if no match
# exists.
sub FindFileInSaveFiles
{
	my $ext = shift;
	my $copycount = 0;
	
	$logger->info ( sprintf "\nAttempting to locate %s files in savefiles directory...\n", uc $ext );

	foreach my $item ( @archivedetails )
	{
		#$logger->info ( sprintf "MD5 is [%s]	Filename is [%s]	Type is [%s]\n", $SaveFilesHash{ $item->{MD5} }->{MD5}, $SaveFilesHash{ $item->{MD5} }->{filename}, $SaveFilesHash{ $item->{MD5} }->{type} );
		my $source = $savefileDirectory.$item->{MD5}.".".$ext;

		if ( -e $source )
		{
			#printf "Filename in hash is %s\n", $SaveFilesHash{ $item->{MD5} }->{filename};
			
			my $destination = "";
			
			$destination = $item->{PDFfilename} 						if ( $ext eq "pdf" );
			$destination = $item->{SpellingAndGrammarErrorsfilename} 	if ( $ext eq "csv" );
			
			copy( $source, $destination ) or warn "Cannot copy from [$source] to [$destination]: $!";
			$copycount++;
			
			$logger->info ( sprintf "Copied [%s] to [%s].\n", $source, $destination ) 			if ( -e $destination );
			$logger->info ( sprintf "Could NOT copy [%s] to [%s].\n", $source, $destination ) 	if ( not -e $destination );
		}
		else
		{
			$logger->info ( sprintf "Could not find [%s file] for [%s] in savefiles directory.  Will create!\n", $ext, $item->{MD5} ) if ( $item->{type} eq 'docx' );
		}
	}
	
	$logger->info ( sprintf "Copied [%i] files from the savefiles directory\n\n", $copycount );
	return;
}


# CheckFilenamesForSpelling
# Checks the display name of each resource against the dictionary and NGSL
# word lists to flag filenames that contain apparent spelling errors or
# non-standard vocabulary. Results are added to the inventory worksheet.
sub CheckFilenamesForSpelling
{
	my $allfiles_input = "./" . $zipbasename . " files.txt";
	my $allfiles_output = "./" . $zipbasename . " files_output.csv";

	unlink $allfiles_input;	# delete old copies (if they exist) before we do anything.
	unlink $allfiles_output;
	
	$logger->info ( sprintf "Checking filenames for spelling mistakes...\n" );
	
	open my $fh, '>', $allfiles_input;		# answers.txt

	foreach my $item ( @archivedetails )
	{
		print "\t" . $item->{FilenameNoPath} . "\n";
		print $fh $item->{FilenameNoPath} . "\n";
	}

	close $fh;
	
	#
		
	if( system ( "/opt/homebrew/bin/python3", "grammar_check.py", "-i", $allfiles_input, "-o", $allfiles_output ) == 0 )	# success!
	{
		if ( -e $allfiles_output )
		{
			# open the file we've just had created, and read contents into @SpellingAndGrammarProblems
			#$i = ReadSpellingAndGrammarFileIntoArray( $keyhash );

			# found the spreadsheet for this document in the savefiles directory
			$logger->info ( sprintf "\t\tGenerated file [%s] with [%i] errors.\n", $allfiles_output, 999 );
		}
		else
		{
			$logger->info ( sprintf "\tExpected to find [%s] but couldn't.\n", $allfiles_output );
		}
	
		#$logger->info ( sprintf "\t\tDone. Found %i possible errors.  Took %is\n", $i, tv_interval ( $Started ) ) if ( -e $keyhash->{SpellingAndGrammarErrorsfilename} );
	}
	
	return;
}

# ConvertResourcesToPDF
# Iterates over all resources that require conversion and invokes LibreOffice
# in headless mode to convert each DOCX or PPTX file to PDF. Manages
# concurrency, checks for conversion failures, and updates the resource hash
# with the path to the resulting PDF.
sub ConvertResourcesToPDF
{
	my $outputformat = shift;
	my $i = 0;
	
	$logger->info ( sprintf "Converting resources; output format is [%s]...\n", $outputformat );
		
	foreach my $item ( @archivedetails )
	{
		$i++;
		my $string = "";
		
		#printf "^^^ FilenameNoPath is [%s]\n", $item->{FilenameNoPath};
		
		if ( ( ( $item->{type} eq "doc" || $item->{type} eq "docx" || $item->{type} eq "pptx" ) && $item->{FilenameNoPath} !~ /^~/ ) && ( -e $item->{filename} && not -e $item->{PDFfilename} ) && ( $item->{subtype} ne 'EXCLUDED' ) )		# && $item->{subtype} ne 'ANSWERS' ) # 		# don't convert presentations: $item->{type} eq "ppt" || $item->{type} eq "pptx" || 
		{
			$logger->info ( sprintf "   %03i: Converting document [%s] to PDF...\n", $i, $item->{filename} );
			
			# scalar @convert == 0 means that the user has specified subtypes to convert
			if ( scalar @convert == 0 || $item->{subtype} eq 'INCLUDED' || ( scalar @convert > 0 && $item->{subtype} ~~ /@convert/ ) )
			{
				my $Started = [gettimeofday];		# start the clock!
				$logger->info ( sprintf "    %03i: ConvertResourcesToPDF [%s]...", $i, $item->{filename} );

				my ( $filename, $directories, $suffix ) = fileparse( $item->{filename}, qr/\.[^.]*/ );
				if ( system ( '/Applications/LibreOffice.app/Contents/MacOS/soffice', '--headless', '--convert-to', $outputformat, '--outdir', $directories, $item->{filename} ) == 0 ) # success!
				#if ( system ( 'unoconv', '--preserve', '-f', $outputformat, $item->{filename} ) == 0 ) # success!   
				{	
					if ( -e $item->{PDFfilename} )
					{						
						$Started = [gettimeofday];		# re-start the clock!
						$logger->info ( sprintf "      %03i: Setting file properties for PDF %s...", $i, $item->{PDFfilename} );
						
						# The PDF object
					    my $pdf = PDF::API2->open( $item->{PDFfilename} ) or warn "Failed to open PDF object: $!\n";

						# Set some of the properties of the document
						$pdf->info( 'Author' => $authorinfo, 'Title' => $item->{PDFfilename} );
	
						# Update the PDF
						$pdf->update();						

						# do this after we've updated the properties
						CopyFileToSaveDirectory( $i, $item->{PDFfilename}, $item->{MD5}, ".pdf" );
					}
					else
					{
						$logger->info ( sprintf "      %03i: Expected to find [%s].  Couldn't :-(", $i, $item->{PDFfilename} );
					}
					#$logger->info ( sprintf " ok! (%i seconds)\n", tv_interval ( $Started ) );
				}
				else
				{
					$logger->info ( "System soffice command failed: $!\n" );
				}
				
				# do this every Nth time in case this main thread craps out and all that hard word goes to shit
				if ( $i % 50 == 0 )
				{
					MakeDirectoryAndCopyFiles();

					FarmResourcesOut();
				}
			}
			else
			{
				$logger->info ( sprintf "   %03i: Missing filename tags.  Skipping [%s]; subtype is [%s]\n", $i, $item->{filename}, $item->{subtype} );
			}				
		}
		elsif ( ( $item->{type} eq 'jpg' || $item->{type} eq 'gif' || $item->{type} eq 'jpeg' ) && $noimages == 0 )
		{
			$logger->info ( sprintf "   %03i: Converting image [%s] to PDF...\n", $i, $item->{filename} );
			
			my $Started = [gettimeofday];		# start the clock!
			$logger->info ( sprintf "   %03i: %s...", $i, $item->{filename} );
			
			# this currently only works with .jpg, .gif (and .jpeg?)
			if ( system ( "/usr/bin/perl", "/usr/local/bin/ConvertImageToJPG.pl", $item->{filename}, ">/dev/null" ) == 0 )	# success!
			{		
				my ( $filename, $directories, $suffix ) = fileparse( $item->{filename}, qr/\.[^.]*/ );	
				my $imagefile = $directories.$filename.".jpg";
				
				# get the image size, and print it out
				( $item->{Image}->{OriginalX}, $item->{Image}->{OriginalY} ) = imgsize( $imagefile );

				# set some of the properties of the document
				my $pdf  = PDF::Create->new( 'filename' => $item->{PDFfilename}, 'Author' => $authorinfo, 'Title'=> $imagefile ) or warn "Failed to create new PDF object: $!\n";
				my $root = $pdf->new_page( 'MediaBox' => $pdf->get_page_size('A4') );
				my $page = $root->new_page;
				my $font = $pdf->font('BaseFont' => 'Helvetica');
					
				# put the filename at the top of the PDF page
				$page->stringc( $font, 20, 595/2, 762, $filename );	# font, size, x, y, text, centre the text

				# add the jpg
				my $imageobj = $pdf->image( $imagefile );

				$item->{Image}->{NewX} = $item->{Image}->{OriginalX};
				$item->{Image}->{NewY} = $item->{Image}->{OriginalY};

				# it's too big and needs scaling down
				# it's wider than comfortable or it's higher than comfortable
				if ( $item->{Image}->{OriginalY} > 642 || $item->{Image}->{OriginalX} > 495 )
				{
					while ( $item->{Image}->{NewY} >= 642 || $item->{Image}->{NewX} >= 495 )		# this is the comfortable border we want
					{
						$item->{Image}->{ScaleFactor} = $item->{Image}->{ScaleFactor} - 0.0001;
						$item->{Image}->{NewX} = $item->{Image}->{OriginalX} * $item->{Image}->{ScaleFactor};
						$item->{Image}->{NewY} = $item->{Image}->{OriginalY} * $item->{Image}->{ScaleFactor};
					}
				}
				# it's too small and needs scaling up
				# it's wider than it is long.
				elsif ( $item->{Image}->{OriginalY} < 642 || $item->{Image}->{OriginalX} < 495 )
				{
					while ( $item->{Image}->{NewY} <= 642 && $item->{Image}->{NewX} <= 495 )		# this is the comfortable border we want
					{
						$item->{Image}->{ScaleFactor} = $item->{Image}->{ScaleFactor} + 0.0001;
						$item->{Image}->{NewX} = $item->{Image}->{OriginalX} * $item->{Image}->{ScaleFactor};
						$item->{Image}->{NewY} = $item->{Image}->{OriginalY} * $item->{Image}->{ScaleFactor};
					}
				}
				else
				{
					$logger->info ( sprintf "\n		Original Image width was %i and height is %i", $item->{Image}->{OriginalX}, $item->{Image}->{OriginalY} );	
				}

				$page->image (
					'image'  => $imageobj,
					'xalign' => 1,				# Alignment of image; 0 is left/bottom, 1 is centered and 2 is right, top
					'yalign' => 1,
					'xscale' => $item->{Image}->{ScaleFactor},	# Scaling of image. 1.0 is original size
					'yscale' => $item->{Image}->{ScaleFactor},
					'xpos'   => 595/2,			# Position of image (required)
					'ypos'   => 842/2 );

				$pdf->close;
						
				#$logger->info ( sprintf " ok! (%i seconds)\n", tv_interval ( $Started ) );
				$logger->info ( sprintf "		Scale ratio is now %.2f\n", $item->{Image}->{ScaleFactor} );
				$logger->info ( sprintf "		Original Image width was %i and height is %i\n", $item->{Image}->{OriginalX}, $item->{Image}->{OriginalY} );
				$logger->info ( sprintf "		New Image width is now %i and height is %i\n", $item->{Image}->{NewX}, $item->{Image}->{NewY} );
							
				my $string .= sprintf "[OLD X&Y:%i & %i]", $item->{Image}->{OriginalX}, $item->{Image}->{OriginalY} 	if ( $item->{Image}->{OriginalX} != 0 );
				$string .= sprintf "[NEW X&Y:%i & %i]", $item->{Image}->{NewX}, $item->{Image}->{NewY}					if ( $item->{Image}->{NewX} != 0 );
				$string .= sprintf "[SCALE:%.2f]", 	$item->{Image}->{ScaleFactor} 										if ( $item->{Image}->{ScaleFactor} != 1 );		
				$item->{CustomFooterText} = $string;
			}
			else
			{
				unlink $item->{PDFfilename};	# delete this if it exists, it's probably corrupt
				$logger->info ( sprintf " failed! (%i seconds)\n", tv_interval ( $Started ) );
			}
			# SaveImageToPDF ( $item->{filename}, $item->{type}, $item->{PDFfilename} );	# this is now in its own perl file.  it sometimes craps out and takes down the whole script
		}
		elsif ( $item->{subtype} eq "EXCLUDED" )
		{
			$logger->info ( sprintf "   %03i: Not converting because [%s] is excluded.\n", $i, $item->{filename} );
		}
		elsif ( $item->{subtype} eq "ANSWERS" )
		{
			$logger->info ( sprintf "   %03i: Not converting because [%s] is an answer key.\n", $i, $item->{filename} );
		}
		elsif ( $item->{type} eq "pdf" && -e $item->{PDFfilename} )
		{
			$logger->info ( sprintf "   %03i: Document [%s] is already a PDF...\n", $i, $item->{filename} );
			
			$logger->info ( sprintf "      %03i: Adding custom footer to native PDF %s...", $i, $item->{PDFfilename} );
			
			my $string .= sprintf "[NMODIN %i D]", $item->{MetaData}->{AgeInDaysSinceLastEdit}		if ( $item->{MetaData}->{AgeInDaysSinceLastEdit} > 1 && $item->{MetaData}->{AgeInDaysSinceLastEdit} < 18390 );
			$string .= sprintf "[MD5 %s SEEN BEFORE]", $item->{MD5} 								if ( IsUniqueMD5( $item->{MD5}, $item->{subtype} ) == 1 );
			$item->{CustomFooterText} = $string;
		}
		elsif ( -e $item->{filename} && -e $item->{PDFfilename} )
		{
			$logger->info ( sprintf "   %03i: Found .docx AND .pdf.  Skipping conversion of [%s]\n", $i, $item->{filename} );
		}
		else
		{
			$logger->info ( sprintf "   %03i: Not an appropriate format.  Skipping [%s]\n", $i, $item->{filename} );
		}
		
		
		# prepare the $item->{CustomFooterText} string
		# we need to do this if (1) it's a fresh PDF or (2) one from the savefiles directory
		if ( $item->{subtype} eq 'TEXT' )
		{
			# Add some text to the page
			$string .= sprintf "[WORDS:%i]", $item->{WordMetaData}->{words} 					if ( $item->{WordMetaData}->{words} != 0 );
			$string .= sprintf "[DESC:%s]", $item->{Readability}->{wordsdescription} 			if ( $item->{Readability}->{wordsdescription} ne '' );
			
			$string .= sprintf "[GF:%.02f]", $item->{Readability}->{GunningFog} 				if ( $item->{Readability}->{GunningFog} ne '' && $item->{Readability}->{GunningFog} > 0 );
			$string .= sprintf "[GF DESC:%s]", $item->{Readability}->{GunningFogDescription} 	if ( $item->{Readability}->{GunningFogDescription} ne '' );
			
			$string .= sprintf "[FL:%.02f]", $item->{Readability}->{Flesch} 					if ( $item->{Readability}->{Flesch} ne '' && $item->{Readability}->{Flesch} > 0 );
			$string .= sprintf "[FL DESC:%s]", $item->{Readability}->{FleschDescription} 		if ( $item->{Readability}->{FleschDescription} ne '' );
			
			$string .= sprintf "[FK:%.02f]", $item->{Readability}->{FleschKincaid} 				if ( $item->{Readability}->{FleschKincaid} ne '' && $item->{Readability}->{FleschKincaid} > 0 );												
			$string .= sprintf "[FK DESC:%s]", $item->{Readability}->{FleschKincaidDescription} if ( $item->{Readability}->{FleschKincaidDescription} ne '' );
		}
		
		# Irrespective of the subtype, add these
		#$string .= sprintf "[CB:%s]", $item->{MetaData}->{createdby} 							if ( defined $item->{MetaData}->{createdby} && $item->{MetaData}->{createdby} ne '' );	
		#$string .= sprintf "[CD:%s]", substr( $item->{MetaData}->{createddate}, 0, 10 ) 		if ( defined $item->{MetaData}->{createddate} && $item->{MetaData}->{createddate} ne '' );	
		#$string .= sprintf "[MB:%s]", $item->{MetaData}->{lastmodifiedby} 						if ( defined $item->{MetaData}->{lastmodifiedby} && $item->{MetaData}->{lastmodifiedby} ne '' );
		$string .= sprintf "[DOCTYPE:%s]", $item->{type} 										if ( $item->{type} eq 'doc' );
		#$string .= sprintf "[MD:%s]", substr( $item->{MetaData}->{lastmodifieddate}, 0, 10 ) 	if ( defined $item->{MetaData}->{lastmodifieddate} && $item->{MetaData}->{lastmodifieddate} ne '' );
		$string .= sprintf "[NMODIN %i D]", $item->{MetaData}->{AgeInDaysSinceLastEdit}			if ( $item->{MetaData}->{AgeInDaysSinceLastEdit} > 1 && $item->{MetaData}->{AgeInDaysSinceLastEdit} < 18390 );	# for some strange reason PDFs we've created have an age of 18390 days.  exclude these.
		$string .= sprintf "[PGSZ:%s]", $item->{WordMetaData}->{PageSize} 						if ( $item->{WordMetaData}->{PageSize} ne 'A4' );		
		$string .= sprintf "[PGORIENT:%s]", $item->{WordMetaData}->{PageOrientation} 			if ( $item->{WordMetaData}->{PageOrientation} ne 'portrait' );
		$string .= sprintf "[MD5 %s SEEN BEFORE]", $item->{MD5} 								if ( IsUniqueMD5( $item->{MD5}, $item->{subtype} ) == 1 );
		#$string .= sprintf "[%i CRLFs @ EOF]", $item->{ErrantCRLFs}								if ( $item->{ErrantCRLFs} > 2 );
		
		$item->{CustomFooterText} = $string;
		
		#
		#
		#
		
		# we get here through a number of routes:
		# 1) we've now got a PDF of the .DOCX we've converted and we've already added a custom footer to it
		# 2) we've now got a PDF of the image we've convered
		# 3) it's a PDF file from our SaveFiles resources
		# 4) it's a PDF resource to start with and we've done nothing with it
		if ( -e $item->{PDFfilename} && $item->{subtype} ne "EXCLUDED" && $item->{subtype} ne "ANSWERS")
		{
			my $Started = [gettimeofday];		# start the clock!
			$logger->info ( sprintf "      %03i: Adding custom and core header and footer information to PDF %s...\n", $i, $item->{PDFfilename} );
		
			#printf "PDF paper size for page #%i is A4 portrait\n";
						
			# The PDF object
		    my $pdf = PDF::API2->open( $item->{PDFfilename} ) or warn "Failed to open PDF object: $!\n";

			# Set some of the properties of the document
			$pdf->info( 'Author' => $authorinfo,'Title' => $item->{PDFfilename} );

			# Add a built-in font to the PDF
			my $stnd_font = $pdf->corefont('Helvetica');				# 'Calibri' is not available		https://metacpan.org/pod/PDF::API2::Resource::Font::CoreFont#STANDARD-FONTS
			my $bold_font = $pdf->corefont('Helvetica-Bold');

			foreach my $page_number ( 1 .. $pdf->pages() )
			{
				my $page = $pdf->openpage( $page_number );
				my ($llx, $lly, $x, $y) = $page->get_mediabox();	# lower left X; lower left Y; upper left X; upper left Y
				my $string = '';
				
				# round the numbers
				$x = sprintf "%d", $x;
				$y = sprintf "%d", $y;
				
				# http://www.printernational.org/iso-paper-sizes.php & https://www.prepressure.com/library/paper-size
				if ( ( $x == 594 || $x == 595 ) && ( $y == 841 || $y == 842 ) )	# A4 portrait
				{
					#printf "PDF paper size for page #%i is A4 portrait\n";
					
					#
					#  H E A D E R 
					#
					
					if ( $page_number == 1 )
					{
						# add the lesson directory name to the top of the document	
						my $header_text = $page->text();
						$header_text->font( $bold_font, 24 );
						#$header_text->fillcolor( '#000000' );	# set the text color to black
						$header_text->fillcolor( '#5B5BA0' );	# set the text color to TeachComputing Purple RGB: 91, 91, 160
				
						$header_text->translate( $x - 559, $y -70 );	# horizontal, vertical.  A4 = 595, 842
						$header_text->text( $item->{LastDirectory} );						
					}
										
					#
					#  F O O T E R  
					#

					# add a white rectangle to the bottom of each page if we haven't already done so
					if ( $item->{AddedWhiteRectangle} < $pdf->pages() )
					{
						# $logger->info ( sprintf "      %03i: Adding white rectangle to page %02i of %s...", $i, $page_number, $item->{PDFfilename} );
						
						my $gfx = $page->gfx();
						my $rect = $gfx->rectxy( 0, 0, 595, 60 );		# x, y, tox, toy.	65 is too high
						$rect->fillcolor( '#ffffff' );		# white
						$rect->fill();
						$item->{AddedWhiteRectangle}++;		# be careful, this will only add the blank rectangle to the first page
					}
					else
					{
						# $logger->info ( "      %03i: Not adding white rectangle.  Already added." );	
					}
					
					# add some custom text to the bottom of the document	
					my $footer_text = $page->text();
					$footer_text->font( $stnd_font, 10 );
					$footer_text->fillcolor( '#000000' );	# set the text color to black
				
					$footer_text->translate( 25, 25 );	# horizontal, vertical.  A4 = 595, 842
					$footer_text->text( $item->{CustomFooterText} );
										
					$string .= sprintf "[%s]", $item->{location} 									if ( $item->{location} ne '' );	# in single-file mode, this is empty
					$string .= sprintf "[SIZE:%s]", 'BI'.'I' x ( $item->{size}/500000 ).'G FILE' 	if ( $item->{size} > 2097152 );
					$string =~ s/$RootDirectory//i;
					$string =~ s/\/For Instructors\///i;
					$footer_text->translate( 25, 10 );	# horizontal, vertical.  A4 = 595, 842
					$footer_text->text( $string );
					
					#$logger->info ( sprintf "      %03i: Core Footer Text is %s", $i, $string );
					#$logger->info ( sprintf "      %03i: Custom Footer Text is %s", $i, $item->{CustomFooterText} );
				}
				elsif ( ( $x == 841 || $x == 842 ) && ( $y == 594 || $y == 595 ) )	# A4 landscape
				{
					#printf "PDF paper size for page #%i is A4 landscape\n", $page_number;					
				}
				elsif ( ( $x == 841 || $x == 842 ) && ( $y == 1190 || $y == 1191 ) )	# A3 portrait
				{
					#printf "PDF paper size for page #%i is A3 portrait\n", $page_number;					
				}
				elsif ( ( $x == 1190 || $x == 1191 ) && ( $y == 841 || $y == 842 ) )	# A3 landscape
				{
					#printf "PDF paper size for page #%i is A3 landscape\n", $page_number;					
				}
				elsif ( $x == 612 && $y == 792 )	# US Letter portrait
				{
					#printf "PDF paper size for page #%i is US Letter portrait\n", $page_number;	
				}
				elsif ( $x == 792 && $y == 612 )	# US Letter landscape
				{
					#printf "PDF paper size for page #%i is US Letter landscape\n", $page_number;
				}
				elsif ( $x == 960 && $y == 540 )
				{
					#printf "PDF paper size for page #%i is Widescreen PowerPoint Slide (16:9 aspect ratio)\n", $page_number;
				}
				elsif ( $x == 720 && $y == 540 )
				{
					#printf "PDF paper size for page #%i is Standard PowerPoint Slide (4:3 aspect ratio)\n", $page_number;
				}
				elsif ( $x == 720 && $y == 405 )
				{
					#printf "PDF paper size for page #%i is Widescreen PowerPoint Slide (16:9 aspect ratio)\n", $page_number;
				}
				elsif ( $x == 432 && $y == 648 )
				{
					#printf "PDF paper size for page #%i is US 1 Catalog envelope\n", $page_number;	# https://github.com/Bubblbu/paper-sizes
				}
				else
				{
					$logger->info ( sprintf "        PDF paper size for page #%i is unknown: x is %s; y is %s", $page_number, $x, $y );
					
					# add some custom text to the bottom of the document
					$string .= sprintf "        PDF paper size for page #%i is unknown: x is %s; y is %s\n", $page_number, $x, $y;
					
					my $text = $page->text();
					$text->font( $stnd_font, 10 );
					$text->fillcolor( '#000000' );	# set the text color to black
					$text->translate( 25, 25 );	# horizontal, vertical.  A4 = 595, 842
					$text->text( $string );					
				}
			}
			
			# Update the PDF
			$pdf->update();
			#$logger->info ( sprintf " ok! (%i seconds)\n", tv_interval ( $Started ) );
		}
	}
	
	return;	
}


# ApplyTemplateToResources
# Applies a DOCX template (cover page, header/footer styles) to each Word
# document resource before conversion, so that the merged PDF output has a
# consistent visual style throughout.
sub ApplyTemplateToResources
{
	my $outputformat = shift;
	my $template = $HomeWorkingDirectory.$templatefile;
	my $i = 0;

	if ( -e $template )
	{
		$logger->info ( sprintf "Applying template to resources; template is [%s]; output format is [%s]...\n", $template, $outputformat );
		
		foreach my $item ( @archivedetails )
		{
			$i++;

			if ( ( $item->{type} eq "doc" || $item->{type} eq "docx" ) && -f $item->{filename} && ( $item->{subtype} ne 'EXCLUDED' && $item->{subtype} ne 'ANSWERS' ) ) 		# don't convert presentations: $item->{type} eq "ppt" || $item->{type} eq "pptx" || 
			{
				if ( scalar @addtemplate == 0 || ( scalar @addtemplate > 0 && $item->{filename} ~~ /@addtemplate/ ) )
				{
					my $Started = [gettimeofday];		# start the timer!
					$logger->info ( sprintf "   %03i: [%s]... ", $i, $item->{filename} );
			
					# always take a backup!!!
					move( $item->{filename}, $item->{UntouchedFilename} ) or warn "Cannot copy from [$item->{filename}] to [$item->{UntouchedFilename}]: $!";  # if $outputformat eq 'pdf';		# save the original
		
					# https://github.com/dagwieers/unoconv/issues/430#issuecomment-464684324
					# This works for me:
					# unoconv -t templatefile.ott -o output.pdf input.docx
					# The input filename is always last. Yeah, that can be confusing.
					
					#if ( system ( 'unoconv', '-t', $template, '--preserve', '-o', $item->{filename}, '-f', $outputformat, $item->{UntouchedFilename} ) == 0 ) # success!   
					
					my ( $filename, $directories, $suffix ) = fileparse( $item->{filename}, qr/\.[^.]*/ );
					
					# https://help.libreoffice.org/6.2/he/text/shared/guide/start_parameters.html
					if ( system ( '/Applications/LibreOffice.app/Contents/MacOS/soffice', '--headless', '--convert-to', $outputformat, '--outdir', $directories, $item->{UntouchedFilename} ) == 0 ) # success!
					{	
						$logger->info ( sprintf "ok! (%is)\n", tv_interval ( $Started ) ) if ( -e $item->{filename} );
					}
					else
					{
						$logger->info ( "System soffice command failed: $!\n" );
					}
					
					# do this every Nth time in case it craps out and all that hard word goes to shit
					if ( $i % 20 == 0 )
					{
						MakeDirectoryAndCopyFiles();

						FarmResourcesOut();
					}					
				}
				else
				{
					$logger->info ( sprintf "   %03i: Missing filename tags.  Skipping [%s]\n", $i, $item->{filename} );
				}
			}
			elsif ( $item->{subtype} eq "EXCLUDED" || $item->{subtype} eq "ANSWERS" )
			{
				$logger->info ( sprintf "   %03i: Not converting because file [%s] is marked for exclusion or is an answer key\n", $i, $item->{filename} );			
			}
			else
			{
				$logger->info ( sprintf "   %03i: Not an appropriate format.  Skipping [%s]\n", $i, $item->{filename} );				
			}
		}	
	}
	else
	{
		$logger->info ( sprintf "Could not find template document ($template) in %s, skipping!\n", $HomeWorkingDirectory );
	}
	
	return;
}


# CountPages
# Returns the number of pages in a PDF file by opening it with PDF::API2 and
# querying the page count. Used to populate page-count fields in the inventory
# and to calculate offsets when building the table of contents.
sub CountPages
{
	my $subtype = shift;
	my $directory = shift;
	my $thisunit = shift;
	
	my $page_count = 0;
	
	foreach my $resource ( @archivedetails )
	{
		if ( $resource->{unit} eq $thisunit )
		{
			my $inputfile = $directory.$resource->{IdAndFilenameNoPathAsPDF};
		
			if ( -e $inputfile && ( $resource->{subtype} eq $subtype || $subtype eq $AllDocumentsDirectory || $resource->{subtype} eq 'INCLUDED' ) )
			{
				my $pdf = PDF::API2->open( $inputfile );
				$page_count += $pdf->pages();				
			}
		}
		
	}	

	$logger->info ( sprintf "Unit [%s] has %i pages.\n", $thisunit, $page_count );
	return $page_count;
}

# CreateTableOfContents
# Builds a PDF table-of-contents document listing all resources with their
# unit, title, and page range. The TOC is inserted at the front of the merged
# PDF and hyperlinked to each section.
sub CreateTableOfContents
{
	my $subtype = shift;
	my $directory = shift;
	my $pattern = shift;
	
	my $toc = $directory.'0000a Table of Contents.txt';
	my $tocPDF = $directory.'0000a Table of Contents.pdf';
	my $template = 'TableOfContents.ott';
	my $cumulative_page_count = 1;
	my $string = "";
	my $returnvalue = 0;
	
	my $oldcourse = '';
	my $thiscourse = '';
	
	my $oldunit = '';
	my $thisunit = '';
	
	my $i = 0;
	my $hasspecial = 0;
	
	# this might exist, delete it; don't delete the PDF version of this.
	unlink $toc;
	open my $fh, '>', $toc;

	$logger->info ( sprintf "\n--> %s -->\nCreating Table of Contents file [%s] in [%s]; pattern is [%s]...\n", $subtype, $toc, $directory, $pattern );

	foreach my $resource ( @archivedetails )
	{
		my $page_count = 0;
		my $indent = '';
		my @focusarray;
		$string = "";		
		
		$thiscourse = $resource->{course};
		$thisunit = $resource->{unit};
						
		if ( $thiscourse ne $oldcourse )
		{
			$logger->info ( sprintf "Old course is [%s].  This course is [%s].  They are different.\n", $oldcourse, $thiscourse );
			$oldcourse = $thiscourse;
			$oldunit = '';	# reset this
			
			$page_count = 1;	# we know this PDF to be one page long
			$string = sprintf "\n%-80s %02i\n", $thiscourse, $cumulative_page_count;
		
			print $fh $string;
			$cumulative_page_count += $page_count;
		}
		
		#
		#
		#
		
		$string = "";	# reset this
		
		if ( $thisunit ne $oldunit )
		{
			# we've just changed units.  if the old unit is 'Course Documents' and we have a vocab file, add this before the Unit change title.
			if ( $oldunit eq 'Course Documents' && $addedvocabfile > 0 )
			{
				$string = sprintf "     %-80s %02i-%02i\n", 'Important Vocabulary At Your Level', $cumulative_page_count, $cumulative_page_count+$addedvocabpages-1;
				$logger->info ( "Added vocab file to ToC.\n" );

				print $fh $string;
				$cumulative_page_count += $addedvocabpages;
			}
			
			$page_count = 1;	# we know this PDF to be one page long
			my $pages_in_unit = CountPages( $subtype, $directory, $thisunit );
			$string = sprintf "%-54s (%02i pages) %02i-%02i\n", $thisunit, $pages_in_unit, $cumulative_page_count, $cumulative_page_count + $pages_in_unit;
			
			print $fh $string;
			$cumulative_page_count += $page_count;
					
			$logger->info ( sprintf "Old unit is [%s].  This unit is [%s].  They are different.\n", $oldunit, $thisunit );
			$oldunit = $thisunit;	
		}

		#
		#
		#
		
		$string = "";	# reset this
		
		my $inputfile = $directory.$resource->{IdAndFilenameNoPathAsPDF};
		my ( $title, $directories, $suffix ) = fileparse( $inputfile, qr/\.[^.]*/ );	
		
		my $oldtitle = $title;				# lets store the original just in case we beautify it to nothing	
		$title =~ s/^\S+\s*//;				# remove non-spaces from start of filename (removes the first word)
		$title =~ s/^$archive_root//i;		# the $archive_root_code is 'Language Awareness 1' or 'Writing 3' or 'Speaking 3'
		$title =~ s/^$archive_root_code\s//i;	# remove any course code (e.g. R3, W1, LA1) from start of filename and any subsequent space
		$title =~ s/^\d{3}//;				# remove any numbers from start of filename.
		$title =~ s/Worksheet //i;			# remove "Worksheet " from the name
		$title =~ s/Activity Sheet //i;		# remove "Activity Sheet " from the name
		#$title =~ s/ – / - /;				# replace  – with - 
		#$title =~ s/\s+/\s/;				# many spaces with one 
		$title =~ s/([\w']+)/\u\L$1/g;		# make the filename Sentence Case
				
		if ( -e $inputfile && ( $resource->{subtype} eq $subtype || $subtype eq $AllDocumentsDirectory || $resource->{subtype} eq 'INCLUDED' ) )
		{
			my $pdf = PDF::API2->open( $inputfile );
			$page_count = $pdf->pages();
			
			# do we have any lesson foci or skills foci?  if so, we'll need this array
			# THIS STOPPED WORKING FOR SOME REASON (AFTER UPGRADE TO BIG SUR).  COMMENTED OUT FOR NOW.
			#if ( defined { $resource->{MetaData}->{LessonFocus} } )
			#{
			#	@focusarray = @{ $resource->{MetaData}->{LessonFocus} };		# dereference this
			#}
			
			# we set the subtype to INCLUDED if the files are in one of our special folders
			# e.g. 'Final Exam Preview'
			# in this case, it's important that we know the subtype, i.e. tasks vs text vs answers
			# don't remove these from the name, instead report them in the ToC 
			# R3 T2 Final Ak	<-- this will not get removed because it is not 'ANSWERS'
			# R3 T2 Final Tasks
			# R3 T2 Final Text
			unless ( $resource->{subtype} eq 'INCLUDED' || $subtype eq $AllDocumentsDirectory )
			{
				# remove subtypes from the end of filename
				foreach my $subtype ( @subtypes )
				{
					# if the filename ends with subtype e.g. "TEXT", remove it
					# $title =~ s/\b$subtype$//gi;
				}				
			}
			
			$title =~ s/^\s+|\s+$//g;	# trim leading and trailing spaces
			$title = $oldtitle if ( $title eq "" );		# if we've stripped away everything there is, reset it
				
			$logger->info ( sprintf "	PDF [%s] exists and it has [%i] pages\n", $title, $page_count );
			$logger->info ( sprintf "		Focus is [%s]\n", join( "\n\t\t\t", @focusarray ) ) if ( scalar( @focusarray ) > 0 );
			
			if ( index ( $resource->{location}, "**" ) != -1 ||index ( $resource->{location}, "^^" ) != -1 )
			{
				$logger->info ( sprintf "		Folder [%s] is marked as special\n", $resource->{location} );
				$title = "** " . $title if ( index ( $resource->{location}, "**" ) != -1 );
				$title = "^^ " . $title if ( index ( $resource->{location}, "^^" ) != -1 );
				$hasspecial++;
			}
			
			my @dirs = split /\//, $resource->{location};
			my $last_directory_plus_title = sprintf( '%s: %s', $dirs[-1], $title );
			
			if ( $page_count > 1 )
			{
				if ( ( $hasunits > 1 && $filename =~ /^Unit\b/ || $filename =~ /^Lesson\b/ ) || $hasunits == 0 )
				{
					$string .= sprintf "%-72s %02i-%02i\n", $last_directory_plus_title, $cumulative_page_count, $cumulative_page_count+$page_count-1;
					foreach my $foci ( @focusarray )
					{
						$string .= sprintf "     %-72s\n", $foci if ( scalar( @focusarray ) > 0 );
					}
				}
				else
				{
					$string .= sprintf "     %-67s %02i-%02i\n", $last_directory_plus_title, $cumulative_page_count, $cumulative_page_count+$page_count-1;					
					foreach my $foci ( @focusarray )
					{
						$string .= sprintf "	      %-67s\n", $foci if ( scalar( @focusarray ) > 0 );
					}			
				}
			}
			else
			{
				if ( ( $hasunits > 1 && $filename =~ /^Unit\b/ || $filename =~ /^Lesson\b/ ) || $hasunits == 0 )
				{
					$string .= sprintf "%-72s %02i\n", $last_directory_plus_title, $cumulative_page_count;
					foreach my $foci ( @focusarray )
					{
						$string .= sprintf "     %-72s\n", $foci if ( scalar( @focusarray ) > 0 );
					}
				}
				else
				{
					$string .= sprintf "     %-67s %02i\n", $last_directory_plus_title, $cumulative_page_count;					
					foreach my $foci ( @focusarray )
					{
						$string .= sprintf "	      %-67s\n", $foci if ( scalar( @focusarray ) > 0 );
					}			
				}
			}
			
			print $fh $string;
			$cumulative_page_count += $page_count;
		}

		#printf "string is [%s]\n", $string if ( $string ne "" );
		$i++;
	}
	
	if ( -e $BLANK_WRITING_PAGE )
	{
		my $pdf = PDF::API2->open( $BLANK_WRITING_PAGE );
		my $page_count = $pdf->pages();
		$string = sprintf "%-67s %02i-%02i\n", 'Blank Writing Paper', $cumulative_page_count, $cumulative_page_count+$page_count-1;
		print $fh $string;
	}
	
	if ( $hasspecial > 0 )
	{
		$string  = "\n\nNOTE:  Lessons marked with ** will be done in class during odd-numbered terms (e.g. Term 26, 28, 30, 32).\n";
		$string .= "\n\nNOTE:  Lessons marked with ^^ will be done in class during even-numbered terms (e.g. Term 27, 29, 31, 33).\n";
		print $fh $string;		
	}
	
	if ( $highlightvocab == 1 )
	{
		$string =  "\n\nVocabulary which has been highlighted in red are key words for you to learn at this level.  Vocabulary which has been highlighted in blue are key academic words for you to learn.  Academic phrases which have been highlighted in purple are important academic collocations for you to learn.\n";
		print $fh $string;	
	}
	
	# Close the Table Of Contents ($toc) file
	close $fh;
	
	#
	#  Apply template to TOC and save as PDF
	#


	if ( $hasTemplateConverter == 1 || $hasPDFConverter == 1 )
	{
		if ( $isUnoconvWorking == 1 )
		{
			# use unoconv
			printf "Applying template to Table of Contents file using Unoconv [%s]... ", $toc;
			
			if ( system ( 'unoconv', '-t', $template, '--preserve', '-o', $tocPDF, '-f', "pdf", $toc ) == 0 ) # success!   
			{	
				if ( -e $tocPDF )
				{
					my $pdf = PDF::API2->open( $tocPDF );
					$returnvalue = $pdf->pages();
					$logger->info ( sprintf "ok! (ToC has %i page(s))\n", $returnvalue );
				}
				else
				{
					$logger->info ( "Output file does not exist\n" );
				}
			}
			else
			{
				$logger->info ( "System unoconv command failed: $!\n" );
			}
		}
		else
		{
			# use soffice
			printf "Applying template to Table of Contents file using soffice [%s]... ", $toc;
			
			my ( $filename, $directories, $suffix ) = fileparse( $toc, qr/\.[^.]*/ );
			if ( system ( '/Applications/LibreOffice.app/Contents/MacOS/soffice', '--headless', '--convert-to', 'pdf', '--outdir', $directories, $toc ) == 0 ) # success!
			{	
				if ( -e $tocPDF )
				{
					my $pdf = PDF::API2->open( $tocPDF );
					$returnvalue = $pdf->pages();
					$logger->info ( sprintf "ok! (ToC has %i page(s))\n", $returnvalue );
				}
				else
				{
					$logger->info ( "Output file does not exist\n" );
				}
			}
			else
			{
				$logger->info ( "System soffice command failed: $!\n" );
			}
		}
	}

	$logger->info ( sprintf "<-- %s <--\n", $subtype );
	return $returnvalue;
}


# DoesArchiveHave
# Checks whether the archive contains at least one resource of the specified
# subtype (e.g. 'docx', 'pdf', 'pptx'). Returns a boolean used to decide
# whether type-specific processing steps should be executed.
sub DoesArchiveHave
{
	my $sub = shift;
	
	foreach my $resource ( @archivedetails )
	{
		return 1 if ( $resource->{subtype} eq $sub )
	}
	
	return 0;
}





# MergePDFs
# Concatenates a list of PDF files into a single output PDF using PDF::API2.
# Preserves bookmarks where possible, inserts unit-change separator pages
# between sections, and writes the merged file to the destination directory.
sub MergePDFs
{
	my $thissubtype = shift;
	my $directory = shift;
	my $outputfile = shift;
	my $ignore_pages = shift;
	my $start_at = shift;
	my $append_subtype = shift;
	my $ignorefile = shift;
	
	my $answers_txt = $destinationDirectory.'ANSWERS.txt';
	my $answers_pdf = $destinationDirectory.'ANSWERS.pdf';
	
	#unlink $answers_txt;	# delete this
	#unlink $answers_pdf;	# delete this
	
	my $template = $HomeWorkingDirectory.$answerstemplatefile;
	
	my $page_count = 0;
	my @addlater;

	# Get a list of all PDFs in the supplied directory
	my @inputfiles = glob "'$directory*.pdf'";		# don't fuck with the quotes here:  https://stackoverflow.com/questions/32260485/using-perl-glob-with-spaces-in-the-pattern
	
	#my @inputfiles = glob "$directory*.pdf";
	my $i = 1;
	
	$logger->info ( sprintf "\n--> %s -->", $thissubtype );

	foreach my $inputfile ( @inputfiles )
	{
		$logger->info ( sprintf "%i.	%s", $i, $inputfile );	
		$i++;
	}
	
	# If we have any PDF files....
	if( scalar( @inputfiles ) > 0 )
	{	
		$logger->info ( sprintf "\nMerging %i PDFs in [%s] into [%s]; ignoring the first [%i] page(s); starting at [%i] (appending [%s]; ignoring [%s])...\n", scalar @inputfiles, $directory, $outputfile, $ignore_pages, $start_at, $append_subtype, $ignorefile );
	
		# The PDF object.  this is the BIG destination file
		my $outputpdf = PDF::API2->new( -file => $outputfile );

		foreach my $inputfile ( @inputfiles )
		{
			next if ( $inputfile eq $ignorefile );	
			my ( $filename, $directories, $suffix ) = fileparse( $inputfile, qr/\.[^.]*/ );
			
			# if we are appending something (e.g. answers), don't add the blank writing page or the back cover
			if ( $filename =~ m/^999/ && $append_subtype ne '' )
			{
				$logger->info ( sprintf "	Will add file [%s] later.", $inputfile );
				push @addlater, $inputfile;
				next;
			}
				
			#$logger->info ( sprintf "	Adding file [%s].", $inputfile );
			my $inputpdf = PDF::API2->open( $inputfile );
			#my $page = $inputpdf->openpage( 0 );	# do we need this????
			
			# Add a built-in font to the PDF
			my $font = $outputpdf->corefont('Helvetica');		# 'Calibri' is not available

		    my @numpages = ( 1..$inputpdf->pages() );
			foreach my $numpage ( @numpages )
			{
		        # add page $numpage from $input_file to the end of the BIG $outputpdf
		        $outputpdf->importpage( $inputpdf, $numpage, 0 );
				print ".";	  # don't use the logger for this
		    }
		}
		
		print "\n";
		
		#
		# do we need to add answers/tapescripts/other at the back 
		if ( $append_subtype ne '' )
		{			
			$logger->info ( sprintf "	Appending [%s] PDFs in [%s] into [%s]...", $append_subtype, $directory, $outputfile );
			
			my $filename = sprintf( '%s%s PAGE BREAK.pdf', $directory, $append_subtype );
			SaveTextAsPDF( $filename, $append_subtype, 0, 0 );
			
			# add the 'ANSWERS' page break to the master file
			my $inputpdf = PDF::API2->open( $filename );
			$outputpdf->importpage( $inputpdf, 1, 0 );

			
			if ( not -e $answers_pdf )
			{
				$logger->info ( sprintf "	Making an answers file [%s]...\n", $answers_txt );			
				open my $fh, '>', $answers_txt;		# answers.txt

				foreach my $resource ( @archivedetails )
				{		
					next if ( $resource->{subtype} ne $append_subtype );	# only append files which match our subtype (exclude covers, ToCs. 'included' files etc )

					print $fh ">>>>>>>>>> ".$resource->{title}.">>>>>>>>>>\n";	
					
					my $answers_text = $resource->{TextofDocument};
					$answers_text =~ s/\[HYPERLINK: \S+\]//g;
					$answers_text =~ s/__+/_/g;		# get rid of lines of underlines on which students should write their answers
			
					print $fh $answers_text;			
					print $fh "<<<<<<<<<< ".$resource->{title}."<<<<<<<<<<\n\n";			
				
					# this code will add each PDF 
					#if ( -e $resource->{CommonFilenamePDF} )
					#{
					#    my $inputpdf = PDF::API2->open( $resource->{CommonFilenamePDF} );
					#	$logger->info ( sprintf "	Appending [%s]...", $resource->{CommonFilenamePDF} );
		
						# Add a built-in font to the PDF
					#	my $font = $outputpdf->corefont('Helvetica');		# 'Calibri' is not available

					#    my @numpages = ( 1..$inputpdf->pages() );
					#	foreach my $numpage ( @numpages )
					#	{
					#        # add page $numpage from $input_file to the end of the file $output_file
					#        $outputpdf->importpage( $inputpdf, $numpage, 0 );
					#    }					
					#}
					#else
					#{
					#	$logger->info ( sprintf "	Expected to find [%s] but didn't...\n", $resource->{CommonFilenamePDF} );				
					#}
				}
			
				close $fh;
			
				$logger->info ( sprintf "Attempting to convert [%s] to [%s]\n", $answers_txt, $answers_pdf );	
				
				my ( $filename, $directories, $suffix ) = fileparse( $answers_txt, qr/\.[^.]*/ );
				if ( system ( '/Applications/LibreOffice.app/Contents/MacOS/soffice', '--headless', '--convert-to', $answers_pdf, '--outdir', $directories, $answers_txt ) == 0 ) # success!
				
				#if ( system ( 'unoconv', '-t', $template, '--preserve', '-o', $answers_pdf, '-f', "pdf", $answers_txt ) == 0 && -e $answers_pdf ) # success!   
				{	
					$logger->info ( sprintf "	Successfully converted [%s] to [%s]!\n", $answers_txt, $answers_pdf );
				
					my $inputpdf = PDF::API2->open( $answers_pdf );

					# Add a built-in font to the PDF
					my $font = $outputpdf->corefont('Helvetica');		# 'Calibri' is not available

					# add a blank white rectangle to the answers.pdf to overwrite the footer in this document
					$logger->info ( sprintf "	Removing the footer from [%s]...", $answers_pdf );
				
					my @numpages = ( 1..$inputpdf->pages() );
					foreach my $numpage ( @numpages )
					{
						my $page = $inputpdf->openpage( $numpage );
	
						my $gfx = $page->gfx();
						my $rect = $gfx->rectxy( 0, 0, 595, 65 );		# x, y, tox, toy.	here this value SHOULD BE 65
						$rect->fillcolor( '#ffffff' );		# white
						$rect->fill();
						print ".";	  # don't use the logger for this
					}
					
					$inputpdf->update();	# save these changes			
				}
				else
				{
					$logger->info ( "System soffice command failed: $!\n" );
				}
				
				print "\n";				
			}
	
			#
			#
			#
			
			# now append each page to the big master file
			$inputpdf = PDF::API2->open( $answers_pdf );
			my @numpages = ( 1..$inputpdf->pages() );
			$logger->info ( sprintf "	Appending %i pages in [%s] to [%s]...", scalar @numpages, $answers_pdf, $outputfile );				

			foreach my $numpage ( @numpages )
			{
				# add page $numpage from $input_file to the end of the file $output_file
				$logger->info ( sprintf "		Adding page [%i] of [%s] to [%s]", $inputpdf, $numpage, $outputpdf );				
				
				$outputpdf->importpage( $inputpdf, $numpage, 0 );
				print ".";	  # don't use the logger for this
			}
			
			#
			#
			#
			
			# nearly done, now add the writing paper and the back cover
			foreach my $addlaterfile ( @addlater )
			{
			    my $inputpdf = PDF::API2->open( $addlaterfile );
				my $font = $outputpdf->corefont('Helvetica');		# 'Calibri' is not available

				$logger->info ( sprintf "\n	Appending [%s]...", $addlaterfile );

			    my @numpages = ( 1..$inputpdf->pages() );
				foreach my $numpage ( @numpages )
				{
			        # add page $numpage from $input_file to the end of the file $output_file
			        $outputpdf->importpage( $inputpdf, $numpage, 0 );
			    }
			}
		}
		
			
		# Set some of the properties of the document
		$outputpdf->info( 'Author' => $authorinfo, 'Title' => $outputfile );

		# Save the PDF
		$outputpdf->save() or warn "Failed to save file: $!\n";
		
		#
		# now we have one giant file, all that's left is to add page numbers to each page and the version info on the front cover
		#
		
		# we've just created this file, it's a merge of lots of smaller files
		my $pdf = new PDF::Report( File => $outputfile );
		
		$pdf->setFont( 'Helvetica-Bold' );		# 'Calibri' is not available
		$pdf->setSize( 40 );
		my ( $width, $height ) = $pdf->getPageDimensions();
		
		$logger->info ( sprintf "\nAdding page numbers to [%s] (pdf has %i pages; ignoring first %i page(s); starting at page %i)... \n", $outputfile, $pdf->pages(), $ignore_pages, $start_at );
		$page_count = $pdf->pages();
		
	    my @numpages = ( 1..$pdf->pages() );
		foreach my $numpage ( @numpages )
		{
			# add version number to front cover (if we have a front cover)
			if ( $numpage == 1 && scalar @frontcovers > 0 )
			{
				printf "Adding text to front cover (x = %i; y = %i)... ", $width, $height;
		
				my $page = $pdf->openpage( $numpage );		

				# add the version number
				my $version_text = sprintf "v%i", $VERSION_NO;
				$pdf->addRawText( $version_text, 500, 775, "black" );
				# A4 $x = 595; $y == 841
				# 0, 0 is the bottom, left hand corner
				# x is horizontal
				# y is vertical				
				
				# if this is a booklet with answers/tapescripts at the back, add this information to the cover
				if ( $append_subtype ne '' )
				{
					$pdf->setSize( 20 );
					$pdf->centerString( 20, $width-20, 60, 'WITH ' . $append_subtype );		# Centers $text between points $a and $b at (vertical) position $yPos											
				} 		
				
				# if it's an elective, add the course code and course name
				if ( $elective_coursecode ne '' && $elective_coursename ne '' )
				{
					$pdf->centerString( 330, 480, 300, $elective_coursecode );
					$pdf->centerString( 100, 480, 230, $elective_coursename );	# start, end, yPos
				}
				
				#$pdf->shadeRect(150, 10, 445, 50, "white");
				
				# add the assembled date
				my $dt = DateTime->today;
				my $created_text = sprintf "Assembled on %s", $dt->date;
				$pdf->setSize( 20 );
				$pdf->centerString( 20, $width-20, 20, $created_text );  # Centers $text between points $a and $b at (vertical) position $yPos
			}
			
			#
			# add the page number and the school logo
			#
			
			if ( $numpage > $ignore_pages && $numpage != $page_count )	# don't add a rectangle or page number to the first pages or the last page
			{
				my $page = $pdf->openpage( $numpage );
				my ($llx, $lly, $x, $y) = $page->get_mediabox();	# lower left X; lower left Y; upper left X; upper left Y
				
				$x = round( $x );	# round this
				$y = round( $y );	# round this
				
				my $string = '';
				
				# add the page number
				$pdf->setSize( 18 );
				
				# http://www.printernational.org/iso-paper-sizes.php & https://www.prepressure.com/library/paper-size
				if ( ( $x == 594 || $x == 595 ) && ( $y == 841 || $y == 842 ) )	# A4 portrait
				{
					#$logger->debug ( sprintf "PDF paper size for page #%3i is A4 portrait", $numpage );
					
					my $number_text = sprintf "%i", $start_at;
					$pdf->addRawText( $number_text, 540, 25 );

					$pdf->addImgScaled($header_image_file, 500, 755, 0.2) if ( $addheader == 1 );
				}
				elsif ( ( $x == 841 || $x == 842 ) && ( $y == 594 || $y == 595 ) )	# A4 landscape
				{
					#$logger->debug ( sprintf "PDF paper size for page #%3i is A4 landscape", $numpage );
					
					my $number_text = sprintf "%i", $start_at;
					$pdf->addRawText( $number_text, 787, 25 );
					
					$pdf->addImgScaled($header_image_file, 750, 520, 0.2) if ( $addheader == 1 );
				}
				elsif ( ( $x == 841 || $x == 842 ) && ( $y == 1190 || $y == 1191 ) )	# A3 portrait
				{
					#$logger->debug ( sprintf "PDF paper size for page #%3i is A3 portrait", $numpage );
				}
				elsif ( ( $x == 1190 || $x == 1191 ) && ( $y == 841 || $y == 842 ) )	# A3 landscape
				{
					#$logger->debug ( sprintf "PDF paper size for page #%3i is A3 landscape", $numpage );
				}
				elsif ( $x == 612 && $y == 792 )	# US Letter portrait
				{
					#$logger->debug ( sprintf "PDF paper size for page #%3i is US letter portrait", $numpage );
				}
				elsif ( $x == 792 && $y == 612 )	# US Letter landscape
				{
					#$logger->debug ( sprintf "PDF paper size for page #%3i is US letter landscape", $numpage );
				} else
				{
					#$logger->debug ( sprintf "PDF paper size for page #%3i is unknown: x is %s; y is %s", $numpage, $numpage, $x, $y );
				}
				
				$start_at++;
			}
		}
		$pdf->saveAs( $outputfile );
	}
	else
	{
		$logger->info ( sprintf "Directory [%s] has no PDF files.  Nothing to do. :-/\n", $directory );
	}

	$logger->info ( sprintf "<-- %s <--\n", $thissubtype );	
	return;
}