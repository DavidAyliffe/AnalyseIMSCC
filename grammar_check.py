#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
grammar_check.py - Check a plain-text file for spelling and grammar errors
using the LanguageTool library and write results to a CSV file.

Usage:
    python grammar_check.py -i <inputfile.txt> -o <outputfile.csv>

Arguments:
    -i  Path to the plain-text input file (UTF-8 encoded)
    -o  Path for the output CSV file (will be created or overwritten)

Output format (!! delimited, one error per line):
    ruleId!!message!!replacements!!context!!offset!!errorLength!!category!!ruleIssueType

Called by AnalyseIMSCC.pl via:
    system("/usr/local/bin/python3.11", "grammar_check.py", "-i", txtfile, "-o", csvfile)

Dependencies:
    language_tool_python  (pip install language-tool-python)

Notes:
    - Uses British English (en-GB) rules.
    - On first run LanguageTool will download its language model (~200MB).
    - The LanguageTool instance is created inside main() so that importing
      this module does not trigger a download or slow initialisation.
"""

import sys
import getopt
import language_tool_python


def main(argv):
    """
    Parse command-line arguments, run LanguageTool on the input file,
    and write all matches to the output CSV file.

    Parameters
    ----------
    argv : list
        Command-line arguments (typically sys.argv[1:])
    """

    inputfile  = ''
    outputfile = ''

    # Parse -i and -o options
    try:
        opts, _ = getopt.getopt(argv, "i:o:")
    except getopt.GetoptError as e:
        print(f"Error: {e}", file=sys.stderr)
        print("Usage: grammar_check.py -i <inputfile> -o <outputfile>", file=sys.stderr)
        sys.exit(2)

    for opt, arg in opts:
        if opt == "-i":
            inputfile = arg
        elif opt == "-o":
            outputfile = arg

    # Validate that both required arguments were provided
    if not inputfile:
        print("Error: No input file specified (-i).", file=sys.stderr)
        sys.exit(2)
    if not outputfile:
        print("Error: No output file specified (-o).", file=sys.stderr)
        sys.exit(2)

    # Read the input file
    # Use errors='replace' to handle any stray non-UTF-8 bytes gracefully
    with open(inputfile, "r", encoding="utf-8", errors="replace") as infile:
        text = infile.read()

    # Initialise LanguageTool with British English rules.
    # NOTE: 'en-GB' is the correct BCP-47 language tag for British English.
    # On first use this downloads the LanguageTool jar (~200MB) automatically.
    tool = language_tool_python.LanguageTool("en-GB")

    # Run the grammar and spelling check
    matches = tool.check(text)

    # Write results to the output CSV file.
    # Delimiter is '!!' (double exclamation) to avoid conflicts with commas
    # in error messages and context strings.
    with open(outputfile, "w", encoding="utf-8") as outfile:
        for match in matches:

            # Build a semicolon-separated list of suggested replacements
            # (strip any trailing '; ' from the final entry)
            if match.replacements:
                str_replacements = "; ".join(match.replacements)
            else:
                str_replacements = ""

            # Write one line per match in the format expected by
            # ReadSpellingAndGrammarFileIntoArray() in AnalyseIMSCC.pl
            outfile.write(
                '{0}!!"{1}!!{2}!!"{3}"!!{4}!!{5}!!{6}!!{7}\n'.format(
                    match.ruleId,
                    match.message,
                    str_replacements,
                    match.context,
                    match.offset,
                    match.errorLength,
                    match.category,
                    match.ruleIssueType,
                )
            )


if __name__ == "__main__":
    main(sys.argv[1:])
