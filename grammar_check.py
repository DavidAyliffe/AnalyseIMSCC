#!/usr/bin/python
# -*- coding: utf-8 -*-

# python grammar_check.py -i "text.txt" -o "out.csv"
# python grammar_check.py -i "text.txt" -o "out.csv"
# python grammar_check.py -i "text.txt" -o "out.csv"

# java -jar languagetool-commandline.jar -l en-GB text.txt

import sys, getopt
#import language_check          # pip install language-check
import language_tool_python     # newest!

tool = language_tool_python.LanguageTool( 'en-UK' )

def main(argv):
    
    inputfile = ''
    outputfile = ''
    strReplacements = ''

    opts, args = getopt.getopt( argv,"i:o:" )

    for opt, arg in opts:
        if opt in ( "-i" ):
            inputfile = arg
        elif opt in ( "-o" ):
            outputfile = arg

#    print ( 'Input file is:', inputfile )
#    print ( 'Output file is:', outputfile )

    infile = open( inputfile, "r", encoding='utf-8' )
    text = infile.read()
    infile.close()

    outfile = open( outputfile, "w+", encoding='utf-8' )
    #outfile.write( 'Rule ID,Message,Replacements,Context,Offset,Error Length,Category,Rule Issue Type\n' )

    matches = tool.check( text )
    #matches = tool.check( "A sentence with a error in the Hitchhiker’s Guide tot he Galaxy" )
    #print ( 'Found ', len( matches ), ' possible errors' )

    #print ( matches[0] )
    #print ( matches[1] )

    for match in matches:
        
#        print( 'Rule ID is         ', match.ruleId )               # e.g. EN_A_VS_AN
#        print( 'Message is         ', match.message )              # e.g. Use “an” instead of ‘a’ if the following word starts with a vowel sound, e.g. ‘an article’, ‘an hour’
#        print( 'Replacements are   ', match.replacements[0] )      # e.g. an
#        print( 'Context is         ', match.context )              # e.g. A sentence with a error in the Hitchhiker’s Guide tot he ...
#        print( 'Offset is          ', match.offset )               # e.g. 16
#        print( 'Error Length is    ', match.errorLength )          # e.g. 1
#        print( 'Category is        ', match.category )             # e.g. MISC
#        print( 'Rule Issue Type is ', match.ruleIssueType )        # e.g. misspelling

        strReplacements = ''
        
        for replacement in match.replacements:
            strReplacements = strReplacements + replacement + "; "

        #strReplacements.removesuffix( " | " )  # doesn't seem to do what I want
        strReplacements = strReplacements[:-2]
        
        #if match.ruleId not in ('WHITESPACE_RULE', 'EN_QUOTES', 'DASH_RULE'):   # ignore these, they are too common
        outfile.write(  "{0}!!\"{1}\"!!{2}!!\"{3}\"!!{4}!!{5}!!{6}!!{7}\n".format( match.ruleId, match.message, strReplacements, match.context, match.offset, match.errorLength, match.category, match.ruleIssueType ) )        # match.replacements[0]

    outfile.close()


if __name__ == "__main__":
   main(sys.argv[1:])