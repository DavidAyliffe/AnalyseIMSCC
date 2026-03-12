#!/usr/bin/perl

=head1 NAME

GetXMLDataTest.pl - Validate that a Schoology XML resource file can be parsed

=head1 SYNOPSIS

    perl GetXMLDataTest.pl <xmlfile>

=head1 DESCRIPTION

A lightweight validation script that attempts to parse an XML resource file
exported by Schoology (or a similar LMS).  It is used by AnalyseIMSCC.pl
as a pre-flight check before calling the heavier GetXMLData() function:

    if ( system( "/usr/bin/perl", "/usr/local/bin/GetXMLDataTest.pl",
                 $keyhash->{filename}, ">/dev/null" ) == 0 )

Returns exit code 0 if the XML parses successfully, non-zero (via die) if it
does not.  The parsed data is not used; this script exists purely to test
parseability and surface any XML errors early.

=head1 DEPENDENCIES

    XML::Simple  (CPAN)

=head1 AUTHOR

David Ayliffe

=cut

use utf8;
use strict;
use warnings;
use v5.10;

use XML::Simple;

# Flush output immediately (useful when output is piped)
$|++;

# ---------------------------------------------------------------------------
# Argument validation
# ---------------------------------------------------------------------------

die "Usage: $0 <xmlfile>\n" unless @ARGV;

my $filename = shift;

die "Error: File '$filename' does not exist or is not readable.\n"
    unless -e $filename && -r $filename;

# ---------------------------------------------------------------------------
# XML parsing
# ---------------------------------------------------------------------------

# Use OO constructor (avoids deprecated indirect-object 'new XML::Simple' syntax)
my $xml = XML::Simple->new();

# Attempt to read the XML file.
# XMLin() dies on parse errors, so a successful return means the file is valid.
# The forcearray option ensures consistent data structure regardless of
# the number of child elements.
my $data = $xml->XMLin( $filename, forcearray => 0 ) or die "Failed to parse '$filename'\n";

# If we reach here the file parsed successfully — exit 0
