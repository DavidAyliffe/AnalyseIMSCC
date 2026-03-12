#!/usr/bin/perl

=head1 NAME

ConvertImageToJPG.pl - Convert an image file to JPEG format using ImageMagick

=head1 SYNOPSIS

    perl ConvertImageToJPG.pl <imagefile>

=head1 DESCRIPTION

Takes a single image file (.png, .gif, .jpg, .jpeg) as input and converts it
to JPEG format using ImageMagick's 'magick' command-line tool.

The output file is written to the same directory as the input, with a .jpg
extension replacing the original extension.

Called by AnalyseIMSCC.pl via:
    system( "/usr/bin/perl", "/usr/local/bin/ConvertImageToJPG.pl", $filename )

Returns exit code 0 on success, dies on failure.

=head1 DEPENDENCIES

    ImageMagick (magick command must be on PATH)
    File::Basename (core Perl module)

=head1 AUTHOR

David Ayliffe

=cut

use utf8;
use strict;
use warnings;
use v5.10;

use File::Basename;

our $version = "v2.00";

# ---------------------------------------------------------------------------
# Argument validation
# ---------------------------------------------------------------------------

# Print usage and exit if no argument was supplied
die "Usage: $0 <imagefile.png|gif|jpg|jpeg>\n" unless @ARGV;

my $imagefile = shift;

# Verify the input file exists and is readable
die "Error: Input file '$imagefile' does not exist or is not readable.\n"
    unless -e $imagefile && -r $imagefile;

# Parse the filename into its components: stem, directory and extension
my ( $filename, $directories, $suffix ) = fileparse( $imagefile, qr/\.[^.]*/ );

# ---------------------------------------------------------------------------
# Conversion
# ---------------------------------------------------------------------------

# Only process recognised image formats (PDF::Create only supports JPEG natively)
if ( $suffix eq '.png' || $suffix eq '.gif' || $suffix eq '.jpg' || $suffix eq '.jpeg' )
{
    my $newfile = $directories . $filename . ".jpg";

    # Use ImageMagick's 'magick' command to convert to JPEG.
    # On older ImageMagick versions this may be 'convert' rather than 'magick'.
    if ( system( 'magick', $imagefile, $newfile ) == 0 )    # exit 0 = success
    {
        printf "    Converted file.  Was [%s] now [%s]\n", $imagefile, $newfile;
    }
    else
    {
        # magick returns non-zero on failure; $! holds the OS error if applicable
        printf "    Could not convert file.  Was [%s] now [%s]\n", $imagefile, $newfile;
        printf "    ImageMagick 'magick' command failed: %s\n", $!;
        die;
    }
}
else
{
    # Unsupported extension — nothing to do
    printf "    Unsupported image format [%s].  Skipping.\n", $suffix;
}
