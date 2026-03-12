#!/usr/bin/perl

=head1 NAME

SaveImageAsPDF.pl - Embed an image file into a single-page A4 PDF

=head1 SYNOPSIS

    perl SaveImageAsPDF.pl <imagefile> <imagetype> <outputfile.pdf>

=head1 DESCRIPTION

Creates an A4 portrait PDF containing:
  - The image filename as a centred text label near the top of the page
  - The image itself, scaled to fit within comfortable A4 margins (495 x 642 pt)
    and centred on the page

The image is scaled proportionally:
  - If wider than tall (landscape/square), it is scaled until width <= 495 pt
  - If taller than wide (portrait), it is scaled until height <= 642 pt
  - If already within bounds, no scaling is applied (scale = 1.0)

Called by AnalyseIMSCC.pl when converting standalone image resources to PDF.

NOTE: PDF::Create only supports JPEG images natively, so the input image
should already have been converted to .jpg by ConvertImageToJPG.pl first.

A4 page dimensions in points: 595 (w) x 842 (h)
Comfortable content area:      495 (w) x 642 (h) leaving ~50 pt margins

=head1 ARGUMENTS

    imagefile   Path to the source image file (JPEG recommended)
    imagetype   Image type string (reserved for future use; currently unused)
    outputfile  Path for the output PDF file

=head1 DEPENDENCIES

    File::Basename  (core Perl module)
    Image::Size     (CPAN) - reads image dimensions
    PDF::Create     (CPAN) - creates the PDF

=head1 AUTHOR

David Ayliffe

=cut

use utf8;
use strict;
use warnings;
use v5.10;

use File::Basename;
use Image::Size;
use PDF::Create;

our $version = "v2.00";

# PDF author string embedded in the output file's metadata
my $authorinfo = sprintf "AnalyseIMSCC%s.pl with PDF::Create by David Ayliffe", $version;

# ---------------------------------------------------------------------------
# Argument handling
# ---------------------------------------------------------------------------

die "Usage: $0 <imagefile> <imagetype> <outputfile.pdf>\n" unless @ARGV >= 3;

my $imagefile  = shift;   # source image path
my $imagetype  = shift;   # image type label (currently unused; reserved)
my $outputfile = shift;   # destination PDF path

die "Error: Image file '$imagefile' does not exist or is not readable.\n"
    unless -e $imagefile && -r $imagefile;

# Parse the image filename into stem, directory and extension
my ( $filename, $directories, $suffix ) = fileparse( $imagefile, qr/\.[^.]*/ );

# ---------------------------------------------------------------------------
# Read image dimensions
# ---------------------------------------------------------------------------

# imgsize() returns (width, height) in pixels; returns undef on failure
my ( $width, $height ) = imgsize( $imagefile );
die "Error: Could not determine image dimensions for '$imagefile'.\n"
    unless defined $width && defined $height;

# ---------------------------------------------------------------------------
# Compute scale factor to fit image within comfortable A4 margins
# A4 usable area: 495 pt wide x 642 pt tall (leaving ~50 pt margins all round)
# ---------------------------------------------------------------------------

my $MAX_WIDTH  = 495;   # maximum image width in points
my $MAX_HEIGHT = 642;   # maximum image height in points

my $scale_ratio = 1.0;

if ( ( $width >= $height && $width > $MAX_WIDTH ) ||
     ( $width == $height && $height > $MAX_WIDTH ) )
{
    # Landscape or square image — scale down until width fits
    $scale_ratio = $MAX_WIDTH / $width;
}
elsif ( $height > $width && $height > $MAX_HEIGHT )
{
    # Portrait image — scale down until height fits
    $scale_ratio = $MAX_HEIGHT / $height;
}
# else: image is already within bounds; scale_ratio stays 1.0

my $newwidth  = $width  * $scale_ratio;
my $newheight = $height * $scale_ratio;

printf "\n        Scale ratio is now %f",        $scale_ratio;
printf "\n        Original Image: width=%i  height=%i", $width,    $height;
printf "\n        Scaled Image:   width=%i  height=%i", $newwidth, $newheight;

# ---------------------------------------------------------------------------
# Create the PDF
# ---------------------------------------------------------------------------

# PDF::Create uses points (1/72 inch).  A4 = 595 x 842 pt.
my $pdf  = PDF::Create->new(
    'filename' => $outputfile,
    'Author'   => $authorinfo,
    'Title'    => $imagefile,
) or die "Failed to create PDF object for '$outputfile': $!\n";

my $root = $pdf->new_page( 'MediaBox' => $pdf->get_page_size('A4') );
my $page = $root->new_page;

# Helvetica is a standard PDF core font, always available without embedding
my $font = $pdf->font(
    'Subtype'  => 'Type1',
    'Encoding' => 'WinAnsiEncoding',
    'BaseFont' => 'Helvetica',
);

# Print the filename as a centred label near the top of the page
# stringc( font, size, x, y, text ) — x=centre of A4 width, y=742 (near top)
$page->stringc( $font, 20, 595 / 2, 742, $filename );

# Load the image into the PDF
my $imageobj = $pdf->image( $imagefile );

# Place the image centred on the page with the computed scale factor
# xpos/ypos are the centre point of the image; xalign/yalign=1 means centred
$page->image(
    'image'  => $imageobj,
    'xalign' => 1,              # 0=left/bottom, 1=centred, 2=right/top
    'yalign' => 1,
    'xscale' => $scale_ratio,   # horizontal scale factor (1.0 = original size)
    'yscale' => $scale_ratio,   # vertical scale factor
    'xpos'   => 595 / 2,        # horizontal centre of A4
    'ypos'   => 842 / 2,        # vertical centre of A4
);

$pdf->close;
