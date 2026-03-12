#!/usr/bin/perl
use utf8;
use strict;
use warnings;
use v5.10;
#use warnings;

our $version = "v2.00";


#
# http://johnbokma.com/mexit/2009/02/24/jpeg-to-pdf-using-perl.html
#
my $imagefile = shift;
my $imagetype = shift;
my $outputfile = shift;
my $image;	# object

# set some of the properties of the document
my $authorinfo = sprintf "AnalyseIMSCC%s.pl with PDF::API2 by David Ayliffe", $version;
	
my ( $filename, $directories, $suffix ) = fileparse( $imagefile, qr/\.[^.]*/ );	

# get the image size, and print it out
my( $width, $height ) = imgsize( $imagefile );

my $pdf  = PDF::Create->new( 'filename' => $outputfile, 'Author' => $authorinfo, 'Title'=> $imagefile ) or warn "Failed to create new PDF object: $!\n";
my $root = $pdf->new_page( 'MediaBox' => $pdf->get_page_size('A4') );
my $page = $root->new_page;
my $font = $pdf->font( 'Subtype' => 'Type1', 'Encoding' => 'WinAnsiEncoding', 'BaseFont' => 'Helvetica' );
	
$page->stringc( $font, 20, 595/2, 742, $filename );	# font, size, x, y, text, centre the text

# include a jpeg image with scaling to 20% size
my $imageobj = $pdf->image( $imagefile );

my $scale_ratio = 1;
my $newwidth = $width;
my $newheight = $height;

if ( ( $newwidth > $newheight && $newwidth > 495 ) || ( $newheight == $newwidth && $newheight > 495 ) )
{
	while ( $newwidth > 495 )		# this is the comfortable border we want
	{
		$newwidth = $width;
		$newheight = $height;

		$scale_ratio = $scale_ratio - 0.001;
		$newwidth = $newwidth * $scale_ratio;
		$newheight = $height * $scale_ratio;
	}
}
elsif ( $newheight > $newwidth && $newheight > 642 )
{
	while ( $newheight > 642 )	# this is the comfortable border we want
	{
		$newwidth = $width;
		$newheight = $height;
		
		$scale_ratio = $scale_ratio - 0.001;
		$newheight = $height * $scale_ratio;
		$newwidth = $newwidth * $scale_ratio;
	}		
}

printf "\n		Scale ratio is now %f", $scale_ratio;
printf "\n		Original Image width was %i and height is %i", $width, $height;
printf "\n		New Image width is now %i and height is %i", $newwidth, $newheight;
	
$page->image (
	'image'  => $imageobj,
	'xalign' => 1,				# Alignment of image; 0 is left/bottom, 1 is centered and 2 is right, top
	'yalign' => 1,
	'xscale' => $scale_ratio,	# Scaling of image. 1.0 is original size
	'yscale' => $scale_ratio,
	'xpos'   => 595/2,			# Position of image (required)
	'ypos'   => 842/2 );

$pdf->close;
return;	