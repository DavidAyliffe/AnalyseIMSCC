#!/usr/bin/perl
use utf8;
use strict;
use warnings;
use v5.10;
#use warnings;

our $version = "v2.00";
use File::Basename;

#
# http://johnbokma.com/mexit/2009/02/24/jpeg-to-pdf-using-perl.html
#
my $imagefile = shift;	
my ( $filename, $directories, $suffix ) = fileparse( $imagefile, qr/\.[^.]*/ );	

if ( $suffix eq '.png' || $suffix eq '.gif' || $suffix eq '.jpg' || $suffix eq '.jpeg' )
{
	my $newfile = $directories.$filename.".jpg";
	if ( system ( 'magick', $imagefile, $newfile ) == 0 ) # success!
	{
		# update this so we make a PDF with the new JPG (PDF::Create only supports JPG)
		printf "	Converted file.  Was [%s] now [%s]\n", $imagefile, $newfile;
	}
	else
	{
		printf "	Could not convert file.  Was [%s] now [%s]\n", $imagefile, $newfile;
		print  "	System unoconv command failed: $!\n";
		die;
	}
}
