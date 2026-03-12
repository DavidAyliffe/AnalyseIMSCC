#!/usr/bin/perl
use utf8;
use strict;
use v5.10;
#use warnings;

$|++;

# use module
use XML::Simple;


my $filename = shift;

my $xml = new XML::Simple;

# attempt to read the XML file
my $data = $xml->XMLin($filename) or die;