#!/usr/bin/perl

use utf8;
use strict;
use v5.10;
use warnings;
no warnings 'experimental::smartmatch';

$|++;

our $version = "v2.01";

# use module
use XML::Simple;
use File::chdir;
use Cwd;
use Archive::Zip;
use Getopt::Long;
use File::Basename;
use File::Copy;
use locale;
use DateTime;
use POSIX qw( locale_h );
use Time::HiRes qw( gettimeofday tv_interval );	# use this to show how long the script took to run
use File::Path qw( make_path );


# command line arguments
my $hasfilename = "";
my $loglevel = "";  # choices: 'DEBUG', 'INFO', 'WARNING', 'ERROR', 'CRITICAL'

# directories
my $destinationDirectory = './temp/';
my $HomeWorkingDirectory = "";
my $RootDirectory = "";
my $zipbasename = '';	# without extension 
my $AllDocumentsDirectory = '';
my @hierarchy;
my $level;

my $folderhasitems = 0;
my $currentunit = "";
my $unitcount = 0;
my $hasunits = 0;
my $archiveitemcounter = 0;
my $xml;
my $parent = 'parent';
my $location = '';
my @archivedetails;				# the master array.  contains references to key hashes, one hash for each item (file)


my $usagestring = "
Usage is:  AnalyseIMSCC.pl -file=XXX.imscc
						
	-file 			the name (and location) of the file containing the data.  Either .imscc or .docx
	\n\n";

GetOptions(
    'file=s'			=> \$hasfilename,
    'loglevel=s'		=> \$loglevel
) or die $usagestring;


my $timeStarted = [gettimeofday];		# start the timer!
my ( $filename, $directories, $suffix ) = fileparse( $hasfilename, qr/\.[^.]*/ );		# set up 'global' filenames/directories

# beautify the filename, make the Lesson Focus Sentence Case
$filename =~ s/([\w']+)/\u\L$1/g;

unless ( ( -e $hasfilename && $suffix eq "\.imscc" ) )
{
     print $usagestring;
}
else
{
	my $dt = DateTime->today;
	printf "Today is: %s\n", $dt->date;
	printf "File is 		[%s]\n", $hasfilename;
	printf "LogLevel is 		[%s]\n", $loglevel;
	printf "Directory is [%s]; Filename is [%s]; Suffix is [%s]\n", $directories, $filename, $suffix;
	
	$zipbasename = $filename;
	$destinationDirectory = './'.$zipbasename.'/';
	$HomeWorkingDirectory = getcwd;
	$RootDirectory = $zipbasename.' Home Folder';
	$AllDocumentsDirectory = $destinationDirectory.'ALLDOCUMENTS/';
	
	if ( -e $hasfilename && $suffix eq "\.imscc" )
	{
		#printf "\nArchive creation time: %s\n", ctime( stat($hasfilename)->ctime );
		
		printf "Unzipping archive %s...", $hasfilename;
		my $zip = Archive::Zip->new( $hasfilename );
		
		foreach my $member ($zip->members)
		{
   			next if $member->isDirectory;
    		#(my $extractName = $member->fileName) =~ s{.*/}{};
			my $extractName = $member->fileName;
			$member->extractToFileNamed("$destinationDirectory/$extractName");
		}
	
		print "done.\n";
		
		# create object
		$xml = new XML::Simple;

		# read XML file
		print "Reading imsmanifest.xml...";
		my $data = $xml->XMLin( $destinationDirectory."imsmanifest.xml", forcearray => 1 );
		print "done.\n";

		Dive( $data->{organizations}[0]{organization}[0]{item} );

		MakeDirectoryAndCopyFiles();		
	}
	
	printf "Took %.2f seconds to complete.\n", tv_interval ( $timeStarted );
}


#
# this builds the home folder (i.e. what we see on Schoology)
#
sub MakeDirectoryAndCopyFiles
{
	print "Building directories and copying resources...\n";
	
	foreach my $item ( @archivedetails )
	{
		$item->{location} =~ s/-->|<--/\//g;
		make_path( $item->{location} ) unless ( -e $item->{location} && -d $item->{location} );
		
		if( -e $item->{filename} )	# this should always exist
		{
			my($filename, $directories, $suffix) = fileparse( $item->{filename}, qr/\.[^.]*/ );				
			my $destination = $item->{location}.'/'.$filename.$suffix;
			
			# don't do this, it causes problems when we come to farm out documents for the nth time
			# we modify this document (e.g. add a template) so this check is problematic
			#unless ( -e $destination )			
			if ( not -e $destination )
			{
				printf "	Copying [%s] to [%s]... \n", $item->{filename}, $destination;
				copy( $item->{filename}, $destination ) or warn "Cannot copy from [$item->{filename}] to [$destination]: $!";
			}
			elsif ( -e $destination && compare ( $item->{filename}, $destination ) != 0 )
			{
				printf "	Found updated [%s].  Updating folder... \n", $item->{filename};
				copy( $item->{filename}, $destination ) or warn "Cannot copy from [$item->{filename}] to [$destination]: $!";
			}
			else
			{
				#printf "	File [%s] has not changed, so not copying!\n", $item->{filename};
			}
		}
	}
	
	print "done.\n";
	return;
}

#
# beautify the filename given.  THIS COULD BE A FULL PATH TO A FILE SO BE CAREFUL
# this could also just be a string
#
sub BeautifyFilename
{
	my $path_and_filename = shift;
	my $old_path_and_filename = $path_and_filename;
	
	my($filename, $directories, $suffix) = fileparse( $path_and_filename, qr/\.[^.]*/ );	
	$directories = "" if $directories eq "./";
	my $oldfilename = $filename;	# save the original
	
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
	return $directories.$filename.$suffix;
}


#
#
#
sub ProcessItem
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
		
	# before we do anything with this, tidy up the filename.
	printf "    Processing item #%04i: %s\n", $archiveitemcounter, $filename;
	
	$filename = BeautifyFilename( $filename );
	
	if ( $OriginalFilename ne $filename )
	{
		move ( $OriginalFilename, $filename );
		#printf "	* Renamed file from [%s] to [%s]\n", $OriginalFilename, $filename;		
	}
		
	my ( $FilenameNoExt, $Directories, $Ext ) = fileparse( $filename, qr/\.[^.]*/ );	
	$Ext =~ s/\.//g;	# remove the dot  (before:	.docx	.pdf	.xml	after:	docx	pdf		xml)
	
	$key{type} = $Ext;
	$key{level} = $level;
	$key{parent} = $parent;
	$key{location} = $location;
	$key{title} = $FilenameNoExt;
	$key{directory} = $directory;
	$key{filename} = $filename;
	$key{FilenameNoPath} = $FilenameNoExt.'.'.$Ext;
	$key{IdAndFilenameNoPath} = sprintf "%04i %s", $archiveitemcounter, $FilenameNoExt.'.'.$Ext;
	$key{IdAndFilenameNoPathAsPDF} = sprintf "%04i %s", $archiveitemcounter, $FilenameNoExt.'.pdf';

	$key{FilenameWithoutExtension} = $Directories.$FilenameNoExt;
	$key{unit} = $currentunit;
				
	push @archivedetails, \%key;
	return;
}


#
#
#
sub Dive
{
	my ($ref) = @_;
		
	$location = join("-->", @hierarchy);
	printf "%s\n", $location;
	my $folderitemcounter = 0;
	
	for my $item (@$ref)
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
					closedir(DIR);
				
					$filename = $directory.'/'.$fileondisk;	
				}	
			}
		
			ProcessItem ( $directory, $filename, $item );
			
		}
		else	# it's a folder
		{
			my $oldunit = $parent;
			my $donotmakePDF = 0;			
			$parent = $item->{title}[0];
			
			$parent = $RootDirectory if ( length( $parent ) == 0 );
			printf "Found a new folder.   Was [%s] is now [%s]\n", $oldunit, $parent;
			
			if ( ( $parent =~ /^Unit\b/i ) && $donotmakePDF == 0 )		# have a space here to avoid catching 'United States'
			{
				printf "Starting a new Unit.  Was [%s] is now [%s]\n", $oldunit, $parent;
				$hasunits++ if ( $parent =~ /^Unit\b/i );
				
				# beautify the unit name, make the Unit Change Page Sentence Case
				$parent =~ s/([\w']+)/\u\L$1/g;
				
				$currentunit = $parent;
				$unitcount++;
			}
		}
		
		if ( defined $item->{item} )
		{
			push @hierarchy, $parent;
			
			$level++;
			Dive($item->{item});
			$level--;
		}
	}
	
	pop @hierarchy;
	$parent = $hierarchy[-1];
	$location = join("<--", @hierarchy);
	printf "[%s]	Parent is [%s]\n", $location, $parent;
	return;
}

