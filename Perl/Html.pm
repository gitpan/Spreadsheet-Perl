
package Spreadsheet::Perl ;

use 5.006 ;

use Carp ;
use strict ;
use warnings ;

require Exporter ;
#~ use AutoLoader qw(AUTOLOAD) ;

our @ISA = qw(Exporter) ;

our %EXPORT_TAGS = 
	(
	'all' => [ qw() ]
	) ;

our @EXPORT_OK = ( @{ $EXPORT_TAGS{'all'} } ) ;

#~ our @EXPORT = qw( ) ;
our @EXPORT ;
push @EXPORT, qw( ) ;

our $VERSION = '0.01' ;

#-------------------------------------------------------------------------------

sub GenerateHtml
{
eval "use Data::Table ;" ;

confess "Data::Table is not installed" if($@) ;

my $ss = shift ;
my ($last_letter, $last_number) = $ss->GetLastIndexes() ;
my ($cols, $rows) = (FromAA($last_letter), $last_number) ;

my $data = 
	[
	map
		{
		[
		map{''} (1 ..$cols + 1)
		]
		} (1 .. $rows)
	
	] ;

my $table1 = new Data::Table($data, [ map{ToAA($_)} (0, 1 .. $cols)]) ;

for (1 .. $rows)
	{
	$table1->setElm($_ - 1, 0, $_) ;
	}
	
for ($ss->GetCellList())
	{
	my ($cellcol, $cellrow) =  ConvertAdressToNumeric($_) ;
	
	$table1->setElm ($cellrow - 1, $cellcol, $ss->Get($_)) ;
	}

return($table1->html()) ;

}

sub GenerateHtmlToFile
{
my $ss = shift ;
my $file_name = shift ;

open(HTML_FILE, '>', $file_name) or confess "Can't open '$file_name' for HTML dump" ;
print HTML_FILE $ss->GenerateHtml() ;
close HTML_FILE
}

#-------------------------------------------------------------------------------

1 ;

__END__
=head1 NAME

Spreadsheet::Perl::Html- HTML output for Spreadsheet::Perl

=head1 SYNOPSIS

  print $ss->GenerateHtml() ;
  $ss->GenerateHtmlToFile('file.html') ;
  
=head1 DESCRIPTION

Part of Spreadsheet::Perl.

=head1 AUTHOR

Khemir Nadim ibn Hamouda. <nadim@khemir.net>

  Copyright (c) 2004 Nadim Ibn Hamouda el Khemir. All rights
  reserved.  This program is free software; you can redis-
  tribute it and/or modify it under the same terms as Perl
  itself.
  
If you find any value in this module, mail me!  All hints, tips, flames and wishes
are welcome at <nadim@khemir.net>.

=cut
