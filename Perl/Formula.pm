
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

our @EXPORT = qw( Formula DefineFunction ) ;
our $VERSION = '0.01' ;

#-------------------------------------------------------------------------------

sub Formula
{
return bless [@_], "Spreadsheet::Perl::Formula" ;
}

sub GenerateFormulaSub
{
#~ use Data::Dumper ;
#~ print Dumper(@_) ;

my ($ss, $current_cell_address, $anchor, $formula) = @_ ;

if($formula =~ /[A-Z]+\]?\[?[0-9]+/)
	{
	my $dh = $ss->{DEBUG}{ERROR_HANDLE} ;
	print $dh "Formula definition (anchor'$anchor' @ cell '$current_cell_address'): $formula\n" if $ss->{DEBUG}{PRINT_FORMULA} ;
	
	my ($column, $row) = $anchor =~ /^([A-Z]+)([0-9]+)/ ;
	my ($column_offset, $row_offset) = $ss->GetCellsOffset("$column$row", $current_cell_address) ;
	
	$formula =~ s/(\[?[A-Z]+\]?\[?[0-9]+\]?(:\[?[A-Z]+\]?\[?[0-9]+\]?)?)/$ss->OffsetAddress($1, $column_offset, $row_offset)/eg ;
	$formula =~ s/\$ss\{('|")(.*)}/\$ss->Get($1$2)/g ;
	$formula =~ s/\$ss\{([^'"].*)}/\$ss->Get("$1")/g ;
	
	print $dh "=> $formula\n" if $ss->{DEBUG}{PRINT_FORMULA} ;
	}

return
	(
	sub 
		{ 
		my $ss = shift ; 
		my $cell = shift ;
		
		my $result = eval $formula ;
		
		if($@)
			{
			my $dh = $ss->{DEBUG}{ERROR_HANDLE} ;
			
			my $ss_name = defined $ss->{NAME} ? "$ss->{NAME}!" : "$ss!" ;
				
			print $dh "At cell '$ss_name$cell' formula: $formula" ;
			print $dh " defined at '@{$ss->{DATA}{$cell}{DEFINED_AT}}'" if(exists $ss->{DATA}{$cell}{DEFINED_AT}) ;
			print $dh ":\n" ;
			print $dh "\t$@" ;
			return($ss->{MESSAGE}{ERROR}) ;
			}
		else
			{
			return($result) ;
			}
		
		}
	, $formula
	) ;
}

#-------------------------------------------------------------------------------

sub GetFormulaText
{
my $self = shift ;
my $address = shift ;

my ($cell, $start_cell, $end_cell) = $self->CanonizeAddress($address) ;

if($cell)
	{
	if(exists $self->{DATA}{$start_cell} && exists $self->{DATA}{$start_cell}{FORMULA})
		{
		return($self->{DATA}{$start_cell}{FORMULA}[0]) ;
		}
	else
		{
		return ;
		}
	}
else
	{
	my $dh = $self->{DEBUG}{ERROR_HANDLE} ;
	print $dh "GetFormula can only return the formula for one cell not '$address'.\n" ;
	
	return ;
	}
}

#-------------------------------------------------------------------------------

1 ;

__END__
=head1 NAME

Spreadsheet::Perl::Formula- Formula support for Spreadsheet::Perl

=head1 SYNOPSIS

  $ss{A1} = Formula('$ss{B1} + $ss{TOTAL}', $arg1, $arg2, ...) ;
  my $formula = $ss->GetFormulaText('A1') ;
  
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
