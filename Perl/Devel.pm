
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

our @EXPORT = qw( ) ;
our $VERSION = '0.02' ;

#-------------------------------------------------------------------------------

sub DumpDependentStack
{
my $ss = shift ;
my $dump ;

my $separator = '-' x 17 . "\n" ;

$dump .= $separator ;
$dump .= "$ss " ;

if(defined $ss->{NAME})
	{
	$dump .= "'$ss->{NAME}'" ;
	}
	
$dump .= " Dependent stack:\n" ;
$dump .= $separator ;

for my $dependent (@{$ss->{DEPENDENT_STACK}})
	{
	my ($spreadsheet, $address, $name) = @$dependent ;
	my $formula = '' ;
	
	if(exists $spreadsheet->{DATA}{$address}{GENERATED_FORMULA})
		{
		$formula = ": $spreadsheet->{DATA}{$address}{GENERATED_FORMULA}" ;
		
		if(exists $ss->{DEBUG}{DEFINED_AT})
			{
			my ($package, $file, $line) = @{$spreadsheet->{DATA}{$address}{DEFINED_AT}} ;
			$formula .= "[$package] $file:$line" ;
			}
		}
		
	$dump .= "$name!$address $formula\n" ;
	}

$dump .= "$separator\n" ;

return($dump) ;
}

#-------------------------------------------------------------------------------

sub Dump
{
my $ss = shift ;
my $address_list  = shift ; # array ref
my $display_setup = shift ;

use Data::Dumper ;
$Data::Dumper::Indent = 1 ;
#~ return(Dumper($ss)) ;

my $use_data_treedumper = 0 ;

eval 
	{
	use Data::TreeDumper ;
	$Data::TreeDumper::Useascii = 0 ;
	#~ return DumpTree($ss, 'ss:') ;
	$use_data_treedumper = 1 ;
	} ;

my $dump ;

$dump .= '-' x 60 . "\n" ;
$dump .= "$ss " ;

if(exists $ss->{NAME} && defined $ss->{NAME})
	{
	$dump .= "'$ss->{NAME}'" ;
	}
	
$dump .= "\n" ;

if($display_setup)
	{
	if($use_data_treedumper)
		{
		my $NoData = sub
				{
				my $s = shift ;
				
				if('Spreadsheet::Perl' eq ref $s)
					{
					return('HASH', undef, sort grep {! /DATA/} keys %$s) ;
					}
					
				return(Data::TreeDumper::DefaultNodesToDisplay($s)) ;
				} ;
		
		$dump .= DumpTree($ss, 'Setup:', FILTER => $NoData, DISPLAY_ADDRESS => 0) ;
		}
	else
		{
		for my $key (sort keys %$ss)
			{
			next if $key =~ /^DATA$/ ;
			
			$dump .= Data::Dumper->Dump([$ss->{$key}], [$key]);
			}
		}
	}
	
$dump .= "\n" ;

if($use_data_treedumper)
	{
	my %cell_filter ;
	
	if(defined $address_list)
		{
		my %cells_to_display ;
		@cells_to_display{$ss->GetAddressList(@$address_list)} = undef ;
		
		my $CellPruner = sub
					{
					my $s = shift ;
					if('HASH' eq ref $s)
						{
						return('HASH', $s, , SortCells(grep {exists $cells_to_display{$_};} keys %$s)); 
						}
						
					die "this filter is to be used on hashes!." ;
					} ;
					
		%cell_filter= (LEVEL_FILTERS => {0 => $CellPruner}) ;
		}
	else
		{
		my $CellSorter= sub
					{
					my $s = shift ;
					if('HASH' eq ref $s)
						{
						return('HASH', $s, SortCells(keys %$s)) ;
						}
						
					die "this filter is to be used on hashes!." ;
					} ;
					
		%cell_filter= (LEVEL_FILTERS => {0 => $CellSorter}) ;
		}
		
	my $NoDependentData = sub
				{
				my $s = shift ;
				
				if('HASH' eq ref $s)
					{
					my $is_dependent_hash = grep {/^Spreadsheet::Perl=HASH\(0x[0-9a-z]+\), [A-Z]/} keys %$s ;
					
					if($is_dependent_hash)
						{
						my @dependents ;
						my @dependents_formulas ;
						
						for my $dependent (keys %$s)
							{
							my ($spreadsheet, $cell, $name) = @{$s->{$dependent}{DATA}} ;
							push @dependents, "$name!$cell" ;
							
							if($ss->{DEBUG}{DEPENDENT})
								{
								push @dependents_formulas, "$s->{$dependent}{FORMULA}[0] [$s->{$dependent}{COUNT}]" ;
								}
							else
								{
								push @dependents_formulas, 1 ;
								}
							}
							
						return ('ARRAY', \@dependents_formulas, map{[$_, $dependents[$_]]} 0 .. $#dependents ) ;
						}
					else
						{
						return(Data::TreeDumper::DefaultNodesToDisplay($s)) ;
						}
					}
					
				return(Data::TreeDumper::DefaultNodesToDisplay($s)) ;
				} ;
			
	$dump .= DumpTree($ss->{DATA}, "Cells:", , DISPLAY_ADDRESS => 0, FILTER => $NoDependentData, %cell_filter) ;
	#~ $dump .= DumpTree($ss->{DATA}, "'Cells':") ;
	}
else
	{
	$dump .= Data::Dumper->Dump([$ss->{DATA}], ['Cells']) ;
	}
	
$dump .= "\n$ss " ;

if(defined $ss->{NAME})
	{
	$dump .= "'$ss->{NAME}'" ;
	}
	
$dump .= " dump end\n" . '-' x 60 . "\n" ;

return($dump) ;
}

#-------------------------------------------------------------------------------

sub GetCellsToUpdateDump
{
my $ss = shift ;
return( "Cells to update: " . (join " - ", $ss->GetCellsToUpdate()) . "\n") ;
}

#-------------------------------------------------------------------------------

1 ;

__END__
=head1 NAME

Spreadsheet::Perl::Devel- Development support for Spreadsheet::Perl

=head1 SYNOPSIS

  print $ss->Dump() ;
  print $ss->DumpDependentStack() ;
  print $ss->GetCellsToUpdateDump() ;
  
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
