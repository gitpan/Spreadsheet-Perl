
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

#~ our @EXPORT = qw( GetCellList GetLastIndexes  GetFormulaText DefineFunction ) ;
our @EXPORT ;
push @EXPORT, qw( GetCellList GetLastIndexesGetFormulaText DefineSpreadsheetFunction ) ;

our $VERSION = '0.01' ;

#-------------------------------------------------------------------------------

sub SetAutocalc
{
my $self = shift ;
my $autocalc = shift ;

if(defined $autocalc)
	{
	$self->{AUTOCALC} = $autocalc ;
	}
else
	{
	$self->{AUTOCALC} = 1 ;
	}
}

#-------------------------------------------------------------------------------

sub GetAutocalc
{
my $self = shift ;
return($self->{AUTOCALC}) ;
}

#-------------------------------------------------------------------------------

sub Recalculate
{
my $self = shift ;

for my $cell_name (keys %{$self->{CELLS}})
	{
	if(exists $self->{CELLS}{$cell_name}{FETCH_SUB})
		{
		$self->Get($cell_name) ;
		}
	}
}

#-------------------------------------------------------------------------------

sub AddSpreadsheet
{
my $self = shift ;
my $name = shift ;
my $reference = shift ;

confess "Invalid spreadsheet name '$name'." unless $name =~ /^[A-Z]+$/ ;

return if(defined $self->{NAME} && $self->{NAME} eq $name) ;

if(exists $self->{OTHER_SPREADSHEETS}{$name})
	{
	if($self->{OTHER_SPREADSHEETS}{$name} != $reference)
		{
		print "AddSpreadsheet: Replacing spreadsheet '$name'\n" ;
		}
	}
	
$self->{OTHER_SPREADSHEETS}{$name} = $reference ;
}

#-------------------------------------------------------------------------------

sub SetName
{
my $self = shift ;
my $name = shift ;

$self->{NAME} = $name ;
}

#-------------------------------------------------------------------------------

sub GetName
{
my $self = shift ;
my $ss = shift ;

return($self->{NAME} || "$self") unless defined $ss ;

my $name ;

if(exists $self->{OTHER_SPREADSHEETS})
	{
	for my $current_name (keys %{$self->{OTHER_SPREADSHEETS}})
		{
		if($self->{OTHER_SPREADSHEETS}{$current_name} == $ss)
			{
			$name = $current_name ;
			last ;
			}
		}
	}
	
return($name) ;
}

#-------------------------------------------------------------------------------

sub GetCellList
{
my $self = shift ;

return(SortCells(keys %{$self->{CELLS}})) ;
}

#-------------------------------------------------------------------------------

sub GetLastIndexes
{
my $self = shift ;

my ($last_letter, $last_number) = ('A', 1) ;

for my $address(keys %{$self->{CELLS}})
	{
	my ($letter, $number) = $address =~ /([A-Z]+)(.+)/ ;
	
	($last_letter) = sort{length($b) <=> length($a) || $b cmp $a} ($last_letter, $letter) ;
	$last_number   = $last_number > $number ? $last_number : $number ;
	}
	
return($last_letter, $last_number) ;
}


#-------------------------------------------------------------------------------

sub GetCellsToUpdate
{
# return the address of all the cells needing an update

my $ss = shift ;

return
	(
	grep 
		{
		   ( exists $ss->{CELLS}{$_}{NEED_UPDATE} && $ss->{CELLS}{$_}{NEED_UPDATE})
		||
			(
			   (exists $ss->{CELLS}{$_}{FORMULA} || exists $ss->{CELLS}{$_}{FETCH_SUB})
			&& (! exists $ss->{CELLS}{$_}{NEED_UPDATE})
			)
		} (SortCells(keys %{$ss->{CELLS}}))
	) ;
}

#-------------------------------------------------------------------------------

sub DefineSpreadsheetFunction
{
my ($name, $function_ref) = @_ ;

confess "Expecting a name!" unless '' eq ref $name && defined $name && $name ne '' ;

#~ my ($package, $filename, $line) = caller ;

#~ *$name = sub {$function_ref->(@_) ;} ; # this has perl generate a <warning but in the wrong context

no strict ;
if(eval "*Spreadsheet::Perl::$name\{CODE}")
	{
	warn "Subroutine Spreadsheet::Perl::$name redefined at @{[join ':', caller()]}\n" ;
	}
	
*$name = $function_ref ; # this doesn't generate a warning but still works
}

#-------------------------------------------------------------------------------

1 ;

__END__
=head1 NAME

Spreadsheet::Perl::QuerySet- Functions at the spreadsheet level

=head1 SYNOPSIS

  SetAutocalc
  GetAutocalc
  Recalculate
  
  SetName
  GetName
  AddSpreadsheet
  GetSpreadsheetReference
  
  GetCellList
  GetLastIndexes
  GetCellsToUpdate
  
  DefineFunction
  
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
