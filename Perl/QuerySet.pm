
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

our @EXPORT = qw( GetCellList GetLastCell  GetFormulaText) ;
our $VERSION = '0.01' ;

#-------------------------------------------------------------------------------

sub SetAutocalc
{
my $self = shift ;
my $autocalc = shift ;
$self->{AUTOCALC} = $autocalc ;
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

local $self->{AUTOCALC} = 1 ;

for my $cell_name (keys %{$self->{DATA}})
	{
	if(exists $self->{DATA}{$cell_name}{SUB})
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

sub GetSpreadsheetReference
{
my $self = shift ;
my $address = shift ;

if($address =~ /^([A-Z]+)!(.+)/)
	{
	if(defined $self->{NAME} && $self->{NAME} eq $1)
		{
		return($self, $2) ;
		}
	else
		{
		if(exists $self->{OTHER_SPREADSHEETS}{$1})
			{
			return($self->{OTHER_SPREADSHEETS}{$1}, $2) ;
			}
		else
			{
			return(undef, $address) ;
			}
		}
	}
else
	{
	return($self, $address) ;
	}
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

return(SortCells(keys %{$self->{DATA}})) ;
}

#-------------------------------------------------------------------------------

sub GetLastIndexes
{
my $self = shift ;

my ($last_letter, $last_number) = ('A', 1) ;

for my $address(keys %{$self->{DATA}})
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
		   ( exists $ss->{DATA}{$_}{NEED_UPDATE} && $ss->{DATA}{$_}{NEED_UPDATE})
		||
			(
			   (exists $ss->{DATA}{$_}{FORMULA} || exists $ss->{DATA}{$_}{SUB})
			&& (! exists $ss->{DATA}{$_}{NEED_UPDATE})
			)
		} (SortCells(keys %{$ss->{DATA}}))
	) ;
}

#-------------------------------------------------------------------------------

sub DefineFunction
{
my ($ss, $name, $function_ref) = @_ ;
#~ my ($package, $filename, $line) = caller ;

{
no strict ;
*$name = sub {$function_ref->(@_) ;} ;
}

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
