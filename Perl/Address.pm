
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

our @EXPORT = qw( SetRangeName SetCellName SortCells) ;
our $VERSION = '0.01' ;

use Spreadsheet::ConvertAA ;

#-------------------------------------------------------------------------------

sub OffsetAddress
{
# this function accept adresses that are fixed ex: [A1]

my $self = shift ;
my $address = shift ;
my $column_offset = shift ;
my $row_offset = shift ;

my ($cell, $start_cell, $end_cell) ;

if($address =~ /(\[?[A-Z]+\]?\[?[0-9]+\]?):(\[?[A-Z]+\]?\[?[0-9]+\]?)/)
	{
	($cell, $start_cell, $end_cell) = (0, $1, $2) ;
	}
else	
	{
	if($address =~ /(\[?[A-Z]+\]?\[?[0-9]+\]?)/)
		{
		($cell, $start_cell, $end_cell) = (1, $1, $1) ;
		}
	else	
		{
		confess "OffsetAddress Error: Invalid address '$address'." ;
		}
	}


if($cell)
	{
	return
		(
		$self->OffsetCellAddress($start_cell, $column_offset, $row_offset)
		) ;
	}
else
	{
	my $lhs = $self->OffsetCellAddress($start_cell, $column_offset, $row_offset) ;
	my $rhs = $self->OffsetCellAddress($end_cell, $column_offset, $row_offset) ;
	
	if(defined $lhs && defined $rhs)
		{
		return("$lhs:$rhs") ;
		}
	else
		{
		return ;
		}
	}
}

sub OffsetCellAddress
{
my $self = shift ;
my $cell_address = shift ;
my $column_offset = shift ;
my $row_offset = shift ;

#~ print "OffsetCellAddress : $cell_address\n" ;
my ($full_column, $column, $full_row, $row) = $cell_address=~ /^(\[?([A-Z]+)\]?)(\[?([0-9]+)\]?)$/ ;

my $column_index = FromAA($column) ;
$column_index += $column_offset if($full_column !~ /[\[\]]/) ;
	
$row += $row_offset if($full_row !~ /[\[\]]/) ;

if($column_index > 0 && $row > 0)
	{
	return(ToAA($column_index) . $row) ;
	}
else
	{
	return ;
	}
}

#-------------------------------------------------------------------------------

sub GetCellsOffset
{
my $self = shift ;
my $cell_address1 = shift ;
my $cell_address2 = shift ;

my ($column1, $row1) = ($self->CanonizeAddress($cell_address1))[1] =~ /^([a-zA-Z]+)([0-9]+)$/ ;
my ($column2, $row2) = ($self->CanonizeAddress($cell_address2))[1] =~ /^([a-zA-Z]+)([0-9]+)$/ ;

my $column1_index = FromAA($column1) ;
my $column2_index = FromAA($column2) ;

return ($column2_index - $column1_index, $row2 - $row1) ;
}

#-------------------------------------------------------------------------------

sub SortCells
{
# returns the addresses, passed as argument, sorted.
return
	(
	sort
		{
		my ($a_spreadsheet_name, $a_letter, $a_number) = $a =~ /^([A-Z]+!)?([A-Z]+)(.+)$/ ;
		my ($b_spreadsheet_name, $b_letter, $b_number) = $b =~ /^([A-Z]+!)?([A-Z]+)(.+)$/ ;
		
		$a_spreadsheet_name ||= '' ;
		$b_spreadsheet_name ||= '' ;
		
		   $a_spreadsheet_name cmp $b_spreadsheet_name 
		|| length($a_letter) <=> length($b_letter) 
		|| $a_letter cmp $b_letter || $a_number <=> $b_number ;
		} @_
	) ;
}

#-------------------------------------------------------------------------------

sub IsAddress
{
my $self = shift ;
my $address = shift ;

eval
	{
	$self->CanonizeAddress($address) ; # dies if address is not valid
	} ;

defined $@ ? return(0) : return(1) ;
}

#-------------------------------------------------------------------------------

sub CanonizeAddress
{
# transform numeric cell index to alphabetic index
# transform symbolic addresses to alphabetic index

my $self = shift ;
my $address = uc(shift) ;

my ($cell, $start_cell, $end_cell) ;

my $spreadsheet = '' ;

if($address =~ /^([A-Z]+!)(.+)/)
	{
	$spreadsheet = $1 ;
	$address = $2 ;
	}

if($address =~ /^(.+):(.+)$/)
	{
	# range
	$start_cell = $self->CanonizeCellAddress($1) ;
	$end_cell   = $self->CanonizeCellAddress($2) ;
	}
else
	{
	# single cell or range name
	my $range = $self->CanonizeRangeName($address) ;
	
	if(defined $range)
		{
		($start_cell, $end_cell) = $range =~ /^(.+):(.+)$/ ;
		}
	else
		{
		$start_cell = $self->CanonizeCellAddress($address) ;
		$end_cell   = $start_cell ;
		$cell++ ;
		}
	}

return($cell, "$spreadsheet$start_cell", "$spreadsheet$end_cell") ;
}

#-------------------------------------------------------------------------------

sub SetRangeName
{
my $self    = shift ;
my $name    = uc(shift) ;
my $address = shift ;

croak "Error: Only Letters allowed in Range names. '$name'" if $name !~ /^[A-Z_]+$/ ; 

if($address =~ /^(.+):(.+)$/)
	{
	my $start_cell = $self->CanonizeCellAddress($1) ;
	my $end_cell   = $self->CanonizeCellAddress($2) ;
	
	$self->{RANGE_NAMES}{$name} = "$start_cell:$end_cell" ;
	}
else
	{
	confess "Error: Invalid Range '$address'." ;
	}
}

#-------------------------------------------------------------------------------

sub CanonizeRangeName
{
my $self    = shift ;
my $name    = uc(shift) ;

return $self->{RANGE_NAMES}{$name} ;
}

#-------------------------------------------------------------------------------

sub SetCellName
{
my $self    = shift ;
my $name    = uc(shift) ;
my $address = shift ;

croak "Error: Only Letters allowed in cell names. '$name'" if $name !~ /^[A-Z_]+$/ ; 

$self->{CELL_NAMES}{$name} = $self->CanonizeCellAddress($address) ;
}

#-------------------------------------------------------------------------------

sub CanonizeCellName
{
my $self = shift ;
my $name = uc(shift) ;

return $self->{CELL_NAMES}{$name} ;
}

#-------------------------------------------------------------------------------

sub CanonizeCellAddress
{
my $self = shift ;
my $address = shift ;

my $cell_address = $self->CanonizeCellName($address) ;

if(defined $cell_address)
	{
	return($cell_address) ;
	}
else
	{
	if($address =~ /^[A-Z]+[0-9]+$/)
		{
		return($address) ;
		}
	else
		{
		if($address =~ /^([0-9]+),([0-9]+)$/)
			{
			return(ConvertNumericToAddress($1, $2)) ;
			}
		else
			{
			croak "Invalid Address '$address'." ;
			}
		}
	}
}

#-------------------------------------------------------------------------------

sub ConvertAdressToNumeric
{
my $address = shift ;

if($address =~ /^([A-Z]+)([0-9]+)$/)
	{
	my $letters = $1 ;
	my $figure = $2 ;
	
	my $converted_letters = FromAA($letters) ;
	
	#~ print "ConvertAdressToNumeric: $address => ($letters => $converted_letters), $figure\n" ;
	
	return($converted_letters, $figure) ;
	}
else
	{
	confess "Invalid Address '$address'." ;
	}
}

#-------------------------------------------------------------------------------

sub ConvertNumericToAddress
{
my ($x, $y) = @_ ;

my $converted_figures = ToAA($x) ;

#~ print "ConvertNumericToAddress: $x,$y => $converted_figures$y\n" ;

return("$converted_figures$y") ;
}

#-------------------------------------------------------------------------------

sub GetAddressList
{
my $self = shift ;
my @addresses_definition = @_;
my @addresses ;

for my $address (@addresses_definition)
	{
	my $spreadsheet = '' ;
	
	if($address =~ /^([A-Z]+!)(.+)/)
		{
		$spreadsheet = $1 ;
		$address = $2 ;
		}
		
	my ($cell, $start_cell, $end_cell) = $self->CanonizeAddress($address) ;
	
	if($cell)
		{
		push @addresses, "$spreadsheet$start_cell" ;
		}
	else
		{
		my ($start_x, $start_y) = ConvertAdressToNumeric($start_cell) ;
		my ($end_x, $end_y) = ConvertAdressToNumeric($end_cell) ;
		
		my @x_list ;
		if($start_x < $end_x)
			{
			@x_list = ($start_x .. $end_x) ;
			}
		else
			{
			@x_list = ($end_x .. $start_x ) ;
			@x_list = reverse @x_list ;
			}
		
		my @y_list ;
		if($start_y < $end_y)
			{
			@y_list = ($start_y .. $end_y) ;
			}
		else
			{
			@y_list = ($end_y .. $start_y ) ;
			@y_list = reverse @y_list ;
			}
			
		for my $x (@x_list)
			{
			for my $y (@y_list)
				{
				push @addresses, $spreadsheet . ConvertNumericToAddress($x, $y) ;
				}
			}
			
		print "GetAddressList '$address': " . (join ' - ', @addresses) . "\n" if($self->{DEBUG}{ADDRESS_LIST});
		}
	}
	
return(@addresses) ;
}

#-------------------------------------------------------------------------------

1 ;

__END__
=head1 NAME

Spreadsheet::Perl::Address - Cell adress manipulation functions

=head1 SYNOPSIS

  $ss->SetRangeName("TestRange", 'A5:B8') ;
  my ($x, $y) = ConvertAdressToNumeric($address) ;
  ...
  
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

=head1 DEPENDENCIES

B<Spreadsheet::ConvertAA>.

=cut
