
package Spreadsheet::Perl ;

use 5.006 ;

use Carp ;
use strict ;
use warnings ;

require Exporter ;

our @ISA = qw(Exporter) ;

our %EXPORT_TAGS = 
	(
	'all' => [ qw() ]
	) ;

our @EXPORT_OK = ( @{ $EXPORT_TAGS{'all'} } ) ;

our @EXPORT ;
push @EXPORT, qw() ;

our $VERSION = '0.02' ;

#-------------------------------------------------------------------------------

sub InsertRows
{
my ($self, $start_row, $number_of_rows_to_insert) = @_ ;

confess "Invalid row '$start_row'\n" unless $start_row =~ /^\s*\d+\s*$/ ;

my (%moved_cell_list, %not_moved_cell_list) ;

for my $cell_address ($self->GetCellList())
	{
	# get all the cells for the rows under the $start_row
	my ($column, $row) = $cell_address =~ /([A-Z]+)(\d+)/ ;

	if( $row >= $start_row)
		{
		push @{$moved_cell_list{$row}}, $cell_address ;
		}
	else
		{
		push @{$not_moved_cell_list{$row}}, $cell_address ;
		}
	}

for my $row (reverse sort keys %moved_cell_list)
	{
	for my $cell_address (@{$moved_cell_list{$row}})
		{
		my $new_address = $self->OffsetAddress($cell_address, 0, $number_of_rows_to_insert) ; 
		
		if(exists $self->{CELLS}{$cell_address}{GENERATED_FORMULA})
			{
			$self->OffsetFormula($cell_address, 0, 0, $start_row, $number_of_rows_to_insert, "A$start_row:AAAA9999") ;
			}
		$self->{CELLS}{$new_address} = $self->{CELLS}{$cell_address} ;
		delete $self->{CELLS}{$cell_address} ;
		}
	}

# note, the cells don't have to be update in a specific order
# we keep the same order as moved cells to create the illusion
# of order
for my $row (reverse sort keys %not_moved_cell_list)
	{
	for my $cell_address (@{$not_moved_cell_list{$row}})
		{
		$self->OffsetFormula($cell_address, 0, 0, $start_row, $number_of_rows_to_insert, "A$start_row:AAAA9999") ;
		}
	}

for my $row_header (reverse sort grep {/^@/} $self->GetCellHeaderList())
	{
	my ($row_index) = $row_header =~ /^@(.+)/ ;
	if($row_index >= $start_row)
		{
		my $new_row = $row_index + $number_of_rows_to_insert ;
		$self->{CELLS}{"\@$new_row"} = $self->{CELLS}{$row_header} ;
		delete $self->{CELLS}{$row_header} ;	
		}
	}
}

sub InsertColumns
{
my ($self, $start_column, $number_of_columns_to_insert) = @_ ;

confess "Invalid w '$start_column'\n" unless $start_column =~ /^\s*[A-Z]{1,4}\s*$/ ;

my (%moved_cell_list, %not_moved_cell_list) ;

for my $cell_address ($self->GetCellList())
	{
	# get all the cells for the rows under the $start_row
	my ($column, $row) = $cell_address =~ /([A-Z]+)(\d+)/ ;

	my $column_index = FromAA($column) ;
	my $start_column_index = FromAA($start_column) ;

	if( $column_index >= $start_column_index)
		{
		push @{$moved_cell_list{$column}}, $cell_address ;
		}
	else
		{
		push @{$not_moved_cell_list{$column}}, $cell_address ;
		}
	}

for my $column (reverse sort keys %moved_cell_list)
	{
	for my $cell_address (@{$moved_cell_list{$column}})
		{
		my $new_address = $self->OffsetAddress($cell_address, $number_of_columns_to_insert, 0) ; 
		
		if(exists $self->{CELLS}{$cell_address}{GENERATED_FORMULA})
			{
			$self->OffsetFormula($cell_address, $start_column, $number_of_columns_to_insert, 0, 0, "${start_column}1:${start_column}9999") ;
			}
		$self->{CELLS}{$new_address} = $self->{CELLS}{$cell_address} ;
		delete $self->{CELLS}{$cell_address} ;
		}
	}

# note, the cells don't have to be update in a specific order
# we keep the same order as moved cells to create the illusion
# of order
for my $column (reverse sort keys %not_moved_cell_list)
	{
	for my $cell_address (@{$not_moved_cell_list{$column}})
		{
		$self->OffsetFormula($cell_address, $start_column, $number_of_columns_to_insert, 0, 0, "${start_column}1:${start_column}9999") ;
		}
	}

#Todo: check that AA A BB sort properly

my $start_column_index = FromAA($start_column) ;

for my $column_header (reverse sort grep {/^[A-Z]+0$/} $self->GetCellHeaderList())
	{
	my ($column_index) = $column_header =~ /^([A-Z]+)0$/ ;
	$column_index = FromAA($column_index) ;

	if($column_index >= $start_column_index)
		{
		my $new_column = $column_index + $number_of_columns_to_insert ;
		$new_column = ToAA($new_column) ;

		$self->{CELLS}{"${new_column}0"} = $self->{CELLS}{$column_header} ;
		delete $self->{CELLS}{$column_header} ;	
		}
	}
}


sub OffsetFormula
{
my 
	(
	$self, $cell_address,
	$start_column, $columns_to_insert,
	$start_row, $rows_to_insert,
	$range
	)  = @_ ;

return unless exists $self->{CELLS}{$cell_address}{GENERATED_FORMULA} ;

my $formula = $self->{CELLS}{$cell_address}{GENERATED_FORMULA} ; 
$formula =~ s/(\[?[A-Z]+\]?\[?[0-9]+\]?(:\[?[A-Z]+\]?\[?[0-9]+\]?)?)/$self->OffsetAddress($1, $columns_to_insert, $rows_to_insert, $range)/eg ;

$self->Set($cell_address, PF($formula)) ;
}


#-------------------------------------------------------------------------------

1 ;

__END__

=head1 NAME

Spreadsheet::Perl::InsertDelete - Columns and rows insertion and deletion

=head1 SYNOPSIS

Part of Spreadsheet::Perl.

=head1 AUTHOR

Khemir Nadim ibn Hamouda. <nadim@khemir.net>

  Copyright (c) 2011 Nadim Ibn Hamouda el Khemir. All rights
  reserved.  This program is free software; you can redis-
  tribute it and/or modify it under the same terms as Perl
  itself.
  
If you find any value in this module, mail me!  All hints, tips, flames and wishes
are welcome at <nadim@khemir.net>.

=cut
