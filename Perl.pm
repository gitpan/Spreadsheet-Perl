
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

our @EXPORT = qw( Reset ) ;
our $VERSION = '0.01' ;

use Spreadsheet::Perl::Address ;
use Spreadsheet::Perl::Function ;
use Spreadsheet::Perl::Formula ;
use Spreadsheet::Perl::Validator ;
use Spreadsheet::Perl::Lock ;
use Spreadsheet::Perl::Devel ;
use Spreadsheet::Perl::QuerySet ;
use Spreadsheet::Perl::Validator ;

#-------------------------------------------------------------------------------

sub GetDefaultData
{ 
return 
	(
	  NAME                => undef
	, OTHER_SPREADSHEETS  => {}
	, DEBUG               => { ERROR_HANDLE => \*STDERR }
	
	, VALIDATORS          => [['Spreadsheet lock validator', \&LockValidator]]
	, AUTOCALC            => 1
	, ERROR_HANDLER       => undef # user registred sub
	, MESSAGE             => 
				{
				ERROR => '#error'
				, NEED_UPDATE => "#need update"
				}
				
	, DATA                => {}
	) ;
}

sub Reset
{
my $self = shift ;
my $data = shift ;

%$self = GetDefaultData() ;
	
if(defined $data)
	{
	$self->{DATA} = $data ;
	}
else
	{
	$self->{DATA} = {} ;
	}
}

#-------------------------------------------------------------------------------

sub TIEHASH 
{
my $class = shift ;

my $self = 
	{
	  GetDefaultData()
	, @_ 
	} ;

return(bless $self, $class) ;
}

#-------------------------------------------------------------------------------

sub FETCH 
{
my $self    = shift ;
my $address = shift;

if($self->{DEBUG}{AUTOCALC})
	{
	my $dh = $self->{DEBUG}{ERROR_HANDLE} ;
	
	if($self->{AUTOCALC})
		{
		print $dh $self->GetName() . " AUTOCALC is ON @ address '$address'.\n" ;
		}
	else
		{
		print $dh $self->GetName() . " AUTOCALC is OFF @ address '$address'.\n" ;
		}
	}
	
#inter spreadsheet references
my $original_address = $address ;
my $ss_reference ;
($ss_reference, $address) = $self->GetSpreadsheetReference($address) ;

if(defined $ss_reference)
	{
	if($ss_reference == $self)
		{
		# fine, it's us
		}
	else
		{
		if($self->{DEBUG}{FETCH_FROM_OTHER})
			{
			my $dh = $self->{DEBUG}{ERROR_HANDLE} ;
			print $dh $self->GetName() . " Fetching from spreadsheet '$original_address'.\n" ;
			}
			
		#handle inter spreadsheet formula references
		my $have_stack = (exists $self->{DEPENDENT_STACK} && @{$self->{DEPENDENT_STACK}}) ;
		push @{$ss_reference->{DEPENDENT_STACK}}, @{$self->{DEPENDENT_STACK}}[-1] if($have_stack) ;
			
		# force recalculation on other spreadsheet
		if($self->{DEBUG}{AUTOCALC} && $self->{AUTOCALC} && ! $ss_reference->{AUTOCALC})
			{
			my $dh = $self->{DEBUG}{ERROR_HANDLE} ;
			print $dh $self->GetName() . " Forcing AUTOCALC for address '$original_address'.\n" ;
			}
		
		local $ss_reference->{AUTOCALC} = 1  ;
		
		my $cell_value = $ss_reference->Get($address) ;
		
		pop @{$ss_reference->{DEPENDENT_STACK}} if($have_stack);
		
		return($cell_value) ;
		}
	}
else
	{
	confess "Can't find Spreadsheet object for address '$address'.\n." ;
	}

my ($cell, $start_cell, $end_cell) = $self->CanonizeAddress($address) ;

if($self->{DEBUG}{FETCH})
	{
	my $dh = $self->{DEBUG}{ERROR_HANDLE} ;
	print $dh "Fetching '$start_cell-$end_cell'\n" ;
	}
	
# Set the value in the current spreadsheet
if($cell)
	{
	my $value ;
	
	#trigger
	if(exists $self->{DEBUG}{FETCH_TRIGGER}{$start_cell})
		{
		if('CODE' eq ref $self->{DEBUG}{FETCH_TRIGGER}{$start_cell})
			{
			$self->{DEBUG}{FETCH_TRIGGER}{$start_cell}->($self, $start_cell) ;
			}
		else
			{
			if(exists $self->{DEBUG}{FETCH_TRIGGER_HANDLER})
				{
				$self->{DEBUG}{FETCH_TRIGGER_HANDLER}->($self, $start_cell) ;
				}
			else
				{
				my $value_text  ;
				if(exists $self->{DATA}{$start_cell})
					{
					$value_text = defined $self->{DATA}{$start_cell} ? "$self->{DATA}{$start_cell}{VALUE}" : 'undef' ;
					}
				else
					{
					$value_text    = "cell doesn't exist!" unless exists $self->{DATA}{$start_cell} ;
					}
				
				my $dh = $self->{DEBUG}{ERROR_HANDLE} ;
				print $dh "Fetching cell '$start_cell' => $value_text\n" ;
				}
			}
		}
		
	if(exists $self->{DATA}{$start_cell})
		{
		my $current_cell = $self->{DATA}{$start_cell} ;
		
		if($self->{DEBUG}{FETCHED})
			{
			$current_cell->{FETCHED}++ ;
			}
			
		my $caller ;
		
		if(exists $current_cell->{CYCLIC_FLAG})
			{
			my $dh = $self->{DEBUG}{ERROR_HANDLE} ;
			
			push @{$self->{DEPENDENT_STACK}}, [$self, $start_cell, , $self->GetName()] ;
			print $dh $self->DumpDependentStack() ;
			
			#~ confess "Found cyclic dependencies!" ;
			die "Found cyclic dependencies!" ;
			}
		else
			{
			$current_cell->{CYCLIC_FLAG}++ ;
			}
		
		$self->FindDependent($current_cell, $start_cell) ;
		push @{$self->{DEPENDENT_STACK}}, [$self, $start_cell, , $self->GetName()] ;
		
		if($self->{DEBUG}{DEPENDENT_STACK})
			{
			my $dh = $self->{DEBUG}{ERROR_HANDLE} ;
			print $dh $self->DumpDependentStack() ;
			}
			
		if(exists $current_cell->{FORMULA} && ! exists $current_cell->{SUB})
			{
			my $formula = $current_cell->{FORMULA} ;
			
			$current_cell->{NEED_UPDATE} = 1 ;
			($current_cell->{SUB}, $current_cell->{GENERATED_FORMULA}) = GenerateFormulaSub
														(
														  $self
														, $address
														, $address
														, $formula->[0]
														, (@$formula)[1 .. (@$formula - 1)]
														) ;
			}
			
		if(exists $current_cell->{SUB})
			{
			if($self->{AUTOCALC})
				{
				if($current_cell->{NEED_UPDATE})
					{
					if($self->{DEBUG}{SUB})
						{
						my $dh = $self->{DEBUG}{ERROR_HANDLE} ;
						my $ss_name = $self->GetName() ;
						
						print $dh "Running Sub @ '$ss_name!$start_cell'" ;
						
						if(exists $current_cell->{FORMULA})
							{
							print $dh " formula: @{$current_cell->{FORMULA}}" ;
							}
							
						print $dh " defined at '@{$current_cell->{DEFINED_AT}}}'" if(exists $current_cell->{DEFINED_AT}) ;
						print $dh "\n" ;
						}
						
					my $args ;
					if(exists $current_cell->{SUB_ARGS} && @{$current_cell->{SUB_ARGS}})
						{
						$args = $current_cell->{SUB_ARGS} ;
						}
					else
						{
						$args = [] ;
						}
						
					$value = ($current_cell->{SUB})->($self, $start_cell, @$args) ;
					
					# handle caching
					if(exists $current_cell->{CACHE} && (! $current_cell->{CACHE}))
						{
						delete $current_cell->{VALUE} ;
						}
					else
						{
						$current_cell->{VALUE} = $value ;
						$current_cell->{NEED_UPDATE} = 0 ;
						}
					}
				else
					{
					$value = $current_cell->{VALUE} ;
					}
				}
			else
				{
				if($current_cell->{NEED_UPDATE})
					{
					$value = $self->{NEED_UPDATE_MESSAGE} ;
					}
				else
					{
					#handle cache
					if(exists $current_cell->{CACHE} && (! $current_cell->{CACHE}))
						{
						$value = $self->{NEED_UPDATE_MESSAGE} ;
						}
					else
						{
						$value = $current_cell->{VALUE} ;
						}
					}
				}
			}
		else
			{
			if(exists $current_cell->{VALUE})
				{
				$value = $current_cell->{VALUE} ;
				}
			else
				{
				$value = undef ;
				}
			}
			
		pop @{$self->{DEPENDENT_STACK}} ;
		delete $current_cell->{CYCLIC_FLAG} ;
		}
	else
		{
		$value = undef ;

		if(exists $self->{DEPENDENT_STACK} && @{$self->{DEPENDENT_STACK}})
			{
			$self->{DATA}{$start_cell} = {} ; # create the cell to hold the dependent
			$self->FindDependent($self->{DATA}{$start_cell}, $start_cell) ;
			}
		
		if($self->{DEBUG}{DEPENDENT_STACK})
			{
			push @{$self->{DEPENDENT_STACK}}, [$self, $start_cell, , $self->GetName()] ;
			
			my $dh = $self->{DEBUG}{ERROR_HANDLE} ;
			print $dh $self->DumpDependentStack() ;
			
			pop @{$self->{DEPENDENT_STACK}} ;
			}
		}
		
	return($value) ;
	}
else
	{
	my $values ;
	for my $current_address ($self->GetAddressList($address))
		{
		push @$values, $self->Get($current_address) ;
		}
		
	return($values) ;
	}
}

*Get = \&FETCH ;

sub FindDependent
{
my ($self, $current_cell, $start_cell) = @_ ;

if(exists $self->{DEPENDENT_STACK} && @{$self->{DEPENDENT_STACK}})
	{
	my $dependent = @{$self->{DEPENDENT_STACK}}[-1] ;
	my ($spreadsheet, $cell_name) = @$dependent ;
	my $dependent_name = "$spreadsheet, $cell_name" ;
	
	if($self->{DEBUG}{DEPENDENT})
		{
		$current_cell->{DEPENDENT}{$dependent_name}{DATA} = $dependent ;
		$current_cell->{DEPENDENT}{$dependent_name}{COUNT}++ ;
		$current_cell->{DEPENDENT}{$dependent_name}{FORMULA} = $spreadsheet->{DATA}{$cell_name}{FORMULA} ;
		}
	else
		{
		$current_cell->{DEPENDENT}{$dependent_name}{DATA} = $dependent ;
		}
	}
}

#-------------------------------------------------------------------------------

sub STORE 
{
my $self    = shift ;
my $address = shift ;
my $value   = shift ;

# inter spreadsheets references
my $original_address = $address ;
my $ss_reference ;
($ss_reference, $address) = $self->GetSpreadsheetReference($address) ;

if(defined $ss_reference)
	{
	if($ss_reference == $self)
		{
		# fine, it's us
		}
	else
		{
		if($self->{DEBUG}{REDIRECTION})
			{
			my $dh = $self->{DEBUG}{ERROR_HANDLE} ;
			print $dh $self->GetName() . " Store redirected to spreadsheet '$original_address'.\n" ;
			}
			
		return($ss_reference->Set($address, $value)) ;
		}
	}
else
	{
	confess "Can't find Spreadsheet object for address '$address'.\n." ;
	}
	
if($self->{DEBUG}{STORE})
	{
	my $dh = $self->{DEBUG}{ERROR_HANDLE} ;
	print $dh "Storing '$address'\n" ;
	}
	
# Set the value in the current spreadsheet
for my $current_address ($self->GetAddressList($address))
	{
	unless(exists $self->{DATA}{$current_address})
		{
		$self->{DATA}{$current_address} = {} ;
		}
	
	my $current_cell = $self->{DATA}{$current_address} ;
	
	if($self->{DEBUG}{STORED})
		{
		$current_cell->{STORED}++ ;
		}
		
	# triggers
	if(exists $self->{DEBUG}{STORE_TRIGGER}{$current_address})
		{
		if('CODE' eq ref $self->{DEBUG}{STORE_TRIGGER}{$current_address})
			{
			$self->{DEBUG}{STORE_TRIGGER}{$current_address}->($self, $current_address, $value) ;
			}
		else
			{
			if(exists $self->{DEBUG}{STORE_TRIGGER_HANDLER})
				{
				$self->{DEBUG}{STORE_TRIGGER_HANDLER}->($self, $current_address, $value) ;
				}
			else
				{
				my $dh = $self->{DEBUG}{ERROR_HANDLE} ;
				my $value_text = "$value" if defined $value ;
				$value_text    = 'undef' unless defined $value ;
				print $dh "Storing cell '$current_address' => $value_text\n" ;
				}
			}
		}
		
	my $value_is_valid = 1 ;
	#~ my @row_validators ;
	#~ my @column_validators ;
	my $cell_validators = $current_cell->{VALIDATORS} if(exists $current_cell->{VALIDATORS}) ;
	
	for my $validator_data (@{$self->{VALIDATORS}}, @$cell_validators)
		{
		if(0 == $validator_data->[1]($self, $current_address, $current_cell, $value))
			{
			$value_is_valid = 0 ;
			last ;
			}
		}
	
	if($value_is_valid)
		{
		$self->MarkDependentForUpdate($current_cell) ;
		$current_cell->{DEFINED_AT} = [caller] if(exists $self->{DEBUG}{DEFINED_AT}) ;
		
		for (ref $value)
			{
			/^Spreadsheet::Perl::Cache$/ && do
				{
				$current_cell->{CACHE} = $$value ;
				last ;
				} ;
				
			/^Spreadsheet::Perl::Formula$/ && do
				{
				delete $current_cell->{VALUE} ;
				
				$current_cell->{SUB_ARGS}      = [(@$value)[1 .. (@$value - 1)]] ;
				$current_cell->{FORMULA}       = $value ;
				$current_cell->{NEED_UPDATE}   = 1 ;
				$current_cell->{ANCHOR}        = $address ;
				($current_cell->{SUB}, , $current_cell->{GENERATED_FORMULA}) = GenerateFormulaSub
																(
																  $self
																, $current_address
																, $address #anchor
																, $value->[0] # formula
																) ;
				last ;
				} ;
				
			/^Spreadsheet::Perl::Format::Add$/ && do
				{
				$current_cell->{FORMAT} = {%{$current_cell->{FORMAT}}, @$value} ;
				last ;
				} ;
			
			/^Spreadsheet::Perl::Format::Set$/ && do
				{
				$current_cell->{FORMAT} = {@$value} ;
				last ;
				} ;
			
			/^Spreadsheet::Perl::Validator::Add$/ && do
				{
				push @{$current_cell->{VALIDATORS}}, [$value->[0], $value->[1]] ;
				last ;
				} ;
			
			/^Spreadsheet::Perl::Validator::Set$/ && do
				{
				$current_cell->{VALIDATORS} = [[$value->[0], $value->[1]]] ;
				last ;
				} ;
			
			/^Spreadsheet::Perl::Function$/ && do
				{
				delete $current_cell->{VALUE} ;
				
				$current_cell->{SUB}         = $value->[0] ;
				$current_cell->{SUB_ARGS}    = [ @$value[1 .. (@$value - 1)] ] ;
				$current_cell->{NEED_UPDATE} = 1 ;
				last ;
				} ;
				
			/^Spreadsheet::Perl::UserData$/ && do
				{
				$current_cell->{USER_DATA} = {@$value} ;
				last
				} ;
				
			# setting a value:
			delete $current_cell->{FORMULA} ;
			delete $current_cell->{NEED_UPDATE} ; 
			delete $current_cell->{CACHE} ; 
			delete $current_cell->{SUB} ;
			delete $current_cell->{SUB_ARGS} ;
			delete $current_cell->{ANCHOR} ;
			# fall through!!!!
			
			/^Spreadsheet::Perl::RangeValues$/ && do
				{
				$current_cell->{VALUE} = shift @$value ;
				last
				} ;
				
			/^Spreadsheet::Perl::RangeValuesSub$/ && do
				{
				$current_cell->{VALUE} = $value->[0]($self, $address, $current_address, @$value[1 .. (@$value - 1)]) ;
				last
				} ;
				
			# DEFAULT:
				$current_cell->{VALUE} = $value ;
			}
		}
	else
		{
		# not validated
		}
	}
}

*Set = \&STORE ;

sub MarkDependentForUpdate
{
my ($self, $current_cell) = @_ ;

return unless exists $current_cell->{DEPENDENT} ;

for my $dependent_name (keys %{$current_cell->{DEPENDENT}})
	{
	my $dependent = $current_cell->{DEPENDENT}{$dependent_name}{DATA} ;
	my ($spreadsheet, $cell_name) = @$dependent ;
	
	if(exists $spreadsheet->{DATA}{$cell_name})
		{
		if(exists $spreadsheet->{DATA}{$cell_name}{SUB})
			{
			$spreadsheet->{DATA}{$cell_name}{NEED_UPDATE}++ ;
			}
		else
			{
			delete $current_cell->{DEPENDENT}{$dependent_name} ;
			}
		}
	else
		{
		delete $current_cell->{DEPENDENT}{$dependent_name} ;
		}
	}
}

#-------------------------------------------------------------------------------

sub DELETE   
{
my $self    = shift ;
my $address = shift ;

for my $current_address ($self->GetAddressList($address))
	{
	delete $self->{DATA}{$current_address} ;
	}
}

sub CLEAR 
{
my $self    = shift ;
my $address = shift ;

delete $self->{DATA} ;
}

sub EXISTS   
{
my $self    = shift ;
my $address = shift ;

for my $current_address ($self->GetAddressList($address))
	{
	unless(exists $self->{DATA}{$current_address})
		{
		return(0) ;
		}
	}
	
return(1) ;
}

sub FIRSTKEY 
{
my $self = shift ;
scalar(keys %{$self->{DATA}}) ;

return scalar each %{$self->{DATA}} ;
}

sub NEXTKEY  
{
my $self = shift;
return scalar each %{ $self->{DATA} }
}

sub DESTROY  
{
}

#-------------------------------------------------------------------------------

sub LockValidator
{
my $self    = shift ;
my $address = shift ;
my $cell    = shift ;
my $value   = shift ;

if($self->IsCellLocked($address))
	{
	carp "While setting '$address': Lock is active" ;
		
	return(0) ;
	}
else
	{
	return(1) ;
	}


}

#-------------------------------------------------------------------------------

1 ;

__END__
=head1 NAME

Spreadsheet::Perl - Pure Perl implementation of a spreadsheet

=head1 SYNOPSIS

  use Spreadsheet::Perl;
  use Spreadsheet::Perl::Formula ;
  use Spreadsheet::Perl::Format ;
  ...

  tie my %ss, "Spreadsheet::Perl"
  my $ss = tied %ss ;

  $ss->SetRangeName("TestRange", 'A5:B8') ;
  $ss{TestRange} = '7' ;
  
  $ss->DefineFunction('AddOne', \&AddOne) ;
  
  $ss{A3} = Spreadsheet::Perl::Formula('$ss->AddOne("A5") + $ss{A5}') ;
  print "A3 formula => " . $ss->GetFormulaText('A3') . "\n" ;
  print "A3 = $ss{A3}\n" ;

  $ss{'ABC1:ABD5'} = '10' ;

  $ss{A4} = Spreadsheet::Perl::Formula('$ss->Sum("A5:B8", "ABC1:ABD5")') ;
  print "A4 = $ss{A4}\n" ;
  
  ...

=head1 DESCRIPTION

Spreadsheet::Perl is a pure Perl implementation of a spreadsheet. 
I you have an application that takes some input and does calculation on them, chances
are that implementing it through a spreadsheet will make it more maintainable and easier to develop. I found
no spreadsheet modules on CPAN (programmers tend to think a spreadsheet is not a programming tool). The idea that it would be
very easy to implement in perl kept going round in my head. I put the limit of 500 lines of code for a functional spreadsheeet.
It took a few days to get something viable and it was just under 5OO lines. When debuggin help kicked in the module became a bit bigger,
still Spreadsheet::Perl is quite minimal in size and can do the the folowwing:

- set and get values from cells or ranges
- cell private data
- cell/range fillers/"wizards"
- set formulas (pure perl)
- compute the dependencies between cells 
- formulas can fetch data from multiple spreadsheets and the dependencies still work
- checks for circular dependencies
- debugging triggers
- has a simple architecture for expansion
- has a simple architecture for debugging (and some flags are already implemented)
- can read it's data from a file
- supports cell naming
- cell and range locking
- input validators
- cell formats (pod, html, ...)
- can define spreadsheet functions from the scripts using it or via a new module of your own
- AUTOCALC ON/OFF, Recalculate()
- value caching to speed up formulas and 'volatile' cells
- cell address offseting functions
- Automatic formula offseting
- Relative and fixed cell addresses
- slice access
- cells can be assigned a sub reference and be non-caching. this could be used for reading from a database
- split in many small modules so you pay only for what you need
- some debugging tool (dump, formula stack trace, ...)

All this under 1500 lines of code, the biggest module being just under 700 lines. Perl rocks.

Look at the 'tests' directory for some examples.
=head1 TODO

Unfortunately there is still a lot to do (the basics are there) and I have the feeling I will not get the time needed.
If someone is willing to help or take over, I'll be glad to step aside.

Here are some of the things that I find missing, this doesn't mean all are good ideas:
- documentation, test (working on it)
- perl debugger support à la PBS
- Row/column/spreadsheet default values.
- R1C1 Referencing
- database interface (a handfull of functions at most)
- WWW interface
- Arithmetic functions (only Sum is implemented), statistic functions
- printing, exporting
- importing from other spreadsheets
- more serious file reading and file writting
- complex stuff (fixing one fixes the other)
	- Insertion of rows and columns
	- Deletion of rows and columns
	- Sorting
- a gui (curses, tk, wxWindows) would be great!
- a nice logo :-)

Some stuff is available on CPAN, just some glue is needed.

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

B<Data::TreeDumper> is used if found (I recommend installing it to get nice dumps).

=cut
