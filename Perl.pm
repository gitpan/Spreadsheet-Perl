
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

#~ our @EXPORT = qw( Reset ) ;
our @EXPORT ;
push @EXPORT, qw( Reset ) ;

our $VERSION = '0.03' ;

use Spreadsheet::Perl::Address ;
use Spreadsheet::Perl::Cache ;
use Spreadsheet::Perl::Devel ;
use Spreadsheet::Perl::Format ;
use Spreadsheet::Perl::Formula ;
use Spreadsheet::Perl::Function ;
use Spreadsheet::Perl::Lock ;
use Spreadsheet::Perl::QuerySet ;
use Spreadsheet::Perl::RangeValues ;
use Spreadsheet::Perl::UserData ;
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
				
	, CELLS                => {}
	) ;
}

sub Reset
{
my $self = shift ;
my $setup = shift ;
my $cell_data  = shift ;

if(defined $setup)
	{
	if('HASH' eq ref $setup)
		{
		confess "Setup data must be a hash reference!" 
		}
	else
		{
		%$self = (GetDefaultData(), %$setup) ;
		}
	}
	
if(defined $cell_data)
	{
	confess "cell data must be a hash reference!" unless 'HASH' eq ref $cell_data ;
	$self->{CELLS} = $cell_data ;
	}
else
	{
	$self->{CELLS} = {} ;
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

my $attribute ;

if($address =~ /(.*)\.(.+)/)
	{
	$address = $1 ;
	$attribute = $2 ;
	}
	
my ($cell, $start_cell, $end_cell) = $self->CanonizeAddress($address) ;

if($self->{DEBUG}{FETCH})
	{
	my $dh = $self->{DEBUG}{ERROR_HANDLE} ;
	print $dh "Fetching '$start_cell:$end_cell'\n" ;
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
			$self->{DEBUG}{FETCH_TRIGGER}{$start_cell}->($self, $start_cell, $attribute) ;
			}
		else
			{
			if(exists $self->{DEBUG}{FETCH_TRIGGER_HANDLER})
				{
				$self->{DEBUG}{FETCH_TRIGGER_HANDLER}->($self, $start_cell, $attribute) ;
				}
			else
				{
				my $value_text  ;
				if(exists $self->{CELLS}{$start_cell})
					{
					$value_text = defined $self->{CELLS}{$start_cell} ? "$self->{CELLS}{$start_cell}{VALUE}" : 'undef' ;
					}
				else
					{
					$value_text    = "cell doesn't exist!" unless exists $self->{CELLS}{$start_cell} ;
					}
				
				my $dh = $self->{DEBUG}{ERROR_HANDLE} ;
				print $dh "Fetching cell '$start_cell' => $value_text\n" ;
				}
			}
		}
		
	if(exists $self->{CELLS}{$start_cell})
		{
		my $current_cell = $self->{CELLS}{$start_cell} ;
		
		if(defined $attribute)
			{
			if(exists $current_cell->{$attribute})
				{
				$value = $current_cell->{$attribute} ;
				}
			else
				{
				$value = undef ;
				}
			}
		else
			{
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
				
				die "Found cyclic dependencies!" ;
				}
			else
				{
				$current_cell->{CYCLIC_FLAG}++ ;
				}
			
			$self->FindDependent($current_cell, $start_cell) ;
			push @{$self->{DEPENDENT_STACK}}, [$self, $start_cell, , $self->GetName()] ;
			
			if($self->{DEBUG}{DEPENDENT_STACK}) #! TODO: dump stack on specific cells
				{
				my $dh = $self->{DEBUG}{ERROR_HANDLE} ;
				print $dh $self->DumpDependentStack() ;
				}
				
			if(exists $current_cell->{FORMULA} && ! exists $current_cell->{FETCH_SUB})
				{
				my $formula = $current_cell->{FORMULA} ;
				
				$current_cell->{NEED_UPDATE} = 1 ;
				($current_cell->{FETCH_SUB}, $current_cell->{GENERATED_FORMULA}) = GenerateFormulaSub
															(
															  $self
															, $address
															, $address
															, $formula->[0]
															, (@$formula)[1 .. (@$formula - 1)]
															) ;
				}
			
			if(exists $current_cell->{FETCH_SUB})
				{
				if($self->{AUTOCALC})
					{
					if($current_cell->{NEED_UPDATE} || ! exists $current_cell->{NEED_UPDATE})
						{
						if($self->{DEBUG}{FETCH_SUB})
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
							
						if(exists $current_cell->{FETCH_SUB_ARGS} && @{$current_cell->{FETCH_SUB_ARGS}})
							{
							$value = ($current_cell->{FETCH_SUB})->($self, $start_cell, @{$current_cell->{FETCH_SUB_ARGS}}) ;
							}
						else
							{
							$value = ($current_cell->{FETCH_SUB})->($self, $start_cell) ;
							}
							
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
		}
	else
		{
		$value = undef ;
			
		if(exists $self->{DEPENDENT_STACK} && @{$self->{DEPENDENT_STACK}})
			{
			$self->{CELLS}{$start_cell} = {} ; # create the cell to hold the dependent
			$self->FindDependent($self->{CELLS}{$start_cell}, $start_cell) ;
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
	my @values ;
	for my $current_address ($self->GetAddressList($address))
		{
		push @values, $self->Get($current_address) ;
		}
		
	return(\@values) ;
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
		$current_cell->{DEPENDENT}{$dependent_name}{CELLS} = $dependent ;
		$current_cell->{DEPENDENT}{$dependent_name}{COUNT}++ ;
		$current_cell->{DEPENDENT}{$dependent_name}{FORMULA} = $spreadsheet->{CELLS}{$cell_name}{FORMULA} ;
		}
	else
		{
		$current_cell->{DEPENDENT}{$dependent_name}{CELLS} = $dependent ;
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
	unless(exists $self->{CELLS}{$current_address})
		{
		$self->{CELLS}{$current_address} = {} ;
		}
	
	my $current_cell = $self->{CELLS}{$current_address} ;
	
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
				
				$current_cell->{FETCH_SUB_ARGS}      = [(@$value)[1 .. (@$value - 1)]] ;
				$current_cell->{FORMULA}       = $value ;
				$current_cell->{NEED_UPDATE}   = 1 ;
				$current_cell->{ANCHOR}        = $address ;
				($current_cell->{FETCH_SUB}, , $current_cell->{GENERATED_FORMULA}) = GenerateFormulaSub
																(
																  $self
																, $current_address
																, $address #anchor
																, $value->[0] # formula
																) ;
				last ;
				} ;
				
			/^Spreadsheet::Perl::Format$/ && do
				{
				@{$current_cell->{FORMAT}}{keys %$value} = values %$value ;
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
			
			/^Spreadsheet::Perl::FetchFunction$/ && do
				{
				delete $current_cell->{VALUE} ;
				
				$current_cell->{FETCH_SUB}         = $value->[0] ;
				$current_cell->{FETCH_SUB_ARGS}    = [ @$value[1 .. (@$value - 1)] ] ;
				$current_cell->{NEED_UPDATE} = 1 ;
				last ;
				} ;
				
			/^Spreadsheet::Perl::StoreFunction$/ && do
				{
				delete $current_cell->{VALUE} ;
				
				$current_cell->{STORE_SUB}         = $value->[0] ;
				$current_cell->{STORE_SUB_ARGS}    = [ @$value[1 .. (@$value - 1)] ] ;
				last ;
				} ;
				
			/^Spreadsheet::Perl::UserData$/ && do
				{
				$current_cell->{USER_DATA} = {@$value} ;
				last
				} ;
			#----------------------
			# setting a value:
			#----------------------
			delete $current_cell->{FORMULA} ;
			delete $current_cell->{NEED_UPDATE} ; 
			delete $current_cell->{CACHE} ; 
			delete $current_cell->{FETCH_SUB} ;
			delete $current_cell->{FETCH_SUB_ARGS} ;
			delete $current_cell->{ANCHOR} ;
			
			my $value_to_store = $value ; # do not modify $value as it is used again when storing ranges
			
			# check for range fillers
			if(/^Spreadsheet::Perl::RangeValues$/)
				{
				$value_to_store  = shift @$value  ;
				}
			else
				{
				if(/^Spreadsheet::Perl::RangeValuesSub$/)
					{
					$value_to_store = $value->[0]($self, $address, $current_address, @$value[1 .. (@$value - 1)]) ;
					}
				#else
					# store the value passed to STORE
				}
			
			if(exists $current_cell->{STORE_SUB})
				{
				if(exists $current_cell->{STORE_SUB_ARGS} && @{$current_cell->{STORE_SUB_ARGS}})
					{
					$current_cell->{STORE_SUB}->($self, $current_address, $value_to_store, @{$current_cell->{STORE_SUB_ARGS}}) ;
					}
				else
					{
					$current_cell->{STORE_SUB}->($self, $current_address, $value_to_store) ;
					}
				}
			else
				{
				$current_cell->{VALUE} = $value_to_store ;
				}
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
	my $dependent = $current_cell->{DEPENDENT}{$dependent_name}{CELLS} ;
	my ($spreadsheet, $cell_name) = @$dependent ;
	
	if(exists $spreadsheet->{CELLS}{$cell_name})
		{
		if(exists $spreadsheet->{CELLS}{$cell_name}{FETCH_SUB})
			{
			$spreadsheet->{CELLS}{$cell_name}{NEED_UPDATE}++ ;
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
	delete $self->{CELLS}{$current_address} ;
	}
}

sub CLEAR 
{
my $self    = shift ;
my $address = shift ;

delete $self->{CELLS} ;
}

sub EXISTS   
{
my $self    = shift ;
my $address = shift ;

for my $current_address ($self->GetAddressList($address))
	{
	unless(exists $self->{CELLS}{$current_address})
		{
		return(0) ;
		}
	}
	
return(1) ;
}

sub FIRSTKEY 
{
my $self = shift ;
scalar(keys %{$self->{CELLS}}) ;

return scalar each %{$self->{CELLS}} ;
}

sub NEXTKEY  
{
my $self = shift;
return scalar each %{ $self->{CELLS} }
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

if($self->{LOCKED})
	{
	carp "While setting '$address': Spreadsheet lock is active" ;
	return(0) ;
	}
else
	{
	if($self->IsCellLocked($address))
		{
		carp "While setting '$address': Cell lock is active" ;
			
		return(0) ;
		}
	else
		{
		return(1) ;
		}
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

Spreadsheet::Perl is minimal in size but can do the the folowwing:

=over 2

=item * set and get values from cells or ranges

=item * cell private data

=item * fetch/store callback

=item * cell attributes access

=item * cell/range fillers/"wizards"

=item * set formulas (pure perl)

=item * compute the dependencies between cells 

=item * formulas can fetch data from multiple spreadsheets and the dependencies still work

=item * checks for circular dependencies

=item * debugging triggers

=item * has a simple architecture for expansion

=item * has a simple architecture for debugging (and some flags are already implemented)

=item * can read it's data from a file

=item * supports cell naming

=item * cell and range locking

=item * input validators

=item * cell formats (pod, html, ...)

=item * can define spreadsheet functions from the scripts using it or via a new module of your own

=item * AUTOCALC ON/OFF, Recalculate()

=item * value caching to speed up formulas and 'volatile' cells

=item * cell address offseting functions

=item * Automatic formula offseting

=item * Relative and fixed cell addresses

=item * slice access

=item * some debugging tool (dump, formula stack trace, ...)

=back

Look at the 'tests' directory for some examples.

=head1 DRIVING FORCE

=head2 Why

I found no spreadsheet modules on CPAN (I see a spreadsheet as a programming tool). The idea that it would be
very easy to implement in perl kept going round in my head. I put the limit at 500 lines of code for a functional spreadsheeet.
It took a few days to get something viable and it was just under 5OO lines.

I you have an application that takes some input and does calculation on them, chances
are that implementing it through a spreadsheet will make it more maintainable and easier to develop.
Here are the reasons (IMO) why:

=over 2

=item * Spreadsheet programming (SP) is data oriented and this is what programming should be more often.

=item * SP is encapsulating. The processing is "hidden"behind the cell value in form of formulas.

=item * SP is encapsulating II. The data dependencies are automatically computed by the spreadsheet, releaving 
you from keeping things in synch

=item * SP is 2 dimensional (or 3 or 4 four that might not be easier for that), specialy if you have a gui  for it.

=item * If you have a gui, SP is visual programming and visual debugging as the 
spreadsheet is the input and the dump of the data. The possibility to to 
show a multi-dimentional dependency is great as is the fact that you don't 
need to look around for where things are defined (this is more about 
visual programming but still fit spreadsheets as they are often gui based)

=item * SP allows for user customization 

=back

=head2 How

I want B<Spreadsheets::Perl> to:

=over 2

=item * Be Perl, be Perl, be fully Perl

=item * Be easy to develop, I try to implement nothing that is already there

=item * Be easy to expand

=item * Be easy to use for Perl programmers

=back 

=head1 CREATING A SPREADSHEET

Spreadsheet perl is implemented as a tie. Remember that you can use hash slices (I 'll give some examples). The
spreadsheet functions are accessed through the tied object.

=head2 Simple creation

  use Spreadsheet::Perl ;
  tie my %ss, "Spreadsheet::Perl" ; 
  my $ss = tied %ss ; # needed to access the spreadsheet functions.

=head2 Setting up data

=head3 Setting the cell data

  use Spreadsheet::Perl ;
  tie my %ss, "Spreadsheet::Perl"
		, CELLS =>
				{
				  A1 =>
						{
						VALUE => 'hi'
						}
					
				, A2 =>
						{
						VALUE => 'there'
						#~ or
						#~ FORMULA => '$ss{A1}'
						}
				} ;


=head3 Setting the cell data, simple way

  use Spreadsheet::Perl ;
  tie my %ss, "Spreadsheet::Perl"
  @ss{'A1', 'B1:C2', 'A8'} = ('A', 'B', 'C');

=head3 Setting the spreadsheet attributes

  use Spreadsheet::Perl ;
  tie my %ss, "Spreadsheet::Perl"
		  , NAME => 'TEST'
		  , AUTOCALC => 0
		  , DEBUG => { PRINT_FORMULA => 1} ;


=head2 reading data from a file

  <- start  of ss_setup.pl ->
  # how to compute the data
  
  sub OneMillion
  {
  return(1_000_000) ;
  }
  
  #-----------------------------------------------------------------
  # the spreadsheet data
  #-----------------------------------------------------------------
  A1 => 120, 
  A2 => sub{1},
  A3 => Formula('$ss->Sum("A1:A2")'),
  
  B1 => 3,
  
  c2 => "hi there",
  
  D1 => OneMillion()
  
  <- end of ss_setup.pl ->

  use Spreadsheet::Perl ;
  tie my %ss, "Spreadsheet::Perl", NAME => 'TEST' ;
  %ss = do "ss_setup.pl" ;

=head2 dumping a spreadsheet

Use the Dump function (see I<Debugging>):

  my $ss = tied %ss ;
  print $ss->Dump() ;

Generates:
  
  ------------------------------------------------------------
  Spreadsheet::Perl=HASH(0x825540c) 'TEST' [3550 bytes]
  
  
  Cells:
  |- A1
  |  `- VALUE = 120
  |- A2
  |  `- VALUE = CODE(0x82554d8)
  |- A3
  |  |- ANCHOR = A3
  |  |- FETCH_SUB = CODE(0x825702c)
  |  |- FETCH_SUB_ARGS
  |  |- FORMULA = Object of type 'Spreadsheet::Perl::Formula'
  |  |  `- 0 = $ss->Sum("A1:A2")
  |  |- GENERATED_FORMULA = $ss->Sum("A1:A2")
  |  `- NEED_UPDATE = 1
  |- B1
  |  `- VALUE = 3
  |- C2
  |  `- VALUE = hi there
  `- D1
     `- VALUE = 1000000
  
  Spreadsheet::Perl=HASH(0x825540c) 'TEST' dump end
  ------------------------------------------------------------

=head1 CELL and RANGE: ADDRESSING, NAMING

Cells are index  with a scheme I call baseAA1 (please let me know if it has a better name).
The cel address is a combinaison of letters and a figure, ie 'A1', 'BB45', 'ABDE15'.

BaseAA figures match /[A-Z]{1,4}/. see B<Spreadsheet::ConvertAA>. There is no limit on the numeric figure.
Spreadsheet::Perl is implemented as a hash thus allowing for sparse spreadsheets.

=head2 Address format

Adresses are composed of:

=over 2

=item * an optional spreadsheet name and '!'. ex: 'TEST!'

=item * a baseAA1 figure. ex 'A1'

=item * a ':' followed by a baseAA1 figure for ranges. ex: ':A5'

=back

The following are valid addresses: A1 TEST!A1 A1:BB5 TESTA5:CE43

the order of the baseAA figures is important!

  $ss{'A1:D5'} = 7; is equivalent to $ss{'D5:A1'} = 7; 

but

  $ss{'A1:D5'} = Formula('$ss{H10}'); is NOT equivalent to $ss{'D5:A1'} = Formula('$ss{H10}'); 
  
because formulas get recalculated for each cell. Spreadsheet::Perl goes from the first baseAA figure
to the second one by iterating the row, then the column.

it is also possible to index cells with numerals only: $ss{"1,7"}. Remember that A is 1 and there are
no zeros.

=head2 Names
It is possible to give a name to a cell or to a range: 

  tie my %ss, "Spreadsheet::Perl" ;
  my $ss = tied %ss ;
  @ss{'A1', 'A2'} = ('cell A1', 'cell A2') ;
  
  $ss->SetCellName("first", "A1") ;
  print  $ss{first} . ' ' . $ss{A2} . "\n" ;
  
  $ss->SetRangeName("first_range", "A1:A2") ;
  print  "First range: @{$ss{first_range}}\n" ;

=head1 OTHER SPREADSHEET

To use interspreadsheet formulas, you need to make the spreadsheet aware of the other spreadsheets by
calling the I<AddSpreadsheet> function.

  tie my %romeo, "Spreadsheet::Perl", NAME => 'ROMEO' ;
  my $romeo = tied %romeo ;

  tie my %juliette, "Spreadsheet::Perl", NAME => 'JULIETTE' ;
  my $juliette = tied %juliette ;

  $romeo->AddSpreadsheet('JULIETTE', $juliette) ;
  $juliette->AddSpreadsheet('ROMEO', $romeo) ;
  
  $romeo{'B1:B5'} = 10 ;
  
  $juliette{A4} = 5 ;
  $juliette{A5} = Formula('$ss->Sum("JULIETTE!A4") + $ss->Sum("ROMEO!B1:B2")') ; 

=head1 SPREADSHEEET Functions

=head2 Locking

=head2 Calculation control

=head2 State queries and debugging

=head1 SETTING CELLS

Setting and reading cells is done in two diffrent ways. I like the way it looks now but it
might change in the (near) future.

=head2 Formulas

=head3 builtin functions

=head3 cell dependencies

=head3 circular dependencies

=head2 Setting a value

=head3 RangeValues

=head2 Setting a formula

=head3 Caching

=head2 Setting a format

=head2 Setting fetch and store callbacks

=head2 Setting Validators

=head2 Setting User data

=head1 READING CELLS

=head2 Reading values

=head2 Reading internal data

=head2 Reading user data

=head1 Debugging

=head2 Dump

The I<Dump> function, err, dumps the spreadsheet. It takes the following arguments:

=over 2

=item * an address list withing an array reference or undef. ex: ['A1', 'B5:B8']

=item * a boolean. When set, the spreadsheet attributes are displayed

=item * an optional hash reference pased as overrides to B<Data::TreeDumper>

=back

If B<Data::TreeDumper> is not installed, Data::Dumper is used.I exclusively use B<Data::TreeDumper> so 
I never look at the dumps generated through Data::Dumper. It will certainly look ugly or might even be broken.
Install B<Data::TreeDumper>, it's worth it (I've written it so I have to force you to try it :-)

=head1 TODO

Unfortunately there is still a lot to do (the basics are there) and I have the feeling I will not get the time needed.
If someone is willing to help or take over, I'll be glad to step aside.

Here are some of the things that I find missing, this doesn't mean all are good ideas:

=over 2

=item * documentation, test (working on it)

=item * perl debugger support à la PBS

=item * Row/column/spreadsheet default values.

=item * R1C1 Referencing

=item * database interface (a handfull of functions at most)

=item * WWW interface

=item * Arithmetic functions (only Sum is implemented), statistic functions

=item * printing, exporting

=item * importing from other spreadsheets

=item * more serious file reading and file writting

=item * complex stuff (fixing one fixes the other)

=over 4

=item * Insertion of rows and columns

=item * Deletion of rows and columns

=item * Sorting

=back

=item * a gui (curses, tk, wxWindows) would be great!

=item * a nice logo :-)

=back

Lots is available on CPAN, just some glue is needed.

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

