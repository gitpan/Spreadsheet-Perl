
use Carp ;
use strict ;
use warnings ;

use Spreadsheet::Perl ;
use Spreadsheet::Perl::Arithmetic ;

#---------------------------------------------------------------------------------

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
						
						#~ FETCH_SUB => \&DoublePrevious
						#~ , FETCH_SUB_ARGS => [ 1, 2, 3]
						
						#~ FORMULA => '$ss{A1}'
						}
				} ;

my $ss = tied %ss ;

print  $ss{A1} . ' ' . $ss{A2} . "\n" ;

#---------------------------------------------------------------------------------

tie %ss, "Spreadsheet::Perl"
		, CELLS =>
				{
				  A1 =>
						{
						VALUE => 1
						}
					
				, A2 =>
						{
						  FETCH_SUB => \&DoublePrevious
						, FETCH_SUB_ARGS => [ 1, 2, 3]
						
						#~ FORMULA => '$ss{A1}'
						}
				} ;

$ss = tied %ss ;
print $ss->Dump() ;

print  $ss{A1} . ' ' . $ss{A2} . "\n" ;

print $ss->Dump() ;

sub DoublePrevious
{
my $ss = shift ;
my $address  = shift ;

my ($x, $y) = ConvertAdressToNumeric($address) ;
my $cell_value = $ss->Get("$x," . ($y - 1)) ;

return($cell_value * 2) ;
}

#---------------------------------------------------------------------------------

tie %ss, "Spreadsheet::Perl"
		, CELLS =>
				{
				  A1 =>
						{
						VALUE => 'hi'
						}
					
				, A2 =>
						{
						  VALUE => 'there'
						, FORMULA => ['$ss{A1}']
						}
				} ;

$ss = tied %ss ;

print  $ss{A1} . ' ' . $ss{A2} . "\n" ;

#---------------------------------------------------------------------------------
# error that we can't cach as of 0.04, we don't differentiate between formula 
# generated FETCH_SUB anf sub comming from setup
#---------------------------------------------------------------------------------

tie %ss, "Spreadsheet::Perl"
		, CELLS =>
				{
				  A1 =>
						{
						VALUE => 'hi'
						}
					
				, A2 =>
						{
						VALUE => 'there'
						
						, FETCH_SUB => \&DoublePrevious
						, FETCH_SUB_ARGS => [ 1, 2, 3]
						, FORMULA => '$ss{A1}'
						}
				} ;

$ss = tied %ss ;

print  $ss{A1} . ' ' . $ss{A2} . "\n" ;


#---------------------------------------------------------------------------------

%ss = do "ss_setup.pl" or confess("Couldn't evaluate setup file 'ss_setup.pl'\n");

print $ss->Dump() ;
$ss->GenerateHtmlToFile('setup_do.html') ;

#---------------------------------------------------------------------------------
