
use Carp ;
use strict ;
use warnings ;

use Spreadsheet::Perl ;

tie my %ss, "Spreadsheet::Perl" ;
my $ss = tied %ss ;
@ss{'A1', 'A2'} = (1 .. 2) ;

$ss->DefineFunction('AddOne', \&AddOne) ;
$ss->DefineFunction('AddOne', \&AddOne) ; # generate a warning

$ss{A3} = Formula('$ss->AddOne("A1") + $ss{A2}') ;
print "A3 => '@{[$ss->GetFormulaText('A3')]}' = $ss{A3}\n" ;

#---------------------------------------------------

sub AddOne
{
my $ss = shift ;
my $address = shift ;

return($ss->Get($address) + 1) ;
}
