
use Carp ;
use strict ;
use warnings ;

use Spreadsheet::Perl ;
use Spreadsheet::Perl::Arithmetic ;

tie my %ss, "Spreadsheet::Perl", NAME => 'TEST' ;
my $ss = tied %ss ;

$ss{A9} = Formula('$ss->Sum("A1:A8") + 100 ') ;
print "$ss{A9}\n" ;

$ss{'A1:A8'} = RangeValues(1 .. 8) ;
print $ss->Dump() ; # show formula dependencies
print "$ss{A9}\n" ;

$ss{A10} = Formula('"$cell => " . (join "-", (ConvertAdressToNumeric($cell)))') ;
print "'A10' Self: " . $ss{A10} . "\n" ;
