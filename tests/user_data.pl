
use Carp ;
use strict ;
use warnings ;

use Spreadsheet::Perl ;
use Spreadsheet::Perl::UserData ;

tie my %ss, "Spreadsheet::Perl" ;
my $ss = tied %ss ;

$ss{'A1:A8'} = '10' ;
$ss{'A1:A8'} = Spreadsheet::Perl::UserData(NAME => 'private data', ARRAY => ['hi']) ;

print "@{$ss{'A1:A8'}}\n" ;

print $ss->Dump() ;
