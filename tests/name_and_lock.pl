
use Carp ;
use strict ;
use warnings ;

use Spreadsheet::Perl ;

tie my %ss, "Spreadsheet::Perl" ;
my $ss = tied %ss ;
@ss{'A1', 'A2'} = ('cell A1', 'cell A2') ;

$ss->SetCellName("First", "1,1") ;
print  $ss{First} . ' ' . $ss{A2} . "\n" ;

$ss->SetCellName("first", "A1") ;
print  $ss{first} . ' ' . $ss{A2} . "\n" ;

$ss->SetRangeName("first_range", "A1:A2") ;
print  "First range: @{$ss{first_range}}\n" ;

$ss->Lock(1) ;
$ss{A1} = 'ho' ;
$ss->Lock(0) ;

$ss{"1,1"} = "cell A1(m)" ; # numeric indexing is also possible

print  $ss{First} . ' ' . $ss{A2} . "\n" ;

$ss->LockRange("A1:B1", 1) ;
$ss{A1} = 'ho' ;
$ss{C1} = 'ho' ; # not locked

$ss->SetRangeName("TestRange", 'A1:B5') ;
$ss{TestRange} = '7' ;

print $ss->Dump() ;
