
use Carp ;
use strict ;
use warnings ;

use Spreadsheet::Perl ;

tie my %ss, "Spreadsheet::Perl", NAME => 'TEST' ;
my $ss = tied %ss ;

$ss{A1} = 1 ;
$ss{A1} = Format(ANSI => {HEADER => "blink"}) ;
$ss{A1} = Formula('$ss{A5} * 5') ;

use Data::TreeDumper ;
print DumpTree($ss{'A1.FORMAT'}, 'format:', USE_ASCII => 0) ;
print DumpTree($ss{'A1.FORMULA'}, 'formula:') ;