
use Carp ;
use strict ;
use warnings ;

use Spreadsheet::Perl ;

tie my %ss, "Spreadsheet::Perl", NAME=> 'TEST' ;
my $ss = tied %ss ;

$ss->{DEBUG}{DEPENDENT}++ ;
$ss->{DEBUG}{FETCH}++ ;
$ss->{DEBUG}{STORE}++ ;
$ss->{DEBUG}{SUB}++ ;

#~ # cyclic error
$ss->{DEBUG}{DEFINED_AT}++ ;
$ss{'A1:A5'} = Formula('$ss{"A2"}') ;
$ss{A6} = Formula('$ss{A1}') ;
print "$ss{A1}\n" ;

