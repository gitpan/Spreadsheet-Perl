
use Carp ;
use strict ;
use warnings ;

use Spreadsheet::Perl ;
use Spreadsheet::Perl::QuerySet ;
use Spreadsheet::Perl::Devel ;
use Spreadsheet::Perl::RangeValues ;

tie my %ss, "Spreadsheet::Perl" ;
my $ss = tied %ss ;

$ss{'A1:A5'} = Spreadsheet::Perl::RangeValues(reverse 1 .. 10) ;
print $ss->Dump(['A1:A5']) ;

$ss{'A1:A5'} = Spreadsheet::Perl::RangeValuesSub(\&Filler, [11, 22, 33]) ;
print $ss->Dump(['A1:A5'], 0) ;

@ss{'A1', 'B1:C2', 'A8'} = ('A', Spreadsheet::Perl::RangeValues(reverse 1 .. 10), -1) ;
print $ss->Dump() ;

@ss{'A1', 'B1:C2', 'A8'} = ('A', 'B', 'C');
print $ss->Dump() ;

#~ use Devel::Size::Report qw/report_size/;
#~ print report_size($ss, { indend => "    " } );


#-------------------------------------------------------------

sub Filler 
{
my ($ss, $anchor, $current_address, $list, @other_args) = @_ ;
return(shift @$list) ;
}

