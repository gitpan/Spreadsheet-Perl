
use Data::Dumper ;
my $complicated_value = Dumper([{1 => [], 2 => {hi => there}}]) ;

sub OneMillion
{
return(1_000_000) ;
}

#-----------------------------------------------------------------
# the spreadsheet data
#-----------------------------------------------------------------

A1 => 120, 
A2 => sub{1},
A3 => Spreadsheet::Perl::Formula('$ss->Sum("A1:A2")'),

B1 => 3,

c2 => "hi there",
c3 => $complicated_value,

D1 => OneMillion()