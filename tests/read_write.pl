
use Carp ;
use strict ;
use warnings ;

use Spreadsheet::Perl ;

my $ss = tie my %ss, "Spreadsheet::Perl" ;

$ss->Read('ss_data.pl') ;

$ss{A3} = PF('$ss{FIRST_CELL}') ;

print $ss->DumpTable() ;
#print $ss->Dump(undef, undef, {USE_ASCII => 1}) ;

$ss->Write('generated_ss_data.pl') ;

undef $ss ;
%ss = () ;
untie %ss ;

$ss = tie %ss, "Spreadsheet::Perl" ;
$ss->Read('generated_ss_data.pl') ;

print $ss->DumpTable() ;
#print $ss->Dump(undef, undef, {USE_ASCII => 1}) ;

for (sort keys %Spreadsheet::Perl::defined_functions)
	{
	print "Found function '$_' in the spreadsheet.\n" ;
	}
	
