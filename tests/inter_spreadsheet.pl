
use Carp ;
use strict ;
use warnings ;

use Spreadsheet::Perl ;
use Spreadsheet::Perl::Arithmetic ;

tie my %romeo, "Spreadsheet::Perl" ;
my $romeo = tied %romeo ;
$romeo->SetName('ROMEO') ;
#~ $romeo->{DEBUG}{SUB}++ ;
#~ $romeo->{DEBUG}{ADDRESS_LIST}++ ;
#~ $romeo->{DEBUG}{FETCH_FROM_OTHER}++ ;
#~ $romeo->{DEBUG}{DEPENDENT_STACK}++ ;
$romeo->{DEBUG}{DEPENDENT}++ ;

tie my %juliette, "Spreadsheet::Perl" ;
my $juliette = tied %juliette ;
$juliette->SetName('JULIETTE') ;
#~ $juliette->{DEBUG}{SUB}++ ;
#~ $juliette->{DEBUG}{PRINT_FORMULA}++ ;
#~ $juliette->{DEBUG}{DEFINED_AT}++ ;
#~ $juliette->{DEBUG}{ADDRESS_LIST}++ ;
#~ $juliette->{DEBUG}{FETCH_FROM_OTHER}++ ;
#~ $juliette->{DEBUG}{DEPENDENT_STACK}++ ;
#~ $juliette->{DEBUG}{DEPENDENT}++ ;

$romeo->AddSpreadsheet('JULIETTE', $juliette) ;
$juliette->AddSpreadsheet('ROMEO', $romeo) ;

$romeo{'B1:B5'} = 10 ;

$juliette{A4} = 5 ;
$juliette{A5} = PerlFormula('$ss->Sum("ROMEO!B1:B5") + $ss{"ROMEO!B2"}') ; 

$romeo{A1} = PerlFormula('$ss->Sum("JULIETTE!A1:A5", "A2")') ;
$romeo{A2} = 100 ;
$romeo{A3} = PerlFormula('$ss{A2}') ;

$romeo->Recalculate() ; #update dependents
#~ # or 
#~ print <<EOP ;  # must access to update dependents
#~ \$romeo{A1} = $romeo{A1}
#~ \$romeo{A3} = $romeo{A3}

#~ EOP

$romeo{A2}++ ; # A1 and A3 need update now
#~ $juliette{A1}++ ; # ROMEO!A1 needs update now


print $romeo->Dump(undef,1) ;
print $juliette->Dump(undef,1) ;

# inter ss cycles
$juliette{A3} = PerlFormula('$ss->Sum("ROMEO!A1")') ;  ;

$juliette{A3} ; # void context, generates warning
