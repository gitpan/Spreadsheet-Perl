
use Carp ;
use strict ;
use warnings ;

use Spreadsheet::Perl ; # tie
use Spreadsheet::Perl::Cache ;
use Spreadsheet::Perl::QuerySet ;
use Spreadsheet::Perl::Validator ;
use Spreadsheet::Perl::Formula ;
use Spreadsheet::Perl::Format ;
use Spreadsheet::Perl::Lock ;
use Spreadsheet::Perl::Devel ;
use Spreadsheet::Perl::Arithmetic ;

use Data::Dumper ;

tie my %romeo, "Spreadsheet::Perl" ;
my $romeo = tied %romeo ;
$romeo->SetName('ROMEO') ;
#~ $romeo->{DEBUG}{AUTOCALC}++ ;
#~ $romeo->{DEBUG}{SUB}++ ;
#~ $romeo->{DEBUG}{ADDRESS_LIST}++ ;
#~ $romeo->{DEBUG}{FETCH_FROM_OTHER}++ ;
#~ $romeo->{DEBUG}{DEPENDENT_STACK}++ ;
$romeo->{DEBUG}{DEPENDENT}++ ;

tie my %juliette, "Spreadsheet::Perl" ;
my $juliette = tied %juliette ;
$juliette->SetName('JULIETTE') ;
#~ $juliette->SetAutocalc(0) ;
#~ $juliette->{DEBUG}{AUTOCALC}++ ;
#~ $juliette->{DEBUG}{SUB}++ ;
#~ $juliette->{DEBUG}{PRINT_FORMULA}++ ;
#~ $juliette->{DEBUG}{DEFINED_AT}++ ;
#~ $juliette->{DEBUG}{ADDRESS_LIST}++ ;
#~ $juliette->{DEBUG}{FETCH_FROM_OTHER}++ ;
#~ $juliette->{DEBUG}{DEPENDENT_STACK}++ ;
#~ $juliette->{DEBUG}{DEPENDENT}++ ;

# inter spreadsheet addresses
$romeo->AddSpreadsheet('JULIETTE', $juliette) ;
$juliette->AddSpreadsheet('ROMEO', $romeo) ;

$romeo{'B1:B5'} = 8 ;

# inter ss cycles
#~ $juliette{A3} = Spreadsheet::Perl::Formula('$ss->Sum("ROMEO!A1")') ;  ;

#~ $romeo{'JULIETTE!A4'} = 8 ; # <=> $juliette{A4}  = 8
#~ $juliette{A1} = 0 ;
$juliette{A4} = 8 ;
$juliette{A5} = Spreadsheet::Perl::Formula('$ss->Sum("ROMEO!B1:B5") + $ss{"ROMEO!B2"}') ; 

$romeo{A1} = Spreadsheet::Perl::Formula('$ss->Sum("JULIETTE!A1:A5", "A2")') ;
$romeo{A3} = Spreadsheet::Perl::Formula('$ss{A2}') ;
$romeo{A2} = 100 ;

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

