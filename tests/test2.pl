
# do not use 

use Carp ;
use strict ;
use warnings ;

use Spreadsheet::Perl ;
use Spreadsheet::Perl::Arithmetic ;

use Data::Dumper ;

tie my %ss, "Spreadsheet::Perl" ;
my $ss = tied %ss ;

#~ $ss->SetRangeName("TestRange", 'A1:A2') ;
#~ $ss{TestRange} = '7' ;

#~ $ss{B1} = Spreadsheet::Perl::Formula('$ss->Sum{"A1:A2"}') ;
#~ $ss{B2} = Spreadsheet::Perl::Formula($ss->GetFormulaText("B1"), SOMETHING => 1) ;

#~ $ss->{DEBUG}{ADDRESS_LIST}++ ;
#~ $ss{'A1:B2'} = 1 ;
#~ $ss{'B2:A1'} = 1 ;
#~ $ss{'A4:B2'} = 1 ;
#~ $ss{'B4:A2'} = 1 ;
#~ print $ss->Dump() ;

#Fetcher/Doer
# dependency OK
#~ $ss{A1} = Spreadsheet::Perl::FetchFunction(sub{ use Data::Dumper; print $ss->Dump() ; return($ss->Get('A2') ) ;}) ;
#~ $ss{A2} = 'hi' ;
#~ print $ss{A1} . "\n" ;
#~ $ss{A2} = 'there' ;
#~ print $ss{A1} . "\n" ;

# cached value is returned even if the ss is changed
#~ $ss{A1} = Spreadsheet::Perl::FetchFunction(sub{ use Data::Dumper; return($ss->Dump()) ;}) ;
#~ $ss{A2} = 'hi' ;
#~ print $ss{A1} . "\n" ;
#~ $ss{A2} = 'there' ;

#~ print $ss{A1} . "\n" ; #!!  cached value is returned

#~ $ss{A1} = Spreadsheet::Perl::NoCache() ;
#~ $ss{A1} = Spreadsheet::Perl::FetchFunction(sub{ use Data::Dumper; return($ss->Dump(undef, 1)) ;}) ;
#~ $ss{A2} = 'hi' ;
#~ print $ss{A1} . "\n" ;
#~ $ss{A2} = 'there' ;
#~ print $ss{A1} . "\n" ; # Ok

# finding out how many time (defined) cells are accessed
#~ $ss->{DEBUG}{FETCH}++ ;
#~ $ss->{DEBUG}{STORE}++ ;
#~ $ss->{DEBUG}{FETCHED}++ ;
#~ $ss->{DEBUG}{STORED}++ ;

# inter spreadsheet addresses
$ss->SetName('ROMEO') ;
my $a  ;

$ss{A1} = 1 ;
$ss{'ROMEO!A2'} = 7 ;
$a = $ss{'ROMEO!A2'} ;
print "\$a = $a\n" ;
print $ss->Dump(undef, 1) ;

# formula relocating
#~ $ss->{DEBUG}{PRINT_FORMULA}++ ; # the formula is displayed when the formula sub is created

$ss{'A1:A3'} = 1 ;

# example 1
#~ $ss{'C1:C2'} = Spreadsheet::Perl::Formula('$ss->Sum("A1:A2")') ;
#~ print "==> @{@ss{'C1:C2'}}\n" ;

# example 2
#~ $ss->Reset() ;
#~ $ss->{DEBUG}{PRINT_FORMULA}++ ;
#~ $ss{'A1:A3'} = 1 ;
#~ $ss{'D1:E2'} = Spreadsheet::Perl::Formula('$ss->Sum("[A]1:A[3]")',) ;
#~ print "==> @{@ss{'D1:E2'}}\n" ;

# example 3, some errors in the range that are caught at run time
#~ $ss->{DEBUG}{DEFINED_AT}++ ;
#~ print "==> @{@ss{'C1:C2'}}\n" ;

#~ $ss{'D1:D2'} = Spreadsheet::Perl::Formula('$ss->Sum("[A1]:[A3]")',) ;
#~ print "==> @{@ss{'D1:D2'}}\n" ;

#~ $ss{'E1:E2'} = Spreadsheet::Perl::Formula('$ss->Sum("[A]1]:[A]3")',) ;
#~ print "==> @{@ss{'E1:E2'}}\n" ;

#~ $ss{'F1:F2'} = Spreadsheet::Perl::Formula('$ss->Sum("[A[1]:A[3]")',) ;
#~ print "==> @{@ss{'F1:F2'}}\n" ;

#~ $ss{'G1:G2'} = Spreadsheet::Perl::Formula('$ss->Sum("[A][1]:A3")',) ;
#~ print "==> @{@ss{'G1:G2'}}\n" ;

#~ $ss{'H1:H2'} = Spreadsheet::Perl::Formula('$ss->Sum("[A]]1]:[A3]")',) ;
#~ print "==> @{@ss{'H1:H2'}}\n" ;

#~ $ss{'I1:I2'} = Spreadsheet::Perl::Formula('$ss->Sum("[A]]1]:")',) ;
#~ print "==> @{@ss{'I1:I2'}}\n" ;

#~ $ss{'J1:J2'} = Spreadsheet::Perl::Formula('$ss->Sum("[A]]1]")',) ;
#~ print $ss->Dump() ;
#~ print "==> @{@ss{'J1:J2'}}\n" ;

#~ # range fetching
#~ print $ss->Dump(undef, 1) ;
#~ $ss->{DEBUG}{FETCH}++ ;
#~ $ss->{DEBUG}{ADDRESS_LIST}++ ;

#~ # data is encapsulated in an array as Fetch forces scalar context
#~ my $array_with_values =$ss{'A1:A3'} ;
#~ my ($a, $b, $c) = @$array_with_values ;
#~ print "$a, $b, $c \n" ;

#~ #slice access
#~ $ss->{DEBUG}{FETCH}++ ;
#~ print Dumper(@ss{'C1:C2', 'A1:A3'}) . "\n" ; 
#~ print $ss->Dump(undef, 1) ;

#~ #slice access
#~ $ss->{DEBUG}{STORE}++ ;
#~ @ss{'C1:C2', 'A1:A3'} = (5, 10) ;
#~ print $ss->Dump(undef, 1) ;

#~ @ss{$ss->GetAddressList('A1:A3')} = (1 .. 3) ;
#~ print $ss->Dump(undef, 1) ;
