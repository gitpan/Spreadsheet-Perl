
use Carp ;
use strict ;
use warnings ;

use Spreadsheet::Perl ;
use Spreadsheet::Perl::Arithmetic ;

tie my %ss, "Spreadsheet::Perl", NAME => 'TEST' ;
my $ss = tied %ss ;

%ss = 
	(
	  A1 => 1
	, A2 => FetchFunction(sub{1})
	, A3 => Formula('$ss->Sum("A1:A2")') 
	
	, B1 => 3
	, c2 => "hi there"
	) ;

print $ss->Dump() ;
print "\$ss{A3} = $ss{A3}\n" ;

%ss = do "ss_setup.pl" ;
print $ss->Dump() ;

print "\$ss{A3} = $ss{A3}\n" ;

print "keys:" . join(', ', keys %ss) . "\n" ;

print "A5 exists\n" if exists $ss{A5} ;
print "B5 doesn't exists\n" unless exists $ss{B5} ;
$ss{B5}++ ;
print "B5 exists\n" if exists $ss{B5} ;

%ss = () ;
print "keys:" . join(', ', keys %ss) . "\n" ;

@ss{'A1', 'B1:C2', 'A8'} = ('A', 'B', 'C');
print $ss->Dump() ;
