
use Carp ;
use strict ;
use warnings ;

use Spreadsheet::Perl ; # tie
use Spreadsheet::Perl::QuerySet ;
use Spreadsheet::Perl::Validator ;
use Spreadsheet::Perl::Formula ;
use Spreadsheet::Perl::Function ;
use Spreadsheet::Perl::Format ;
use Spreadsheet::Perl::Lock ;
use Spreadsheet::Perl::Devel ;
use Spreadsheet::Perl::Arithmetic ;

tie my %ss, "Spreadsheet::Perl"
		, DATA =>
				{
				  A1 =>
						{
						VALUE => 'hi'
						}
					
				, A2 =>
						{
						VALUE => 'there'
						
						#~ SUB => \&DoublePrevious
						#~ , SUB_ARGS => [ 1, 2, 3]
						
						#~ FORMULA => '$ss{A1}'
						}
				} ;

my $ss = tied %ss ;

#~ $ss->{DEBUG}{DEPENDENT}++ ;
#~ $ss->{DEBUG}{FETCH}++ ;
#~ $ss->{DEBUG}{STORE}++ ;
#~ $ss->{DEBUG}{SUB}++ ;

# cyclic error
#~ $ss->{DEBUG}{DEFINED_AT}++ ;
#~ $ss{'A1:A5'} = Spreadsheet::Perl::Formula('$ss{"A2"}') ;
#~ $ss{A6} = Spreadsheet::Perl::Formula('$ss{A1}') ;
#~ print "$ss{A1}\n" ;

#~ print  $ss{A1} . ' ' . $ss{A2} . "\n" ;

#~ $ss->SetCellName("First", "1,1") ;
#~ print  $ss{First} . ' ' . $ss{A2} . "\n" ;

#~ $ss->Lock(1) ;
#~ $ss{A1} = 'ho' ;
#~ $ss->Lock(0) ;

#~ $ss->LockRange("A1:B1", 1) ;
#~ $ss{A1} = 'ho' ;
#~ $ss{C1} = 'ho' ;
#~ $ss->LockRange("A1:B1", 0) ;
#~ $ss{A1} = 'hej' ;

$ss->SetRangeName("TestRange", 'A5:B8') ;
$ss{TestRange} = '7' ;

$ss->DefineFunction('AddOne', \&AddOne) ;

#~ $ss{A3} = Spreadsheet::Perl::Formula('$ss->AddOne("A5") + $ss{A5}') ;
#~ print "A3 => '@{[$ss->GetFormulaText('A3')]}' = $ss{A3}\n" ;

$ss{'ABC1:ABD5'} = '10' ;

$ss{A4} = Spreadsheet::Perl::Formula('$ss->Sum("A5:B8", "ABC1:ABD5")') ;
print "A4 => '@{[$ss->GetFormulaText('A4')]}' = $ss{A4}\n" ;

#~ $ss{A9} = Spreadsheet::Perl::Function(\&SumRowsAbove) ;
#~ print "'A9' SumRowsAbove: " . $ss{A9} . "\n" ;

#~ $ss{A10} = Spreadsheet::Perl::Formula('"$cell => " . (join "-", (ConvertAdressToNumeric($cell)))') ;
#~ print "'A10' Self: " . $ss{A10} . "\n" ;

#~ print "Cells: " . (join " - ", $ss->GetCellList()) . "\n" ;
#~ print "Last Indexes: " . (join " - ", $ss->GetLastIndexes()) . "\n" ;

#~ $ss{A5}++ ;
#~ print $ss->GetCellsToUpdateDump() ;

#~ $ss->{DEBUG}{VALIDATOR}++ ;
#~ $ss{A5} = Spreadsheet::Perl::Validator("Test validator", sub{return(1) ;}) ;
#~ $ss{A5}++ ;

#~ print $ss->Dump() ;

#~ $ss->Reset() ;

#~ $ss{A1} = Spreadsheet::Perl::Format(ANSI => ["red_on_black"]) ;
#~ print "\n---------------- Dump ----------------\n" ;
#~ print $ss->Dump() ;

#~ $ss{A1} = Spreadsheet::Perl::AddFormat(ANSI => ["blink"]) ;
#~ print "\n---------------- Dump ----------------\n" ;
#~ print $ss->Dump() ;

#-------------------------------------------------------------------------------
#~ dummy subs
#-------------------------------------------------------------------------------

sub AddOne
{
my $ss = shift ;
my $address = shift ;

return($ss->Get($address) + 1) ;
}

#-------------------------------------------------------------------------------

sub DoublePrevious
{
my $ss = shift ;
my $address  = shift ;

my ($x, $y) = Spreadsheet::Perl::ConvertAdressToNumeric($address) ;
my $cell_value = $ss->Get("$x," . ($y - 1)) ;

return("$cell_value$cell_value") ;
}

#-------------------------------------------------------------------------------

sub SumRowsAbove
{
my $ss = shift ;
my $address  = shift ;

my ($x, $y) = Spreadsheet::Perl::ConvertAdressToNumeric($address) ;

my $sum = 0 ;

for my $current_y (1 .. ($y - 1))
	{
	my $cell_value = $ss->Get("$x,$current_y") ;
	
	$sum += $cell_value if (is_numeric($cell_value)) ;
	}
	
return($sum) ;
}

#-------------------------------------------------------------------------------
# from Perl Cookbook, doesn't seem to work
#-------------------------------------------------------------------------------

sub getnum 
{
use POSIX qw(strtod);
my $str = shift;

return unless (defined $str) ;

$str =~ s/^\s+//;
$str =~ s/\s+$//;
$! = 0;

my($num, $unparsed) = strtod($str);
if (($str eq '') || ($unparsed != 0) || $!) 
	{
	return;
	}
else 
	{
	return $num;
	} 
} 

sub is_numeric
{
defined scalar &getnum ;
} 




