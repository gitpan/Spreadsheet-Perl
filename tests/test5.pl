
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
						SUB => \&DoublePrevious
						, SUB_ARGS => [ 1, 2, 3]
						}
				} ;

my $ss = tied %ss ;

$ss{'A1:A8'} = '10' ;
$ss{A9} = Spreadsheet::Perl::Function(\&SumRowsAbove) ;

print "$_ = $ss{$_}\n" for($ss->GetAddressList('A1:A8')) ;
print "'A9' SumRowsAbove: " . $ss{A9} . "\n" ;

#~ print $ss->Dump() ;

sub SumRowsAbove
{
my $ss = shift ;
my $address  = shift ;
my $extra_arg = shift || 0 ;

my ($x, $y) = Spreadsheet::Perl::ConvertAdressToNumeric($address) ;

my $sum = 0 ;

for my $current_y (1 .. ($y - 1))
	{
	my $cell_value = $ss->Get("$x,$current_y") ;
	
	$sum += $cell_value if (is_numeric($cell_value)) ;
	}
	
return($sum + $extra_arg) ;
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




