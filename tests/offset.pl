
use Carp ;
use strict ;
use warnings ;

use Spreadsheet::Perl ; 
use Spreadsheet::Perl::QuerySet ;

use Data::Dumper ;

tie my %ss, "Spreadsheet::Perl" ;
my $ss = tied %ss ;

for
	(
	  ['A1', 1, 1]
	, ['Z9', 1, 0]
	, ['Z9', 1, 1]
	, ['ZZ1', 1, 1]
	, ['AAA1', -1, 0]
	, ['AAA1', -1, 0]
	, ['ABC5', 25, 3]
	, ['Z1', -25, 0]
	, ['Z1', -26, 0]
	, ['AA2', -26, -1]
	, ['AA2', -26, -2]
	)
	{
	my $offset_cell = $ss->OffsetAddress(@$_) ;
	my $offset_string = "Can't compute!" ;
	
	if(defined $offset_cell)
		{
		$offset_string = join ", ", $ss->GetCellsOffset($_->[0], $offset_cell) ;
		}
	else
		{
		$offset_cell = "Can't offset!" ;
		}
	
	print '' . (join(", ", @$_)) . " => " . $offset_cell . " offset: " . $offset_string  . "\n" ;
	}

