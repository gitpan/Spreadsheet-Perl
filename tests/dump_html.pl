
use Carp ;
use strict ;
use warnings ;

use Spreadsheet::Perl ;

#autofill is just a construction of the mind!

tie my %ss, "Spreadsheet::Perl" ;
my $ss = tied %ss ;

my $week_days = [qw(dimanche lundi mardi mercredi jeudi vendredi samedi)] ;

$ss{'B2:G5'} = RangeValuesSub(\&WeekDayFiller, $week_days, 'mardi') ;
$ss{K10} = 'last cell' ;

DumpHtmlToFile($ss, './html_dump.html') ;

sub DumpHtmlToFile
{
my $ss = shift ;
my $file_name = shift ;

open(HTML_FILE, '>', $file_name) ;
print HTML_FILE HtmlDump($ss) ;
close HTML_FILE
}

#-------------------------------------------------------------

sub HtmlDump2
{
my $ss = shift ;

use Spreadsheet::ConvertAA ;
use HTML::Table;

my ($last_letter, $last_number) = $ss->GetLastIndexes() ;

my ($rows, $cols) = (FromAA($last_letter), $last_number) ;

my $table1 = new HTML::Table($rows, $cols) ;

for ($ss->GetCellList())
	{
	my ($cellcol, $cellrow) =  ConvertAdressToNumeric($_) ;
	
	$table1->setCell($cellrow, $cellcol, $ss{$_}) ;
	}

return($table1->getTable()) ;
}

sub HtmlDump
{
use Data::Table ;

my $ss = shift ;
my ($last_letter, $last_number) = $ss->GetLastIndexes() ;
my ($cols, $rows) = (FromAA($last_letter), $last_number) ;

my $data = 
	[
	map
		{
		[
		map{''} (1 ..$cols + 1)
		]
		} (1 .. $rows)
	
	] ;

my $table1 = new Data::Table($data, [ map{ToAA($_)} (0, 1 .. $cols)]) ;

for (1 .. $rows)
	{
	$table1->setElm($_ - 1, 0, $_) ;
	}
	
for ($ss->GetCellList())
	{
	my ($cellcol, $cellrow) =  ConvertAdressToNumeric($_) ;
	
	$table1->setElm ($cellrow - 1, $cellcol, $ss->{CELLS}{$_}) ;
	}

return($table1->html()) ;

}

#-------------------------------------------------------------

sub WeekDayFiller
{
my ($ss, $anchor, $current_address, $week_days, $start_day) = @_ ;

my $day_offset = 0 ;

for (@$week_days)
	{
	last if $_ eq $start_day ;
	$day_offset++ ;
	}
	
my ($cell_offset_x, $cell_offset_y) = $ss->GetCellsOffset($anchor, $current_address) ;

my $day = ($cell_offset_x + $cell_offset_y + $day_offset) % @$week_days ;

return($week_days->[$day]) ;
}

