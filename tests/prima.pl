
=head1 NAME

Spreadsheet::Perl Prima GUI

base on Prima's grid exemple.

=cut

use strict;
use warnings ;


package PrimaOut;

use strict;

sub TIEHANDLE 
{
my $class = shift;
my $window = shift ;

bless {OUTPUT_WINDOW => $window, BUFFER => '' }, $class;
}

sub PRINT 
{
my $self = shift;

my $input = $self->{BUFFER} . $_[0] ;
$self->{BUFFER} = '' ;

for my $line (split /(.*?\n)/, $input)
	{
	if(chomp $line)
		{
		$line =~ s/\t/    /g ;
		$self->{OUTPUT_WINDOW}->add_items($line);
		}
	else
		{
		$self->{BUFFER} .= $line;
		}
	}
}


package main ;

use Prima qw(Application FrameSet Edit ComboBox MsgBox Grids) ;

use Spreadsheet::Perl ;
use Spreadsheet::ConvertAA ;

use Spreadsheet::Perl ;

my $ss = tie my %ss, "Spreadsheet::Perl" ;

my $file = $ARGV[0] || 'ss_data.pl' ;
$ss->Read($file) ;
$ss{C3} = PF('$ss{C2} + 5') ;

my $grid;

my $window = Prima::MainWindow->create
		(
		text => 'Spreadsheet::Perl example',
		packPropagate => 0,
		menuItems => 
			[
			['~Grid' => 
				[
				['*dhg', 'Draw HGrid'=> sub { $grid->drawHGrid( $_[0]->menu->toggle( $_[1])) }],
				['*dvg', 'Draw VGrid'=> sub { $grid->drawVGrid( $_[0]->menu->toggle( $_[1])) }],
				['mse', 'Multi select'=> sub { $grid->multiSelect( $_[0]->menu->toggle( $_[1])) }],
				['ccw', 'Constant cell width' => sub { $grid->constantCellWidth($_[0]->menu->toggle( $_[1]) ? 100 : undef); }],
				['cch', 'Constant cell height' => sub { $grid->constantCellHeight($_[0]->menu->toggle( $_[1]) ? 100 : undef); }],
				]
			],
			['~Spreadsheet' => 
				[
				['c', 'Clear All'=> sub {}],
				['f', 'Functions'=> sub {}],
				['n', 'Names'=> sub {}],
				['Aa', 'About'=> sub {message_box('About', "Spreadsheet::Perl && Prima", mb::Ok)}]
				]
			]
			],
		);

my @user_breadths=({},{});

my $frame = $window->insert
		(		    
		FrameSet =>
			size => [$window->size],
			origin => [0, 0],
			arrangement => fra::Vertical,
			frameSizes => [qw(25%)],
		);


my $bottom_frame = $frame->insert_to_frame
			(
			0,
			FrameSet =>
				#~ size => [$window->size],
				size => [$frame->frames->[0]->size],
				origin => [0, 0],
				#~ arrangement => fra::Vertical,
				frameSizes => [qw(60% 40%)],
			);

my $output = $bottom_frame->insert_to_frame
			(
			1,
			"ListBox",
			hScroll        => 1,
			multiSelect    => 0,
			extendedSelect => 1,
			name            => 'output',
			pack            => { side => 'left', expand => 1, fill => 'both', padx => 20, pady => 20},
			);

my $editor = $bottom_frame->insert_to_frame
			(
			0,
			"Edit", 
			maxLen         => 200,
			name           => '',
			hScroll        => 1,
			vScroll        => 1,
			wantReturns    => 1,
			pack            => { side => 'left', expand => 1, fill => 'both', padx => 20, pady => 20},
			);

tie *OUT, "PrimaOut", $output ;
$ss->{DEBUG}{ERROR_HANDLE} = \*OUT ;

$grid = $frame->insert_to_frame
	(
	1,
	'AbstractGrid', 
	 
	backColor => cl::White,
	
	onSelectCell => sub
			{
			my ($self, $column, $row) = @_ ;
			
			$editor->select_all() ;
			$editor->delete_block() ;
			
			my $formula = $ss->GetFormulaText("$column,$row") ;
			
			if(defined $formula)
				{
				$editor->insert_text($formula) ;
				}
			else
				{
				$editor->insert_text($ss{"$column,$row"}) ;
				}
			},
			
	#~ onKeyDown => sub
			#~ {
			#~ my ($self, $key) = @_ ;
			
			#~ message_box('OnButtonDown', $key, mb::Ok) ;
			#~ } ,
	
	onMeasure => sub 
			{
			my ($self, $axis, $index, $ref) = @_;
			
			if(exists $user_breadths[$axis]{$index}) # use exists saves memory
				{
				$$ref = $user_breadths[$axis]{$index};
				}
			else
				{
				$$ref = $axis ? 40 : 80 ;
				}
			},
			
	onSetExtent => sub 
			{
			my ($self, $axis, $index, $breadth) = @_;
			$user_breadths[$axis]{$index} = $breadth;
			},
			
	onStringify => sub 
			{
			my ($self, $col, $row, $ref) = @_ ;
			
			$$ref = $ss{"$col,$row"} ;
			},
	
	allowChangeCellWidth => 1,
	allowChangeCellHeight => 1,
	clipCells => 1,
	
	pack => { expand => 1, fill => 'both' },
	);

$grid->columns(100);
$grid->rows(100);
$grid->cellIndents(1,1);

Prima->run() ;
