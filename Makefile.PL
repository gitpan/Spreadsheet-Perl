use ExtUtils::MakeMaker;
# See lib/ExtUtils/MakeMaker.pm for details of how to influence
# the contents of the Makefile that is written.
WriteMakefile(
    'NAME'		=> 'Spreadsheet::Perl',
    'VERSION_FROM'	=> 'Perl.pm', # finds $VERSION
    #'DIR'               => ['Perl'],
    'PREREQ_PM'		=>
    	{
	Spreadsheet::ConvertAA => 0.01,
        Data::TreeDumper       => 0.12,
	},

    ($] >= 5.005 ?    ## Add these new keywords supported since 5.005
      (ABSTRACT_FROM => 'Perl.pm', # retrieve abstract from module
       AUTHOR     => '<Khemir Nadim ibn Hamouda. <nadim@khemir.net>') : ()),
);