
package Spreadsheet::Perl ;

use 5.006 ;

use Carp ;
use strict ;
use warnings ;

require Exporter ;
#~ use AutoLoader qw(AUTOLOAD) ;

our @ISA = qw(Exporter) ;

our %EXPORT_TAGS = 
	(
	'all' => [ qw() ]
	) ;

our @EXPORT_OK = ( @{ $EXPORT_TAGS{'all'} } ) ;

#~ our @EXPORT = qw( Formula ) ;
our @EXPORT ;
push @EXPORT, qw( Formula ) ;

our $VERSION = '0.01' ;

#-------------------------------------------------------------------------------

sub Formula
{
confess "unimplemented!" ;

my $perl_code ;
return(PerlFormula($perl_code)) ;
# or
#return bless [\&GenerateXXXFormulaSub, @_], "Spreadsheet::Perl::PerlFormula" ;
}

#-------------------------------------------------------------------------------

1 ;

__END__
=head1 NAME

Spreadsheet::Perl::Formula - Formula support for Spreadsheet::Perl

=head1 SYNOPSIS

  $ss{A1} = Formula('B1 + Sum(A1:A6)') ;
  
=head1 DESCRIPTION

Part of Spreadsheet::Perl.

=head1 AUTHOR

Khemir Nadim ibn Hamouda. <nadim@khemir.net>

  Copyright (c) 2004 Nadim Ibn Hamouda el Khemir. All rights
  reserved.  This program is free software; you can redis-
  tribute it and/or modify it under the same terms as Perl
  itself.
  
If you find any value in this module, mail me!  All hints, tips, flames and wishes
are welcome at <nadim@khemir.net>.

=cut
