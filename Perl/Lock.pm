
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

our @EXPORT = qw( ) ;
our $VERSION = '0.01' ;

#-------------------------------------------------------------------------------

sub Lock
{
my $self = shift ;
my $address = shift ;
my $lock = shift ;                    

$lock = $address unless $self->IsAddress($address) ;

$self->{LOCKED} = $lock ;
}

#-------------------------------------------------------------------------------

sub LockRange
{
my $self = shift ;
my $address = shift ;
my $lock = shift ;                    

for my $current_address ($self->GetAddressList($address))
	{
	$self->{DATA}{$current_address}{LOCKED} = $lock ;
	}
}

#-------------------------------------------------------------------------------

sub IsLocked
{
my $self = shift ;
return($self->{LOCKED}) ;
}

#-------------------------------------------------------------------------------

sub IsRangeLocked
{
my $self = shift ;
my $address = shift ;

confess "Unimplemented" ;
}

#-------------------------------------------------------------------------------

sub IsCellLocked
{
my $self = shift ;
my $address = $self->CanonizeCellAddress(shift) ;

if($self->{LOCKED})
	{
	return(1) ;
	}
else
	{
	if(exists $self->{DATA}{$address})
		{
		if(exists $self->{DATA}{$address}{LOCKED})
			{
			return($self->{DATA}{$address}{LOCKED}) ;
			}
		else
			{
			return(0) ;
			}
		}
	else
		{
		return(0) ;
		}
	}
}

#-------------------------------------------------------------------------------

1 ;

__END__
=head1 NAME

Spreadsheet::Perl::Lock- Lock support for Spreadsheet::Perl

#~ =head1 SYNOPSIS
=head1 TODO

Locking functinos need some work to look unified!

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
