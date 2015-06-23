#!/usr/bin/perl

use warnings;
use strict;

use open ':encoding(utf8)';

use lib "/home/uals/lib";
use UTILS::Config 'getConfig';
use UTILS::LOG;

use Data::Dumper;
$Data::Dumper::Indent = 1;

use XML::Simple;
$XML::Simple::PREFERRED_PARSER = 'XML::Parser';

use Getopt::Long;
use Mail::Sender;
use Pod::Usage;

my $DEBUG = 0;
my $help = 0;
GetOptions( "debug!" => \$DEBUG, "help!" => \$help );
pod2usage(-exitval => 0, -verbose => 2) if $help;

my $config = getConfig( "/home/uals/etc/invoices.conf", "perl");
my $log = UTILS::LOG->init(
        path => $config->{logfile},
	withtime => 1,
	separator => '', level => 'INFO' );

my $first = 1; ## flag for first-voucher exception
my $count = 0;

$log->info( 'Processing Alma Invoices ...' );

my $invoices = getInvoices();

if ( $DEBUG ) {
	$log->withtime(0);
	$log->trace( Dumper($invoices) );
	exit;
}

openRTF();

foreach my $id ( keys %{ $invoices } ) {

    ## process each invoice

    my $invoice = $invoices->{$id};
    
    print
	_startSection(), 
	vendorSection($invoice),
	poSection($invoice);

    $count++;

}

closeRTF();

mailLog( {
	from    => $config->{from},
	to      => $config->{to},
	subject => 'Vouchers',
	msg     => 'Attached file contains the latest vouchers',
	file    => $config->{file}
    } ) unless $DEBUG;

$log->info( "Processed $count invoices\n" );

sub getInvoices {
	local $/;
	my $xml = <>;
	my $invoices = XMLin($xml, 
	    ForceArray => [ 'fund_info', 'invoice', 'invoice_line', 'payment_address' ],
	    KeyAttr => { 'invoice' => 'unique_identifier' });
	return $invoices->{invoice_list}->{invoice};
}

##----------------------------------------------------------------------

sub vendorSection {
    my $invoice = shift;
    return _vendorSection(
	(join "\\line\n", ( $invoice->{vendor_name}, address($invoice) ) ),
	$invoice->{invoice_number},
	$invoice->{invoice_date},
	(sprintf "%s %.2f", $invoice->{invoice_amount}->{currency}, $invoice->{invoice_amount}->{sum} )
    );
}

sub address{
    ## get the vendor's address and return as an array
    my $invoice = shift;
    foreach my $address ( @{$invoice->{vendor_payment_address_list}->{payment_address}} ) {
	if ( defined $address->{preferred} and $address->{preferred} eq 'true' ) {
	    return _addr($address);
	}
    }
    ## no preferred, so we'll use the first one ...
    return _addr( $invoice->{vendor_payment_address_list}->{payment_address}->[0] );
}

sub _addr {
    my $address = shift;
    my @a = ();
    my ($line1, $line2, $line3, $line4, $line5, $city, $state, $pcode, $country, $note);
    $line1   = ref($address->{line1}) ? '' : $address->{line1};
    $line2   = ref($address->{line2}) ? '' : $address->{line2};
    $line3   = ref($address->{line3}) ? '' : $address->{line3};
    $line4   = ref($address->{line4}) ? '' : $address->{line4};
    $line5   = ref($address->{line5}) ? '' : $address->{line5};
    $city    = ref($address->{city}) ? '' : $address->{city};
    $state   = ref($address->{stateProvince}) ? '' : $address->{stateProvince};
    $pcode   = ref($address->{postalCode}) ? '' : $address->{postalCode};
    $country = ref($address->{country}) ? '' : $address->{country};
    $note    = ref($address->{note}) ? '' : $address->{note};

    push @a, $line1 if $line1;
    push @a, $line2 if $line2;
    push @a, $line3 if $line3;
    push @a, $line4 if $line4;
    push @a, $line5 if $line5;
    push @a, $city if $city;
    if ($state and $pcode) {
	push @a, "$state $pcode";
    } else {
	push @a, $state if $state;
	push @a, $pcode if $pcode;
    }
    push @a, $country if $country;
    push @a, $note if $note;
    return @a;
}

##----------------------------------------------------------------------

sub poSection {
    my $invoice = shift;
    my $invoice_lines = $invoice->{invoice_line_list}->{invoice_line};
    my @rtf = ();
    my $accounts = {};

    push @rtf, _lineItemHeader();

    ## line items
    ## each invoice line has one PO line and one or more funds
    ## (in practice is there only one fund per line ?)
    foreach my $line ( @{ $invoice_lines } ) {
	my $line_number = $line->{line_number};
	if ( exists $line->{po_line_info} ) {

	    push @rtf, _line(
		#$line->{line_number},
		$line->{po_line_info}->{po_line_number},
		$line->{po_line_info}->{po_line_title},
		sprintf( "%.2f", $line->{total_price} )
		#sprintf( "%.2f", $line->{po_line_info}->{po_line_price} )
		);
	    if ( exists $line->{fund_info_list} ) {
		foreach my $fund_info ( @{ $line->{fund_info_list}->{fund_info} } ){
		    my $account = $fund_info->{external_id};
		    my $currency = $fund_info->{amount}->{currency};
		    $accounts->{$account}->{$currency} += $fund_info->{amount}->{sum};
		}
	    }

	}
    }

    ## Funds - probably only one, but being safe, and allowing for multiple currencies
    foreach my $account ( sort keys %{$accounts} ) {
	foreach my $currency ( sort keys %{$accounts->{$account}} ) {
	    push @rtf, _fund( $account, $currency, $accounts->{$account}->{$currency} );
	}
    }
    push @rtf, _blankLine();

    ## other charges (e.g. GST)
    foreach my $line ( @{ $invoice_lines } ) {
	my $line_number = $line->{line_number};
	unless ( exists $line->{po_line_info} or $line->{total_price} eq '0.0' ) {
	    $line->{line_type} = "GST" if $line->{line_type} eq 'ADJUSTMENT';
	    push @rtf, _line( '*', $line->{line_type}, $line->{total_price} ), _blankLine();
	    if ( exists $line->{fund_info_list} ) {
		push @rtf, _funds($line);
	    }
	}
    }

    return @rtf;
}

sub _funds {
    my $line = shift;
    my @rtf = ();
    foreach my $fund_info ( @{ $line->{fund_info_list}->{fund_info} } ){
	push @rtf, _fund( $fund_info->{external_id}, $fund_info->{amount}->{currency}, $fund_info->{amount}->{sum} );
    }
    return @rtf;
}

##----------------------------------------------------------------------

sub _startSection {
    ## check for and close a previous section
    my $x = ( $first ? "" : "\\sect");
    $first = 0;

    return $x,
q|\sectd\titlepg\sbkpage\sectunlocked1|,
q|\pgndec\pgnrestart|,
q|\pgwsxn11905\pghsxn16837|,
q|\marglsxn1134\margrsxn1134\margtsxn2328\headery1134\margbsxn1685\footery1134|,
    _header(), _footer(),
q|\pard\plain \ql{\fs22\f0
\par }|;
}

sub _header {
    return q|{\headerf\pard\plain \qc{\fs20\f0
The University of Adelaide}
\par 
\pard\plain \qc{\b\fs48\f0
VOUCHER}
\par 
\pard\plain \ql{\fs22\f0
The following amount should be paid to the indicated vendor for the listed invoice which apply to the displayed purchase orders:}
\par }|;
}

sub _footer {
    return q|{\footerf|,
    authorisationSection(),
q|\pard\plain \qc{\fs20\f0
Page }{\fs20\f0
{\field{\*\fldinst  PAGE }{\fldrslt 1}}}
\par 
}|;
# numpages is not working, gives the total pages rather than per voucher
#{\fs20\f0 of }{\fs20\f0 {\field{\*\fldinst  NUMPAGES }{\fldrslt 1}}}
}

sub _blankLine {
    return
q|\pard\plain {\fs22\f0
}
\par
|;
}

##----------------------------------------------------------------------

sub _vendorSection {
    my ($address, $number, $date, $amount) = @_;
    return
q|\trowd\trql\ltrrow|,
q|\trpaddft3\trpaddt0\trpaddfl3\trpaddl108\trpaddfb3\trpaddb0\trpaddfr3\trpaddr108|,
q|\clbrdrb\brdrs\brdrw20\brdrcf3|,
q|\clvertalc\cellx4880|,
q|\clbrdrb\brdrs\brdrw20\brdrcf3|,
q|\clvertalc\cellx9636|,
q|\pard\plain \intbl{\sl0\fs22\f0
|, $address,
q|}\cell
\pard\plain \intbl\qr{\b\fs22\f0 Invoice number: |, $number, q|}\par
\pard\plain \intbl\qr{\b\fs22\f0 Invoice date: |, $date, q|}\par
\pard\plain \intbl\qr{\fs22\f0
}
\par
\pard\plain \intbl\qr{\b\fs22\f0 Invoice total: |, $amount,
q|}
\cell
\row
\pard\plain {\fs22\f0
}
\par
|;
}

##----------------------------------------------------------------------

sub _lineItemHeader {
    return
q|\trowd\trqc|,
q|\trpaddft3\trpaddt0\trpaddfl3\trpaddl57\trpaddfb3\trpaddb0\trpaddfr3\trpaddr57|,
q|\cellx1440\cellx8208\cellx9636|,
q|\pard\plain \intbl\ql{\b\fs20\f0
Line Item}\cell
\pard\plain \intbl\ql{\b\fs20\f0
Title}\cell
\pard\plain \intbl\qr{\b\fs20\f0
Amount}\cell\row|;
}

sub _line {
    my ($number, $desc, $amount) = @_;
    ## escape unicode chars.
    $desc =~ s/([^[:ascii:]])/sprintf("\\u%d\\'5f",ord($1))/eg;
    return 
q|\trowd\trqc|,
q|\trpaddft3\trpaddt0\trpaddfl3\trpaddl57\trpaddfb3\trpaddb0\trpaddfr3\trpaddr57|,
q|\cellx1440\cellx8208\cellx9636|,
q|\pard\plain \intbl\ql{\fs20\f0
|, $number, q|}\cell
\pard\plain \intbl\ql{\sl0\fs20\f0
|, $desc, q|}\cell
\pard\plain \intbl\qr{\fs20\f0
|, sprintf( "%.2f", $amount), q|}\cell\row|;
}

sub _fund {
    my ($fund, $currency, $sum) = @_;
    return
q|\trowd\trql|,
q|\trleft0\ltrrow|,
q|\trpaddft3\trpaddt0\trpaddfl3\trpaddl0\trpaddfb3\trpaddb0\trpaddfr3\trpaddr0|,
q|\cellx5760\cellx9637
\pard\plain \intbl{\fs22\f0
Fund: |,
(ref($fund) ? '' : $fund),
q|}\cell
\pard\plain \intbl\qr{\fs22\f0
Fund total: |,
sprintf("%s %.2f", $currency, $sum ),
q|}\cell\row
|;
}

sub authorisationSection {
    my $address = join "\\line\n", @{ $config->{address} };
    my $contact = join "\\line\n", @{ $config->{contact} };
    my $agent   = join "\\line\n", @{ $config->{agent} };
    return 
q|\trowd\trql\ltrrow\trpaddft3\trpaddt0\trpaddfl3\trpaddl108\trpaddfb3\trpaddb0\trpaddfr3\trpaddr108|,
q|\clbrdrt\brdrs\brdrw20\brdrcf3\clbrdrl\brdrs\brdrw20\brdrcf3|,
q|\clvertalc\cellx4877|,
q|\clbrdrt\brdrs\brdrw20\brdrcf3\clbrdrr\brdrs\brdrw20\brdrcf3|,
q|\clvertalc\cellx9635
\pard\plain \intbl{\sl0\fs20\f0
|, $address, q|}
\cell
\pard\plain \intbl\qr{\sl0\fs20\f0
|, $contact, q|}
\cell
\row
\trowd\trql\ltrrow\trpaddft3\trpaddt0\trpaddfl3\trpaddl108\trpaddfb3\trpaddb0\trpaddfr3\trpaddr108|,
q|\clbrdrl\brdrs\brdrw20\brdrcf3\clbrdrb\brdrs\brdrw20\brdrcf3|,
q|\clvertalc\cellx4877|,
q|\clbrdrb\brdrs\brdrw20\brdrcf3\clbrdrr\brdrs\brdrw20\brdrcf3|,
q|\clvertalc\cellx9635
\pard\plain \intbl{\afs22 \fs22\f0
}\cell
\pard\plain \intbl{\b\fs22\f0 Authorised by} \par 
\pard\plain \intbl{\b\fs22\f0 } \par 
\pard\plain \intbl{\b\fs22\f0 } \par 
\pard\plain \intbl{\b\fs22\f0
|, $agent, q|}
\cell
\row|;
}

##----------------------------------------------------------------------

sub openRTF {
    open RTF, ">", $config->{file} or die "Cannot open rtf file, $!\n";
    select RTF;

    print 
q|{\rtf1\ansi\deff0
{\fonttbl{\f0\fswiss\fprq2\fcharset128 Arial Unicode MS;}}
{\colortbl;\red255\green0\blue0;\red0\green0\blue255;}
{\info 
{\title Vouchers}
}|,
q|\paperh16837\paperw11905|, # A4
q|\margl1134\margr1134\margt1134\margb1134|; # 2cm margins
#    print _header(), _footer();
}

sub closeRTF {
    print q|}|;
    close RTF;
}

##----------------------------------------------------------------------

sub mailLog {

    my $email = shift;

    my $sender = new Mail::Sender { smtp => 'smtp.adelaide.edu.au' };
    $sender->MailFile(
	{
	    from    => $email->{from},
	    to      => $email->{to},
	    subject => $email->{subject},
	    msg     => $email->{msg},
	    file    => $email->{file}
	}
    );
}

__END__

=head1 NAME

Alma Invoice processing

=head1 SYNOPSIS

$ invoices.pl [ --debug ] invoices.xml

$ invoices.pl --help

=head1 DESCRIPTION

Invoices are exported from Alma in XML format. This script takes the XML file 
and produces vouchers for printing as an RTF format file.

The RTF file is emailed to the address in config.

Processing is logged in /home/uals/log/

RTF is ugly, so all that formatting is hidden in separate subroutines.

=head1 OPTIONS

--debug disables emailing of the RTF file

--help displays this

=head1 VERSION

This is version 2014.05.27

=cut

