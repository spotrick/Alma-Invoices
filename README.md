NAME
    Alma Invoice processing

SYNOPSIS
    $ invoices.pl [ --debug ] invoices.xml

    $ invoices.pl --help

DESCRIPTION
    Invoices are exported from Alma in XML format. This script takes the XML
    file and produces vouchers for printing as an RTF format file.

    The RTF file is emailed to the address in config.

    Processing is logged in /home/uals/log/

    RTF is ugly, so all that formatting is hidden in separate subroutines.

OPTIONS
    --debug disables emailing of the RTF file

    --help displays this

VERSION
    This is version 2014.05.27

