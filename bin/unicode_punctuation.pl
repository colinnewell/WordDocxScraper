=head1 NAME

unicode_punctuation.pl - a quick and simple script to show unicode chars in use

=head1 SYNOPSIS

This script takes the xml from a Microsoft Office docx document and displays
the unicode characters in use.

Run this like,

  unzip mydoc.docx
  ./unicode_punctuation.pl word/document.xml word/_rels/document.xml.rels > my.pod

=head1 AUTHOR

Colin Newell, C<< <colin.newell at gmail.com> >>

=head1 BUGS

This module has only been very lightly tested to work against a single document. 
It may well not work against general word docx documents.

=head1 LICENSE AND COPYRIGHT

Copyright 2011 Colin Newell.

This program is free software; you can redistribute it and/or modify it
under the terms of either: the GNU General Public License as published
by the Free Software Foundation; or the Artistic License.

See http://dev.perl.org/licenses/ for more information.


=cut

use strict;
use feature "switch";
use List::Util qw/reduce/;
use Text::Wrap;
use WordDocxScraper;
use XML::LibXML;

binmode STDOUT, ":utf8";

# FIXME: ought to add help on parameters etc.
my $file = shift;
my $rels = shift;
my $dom = XML::LibXML->load_xml(location => $file);
my $relationships = XML::LibXML->load_xml(location => $rels);
my $doc = WordDocxScraper::read_doc($dom, $relationships);

my $unicode_chars = {};
for my $para (@$doc)
{
    for my $line (@{$para->{lines}})
    {
        for my $frag (@$line)
        {
            while($frag =~ /(\p{P})/g)
            {
                my $ord = ord($1);
                if($ord > 255)
                {
                    $unicode_chars->{$ord} = $1;
                }
            }
        }
    }
}
for my $char (sort keys %$unicode_chars)
{
    printf "Punctuation: %s - %x\n", $unicode_chars->{$char} ,$char;
}
