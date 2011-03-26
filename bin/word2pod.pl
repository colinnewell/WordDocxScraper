=head1 NAME

word2pod.pl - a quick and simple script to turn a docx to pod

=head1 SYNOPSIS

This script takes the xml from a Microsoft Office docx document and turns it into pod.

Why you'd want to do this is anyones guess but we want to turn a single document
into pod so I wrote this script.  It does a reasonable job on that one document
so I figured I'd share it. 

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
use XML::LibXML;
use WordDocxScraper;
use Data::Dumper;
use feature "switch";
use Text::Wrap;
use List::Util qw/reduce/;

binmode STDOUT, ":utf8";

my $file = shift;
my $rels = shift;
my $dom = XML::LibXML->load_xml(location => $file);
my $relationships = XML::LibXML->load_xml(location => $rels);
my $doc = WordDocxScraper::read_doc($dom, $relationships);

for my $para (@$doc)
{
    for my $links (@{$para->{img}})
    {
        print "L<$links>\n\n";
    }
    given($para->{style})
    {
        when ('Bullet')
        {
            print "=over\n\n";
            print reduce { $a . $b } map { "\n=item * " . join ( '', @$_ ) ."\n" } @{$para->{lines}}; 
            print "\n=back\n\n";
        }
        when ('Code')
        {
            print reduce { $a . $b } map { "  " . join ( '', @$_ ). "\n" } @{$para->{lines}};
            print "\n";
        }
        when ('Heading1')
        {
            print "=head1 " . join ("", @{$para->{lines}->[0]}) . "\n";
            my @rest = splice @{$para->{lines}}, 1;
            print reduce { $a . $b } map { "\n" . join ( '', @$_ ) } @rest;
            print "\n";
        }
        when ('Heading2')
        {
            print "=head2 " . join ("", @{$para->{lines}->[0]}) . "\n";
            my @rest = splice @{$para->{lines}}, 1;
            print reduce { $a . $b } map { "\n" . join ( '', @$_ ) } @rest;
            print "\n";
        }
        when ('Heading3')
        {
            print "=head3 " . join ("", @{$para->{lines}->[0]}) . "\n";
            my @rest = splice @{$para->{lines}}, 1;
            print reduce { $a . $b } map { "\n" . join ( '', @$_ ) } @rest;
            print "\n";
        }
        when ('normal')
        {
            print wrap ('', '', reduce { $a . $b } map { join ( '', @$_ ). "\n" } @{$para->{lines}});
            print "\n";
        }
    }

}

