=head1 NAME

word2pod.pl - a quick and simple script to turn a docx to pod

=head1 SYNOPSIS

This script takes the xml from a Microsoft Office docx document and turns it into pod.

Run this like,

  unzip mydoc.docx
  ./word2pod.pl word/document.xml word/_rels/document.xml.rels > my.pod

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

# remove all 'smart' punctuation.
for my $para (@$doc)
{
    for my $line (@{$para->{lines}})
    {
        for my $bit (@$line)
        {
            my $frag = $bit->{text};
            $frag =~ s/\N{U+2013}/-/g;
            $frag =~ s/\N{U+2018}/'/g;
            $frag =~ s/\N{U+2019}/'/g;
            $frag =~ s/\N{U+201c}/"/g;
            $frag =~ s/\N{U+201d}/"/g;
            $frag =~ s/\N{U+2026}/.../g;
        }
    }
}
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
            print reduce { $a . $b } map { "\n=item * " . join ( '', map { render_chunk ($_) } @$_ ) ."\n" } @{$para->{lines}}; 
            print "\n=back\n\n";
        }
        when ('Code')
        {
            my $paragraph = reduce { $a . $b } map { "  " . join ( '', map { render_chunk ($_) } @$_ ). "\n" } @{$para->{lines}};
            print $paragraph . "\n";
        }
        when ('Heading1')
        {
            print "=head1 " . join ("", map { render_chunk ($_) } @{$para->{lines}->[0]}) . "\n";
            my @rest = splice @{$para->{lines}}, 1;
            print reduce { $a . $b } map { "\n" . join ( '', map { render_chunk ($_) } @$_ ) } @rest;
            print "\n";
        }
        when ('Heading2')
        {
            print "=head2 " . join ("", map { render_chunk ($_) } @{$para->{lines}->[0]}) . "\n";
            my @rest = splice @{$para->{lines}}, 1;
            print reduce { $a . $b } map { "\n" . join ( '', map { render_chunk ($_) } @$_ ) } @rest;
            print "\n";
        }
        when ('Heading3')
        {
            print "=head3 " . join ("", @{$para->{lines}->[0]}) . "\n";
            my @rest = splice @{$para->{lines}}, 1;
            print reduce { $a . $b } map { "\n" . join ( '', map { render_chunk ($_) } @$_ ) } @rest;
            print "\n";
        }
        when ('normal')
        {
            print wrap ('', '', reduce { $a . $b } map { join ( '', map { render_chunk ($_) } @$_ ). "\n" } @{$para->{lines}});
            print "\n";
        }
    }

}

sub render_chunk
{
    my $bit = shift;
    return if !$bit;
    if($bit->{style} eq 'bold')
    {
        return 'B<' . $bit->{text} . '>';
    }
    elsif($bit->{style} eq 'italic')
    {
        return 'I<' . $bit->{text} . '>';
    }
    else
    {
        return $bit->{text};
    }
}
