package WordDocxScraper;

use warnings;
use strict;

use Exporter 'import';
our @EXPORT_OK = qw/read_doc/;
use XML::LibXML;


=head1 NAME

WordDocxScraper - A quick and dirty Microsoft Word DocX document scraper

=head1 VERSION

Version 0.02

=cut

our $VERSION = '0.02';


=head1 SYNOPSIS

This module reads the main document xml and relationship xml and turns it into a hash
of some of the document information.

It drops a lot of the information.  Currently it tries to get the rough style, e.g.
'Heading1', 'normal' and 'Code' for example but it doesn't get fine grained style info
like bold or italic.

It also tries to pull out image references.

Perhaps a little code snippet.

    use WordDocxScraper;
    use XML::LibXML;


    my $dom = XML::LibXML->load_xml(location => 'word/document.xml');
    my $rel = XML::LibXML->load_xml(location => 'word/_rels/document.xml.rels');
    my $doc_info = WordDocxScraper::read_doc($doc, $rel);
    ...

=head1 EXPORT

=head2 read_doc($main_doc, $relationships)

=head1 SUBROUTINES/METHODS

=head2 read_doc

This function takes two XML::LibXML documents, one corresponding to word/document.xml
and the other corresponding to word/_rels/document.xml.rels.  It then returns the 
document info it has scraped.


=cut

sub read_doc
{
    my ($main_doc, $rels) = @_;
    my $relation_ships = XML::LibXML::XPathContext->new($rels);
    $relation_ships->registerNs('a', 'http://schemas.openxmlformats.org/package/2006/relationships');

    my $xc = XML::LibXML::XPathContext->new($main_doc);

    my $paths = {
        a   => "http://schemas.openxmlformats.org/drawingml/2006/main" ,
        r   => "http://schemas.openxmlformats.org/officeDocument/2006/relationships" ,
        w   => "http://schemas.openxmlformats.org/wordprocessingml/2006/main" ,
        m   => "http://schemas.openxmlformats.org/officeDocument/2006/math" ,
        ma  => "http://schemas.microsoft.com/office/mac/drawingml/2008/main" ,
        mo  => "http://schemas.microsoft.com/office/mac/office/2008/main" ,
        mv  => "urn:schemas-microsoft-com:mac:vml" ,
        o   => "urn:schemas-microsoft-com:office:office" ,
        pic => "http://schemas.openxmlformats.org/drawingml/2006/picture" ,
        v   => "urn:schemas-microsoft-com:vml" ,
        ve  => "http://schemas.openxmlformats.org/markup-compatibility/2006" ,
        w10 => "urn:schemas-microsoft-com:office:word" ,
        wne => "http://schemas.microsoft.com/office/word/2006/wordml" ,
        wp  => "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    };
    for my $key (keys %$paths)
    {
        $xc->registerNs($key, $paths->{$key});
    }

    my @doc;

    my @paragraphs = $xc->findnodes('//w:p');
    my $prev_style;
    my $current_paragraph = {};
    for my $p (@paragraphs)
    {
        my $style = get_style($xc, $p);
        my $page_break = $xc->findnodes('w:r/w:br[@w:type="page"]', $p) ? 1 : 0;
        if(!defined $current_paragraph->{style} || $style ne $current_paragraph->{style})
        {
            push @doc, $current_paragraph if(defined $current_paragraph->{style});

            $current_paragraph = { style => $style, lines => [], page_break => $page_break };
        }
        my @pictures = $xc->findnodes('*/w:drawing', $p);
        for my $img (@pictures)
        {
            my @refs = $xc->findnodes('descendant::a:blip', $img);
            for my $ref (@refs)
            {
                my $id = $ref->getAttribute('r:embed');
                my ($node) = $relation_ships->findnodes("//a:Relationship[\@Id='$id']");
                my $url = $node->getAttribute('Target');
                $current_paragraph->{img} = [] if !defined $current_paragraph->{img};
                push @{$current_paragraph->{img}}, $url;
            }
        }

        my @bits = $xc->findnodes('descendant::w:r', $p);
        my @line;
        for my $b (@bits)
        {
            my $s = 'normal';
            $s = 'bold' if $xc->findnodes('w:rPr/w:b', $b);
            $s = 'italic' if $xc->findnodes('w:rPr/w:i', $b);
            $s = 'link' if $b->parentNode->tagName eq 'w:hyperlink';
            my @nodes = $xc->findnodes('w:t', $b);
            my @text = map { $_->textContent } @nodes;
            for my $t (@text)
            {
                $DB::single = 1 if $t eq 'http://search.cpan.org';
                push @line, { style => $s, text => $t };
            }
        }
        push @{$current_paragraph->{lines}}, \@line;
    }
    push @doc, $current_paragraph if(defined $current_paragraph->{style});
    return \@doc;
}

sub get_style
{
    my ($xc, $p) = @_;

    my @styles = $xc->findnodes('*/w:pStyle', $p);
    for my $s (@styles)
    {
        my $style;
        $style = $s->getAttribute('w:val') if $s->hasAttribute('w:val');
        return $style;
    }
    return 'normal';
}

=head1 AUTHOR

Colin Newell, C<< <colin.newell at gmail.com> >>

=head1 BUGS

This module has only been very lightly tested to work against a single document. 
It may well not work against general word docx documents.

=cut

#Please report any bugs or feature requests to C<bug-worddocxscraper at rt.cpan.org>, or through
#the web interface at L<http://rt.cpan.org/NoAuth/ReportBug.html?Queue=WordDocxScraper>.  I will be notified, and then you'll
#automatically be notified of progress on your bug as I make changes.
#



=head1 SUPPORT

You can find documentation for this module with the perldoc command.

    perldoc WordDocxScraper


You can also contact the author.

=cut

# =over 4
# 
# =item * RT: CPAN's request tracker
# 
# L<http://rt.cpan.org/NoAuth/Bugs.html?Dist=WordDocxScraper>
# 
# =item * AnnoCPAN: Annotated CPAN documentation
# 
# L<http://annocpan.org/dist/WordDocxScraper>
# 
# =item * CPAN Ratings
# 
# L<http://cpanratings.perl.org/d/WordDocxScraper>
# 
# =item * Search CPAN
# 
# L<http://search.cpan.org/dist/WordDocxScraper/>
# 
# =back
# 

=head1 ACKNOWLEDGEMENTS


=head1 LICENSE AND COPYRIGHT

Copyright 2011 Colin Newell.

This program is free software; you can redistribute it and/or modify it
under the terms of either: the GNU General Public License as published
by the Free Software Foundation; or the Artistic License.

See http://dev.perl.org/licenses/ for more information.


=cut

1; # End of WordDocxScraper
