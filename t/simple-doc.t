use Test::Most;

use WordDocxScraper;
use XML::LibXML;
use FindBin;
use lib "$FindBin::Bin/../lib";


binmode STDOUT, ":utf8";

my $file = "$FindBin::Bin/document.xml";
my $rels = "$FindBin::Bin/document.xml.rels";
my $dom = XML::LibXML->load_xml(location => $file);
my $relationships = XML::LibXML->load_xml(location => $rels);
my $doc = WordDocxScraper::read_doc($dom, $relationships);

eq_or_diff($doc, 
        [
        {
        lines => [
        [
        {
        style => 'normal',
        text => 'A Heading'
        }
        ]
        ],
        page_break => 0,
        style => 'style1'
        },
        {
        lines => [
        [],
        [
        {
        style => 'normal',
        text => 'What the heck'
        }
        ],
            [],
            [
            {
                style => 'link',
                text => 'http://www.google.com'
            }
        ],
            []
                ],
            page_break => 0,
            style => 'style0'
        },
{
    lines => [
        [
        {
            style => 'normal',
            text => 'Test this'
        }
    ]
        ],
        page_break => 0,
        style => 'style2'
},
{
    lines => [
        [],
    [
    {
        style => 'normal',
        text => '# a bit of code'
    }
    ],
        [
        {
            style => 'normal',
            text => 'my $test = \'blah\';'
        }
    ],
        []
            ],
        page_break => 0,
        style => 'style0'
}
]
);

done_testing;

