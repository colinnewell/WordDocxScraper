#!perl -T

use Test::More tests => 1;

BEGIN {
    use_ok( 'WordDocxScraper' ) || print "Bail out!
";
}

diag( "Testing WordDocxScraper $WordDocxScraper::VERSION, Perl $], $^X" );
