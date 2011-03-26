use inc::Module::Install;

name     'WordDocxScraper';
all_from 'lib/WordDocxScraper.pm';
author   q{Colin Newell <colin.newell@gmail.com>};
license  'perl';

requires 'XML::LibXML';
requires 'Text::Wrap';

build_requires 'Test::More';

auto_install;

WriteAll;

