#!/usr/bin/perl -w
use Win32::SAPI5;
use Test::More tests => 1;

my $s = Win32::SAPI5::SpVoice->new();
isa_ok ($s, 'Win32::SAPI5::SpVoice');
