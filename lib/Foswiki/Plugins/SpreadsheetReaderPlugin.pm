# Plugin for Foswiki - The Free and Open Source Wiki, https://foswiki.org/
#
# SpreadsheetReaderPlugin is Copyright (C) 2018 Michael Daum http://michaeldaumconsulting.com
#
# This program is free software; you can redistribute it and/or
# modify it under the terms of the GNU General Public License
# as published by the Free Software Foundation; either version 2
# of the License, or (at your option) any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
# GNU General Public License for more details, published at
# http://www.gnu.org/copyleft/gpl.html

package Foswiki::Plugins::SpreadsheetReaderPlugin;

use strict;
use warnings;

use Foswiki::Func ();

our $VERSION = '1.10';
our $RELEASE = '28 May 2018';
our $SHORTDESCRIPTION = 'Read data from a spreadsheet';
our $NO_PREFS_IN_TOPIC = 1;
our $core;

sub initPlugin {

  Foswiki::Func::registerTagHandler('SPREADSHEET', sub { return getCore()->SPREADSHEET(@_); });

  return 1;
}

sub getCore {
  unless (defined $core) {
    require Foswiki::Plugins::SpreadsheetReaderPlugin::Core;
    $core = Foswiki::Plugins::SpreadsheetReaderPlugin::Core->new();
  }
  return $core;
}

sub finishPlugin {
  undef $core;
}

1;
