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

package Foswiki::Plugins::SpreadsheetReaderPlugin::Core;

use strict;
use warnings;
use Spreadsheet::Read ();
use Foswiki::Func ();
#use Data::Dump qw(dump);
use Error qw(:try);

use constant TRACE => 0; # toggle me

sub new {
  my $class = shift;

  my $this = bless({
    @_
  }, $class);

  return $this;
}

sub SPREADSHEET {
  my ($this, $session, $params, $topic, $web) = @_;

  _writeDebug("called SPREADSHEET()");
  my $attachment = $params->{_DEFAULT} || $params->{attachment};

  return _inlineError("no attachment specified") unless defined $attachment;

  my ($theWeb, $theTopic) = Foswiki::Func::normalizeWebTopicName($params->{web} || $web, $params->{topic} || $topic);

  return _inlinError("access denied") unless
        Foswiki::Func::checkAccessPermission("VIEW", $session->{user}, undef, $theTopic, $theWeb);

  return _inlineError("topic not found") unless Foswiki::Func::topicExists($theWeb, $theTopic);

  my ($fileName) = Foswiki::Func::sanitizeAttachmentName($attachment);
  return _inlineError("attachment not found") unless Foswiki::Func::attachmentExists($theWeb, $theTopic, $fileName);

  my $filePath = $Foswiki::cfg{PubDir} . '/' . $web . '/' . $topic . '/' . $fileName;
  _writeDebug("filePath=$filePath");

  return _inlineError("file not found") unless -e $filePath;

  my $opts = {
    strip => 1,
    #dtfmt => "dd.mm.yyyy",  # SMELL: doesn't seem to make any difference
    attr => 1,
  };
  $opts->{password} = $params->{password} if defined $params->{password};
  $opts->{sep} = $params->{sep} if defined $params->{sep};
  $opts->{quote} = $params->{quote} if defined $params->{quote};

  my $error;
  my $book;
  try {
    $book = Spreadsheet::Read->new($filePath, $opts);
  } catch Error with {
    $error = shift;
    $error =~ s/ at .*$//;
  };
  return _inlineError($error) if defined $error;


  my $theSheets = $params->{sheets} || $params->{sheet};
  my $maxSheets = $book->sheets;
  _writeDebug("maxSheets=$maxSheets");

  return _inlineError("spreadsheet contains no sheets") unless $maxSheets > 0;

  my $theShowAttrs = Foswiki::Func::isTrue($params->{showattrs}, 1);
  my $theShowIndex = Foswiki::Func::isTrue($params->{showindex}, 0);

  my @selectedSheets = ();
  if (defined $theSheets) {
    if ($theSheets eq 'all') {
      push @selectedSheets, $_ foreach 1..$maxSheets;
    } else {
      @selectedSheets = split(/\s*,\s*/, $theSheets);
    }
  } else {
    push @selectedSheets, 1;
  }

  my @selectedRows = ();
  my $theRows = $params->{rows} || 'all';
  if ($theRows ne "all") {
    @selectedRows = split(/\s*,\s*/, $theRows);
  }

  my @selectedCols = ();
  my $theCols = $params->{cols} || 'all';
  if ($theCols ne "all") {
    @selectedCols = split(/\s*,\s*/, $theCols);
  }
  

  #_writeDebug("selectedSheets=@selectedSheets");
  my $class = $params->{class} // "foswikiTable";

  my @result = ();
  foreach my $sheetId (@selectedSheets) {
    my $sheet = $book->sheet($sheetId);
    my $maxRows = $sheet->maxrow;
    my $maxCols = $sheet->maxcol;
    _writeDebug("reading sheet $sheetId, maxRows=$maxRows, maxCols=$maxCols");

    if ($theRows eq "all") {
      @selectedRows = ();
      push @selectedRows, $_ foreach 1..$maxRows;
    }

    if ($theCols eq "all") {
      @selectedCols = ();
      push @selectedCols, $_ foreach 1..$maxCols;
    }

    my $isHeader = 1;
    push @result, "<table class='$class'><thead>";
    my $index = 1;
    foreach my $row (@selectedRows) {
      my @cells = ();
      if ($theShowIndex) {
        if ($isHeader) {
          push @cells, "<th>#</th>";
        } else {
          push @cells, "<th>$index</th>";
          $index++;
        }
      }
      foreach my $col (@selectedCols) {
	my $key = $book->cr2cell($col, $row);
	my $val = $sheet->cell($key);
	my $rawVal = $sheet->cell($col, $row) || '';
	my $attr = $sheet->attr($key);
	my $type = $attr->{type} || '';
	_writeDebug("$key=$val, raw=$rawVal, type=$type");
	#_writeDebug(dump($attr));
	my $align = $attr->{halign};	
	$align = "right" if !defined $align && $type eq "numeric";
        my $styles = "";
        if ($theShowAttrs) {
          my @styles = ();
          push @styles, "background-color:$attr->{bgcolor}" if $attr->{bgcolor};
          push @styles, "color:$attr->{fgcolor}" if $attr->{fgcolor};
          push @styles, "font-weight:bold" if $attr->{bold};
          push @styles, "font-style:italic" if $attr->{italic};
          push @styles, "text-align:$align" if $align;
          $styles = @styles?"style='".join(";", @styles)."'":"";
        }
	push @cells, $isHeader?"<th $styles> $val </th>":"<td $styles> $val </td>";
      }
      my $line = "<tr>". join("", @cells). "</tr>";
      #_writeDebug("line=$line");
      push @result, $line;
      push @result, "</thead><tbody>" if $isHeader;
      $isHeader = 0;
    }
  }
  push @result, "</tbody></table>";

  return join("\n", @result);
}

sub _inlineError {
  return "<span class='foswikiAlert'>ERROR: @_</span>";
}

sub _writeDebug {
  return unless TRACE;
  #Foswiki::Func::writeDebug("SpreadsheetReaderPlugin::Core - $_[0]");
  print STDERR "SpreadsheetReaderPlugin::Core - $_[0]\n";
}


1;
