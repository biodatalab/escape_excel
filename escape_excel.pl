#!/usr/bin/perl -w

# Copyright (c) 2016, 2017, 2018, 2019 Eric A. Welsh. All Rights Reserved.
#
# Escape Excel is distributed under the following BSD-style license:
#
# Redistribution and use in source and binary forms, with or without
# modification, are permitted provided that the following conditions are met:
#
# 1. Redistributions of source code must retain the above copyright notice,
#    this list of conditions and the following disclaimer.
#
# 2. Redistributions in binary form must reproduce the above copyright notice,
#    this list of conditions and the following disclaimer in the documentation
#    and/or other materials provided with the distribution.
#
# 3. The name of the author may not be used to endorse or promote products
#    derived from this software without specific prior written permission.
#
# THIS SOFTWARE IS PROVIDED BY THE CONTRIBUTORS "AS IS" AND ANY EXPRESS OR
# IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES
# OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED.
# IN NO EVENT SHALL THE AUTHOR BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL,
# SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO,
# PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS;
# OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY,
# WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR
# OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF
# ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

# 2024-03-21  treat spaces after enclosing "" as part of the enclosing ""
# 2024-03-21  only strip ="" if first or immediately inside previous =""
# 2024-03-20  improved handling of double quotes, new double quote options
# 2024-03-20  update csv2tab function, no line wrap support yet
# 2023-06-27  update is_number() to not treat NaNs as numbers
# 2021-11-02  add --csv input comma-delimited file (still output tab-delimited)
# 2020-09-25  do not strip empty fields from the ends of lines
# 2019-05-20  fixed NNeXX reversion; looks like I broke escaping 2 digits
#             before the decimal at some point
# 2018-08-10  move decimal point so as not to trigger auto-format truncation to
#             scientific notation with 2 digits after the decimal
# v1.2 2017-09-14  added --unstrip flag to restore stripped characters
# v1.1 2017-06-26  PLOS ONE publication; a few minor bug fixes, better help
# v1.0 2017-01-27  initial public release on BioRxiv
#
# TODO: decide whether we should we escape dates such as "Sep 22, 2016",
#       since that currently isn't checked for, but we may not care, since
#       if it gets autoconverted into a date it really is a date anyways?

use Scalar::Util qw(looks_like_number);
use POSIX;
use File::Basename;

$date_abbrev_hash{'jan'} = 'january';
$date_abbrev_hash{'feb'} = 'february';
$date_abbrev_hash{'mar'} = 'march';
$date_abbrev_hash{'apr'} = 'april';
$date_abbrev_hash{'may'} = 'may';
$date_abbrev_hash{'jun'} = 'june';
$date_abbrev_hash{'jul'} = 'july';
$date_abbrev_hash{'aug'} = 'august';
$date_abbrev_hash{'sep'} = 'september';
$date_abbrev_hash{'oct'} = 'october';
$date_abbrev_hash{'nov'} = 'november';
$date_abbrev_hash{'dec'} = 'december';


sub print_usage_statement
{
    $program_name = basename($0);

    printf STDERR "Usage: $program_name [options] tab_delimited_input.txt [output.txt]\n";
    printf STDERR "   Options:\n";
    printf STDERR "      --csv        input CSV file instead of tab-delimited, still outputs tsv\n";
    printf STDERR "      --escape-dq  \"smart\" escaping of \" to better preserve them (default)\n";
    printf STDERR "      --no-dates   Do not escape text that looks like dates and/or times\n";
    printf STDERR "      --no-dq      Disable \"smart\" handling of double quotes\n";
    printf STDERR "      --no-sci     Do not escape >= ##E (ex: 12E4) or >11 digit integer parts\n";
    printf STDERR "      --no-zeroes  Do not escape leading zeroes (ie. 012345)\n";
    printf STDERR "      --unstrip    restore auto-stripped field when not escaped\n";
    printf STDERR "      --paranoid   Escape *ALL* non-numeric text (overrides --no-dates)\n";
    printf STDERR "                   WARNING -- Excel can take a LONG time to import\n";
    printf STDERR "                   text files where most fields are escaped.\n";
    printf STDERR "                   Copy / Paste Values can become near unusuable....\n";
    printf STDERR "\n";
    printf STDERR "   Reads input from STDIN if input file name is - or no file name is given.\n";
    printf STDERR "   Input file must be tab-delimited.\n";
    printf STDERR "   Fields will be stripped of existing =\"\" escapes, enclosing \"\", leading \",\n";
    printf STDERR "    and leading/trailing spaces, as they may all cause problems.\n";
    printf STDERR "\n";
    printf STDERR "   Defaults to escaping most Excel mis-imported fields.\n";
    printf STDERR "   Escapes a few extra date-like formats that Excel does not consider dates.\n";
    printf STDERR "   Please send unhandled mis-imported field examples (other than gene symbols\n";
    printf STDERR "    with 1-digit scientific notation, such as 2e4) to Eric.Welsh\@moffitt.org.\n";
    printf STDERR "\n";
    printf STDERR "   Does not currently support or detect line wraps.  To best preserve line\n";
    printf STDERR "   wraps, use --no-dq and enable --unstrip to better preserve spacing.\n";
    printf STDERR "\n";
    printf STDERR "   Copy / Paste Values in Excel, after importing, to de-escape back into text.\n";
}


sub is_number
{
    # use what Perl thinks is a number first
    # this is purely for speed, since the more complicated REGEX below should
    #  correctly handle all numeric cases
    if (looks_like_number($_[0]))
    {
        # Perl treats infinities as numbers, Excel does not.
        #
        # Perl treats NaN or NaNs, and various mixed caps, as numbers.
        # Weird that not-a-number is a number... but it is so that
        # it can do things like nan + 1 = nan, so I guess it makes sense
        #
        if ($_[0] =~ /^[-+]*(Inf|NaN)/i)
        {
            return 0;
        }
        
        return 1;
    }

    # optional + or - sign at beginning
    # then require either:
    #  a number followed by optional comma stuff, then optional decimal stuff
    #  mandatory decimal, followed by optional digits
    # then optional exponent stuff
    #
    # Perl cannot handle American comma separators within long numbers.
    # Excel does, so we have to check for it.
    # Excel doesn't handle European dot separators, at least not when it is
    #  set to the US locale (my test environment).  I am going to leave this
    #  unsupported for now.
    #
    if ($_[0] =~ /^([-+]?)([0-9]+(,[0-9]{3,})*\.?[0-9]*|\.[0-9]*)([Ee]([-+]?[0-9]+))?$/)
    {
        # current REGEX can treat '.' as a number, check for that
        if ($_[0] eq '.')
        {
            return 0;
        }
        
        return 1;
    }
    
    return 0;
}


sub csv2tsv_not_excel
{
    my $line = $_[0];
    my $i;
    my $n;
    my @temp_array;

    # placeholder strings unlikely to ever be encountered normally
    #    my $tab = '___TaBChaR___';
    #    my $dq  = '___DqUOteS___';
    #
    my $tab = "\x19";    # need single char for split regex
    my $dq  = "\x16";    # need single char for split regex
    
    # remove null characters, since I've only seen them randomly introduced
    #  during NTFS glitches; "Null characters may be inserted into or removed
    #  from a stream of data without affecting the information content of that
    #  stream."
    $line =~ s/\x00//g;
    
    # remove UTF8 byte order mark, since it corrupts the first field
    # also remove some weird corruption of the UTF8 byte order mark (??)
    #
    # $line =~ s/^\xEF\xBB\xBF//;      # actual BOM
    # $line =~ s/^\xEF\x3E\x3E\xBF//;  # corrupted BOM I have seen in the wild
    $line =~ s/^(\xEF\xBB\xBF|\xEF\x3E\x3E\xBF)//;

    # replace any (incredibly unlikely) instances of $dq with $tab
    $line =~ s/$dq/$tab/g;
    
    # replace embedded tabs with placeholder string, to deal with better later
    $line =~ s/\t/$tab/g;
    
    # HACK -- handle internal \r and \n the same way we handle tabs
    $line =~ s/[\r\n]+(?!$)/$tab/g;
    
    # further escape ""
    $line =~ s/""/$dq/g;

    # only apply slow tab expansion to lines still containing quotes
    if ($line =~ /"/)
    {
        # convert commas only if they are not within double quotes
        # incrementing $i within array access for minor speed increase
        #   requires initializing things to -2
        #@temp_array = split /((?<![^, $tab$dq])"[^\t"]+"(?![^, $tab$dq\r\n]))/, $line;
        @temp_array = split /((?:^|(?<=,))[ $tab$dq]*\"[^\t"]+\"[ $tab$dq\r\n]*(?:$|(?=,)))/, $line, -1;
        $n = @temp_array - 2;
        for ($i = -2; $i < $n;)
        {
            $temp_array[$i += 2] =~ tr/,/\t/;
        }
        $line = join '', @temp_array;
        
        # slightly faster than split loop on rows with many quoted fields,
        #  but *much* slower on lines containing very few quoted fields
        # use split loop instead
        #
        # /e to evaluates code to handle different capture cases correctly
        #$line =~ s/(,?)((?<![^, $tab$dq])"[^\t"]+"(?![^, $tab$dq\r\n]))|(,)/defined($3) ? "\t" : ((defined($1) && $1 ne '') ? "\t$2" : $2)/ge;
    }
    else
    {
        $line =~ tr/,/\t/;
    }

    # unescape ""
    $line =~ s/$dq/""/g;

    # finish dealing with embedded tabs
    # remove tabs entirely, preserving surrounding whitespace
    $line =~ s/(\s|^)($tab)+/$1/g;
    $line =~ s/($tab)+(\s|$)/$2/g;
    # replace remaining tabs with spaces so text doesn't abutt together
    $line =~ s/($tab)+/ /g;

    # Special case "" in a field by itself.
    #
    # This generally results from lazily-coded csv writers that enclose
    #  every single field in "", even empty fields, whether they need them
    #  or not.
    #
    # \K requires Perl >= v5.10.0 (2007-12-18)
    #   (?: *)\K is faster than replacing ( *) with $1
    $line =~ s/(?:(?<=\t)|^)(?: *)\K""( *)(?=\t)/$1/g;  # start|tabs "" tabs
    $line =~    s/(?<=\t)(?: *)\K""( *)(?=\t|$)/$1/g;   # tabs  "" tabs|end
    $line =~          s/^(?: *)\K""( *)$/$1/g;          # start "" end

    # strip enclosing double-quotes, preserve leading/trailing spaces
    #
    # \K requires Perl >= v5.10.0 (2007-12-18)
    #   (?: *)\K is faster than replacing ( *) with $1
    #
    #$line =~ s/(?:(?<=\t)|^)(?: *)\K"([^\t]+)"( *)(?=\t|[\r\n]*$)/$1$2/g;

    # remove enclosing spaces, to support space-justified quoted fields
    #   ( *) might be faster without () grouping, but left in for clarity
    $line =~ s/(?:(?<=\t)|^)( *)"([^\t]+)"( *)(?=\t|[\r\n]*$)/$2/g;

    # unescape escaped double-quotes
    $line =~ s/""/"/g;

    
    return $line;
}


sub has_text_month
{
    my $date_str = $_[0];
    my $abbrev;
    my $full;
    my $xor;
    my $prefix_length;

    $candidate = '';
    if ($date_str =~ /^([0-9]{1,4}[- \/]*)?([A-Za-z]{3,9})/)
    {
        $candidate = lc $2;
    }

    if ($candidate eq '')
    {
        return 0;
    }

    $abbrev = substr $candidate, 0, 3;
    $full = $date_abbrev_hash{$abbrev};

    # first three letters are not the start of a month
    if (!defined($full))
    {
        return 0;
    }

    # find common prefix
    $xor = "$candidate" ^ "$full";
    $xor =~ /^\0*/;
    $prefix_length = $+[0];
    
    # if the common prefix is the same as the full candidate, it is real
    if (length $candidate eq $prefix_length)
    {
        return 1;
    }
    
    return 0;
}

$opt_csv = 0;
$opt_escape_excel_paranoid = 0;
$opt_escape_sci            = 1;
$opt_escape_zeroes         = 1;
$opt_escape_dates          = 1;
$opt_escape_dq             = 1;
$opt_unstrip               = 0;

# read in command line arguments
$num_files = 0;
for ($i = 0; $i < @ARGV; $i++)
{
    $field = $ARGV[$i];

    if ($field =~ /^--/)
    {
        if ($field eq '--paranoid')
        {
            if ($opt_escape_excel_paranoid == 0)
            {
                $opt_escape_excel_paranoid = 1;
            }
            else
            {
                $opt_escape_excel_paranoid = 0;
            }
        }
        elsif ($field eq '--no-sci')
        {
            $opt_escape_sci = 0;
        }
        elsif ($field eq '--no-zeroes')
        {
            $opt_escape_zeroes = 0;
        }
        elsif ($field eq '--no-dates')
        {
            $opt_escape_dates = 0;
        }
        elsif ($field eq '--no-dq')
        {
            $opt_escape_dq = 0;
        }
        elsif ($field eq '--escape-dq')
        {
            $opt_escape_dq = 1;
        }
        elsif ($field eq '--unstrip')
        {
            $opt_unstrip = 1;
        }
        elsif ($field eq '--csv')
        {
            $opt_csv = 1;
        }
        else
        {
            printf "ABORT -- unknown option %s\n", $field;
            $syntax_error_flag = 1;
        }
    }
    else
    {
        if ($num_files == 1)
        {
            $outname = $field;
            $num_files++;
        }
        if ($num_files == 0)
        {
            $filename = $field;
            $num_files++;
        }
    }
}

# default to stdin if no filename given
if ($num_files == 0)
{
    $filename = '-';
    $num_files = 1;
}


# print syntax error message
if ($num_files == 0 || $syntax_error_flag)
{
    print_usage_statement();
    exit(1);
}


# output to STDOUT
if ($num_files == 1)
{
    *OUTFILE = STDOUT;
}
# output to specified file name
if ($num_files == 2)
{
    # do not allow output file names starting with -
    if ($outname =~ /^-/)
    {
        print STDERR "ABORT -- output file names beginning with hyphens are not allowed\n";
        print_usage_statement();

        exit(1);
    }

    $temp_flag = open OUTFILE, ">$outname";

    if (!$temp_flag)
    {
        print STDERR "ABORT -- can't open output $outname\n";

        exit(3);
    }
}


$temp_flag = open INFILE, "$filename";

if (!$temp_flag)
{
    print STDERR "ABORT -- can't open input $filename\n";

    if ($filename =~ /^-/)
    {
        print_usage_statement();
    }

    exit(2);
}

# read in, escape, and print escaped lines
while(defined($line=<INFILE>))
{
    # strip newline characters
    $line =~ s/[\r\n]+//g;
    
    # convert comma-delimited to tab-delimited
    if ($opt_csv)
    {
        $line = csv2tsv_not_excel($line);
    }

    # do NOT strip empty fields from the end of the line
    @array = split /\t/, $line, -1;

    # Strip any leading UTF-8 byte order mark so it won't corrupt the
    #  first field, since regular Perl I/O is not byte order mark aware.
    #
    # https://en.wikipedia.org/wiki/Byte_order_mark
    #
    # Various Microsoft products can emit these and screw things up....
    #
    for ($i = 0; $i < @array; $i++)
    {
        $line =~ s/^﻿//;
    }
    
    for ($i = 0; $i < @array; $i++)
    {
        $original_field      = $array[$i];
        $needs_escaping_flag = 0;
        $was_enclosed_flag   = 0;
        $last_strip_rule     = '';

        # NOTE -- Decide how best to handle ""
        #         Do we treat it as 2 double-quotes, or as an empty field?
        #         Pure tsv without quotes would expect it to be 2 double-quotes
        #         Excel/Python interpret it as empty field, but would not export such
        #
        #         For now, treat it as 2 double-quotes, since some software may
        #         export "" as meaning 2 double-quotes, but nothing should be
        #         exporting "" as meaning blank?

        # Remove only outer layer of enclosing double quotes
        # Excel treats spaces *after* the "" as still enclosed in ""
        # However, it does not treat spaces *before* the "" as enclosed in ""
        #
        # Treat "" as 2 double-quotes, rather than an empty field
        if ($array[$i] =~ s/^\"(.+?)\" *$/$1/)
        {
            $was_enclosed_flag = 1;
            $last_strip_rule   = '""';
        }
        
        # continue stripping problematic stuff until all has been stripped
        do
        {
            $changed_flag = 0;
            
            # remove pre-existing escapes
            if ($last_strip_rule eq '' ||
                $last_strip_rule eq '=""')
            {
                while ($array[$i] =~ s/^\=\"(.*?)\"$/$1/)
                {
                    $was_enclosed_flag = 1;
                    $changed_flag      = 1;
                    $last_strip_rule   = '=""';
                }
            }
        } while ($changed_flag);

        # remove leading spaces, since they won't protect long numbers,
        # and will cause various REGEX to fail
        #
        # remove trailing spaces, since they won't protect dates,
        # and will cause various REGEX to fail
        $array[$i] =~ s/^\s+//;
        $array[$i] =~ s/\s+$//;


        # escape fields
        #
        # Strange but true -- 'text doesn't escape text properly in Excel
        # when you try to use it in a text file to import.  It will not
        # auto-strip the leading ' like it does when you type it in a live
        # spreadsheet.  "text" doesn't, either.  Oddly, ="text" DOES work,
        # but an equation containing just a text string and no actual
        # equation doesn't make much sense.  However, it works, so that's
        # what I use here to escape fields into mangle-protected text.

        # escape numeric problems
        if (is_number($array[$i]))
        {
          # keep leading zeroes for >1 digit before the decimal point
          if ($opt_escape_zeroes && $array[$i] =~ /^([-+]?)0[0-9]/)
          {
              $needs_escaping_flag = 1;
          }

          # Escape scientific notation with >= 2 digits before the E,
          #  since they are likely accessions or plate/well identifiers.
          #
          # Also escape whole numbers with >11 digits before the decimal.
          # >11 is when it displays scientific notation in General format,
          #  which can result in corruption when saved to text.
          # >15 would be the limit at which it loses precision internally.
          #
          # NOTE -- if there is a + or - at the beginning, this rule
          #         will not trigger.  Undecided if this is desired or not.
          #         Probably desired behavior, since +/- would indicate that
          #         it is probably a true number, and not an accession or
          #         plate/well identifier.
          #
          elsif ($opt_escape_sci)
          {
              # strip commas before counting digits
              $temp = $array[$i];
              $temp =~ s/\,//g;
              $temp = abs($temp);
              
              if ($temp =~ /^([1-9][0-9]{11,}|[0-9]{2,}[eE])/)
              {
                  $needs_escaping_flag = 1;
              }
              elsif ($temp >= 1E11 &&
                  !($temp =~ /^[-+]/) &&
                  $temp - floor($temp + 0.5) == 0)
              {
                  # count number of significant digits
                  $len = length $temp;
                  $sigdigits = 0;
                  $k = 0;
                  for ($j = 0; $j < $len; $j++)
                  {
                      $c = substr $temp, $j, 1;
                      
                      if ($c ne '.')
                      {
                          $k++;
                      }
                      
                      if ($c ne '0')
                      {
                          $sigdigits = $k + 1;
                      }
                  }
                  
                  if ($sigdigits >= 11)
                  {
                      $needs_escaping_flag = 1;
                  }
              }
          }
        }
        # escape all text if paranoid
        elsif ($opt_escape_excel_paranoid)
        {
          $needs_escaping_flag = 1;
        }
        # escape only text that might be corrupted
        else
        {
          # escape single quote at beginning of line
          if ($array[$i] =~ /^'/)
          {
              $needs_escaping_flag = 1;
          }

          # prevent conversion into formulas
          elsif ($array[$i] =~ /^\=/)
          {
              $needs_escaping_flag = 1;
          }
          # Excel is smart enough to treat all +/- as not an equation
          #  but, otherwise, it will convert anything starting with +/-
          #  into "#NAME?" as a failed invalid equation
          elsif ($array[$i] =~ /^[-+]/ && !($array[$i] =~ /^[-+]+$/))
          {
              $needs_escaping_flag = 1;
          }

          # check for time and/or date stuff
          elsif ($opt_escape_dates)
          {
              $time = '';
              $date = '';
          
              # attempt to guess at how excel might autoconvert into time
              # allow letter/punctuation at end if it could be part of a date
              #  it would get too complicated to handle date-ness correctly,
              #  since I'm already resorting to negative look-ahead
              if ($array[$i] =~ /\b(([0-9]+\s+(AM|PM|A|P)|[0-9]+:[0-9]+(:[0-9.]+)?)(\s+(AM|PM|A|P))?)(?!([^-\/, 0-9ADFJMNOSadfjmnos]))/)
              {
                  $time = $1;
              }
              
              $strip_time = $array[$i];
              if ($time =~ /\S/)
              {
                  $strip_time =~ s/\Q$time\E//;
                  $strip_time =~ s/^\s+//;
                  $strip_time =~ s/\s+$//
              }

              # text date, month in the middle
              if ($strip_time =~ /\b([0-9]{1,4}[- \/]*Jan[A-Za-z]{0,6}([- \/]*[0-9]{1,4})?)\b/i ||
                  $strip_time =~ /\b([0-9]{1,4}[- \/]*Feb[A-Za-z]{0,6}([- \/]*[0-9]{1,4})?)\b/i ||
                  $strip_time =~ /\b([0-9]{1,4}[- \/]*Mar[A-Za-z]{0,6}([- \/]*[0-9]{1,4})?)\b/i ||
                  $strip_time =~ /\b([0-9]{1,4}[- \/]*Apr[A-Za-z]{0,6}([- \/]*[0-9]{1,4})?)\b/i ||
                  $strip_time =~ /\b([0-9]{1,4}[- \/]*May[A-Za-z]{0,6}([- \/]*[0-9]{1,4})?)\b/i ||
                  $strip_time =~ /\b([0-9]{1,4}[- \/]*Jun[A-Za-z]{0,6}([- \/]*[0-9]{1,4})?)\b/i ||
                  $strip_time =~ /\b([0-9]{1,4}[- \/]*Jul[A-Za-z]{0,6}([- \/]*[0-9]{1,4})?)\b/i ||
                  $strip_time =~ /\b([0-9]{1,4}[- \/]*Aug[A-Za-z]{0,6}([- \/]*[0-9]{1,4})?)\b/i ||
                  $strip_time =~ /\b([0-9]{1,4}[- \/]*Sep[A-Za-z]{0,6}([- \/]*[0-9]{1,4})?)\b/i ||
                  $strip_time =~ /\b([0-9]{1,4}[- \/]*Oct[A-Za-z]{0,6}([- \/]*[0-9]{1,4})?)\b/i ||
                  $strip_time =~ /\b([0-9]{1,4}[- \/]*Nov[A-Za-z]{0,6}([- \/]*[0-9]{1,4})?)\b/i ||
                  $strip_time =~ /\b([0-9]{1,4}[- \/]*Dec[A-Za-z]{0,6}([- \/]*[0-9]{1,4})?)\b/i)
              {
                  $temp = $1;
              
                  if (has_text_month($temp))
                  {
                      $date = $temp;
                  }
              }

              # text date, month first
              elsif ($strip_time =~ /\b(Jan[A-Za-z]{0,6}[- \/]*[0-9]{1,4}([- \/]+[0-9]{1,4})?)\b/i ||
                     $strip_time =~ /\b(Feb[A-Za-z]{0,6}[- \/]*[0-9]{1,4}([- \/]+[0-9]{1,4})?)\b/i ||
                     $strip_time =~ /\b(Mar[A-Za-z]{0,6}[- \/]*[0-9]{1,4}([- \/]+[0-9]{1,4})?)\b/i ||
                     $strip_time =~ /\b(Apr[A-Za-z]{0,6}[- \/]*[0-9]{1,4}([- \/]+[0-9]{1,4})?)\b/i ||
                     $strip_time =~ /\b(May[A-Za-z]{0,6}[- \/]*[0-9]{1,4}([- \/]+[0-9]{1,4})?)\b/i ||
                     $strip_time =~ /\b(Jun[A-Za-z]{0,6}[- \/]*[0-9]{1,4}([- \/]+[0-9]{1,4})?)\b/i ||
                     $strip_time =~ /\b(Jul[A-Za-z]{0,6}[- \/]*[0-9]{1,4}([- \/]+[0-9]{1,4})?)\b/i ||
                     $strip_time =~ /\b(Aug[A-Za-z]{0,6}[- \/]*[0-9]{1,4}([- \/]+[0-9]{1,4})?)\b/i ||
                     $strip_time =~ /\b(Sep[A-Za-z]{0,6}[- \/]*[0-9]{1,4}([- \/]+[0-9]{1,4})?)\b/i ||
                     $strip_time =~ /\b(Oct[A-Za-z]{0,6}[- \/]*[0-9]{1,4}([- \/]+[0-9]{1,4})?)\b/i ||
                     $strip_time =~ /\b(Nov[A-Za-z]{0,6}[- \/]*[0-9]{1,4}([- \/]+[0-9]{1,4})?)\b/i ||
                     $strip_time =~ /\b(Dec[A-Za-z]{0,6}[- \/]*[0-9]{1,4}([- \/]+[0-9]{1,4})?)\b/i)
              {
                  $temp = $1;

                  if (has_text_month($temp))
                  {
                      $date = $temp;
                  }
              }

              # possibly a numeric date
              elsif ($strip_time =~ /\b([0-9]{1,4}[- \/]+[0-9]{1,2}[- \/]+[0-9]{1,2})\b/ ||
                     $strip_time =~ /\b([0-9]{1,2}[- \/]+[0-9]{1,4}[- \/]+[0-9]{1,2})\b/ ||
                     $strip_time =~ /\b([0-9]{1,2}[- \/]+[0-9]{1,2}[- \/]+[0-9]{1,4})\b/ ||
                     $strip_time =~ /\b([0-9]{1,2}[- \/]+[0-9]{1,4})\b/ ||
                     $strip_time =~ /\b([0-9]{1,4}[- \/]+[0-9]{1,2})\b/)
              {
                  $date = $1;
              }
              
              # be sure that date and time anchor the ends
              # mix of time and date
              if ($time =~ /\S/ && $date =~ /\S/)
              {
                  if ($array[$i] =~ /^\Q$time\E(.*)\Q$date\E$/ ||
                      $array[$i] =~ /^\Q$date\E(.*)\Q$time\E$/)
                  {
                      $middle = $1;

                      # allow blank
                      # allow for purely whitespace
                      # allow for a single hyphen, slash, comma
                      #  allow for multiple spaces before and/or after
                      if ($middle eq '' ||
                          $middle =~ /^\s+$/ ||
                          $middle =~ /^\s*[-\/,]\s*$/)
                      {
                          $needs_escaping_flag = 1;
                      }
                  }
              }
              # only time
              elsif ($time =~ /\S/)
              {
                  if ($array[$i] =~ /^\Q$time\E$/)
                  {
                      $needs_escaping_flag = 1;
                  }
              }
              # only date
              elsif ($date =~ /\S/)
              {
                  if ($array[$i] =~ /^\Q$date\E$/)
                  {
                      $needs_escaping_flag = 1;
                  }
              }
          }
        }
        
        # n.nnEn, nn.nEn, etc.
        #
        # If it looks like scientific notation, Excel will automatically
        #  format it as scientific notation.  However, if the magnitude is
        #  > 1E7, it will also automatically display it set to only 2 digits
        #  to the right of the decimal (it is still fine internally).  If the
        #  file is re-exported to text, truncation to 3 significant digits
        #  will occur!!!
        #
        # Reformat the number to (mostly) avoid this behavior.
        #
        # Unfortunately, the >= 11 significant digits behavior still
        #  triggers, so it still truncates to 10 digits when re-exporting
        #  General format.  10 digits is still better than 3....
        #
        # The re-export truncation behavior can only be more fully avoided by
        #  manually setting the format to Numeric and specifying a large
        #  number of digits after the decimal place for numbers with
        #  fractions, or 0 digits after the decimal for whole numbers.
        #
        # Ugh.
        #
        # There is no fully fixing this brain-damagedness automatically,
        #  I can only decrease the automatic truncation of significant
        #  digits from 3 to 10 digits :(  Any precision beyond 10 digits
        #  *WILL* be lost on re-export if the format is set to General.
        #
        # NOTE -- We truncate to 16 significant digits by going through
        #         a standard IEEE double precision intermediate.
        #         However, Excel imports numbers as double precision
        #         anyways, so we aren't losing any precision that Excel
        #         wouldn't already be discarding.
        #
        if ($needs_escaping_flag == 0 && is_number($array[$i]))
        {
              # strip commas
              $temp = $array[$i];
              $temp =~ s/\,//g;
              
              if (abs($temp) >= 1 &&
                  $temp =~ /^([-+]?[0-9]*\.*[0-9]*)[Ee]([-+]?[0-9]+)$/)
              {
                  #$number   = $1 + 0;
                  #$exponent = $2 + 0;
                  
                  $temp /= 1;
                  
                  # overwrite saved unstripped version, if it wasn't stripped
                  if ($opt_unstrip &&
                      $array[$i] eq $original_field)
                  {
                      $original_field = $temp;
                  }

                  # replace original with new scientific notation format
                  $array[$i] = $temp;
              }
        }


        # quotes need to be escaped to protect them from line wraps
        $count_dq1 = () = $array[$i] =~ m/\"/g;
        $count_dq2 = () = $array[$i] =~ m/\"\"/g;
        $unescaped_dq_flag           =  0;
        if ($count_dq1 &&
            ($was_enclosed_flag == 0 ||
             $count_dq1 != 2 * $count_dq2))
        {
            $unescaped_dq_flag = 1;
        }
        
        # de-escape already properly escaped double quotes
        if ($opt_escape_dq && $was_enclosed_flag && $unescaped_dq_flag == 0)
        {
            $array[$i] =~ s/\"\"/\"/g;
        }
        
        ## restore double-quoted empty field to double quotes
        #if ($original_field eq '""' && $opt_escape_dq == 0)
        #{
        #    $array[$i] = '""';
        #}

        if ($needs_escaping_flag)
        {
            if ($opt_escape_dq && $count_dq1)
            {
                $array[$i] =~ s/\"/\"\"/g;
            }

            $array[$i] = sprintf "=\"%s\"", $array[$i];
        }
        else
        {
            if ($opt_unstrip)
            {
                $array[$i] = $original_field;
            }

            if ($opt_escape_dq && $count_dq1)
            {
                # Remove pre-existing escape or start/end double quotes
                # Include space after terminal " in the pattern, since
                # Excel still treats such fields as enclosed in ""
                if ($opt_unstrip && $was_enclosed_flag)
                {
                    $array[$i] =~ s/^\=*\"(.*?)\" *$/$1/;
                }

                # escape " as "" only if they need it
                # properly escaped original fields don't need re-escaping
                if ($opt_unstrip == 0 || $unescaped_dq_flag)
                {
                    $array[$i] =~ s/\"/\"\"/g;
                }

                $array[$i] =  sprintf "\"%s\"", $array[$i];
                
                # re-insert the equal sign
                if ($opt_unstrip && $was_enclosed_flag &&
                    $original_field =~ /^\=/)
                {
                    $array[$i] = '=' . $array[$i];
                }
            }
            # restore enclosing quotes for properly escaped fields
            elsif ($count_dq1 && $was_enclosed_flag && $opt_unstrip == 0)
            {
                $array[$i] =  sprintf "\"%s\"", $array[$i];
            }
        }
    }
    
    # make the new escaped line
    $line_escaped = join "\t", @array;
    
    # print it
    print OUTFILE "$line_escaped\n";
}
close INFILE;

close OUTFILE;
