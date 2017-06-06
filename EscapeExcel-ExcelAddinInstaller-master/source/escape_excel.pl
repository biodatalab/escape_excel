#!/usr/bin/perl -w

# Copyright (c) 2016, 2017 Eric A. Welsh. All Rights Reserved.
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


use Scalar::Util qw(looks_like_number);

$date_abbrev_hash{'jan'} = 'january';
$date_abbrev_hash{'feb'} = 'february';
$date_abbrev_hash{'mar'} = 'march';
$date_abbrev_hash{'apr'} = 'april';
$date_abbrev_hash{'may'} = 'may';
$date_abbrev_hash{'jun'} = 'jun';
$date_abbrev_hash{'jul'} = 'july';
$date_abbrev_hash{'aug'} = 'august';
$date_abbrev_hash{'sep'} = 'september';
$date_abbrev_hash{'oct'} = 'october';
$date_abbrev_hash{'nov'} = 'november';
$date_abbrev_hash{'dec'} = 'december';


sub print_usage_statement
{
    printf STDERR "Syntax: escape_excel.pl [options] tab_delimited_input.txt [output.txt]\n";
    printf STDERR "   Options:\n";
    printf STDERR "      --no-dates   Do not escape text that looks like dates and/or times\n";
    printf STDERR "      --no-sci     Do not escape > #E (ex: 12E4) or >11 digit integer parts\n";
    printf STDERR "      --no-zeroes  Do not escape leading zeroes (ie. 012345)\n";
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
    printf STDERR "   Copy / Paste Values in Excel, after importing, to de-escape back into text.\n";
}


sub is_number
{
    # use what Perl thinks is a number first
    # this is purely for speed, since the more complicated REGEX below should
    #  correctly handle all numeric cases
    if (looks_like_number($_[0]))
    {
        # Perl treats infinities as numbers, Excel does not
        if ($_[0] =~ /^[+-]*inf/)
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
    return ($_[0] =~ /^([+-]?)([0-9]+(,[0-9]{3,})*\.?[0-9]*|\.[0-9]*)([Ee]([+-]?[0-9]+))?$/);
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

$escape_excel_paranoid_flag = 0;
$escape_sci_flag = 1;
$escape_zeroes_flag = 1;
$escape_dates_flag = 1;

# read in command line arguments
$num_files = 0;
for ($i = 0; $i < @ARGV; $i++)
{
    $field = $ARGV[$i];

    if ($field =~ /^--/)
    {
        if ($field eq '--paranoid')
        {
            if ($escape_excel_paranoid_flag == 0)
            {
                $escape_excel_paranoid_flag = 1;
            }
            else
            {
                $escape_excel_paranoid_flag = 0;
            }
        }
        elsif ($field eq '--no-sci')
        {
            $escape_sci_flag = 0;
        }
        elsif ($field eq '--no-zeroes')
        {
            $escape_zeroes_flag = 0;
        }
        elsif ($field eq '--no-dates')
        {
            $escape_dates_flag = 0;
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

    @array = split /\t/, $line;

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
        # continue stripping problematic stuff until all has been stripped
        do
        {
            $changed_flag = 0;
            
            # remove pre-existing escapes or start/end double quotes,
            # since either messes up ="" escapes
            while ($array[$i] =~ s/^\=*\"(.*?)\"$/$1/)
            {
                $changed_flag = 1;
            }
        
            # remove leading ", since they mess up Excel in general
            #
            # this must be done after "", but before leading/trailing spaces,
            # since removing leading/trailing spaces could result in more
            # full "" enclosures, which would then be messed up by removing
            # only the leading "
            #
            while ($array[$i] =~ s/^\"//)
            {
                $changed_flag = 1;
            }

            # remove leading spaces, since they won't protect long numbers,
            # and will cause various REGEX to fail
            if ($array[$i] =~ s/^\s+//)
            {
                $changed_flag = 1;
            }

            # remove trailing spaces, since they won't protect dates,
            # and will cause various REGEX to fail
            if ($array[$i] =~ s/\s+$//)
            {
                $changed_flag = 1;
            }
        } while ($changed_flag)
    }
    
    # escape fields
    for ($i = 0; $i < @array; $i++)
    {
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
          if ($escape_zeroes_flag && $array[$i] =~ /^([+-]?)0[0-9]/)
          {
              $array[$i] = sprintf "=\"%s\"", $array[$i];
          }

          # Escape scientific notation with >= 2 digits before the E,
          #  since they are likely accessions or plate/well identifiers.
          #
          # Also escape numbers with >11 digits before the decimal point.
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
          elsif ($escape_sci_flag)
          {
              # strip commas before counting digits
              $temp = $array[$i];
              $temp =~ s/\,//g;
              
              if ($temp =~ /^([1-9][0-9]{11,}|[0-9]{2,}[eE])/)
              {
                  $array[$i] = sprintf "=\"%s\"", $array[$i];
              }
          }
        }
        # escape all text if paranoid
        elsif ($escape_excel_paranoid_flag)
        {
          $array[$i] = sprintf "=\"%s\"", $array[$i];
        }
        # escape only text that might be corrupted
        else
        {
          # escape single quote at beginning of line
          if ($array[$i] =~ /^'/)
          {
              $array[$i] = sprintf "=\"%s\"", $array[$i];
          }

          # prevent conversion into formulas
          elsif ($array[$i] =~ /^\=/)
          {
              $array[$i] = sprintf "=\"%s\"", $array[$i];
          }
          # Excel is smart enough to treat all +/- as not an equation
          #  but, otherwise, it will convert anything starting with +/-
          #  into "#NAME?" as a failed invalid equation
          elsif ($array[$i] =~ /^[+-]/ && !($array[$i] =~ /^[+-]+$/))
          {
              $array[$i] = sprintf "=\"%s\"", $array[$i];
          }

          # check for time and/or date stuff
          elsif ($escape_dates_flag)
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
                          $array[$i] = sprintf "=\"%s\"", $array[$i];
                      }
                  }
              }
              # only time
              elsif ($time =~ /\S/)
              {
                  if ($array[$i] =~ /^\Q$time\E$/)
                  {
                      $array[$i] = sprintf "=\"%s\"", $array[$i];
                  }
              }
              # only date
              elsif ($date =~ /\S/)
              {
                  if ($array[$i] =~ /^\Q$date\E$/)
                  {
                      $array[$i] = sprintf "=\"%s\"", $array[$i];
                  }
              }
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
