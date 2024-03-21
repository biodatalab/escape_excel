# Escape Excel, a tool for preventing gene symbol and accession conversion errors

## Escape Excel Perl script on the Command Line

### Quick Start

<pre>
Usage: escape_excel.pl [options] tab_delimited_input.txt [output.txt]
   Options:
      --all-commas  Escape *ALL* numeric-looking fields with commas in them
      --csv         input CSV file instead of tab-delimited, still outputs tsv
      --escape-dq   "smart" escaping of " to better preserve them (default)
      --no-commas   Do not escape ,#### and ####,###
      --no-dates    Do not escape text that looks like dates and/or times
      --no-dq       Disable "smart" handling of double quotes
      --no-sci      Do not escape >= ##E (ex: 12E4) or >11 digit integer parts
      --no-zeroes   Do not escape leading zeroes (ie. 012345)
      --unstrip     restore auto-stripped field when not escaped
      --paranoid    Escape *ALL* non-numeric text (overrides --no-dates)
                    WARNING -- Excel can take a LONG time to import
                    text files where most fields are escaped.
                    Copy / Paste Values can become near unusuable....

   Reads input from STDIN if input file name is - or no file name is given.
   Input file must be tab-delimited.
   Fields will be stripped of existing ="" escapes, enclosing "", leading ",
    and leading/trailing spaces, as they may all cause problems.

   Defaults to escaping most Excel mis-imported fields.
   Escapes a few extra date-like formats that Excel does not consider dates.
   Please send unhandled mis-imported field examples (other than gene symbols
    with 1-digit scientific notation, such as 2e4) to Eric.Welsh@moffitt.org.

   Does not currently support or detect line wraps.  To best preserve line
   wraps, use --no-dq and enable --unstrip to better preserve spacing.

   Copy / Paste Values in Excel, after importing, to de-escape back into text.
</pre>

Use the provided text file (`test_excel_import.txt`) to test the script.



### Tutorial

To run Escape Excel on the command line, you first need to see if Perl is
installed on your operating system.  MacOS and Linux operating systems should
have Perl installed by default, while Windows users will need to download and
install Perl.  To see if Perl is installed on your system, open a terminal
(or "Command Prompt" in Windows) and type `perl -v`.  If the terminal returns
a Perl version, then Perl is already installed on your system. If you see an
error that the command can't be found, then you likely need to install Perl.
To install Perl, visit the [Perl download site]
(https://www.perl.org/get.html), select your appropriate operating system,
and follow the instructions.

Once Perl is installed, open a terminal.  Make sure the Escape Excel file
(escape\_excel.pl) and the tab delimited data you wish to escape are in the
same folder. Change directory to the location of these files using the `cd`
command.  For example, `cd /home/pstew/mydata/`.  If you are using Mac or
Linux, then you may first need to make the escape\_excel.pl file executable.
To do this, type the command `chmod u+x escape_excel.pl`.  This will allow you
to type `escape_excel.pl` without having to precede it every time with the
`perl` command.

The syntax for using Escape Excel can be found above in the quick start guide.
The usage statement will be displayed with the --help option, or any other
invalid option.



## Escape Excel Windows .exe file
This is no longer supported, updated, or otherwise maintained.

You can, however, generate your own .exe file from escape\_excel.pl using
PAR Packager (`pp`): https://metacpan.org/pod/pp



## Escape Excel MacOS App
This is no longer supported, updated, or otherwise maintained.

## Escape Excel Web Server
This is no longer supported, updated, or otherwise maintained.

## Escape Excel on Galaxy
This is no longer supported, updated, or otherwise maintained.

## Excel Plugin for Escape Excel
This is no longer supported, updated, or otherwise maintained.
