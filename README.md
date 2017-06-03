# Escape Excel: a tool for preventing gene symbol and accession conversion errors

## Escape Excel on the Command Line

### Quick Start

Syntax: `escape_excel.pl [options] tab_delimited_input.txt [output_file_name.txt]`

Options:

      --no-dates   Do not escape text that looks like dates
      --no-sci     Do not escape > #E (ex: 12E4) or >11 digit integer parts
      --no-zeroes  Do not escape leading zeroes (ie. 012345)
      --paranoid   Escape *ALL* non-numeric text
                   WARNING -- Excel can take a LONG time to import
                   text files where most fields are escaped.
                   Copy / Paste Values can become near unusuable....

Input file **must** be tab-delimited.

Fields will be stripped of existing ="" escapes, enclosing "", leading ", and leading/trailing spaces, as they may all cause problems.

Defaults to escaping most Excel mis-imported fields.

Escapes a few extra date-like formats that Excel does not consider dates.

Please file an issue for unhandled fields (other than gene symbols with 1-digit scientific notation, such as 2e4) or email the corresponding author at [Eric.Welsh@moffitt.org](mailto:Eric.Welsh@moffitt.org) directly.

Copy / Paste Values in Excel, after importing, to de-escape back into text.

Use the provided text file (_test_excel_import.txt_) to test the script.

### Tutorial
To run Escape Excel on the command line, you first need to see if Perl is installed on your operating system. MacOS and Linux operating systems should have Perl installed by default while Windows users will need to download and install Perl. To see if Perl is installed on your system, open a terminal (or "Command Prompt" in Windows) and type `perl -v`. If the terminal returns a Perl version, then Perl is already installed on your system. If you see an error that the command can't be found, then you likely need to install Perl. To install Perl, visit the [Perl download site](https://www.perl.org/get.html), select your appropriate operating system, and follow the instructions.

Once Perl is installed, open a terminal. Make sure the Escape Excel file (escape\_excel.pl) and the tab delimited data you wish to escape are in the same folder. Change directory to the location of these files using the `cd` command. For example, `cd /home/pstew/mydata/`. If you are using Mac or Linux, then you may first need to make escape\_excel.pl file executable. To do this, type the command `chmod u+x escape_excel.pl`. This will allow you to type `escape_excel.pl` without having to precede it every time with the `perl` command.

The syntax for using Escape Excel can be found above in the quick start guide, namely you type `escape_excel.pl [options] tab_delimited_input.txt [output_file_name.txt]` with appropriate options and the name of your output file name inserted. Default usage without specifying any options should be fine for most applications. Please note that in our hands, the `--paranoid` flag should be used cautiously as Excel has a hard time importing data where all fields have been escaped. Once run, the output file name you provided will be created and populated with escaped data that can now be loaded safely into Excel. If you wish to sort this escaped data or any sort of text manipulation, then it is a good idea to select all data, copy, and paste the data back on top of itself. This replaces the escaped values that contain escape characters (e.g. `"=value"` with the actual value (e.g. `value`).

## Escape Excel on Galaxy

### Galaxy Interface

The Moffitt Cancer Center generously provides a Galaxy server to run Escape Excel through a web based interface at [http://apostl.moffitt.org](http://apostl.moffitt.org). This provides a means to use Escape Excel through a point-and-click web-based interface without needing to know how to use the command line. Use the tools menu on the left to navigate to the tool interface, select your file, choose appropriate options, and click "Execute".

### Galaxy Tool Shed

Escape Excel is [available to install](https://toolshed.g2.bx.psu.edu/view/pstew/escape_excel/482c23a5abfe) from the Galaxy Tool Shed. This option is for Galaxy administrators and is not meant for end users. For end users wishing to escape their data, please see the **Command Line - Quick Start** and **Command Line - Tutorial** sections above.

## Press
Escape Excel is [now available](http://biorxiv.org/content/early/2017/01/27/103820) as a preprint on biorxiv.

Escape Excel was featured in a [blog article](http://blogs.nature.com/naturejobs/2017/02/27/escape-gene-name-mangling-with-escape-excel/) at Nature.
