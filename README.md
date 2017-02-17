##Escape Excel: a tool for preventing gene symbol and accession conversion errors

###Syntax: escape_excel.pl [options] tabDelimitedInput.txt [output.txt]
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

###Galaxy Server
Escape Excel is available for testing on a Galaxy server (web-based interface; no command line needed) at [http://apostl.moffitt.org](http://apostl.moffitt.org). Use the menu on the left to navigate to the tool interface.

###Galaxy Tool Shed
Escape Excel is [available to install](https://toolshed.g2.bx.psu.edu/view/pstew/escape_excel/482c23a5abfe) from the Galaxy Tool Shed (for Galaxy administrators).

###Preprint
Escape Excel is [now available](http://biorxiv.org/content/early/2017/01/27/103820) as a preprint on biorxiv.
