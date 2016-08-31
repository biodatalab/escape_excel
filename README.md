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

Input file must be tab-delimited.

Fields will be stripped of existing ="" escapes, enclosing "", leading ", and leading/trailing spaces, as they may all cause problems.

Defaults to escaping most Excel mis-imported fields.

Escapes a few extra date-like formats that Excel does not consider dates.

Please send unhandled mis-imported field examples (other than gene symbols with 1-digit scientific notation, such as 2e4) to [Eric.Welsh@moffitt.org](mailto:Eric.Welsh@moffitt.org).

Copy / Paste Values in Excel, after importing, to de-escape back into text.

Use the provided text file (_test_excel_import.txt_) to test the script.

Also available on a Galaxy test server at [http://apostl.moffitt.org](http://apostl.moffitt.org).
