# Escape Excel: a tool for preventing gene symbol and accession conversion errors

## Escape Excel Plugin
Download and run setupEscapeExcel.exe on windows to install the plugin. You can download the setup program under the releases tab.

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

Use the provided text file (`test_excel_import.txt`) to test the script.

### Tutorial
To run Escape Excel on the command line, you first need to see if Perl is installed on your operating system. MacOS and Linux operating systems should have Perl installed by default while Windows users will need to download and install Perl. To see if Perl is installed on your system, open a terminal (or "Command Prompt" in Windows) and type `perl -v`. If the terminal returns a Perl version, then Perl is already installed on your system. If you see an error that the command can't be found, then you likely need to install Perl. To install Perl, visit the [Perl download site](https://www.perl.org/get.html), select your appropriate operating system, and follow the instructions.

Once Perl is installed, open a terminal. Make sure the Escape Excel file (escape\_excel.pl) and the tab delimited data you wish to escape are in the same folder. Change directory to the location of these files using the `cd` command. For example, `cd /home/pstew/mydata/`. If you are using Mac or Linux, then you may first need to make escape\_excel.pl file executable. To do this, type the command `chmod u+x escape_excel.pl`. This will allow you to type `escape_excel.pl` without having to precede it every time with the `perl` command.

The syntax for using Escape Excel can be found above in the quick start guide, namely you type `escape_excel.pl [options] tab_delimited_input.txt [output_file_name.txt]` with appropriate options and the name of your output file name inserted. Default usage without specifying any options should be fine for most applications. Please note that in our hands, the `--paranoid` flag should be used cautiously as Excel has a hard time importing data where all fields have been escaped. Once run, the output file name you provided will be created and populated with escaped data that can now be loaded safely into Excel. If you wish to sort this escaped data or any sort of text manipulation, then it is a good idea to select all data, copy, and paste the data back on top of itself. This replaces the escaped values that contain escape characters (e.g. `"=value"` with the actual value (e.g. `value`).

## Escape Excel on Galaxy

### Galaxy Interface

Galaxy is a web-based platform that provides a point-and-click interface to bioinformatics tools. Galaxy allows for open and reproducible science by allowing files, workflows, and user histories to be easily shared. Escape Excel is provided within a Galaxy interface so that, after installation by the server adminstrator, no installation or technical expertise is required by the user to run the software.

### Installing Escape Excel in Galaxy

Escape Excel can be installed into a Galaxy instance by an administrator from the [Galaxy Tool Shed](https://toolshed.g2.bx.psu.edu/view/pstew/escape_excel/482c23a5abfe). Please note that this is meant for Galaxy administrators and not end users. For end users wishing to escape their data, please see the **Galaxy User Tutorial** below.

### Galaxy User Tutorial
 
To use the Escape Excel Galaxy interface, open a web browser and navigate to the user's Galaxy web server (for example, [http://apostl.moffitt.org](http://apostl.moffitt.org)). The user must first upload a delimited text file to be escaped by clicking “Upload Data” and then “Upload File”. A file upload window will appear, and a delimited text file can then be specified from a user’s hard drive by clicking “Choose local file”. Alternatively, data can be provided as a uniform resource locator (URL; e.g. a web link) by clicking the “Paste/Fetch data” button and pasting in the appropriate link.
 
Clicking “Start” will begin the upload process. Once the file upload process is complete, the file should appear in the user “History” on the right hand side of the screen. Click “Close” to close the file upload window. Find “Escape Excel” under the “Tools” menu on the left side of the screen and click it. Click “Escape Excel” again in the expanded menu to open the tool interface. The first dropdown in the tool menu allows the user to specify the uploaded delimited text file. Default behavior (all options set to “No”) should be sufficient for most uses.
 
Once done selecting the appropriate file and selecting options, click “Execute”. The processed data will appear in the user “History” and it will turn green once the file is processed. Click on the name of the processed data (“Escape Excel on data #”) in the “History” to expand a menu. Find the disk icon (moving the mouse over the icon reveals the text “Download”) and left click to download the escaped file. This file can now be loaded into Excel without auto-formatting.
 
If you have trouble getting Escape Excel to successfully run, make sure you are provided a tab-delimited text file. Escape Excel only works on tab-delimited text, and will not work on an Excel format file such as .csv, .xls, or .xlsx. In some instances, Galaxy may have trouble automatically detecting the type of file provided for upload. Try manually setting the file type to “txt” in the file upload window to correct for this.


## Escape Excel MacOS App

This zip file contains a Mac application allowing for easy use of Escape Excel on OS X. 

### Instructions for use
- Unzip the provided zip file using the Archive Utility
- Start escaping files using one of two methods:
	- Drag and drop files to be escaped directly onto application icon. 
	- Open the Escape Excel application to reveal a droplet for dropping files.
- After dropping a file, newly escaped files will be created in the same directory.


## Press
Escape Excel is [now available](http://biorxiv.org/content/early/2017/01/27/103820) as a preprint on biorxiv.

Escape Excel was featured in a [blog article](http://blogs.nature.com/naturejobs/2017/02/27/escape-gene-name-mangling-with-escape-excel/) at Nature.
