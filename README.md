# sol-ipg

This file and all others are my own work and are offered with no warranty on there function or safety if used - by any one. If this breaks something, you get to keep both parts. I know it's not a pretty script, YOU know it's not pretty but it works (for me anyway) so feel free to pretty things up if you want.

The purpose of this script is to convert a Solarwinds IPAM All Subnets report exported to an Excel file to create an xml file that can be imported into the Solalrwinds NTA module with each IP Subnet listed as a range to allow for more granular Netflow identities in Solarwinds graphs e.g. Top 10 IP Ranges could be Data_Centre Sub1, Guest_Wireless etc. This all makes looking at the dat a bit more nuanced.

In order to make this work you need a few extra modules via pip:

  1)Prerequisites - Solarwinds NTA and IPAM modules.
  2)IPy to do easy manipulation of dotted decimal IPv4 address manipulation.
  3)openpyxl to read xlsx files. 

Export the IP Groups file from the NTA module Admin page and you will get an xml file called ipgroups.xml - save it in case things go wrong and you want to restore the original.
N.B. for some reason this is exported as UTF-16 so some editors may struggle to read it - i don't know why it is UTF-16 it just is. When we create the new xml file we will be using UTF-8 and this has zero dertimental effect on import but make life a lot easier.

Run the built in report called 'IPAM - All subnets' and export it to excel. Then, delete the top rows of the file so that the data starts in row 1 - no headers or anything - and re-save the file.
Check the script either is present in the same dir as the excel file (with read/write permissions) or that you have the full path to both the excel file and the xml file location in the relevant script variables.
