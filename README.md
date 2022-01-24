# SpreadCheck
A program which can be used to verify data and produce a report within Excel documents. Writen in C# utilising the Excel Interop layer

This program is able to validate spreadsheet data for a set of conditions, and produce a report of failed checks.

# Possible Checks

- If a cell is empty
- If a cell contains spaces
- If a cell contrains Non-Alphanumeric characters
- If a cell contains numbers
- If a cell contains letters
- Check Date & Time (incomplete)
- If a cell begins with a certain character string (exact match only customisation planned)
- If a cell ends with a certain character string (exact match only customisation planned)
- Check length
- Check for values in a list (exact match only customisation planned)
- If more than (incomplete)
- If less than (incomplete)

# Program Features

* Customisable error report strings
* Auto-detection of amount of columns / rows within spreadsheet
* Highlight Cell backcolor on cells which fail verification
* Trim data within Cells
* Reverse data within Cells
* Save / Load of programmed verification checks. 
* Fully configurable checks for each column of data within the spreadsheet
* Multiple Sheet support

# Future Updates

- [ ] Attempting to auto-correct failed data
- [ ] Drag / Drop & command line support
- [ ] Inclusion of Pivot-Table within report.
- [ ] Saving of report to various formats, HTML, XML, TXT, CSV ect. 
