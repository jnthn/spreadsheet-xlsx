# Spreadsheet::XLSX

A Raku module for working with Excel spreadsheets (XLSX format). Provides a
mutable document object model, which can be created from scratch or parsed
from an existing file. This can then be written out to an XLSX file also.

This module is currently in development, and supports the subset of XLSX
format features that were immediately needed for the use-case it was built
for.

## Synopsis

```raku
# Create a new spreadsheet.
my $workbook = Spreadsheet::XLSX.new;

# Or read a workbook from an existing file.
my $parsed-workbook = Spreadsheet::XLSX.load('accounts.xlsx'); 
```

