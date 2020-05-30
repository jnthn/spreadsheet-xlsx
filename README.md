# Spreadsheet::XLSX

A Raku module for working with Excel spreadsheets (XLSX format), both
reading existing files, creating new files, or modifying existing files
and saving the changes. Of note, it:

* Knows how to lazily load sheet content, so if you don't look at a sheet
  then time won't be spent deserializing it (down to a cell level, even)
* In the modification scenario, tries to leave as much intact as it can,
  meaning that it's possible to poke data into a sheet more complex than
  could be produced by the module from scratch
* Only depends on the Raku LibXML and Libarchive modules (and their
  respective native dependencies)

This module is currently in development, and supports the subset of XLSX
format features that were immediately needed for the use-case it was built
for. That isn't so much, for now, but it will handle the most common needs:

* Enumerating worksheets
* Reading text and numbers from cells on a worksheet
* Creating new workbooks with worksheets with text and number cells
* Setting basic styles and number formats on cells in newly created
  worksheets
* Reading a workbook, making modifications, and saving it again
* Reading and writing column properties (such as column width)

## Synopsis

### Reading existing workbooks

```
# Read a workbook from an existing file (can pass IO::Path or a
# Blob in the case it was uploaded).
my $workbook = Spreadsheet::XLSX.load('accounts.xlsx');

# Get worksheets.
say "Workbook has {$workbook.worksheets.elems} sheets";

# Get the name of a worksheet.
say $workbook.worksheets».name;

# Get cell values (indexing is zero-based, done as a multi-dimensional array
# indexing operation [row ; column].
my $cells = $workbook.worksheets[0].cells;
say .value with $cells[0;0];      # A1
say .value with $cells[0;1];      # B1
say .value with $cells[1;0];      # A2
say .value with $cells[1;1];      # B2
```

### Creating new workbooks

```raku
# Create a new workbook and add some worksheets to it.
my $workbook = Spreadsheet::XLSX.new;
my $new-sheet-a = $workbook.create-worksheet('Ingredients');
my $sheet-b = $workbook.create-worksheet('Matching Drinks');

# Put some data into a worksheet and style it. This is how the model
# actually works (useful if you want to add styles later)...
$new-sheet-a.cells[0;0] = Spreadsheet::XLSX::Cell::Text.new(value => 'Ingredient');
$new-sheet-a.cells[0;0].style.bold = True;
$new-sheet-a.cells[0;1] = Spreadsheet::XLSX::Cell::Text.new(value => 'Quantity');
$new-sheet-a.cells[0;1].style.bold = True;
$new-sheet-a.cells[1;0] = Spreadsheet::XLSX::Cell::Text.new(value => 'Eggs');
$new-sheet-a.cells[1;1] = Spreadsheet::XLSX::Cell::Number.new(value => 6);
$new-sheet-a.cells[1;1].style.number-format = '#,###';

# However, there is a convenience form too.
$new-sheet-a.set(0, 0, 'Ingredient', :bold);
$new-sheet-a.set(0, 1, 'Quantity', :bold);
$new-sheet-a.set(1, 0, 'Eggs');
$new-sheet-a.set(1, 1, 6, :number-format('#,###'));

# Save it to a file (string or IO::Path name).
$workbook.save("foo.xlsx");

# Or get it as a blob, e.g. for a HTTP response.
my $blob = $workbook.to-blob();
```

## Credits

Thanks goes to [Agrammon](https://agrammon.ch/) for making the development of
this module possible. If you need further development on the module and are
willing to fund it (or other Raku ecosystem work), you can get in contact with
[Edument](https://www.edument.se/en) or [Oetiker+Partner](https://www.oetiker.ch/en/).
