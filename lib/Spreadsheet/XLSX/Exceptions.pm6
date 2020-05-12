#| Marker role for all kinds of Spreadsheet::XLSX exceptions.
role X::Spreadsheet::XLSX is Exception {}

#| A problem with parsing data out of XLSX file.
class X::Spreadsheet::XLSX::Format does X::Spreadsheet::XLSX {
}
