use Spreadsheet::XLSX::Root;

#| A particular worksheet within an Excel workbook.
class Spreadsheet::XLSX::Worksheet {
    #| The root, used for resolutions at the document level.
    has Spreadsheet::XLSX::Root $.root is required;

    #| The name of the worksheet.
    has Str $.name is rw is required;
}
