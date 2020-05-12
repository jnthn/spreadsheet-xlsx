use Spreadsheet::XLSX::Workbook;

#| A particular worksheet within an Excel spreadsheet.
class Spreadsheet::XLSX::Worksheet {
    #| The workbook that this worksheet belongs to.
    has Spreadsheet::XLSX::Workbook $.workbook;

    #| The name of the worksheet.
    has Str $.name is rw;
}
