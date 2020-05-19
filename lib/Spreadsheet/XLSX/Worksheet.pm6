use Spreadsheet::XLSX::Root;

#| A particular worksheet within an Excel workbook.
class Spreadsheet::XLSX::Worksheet {
    #| The root, used for resolutions at the document level.
    has Spreadsheet::XLSX::Root $.root is required;

    #| The name of the worksheet.
    has Str $.name is rw is required;

    #| If the sheet is loaded from an existing file, this is its path.
    #| We lazily load from this (meaning if we have a workbook with many
    #| sheets, we only load those we need to).
    has Str $!backing-path;

    submethod TWEAK(:$!backing-path --> Nil) {}
}
