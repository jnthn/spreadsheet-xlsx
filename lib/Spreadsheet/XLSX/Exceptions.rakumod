#| Marker role for all kinds of Spreadsheet::XLSX exceptions.
role X::Spreadsheet::XLSX is Exception {}

#| A problem with parsing data out of XLSX file.
class X::Spreadsheet::XLSX::Format does X::Spreadsheet::XLSX {
    has Str $.message is required;
}

#| A problem with a missing relationship.
class X::Spreadsheet::XLSX::NoSuchRelationship does X::Spreadsheet::XLSX {
    has Str $.id is required;
    method message() {
        "No such relationship '$!id'"
    }
}

#| A non-standard property name
class X::Spreadsheet::XLSX::NoSuchProperty does X::Spreadsheet::XLSX {
    has Str $.property is required;
    method message() {
        "No such property '$!property'"
    }
}
