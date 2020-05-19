use LibXML::Document;

#| Commonalities of all kinds of cell.
role Spreadsheet::XLSX::Cell {
}

#| A simple text cell.
class Spreadsheet::XLSX::Cell::Text does Spreadsheet::XLSX::Cell {
    #| The textual value of the cell.
    has Str $.value is required;

    #| Get a string representing the cell's value.
    multi method Str(::?CLASS:D:) {
        $!value
    }
}

#| Takes an XML node and produces the correct kind of element from it.
sub cell-from-xml(LibXML::Element $element) is export {
    given $element.nodeName {
        when 't' {
            Spreadsheet::XLSX::Cell::Text.new(value => $element.string-value)
        }
        default {
            die X::NYI.new(feature => "Excel cells of type '$_'");
        }
    }
}

