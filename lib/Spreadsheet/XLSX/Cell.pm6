use LibXML::Document;
use Spreadsheet::XLSX::Exceptions;

#| Commonalities of all kinds of cell.
role Spreadsheet::XLSX::Cell {
}

#| A number cell.
class Spreadsheet::XLSX::Cell::Number does Spreadsheet::XLSX::Cell {
    #| The numeric value of the cell.
    has Real $.value is required;

    #| Get a string representing the cell's value.
    multi method Str(::?CLASS:D:) {
        ~$!value
    }
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

#| Takes an XML node from shared strings and produces the correct kind of
#| Cell object from it.
sub cell-from-xml(LibXML::Element $element) is export {
    my LibXML::Attr $type-node = $element.getAttributeNode('t');
    given $type-node ?? $type-node.string-value !! '' {
        when '' {
            # Empty means number.
            my LibXML::Element $value-node = $element.first;
            unless $value-node.nodeName eq 'v' {
                die X::Spreadsheet::XLSX::Format.new:
                    message => 'Number cell node missing v value tag';
            }
            Spreadsheet::XLSX::Cell::Number.new(value => +$value-node.string-value)
        }
        default {
            die X::NYI.new(feature => "Excel cells of type '$_'");
        }
    }
}

#| Takes an XML node from shared strings and produces the correct kind of
#| Cell object from it.
sub shared-cell-from-xml(LibXML::Element $element) is export {
    given $element.nodeName {
        when 't' {
            Spreadsheet::XLSX::Cell::Text.new(value => $element.string-value)
        }
        default {
            die X::NYI.new(feature => "Excel shared cells of type '$_'");
        }
    }
}

