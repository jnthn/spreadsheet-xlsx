use LibXML::Document;
use Spreadsheet::XLSX::CellStyle;
use Spreadsheet::XLSX::Exceptions;

#| Commonalities of all kinds of cell.
role Spreadsheet::XLSX::Cell {
    #| Cell styling information.
    has Spreadsheet::XLSX::CellStyle $.style .= new;

    #| Sync the value to XML.
    method sync-value-xml(LibXML::Document $document, LibXML::Element $col --> Nil) { ... }
}

#| A number cell.
class Spreadsheet::XLSX::Cell::Number does Spreadsheet::XLSX::Cell {
    #| The numeric value of the cell.
    has Real $.value is required;

    #| Get a string representing the cell's value.
    multi method Str(::?CLASS:D:) {
        ~$!value
    }

    #| Sync the value to XML.
    method sync-value-xml(LibXML::Document $document, LibXML::Element $col --> Nil) {
        # A number node has no type.
        $col.removeAttributeNode($_) with $col.getAttributeNode('t');

        # Number goes within a v element.
        $col.removeChildNodes();
        my LibXML::Element $v = $document.createElement('v');
        $v.nodeValue = ~$!value;
        $col.add($v);
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

    #| Sync the value to XML.
    method sync-value-xml(LibXML::Document $document, LibXML::Element $col --> Nil) {
        # Make sure the type is inlineStr.
        with $col.getAttributeNode('t') {
            .setValue('inlineStr')
        }
        else {
            $col.add($document.createAttribute('t', 'inlineStr'));
        }

        # Set the content (<is><t>content</t></is>).
        $col.removeChildNodes();
        my LibXML::Element $is = $document.createElement('is');
        my LibXML::Element $t = $document.createElement('t');
        $t.appendText($!value);
        $is.add($t);
        $col.add($is);
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
        when 'inlineStr' {
            my LibXML::Element $is-node = $element.first;
            unless $is-node.nodeName eq 'is' {
                die X::Spreadsheet::XLSX::Format.new:
                        message => 'inlineStr cell missing is tag';
            }
            my LibXML::Element $content-node = $is-node.first;
            unless $content-node.nodeName eq 't' {
                die X::NYI.new(feature => "Node of type $content-node.nodeName() in inline string");
            }
            Spreadsheet::XLSX::Cell::Text.new(value => $content-node.string-value);
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

