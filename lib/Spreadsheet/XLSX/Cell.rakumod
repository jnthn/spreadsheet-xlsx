use LibXML::Document;
use LibXML::Element;
use Spreadsheet::XLSX::CellStyle;
use Spreadsheet::XLSX::Exceptions;
use Spreadsheet::XLSX::Root;
use Spreadsheet::XLSX::XMLHelpers;
use Spreadsheet::XLSX::Exceptions;
use Spreadsheet::XLSX::Styles;
use Spreadsheet::XLSX::Types;

enum Spreadsheet::XLSX::Cell::Types (:CTBoolean<b>, :CTDate<d>, :CTError<e>, :CTInlineStr<inlineStr>,
                                     :CTNumber<n>, :CTSharedStr<s>, :CTFormulaStr<str>);

enum Spreadsheet::XLSX::Cell::FormulaType (:CFTNormal<normal>, :CFTArray<array>,
                                           :CFTDataTable<dataTable>, :CFTShared<shared>);

# <v> element might have an attribute. We don't use it but need to preserve it.
my class Spreadsheet::XLSX::Cell::Value does XMLRepresentation {
    has Str $.value handles <Str gist Int Numeric Num Rat Rational Real>
                    is xml-text;
    method Bool { $!value eq 'true' | '1' }
}

my class Spreadsheet::XLSX::Cell::Formula does XMLRepresentation {
    has Str $.expression        is xml-text;
    has Spreadsheet::XLSX::Cell::FormulaType $.type  is xml-attr<t>;
    has Str $.ref               is xml-attr;
    has Str $.r1                is xml-attr;
    has Str $.r2                is xml-attr;
    has Bool $.aca              is xml-attr;
    has Bool $.dt2D             is xml-attr;
    has Bool $.dtr              is xml-attr;
    has Bool $.del1             is xml-attr;
    has Bool $.del2             is xml-attr;
    has Bool $.ca               is xml-attr;
    has Bool $.bx               is xml-attr;
    has UInt $.si               is xml-attr;
}

#| Commonalities of all kinds of cell.
role Spreadsheet::XLSX::Cell does XMLRepresentation {
    has Spreadsheet::XLSX::Root $.root;
    #| Cell styling information.
    has Spreadsheet::XLSX::CellStyle $!style;
    has UInt $!style-id                    is built is xml-attr<s>;
    has Bool $.show-phonetic                        is xml-attr<ph>;
    has Str $.reference                             is xml-attr<r>;
    has Spreadsheet::XLSX::Cell::Types $.type       is xml-attr<t>;
    has UInt $.value-metadata-idx                   is xml-attr<vm>;
    has Spreadsheet::XLSX::Cell::Value $!v is built is xml-elem<v>;
    has Spreadsheet::XLSX::Cell::Formula $.formula  is xml-elem<f>;
    has CT_Rst $.is                                 is xml-elem;

    #| Sync the value to XML.
    method sync-value-xml(LibXML::Document $document, LibXML::Element $col --> Nil) { ... }

    method style {
        $!style //= do {
            if $!root && $!style-id && !$!root.styles.cell-formats.EXISTS-POS($!style-id) {
                die X::Spreadsheet::XLSX::Format.new:
                    message => "Cell is referencing missing style index $!style-id"
            }
            Spreadsheet::XLSX::CellStyle.new(:$!root, :$!style-id);
        }
    }
}

#| A number cell.
class Spreadsheet::XLSX::Cell::Number does Spreadsheet::XLSX::Cell {
    #| The numeric value of the cell.
    has Real $.value;

    submethod TWEAK {
        without $!value {
            without $!v {
                die X::Spreadsheet::XLSX::Format.new:
                    message => 'Number cell node missing <v> value tag';
            }
            $!value = $!v.Real;
        }
    }

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
    has Str $.value;

    submethod TWEAK {
        without $!value {
            given $!type {
                when CTInlineStr {
                    with $!is {
                        # According to SpreadsheetML schema, <t> is not mandatory under <is>. Hence we can end up having
                        # undefined value.
                        $!value = $_ with .t;
                    }
                    else {
                        die X::Spreadsheet::XLSX::Format.new:
                            message => 'inlineStr cell missing <is> child';
                    }
                }
                when CTSharedStr {
                    with $!v {
                        $!value = $!root.shared-strings[+$_].t;
                    }
                    else {
                        die X::Spreadsheet::XLSX::Format.new:
                            message => "Missing <v> node for shared cell value";
                    }
                }
                default {
                    die X::Spreadsheet::XLSX::Format.new:
                        message => "Unsupported text node type '" ~ $_.value ~ "'"
                }
            }
        }
    }

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

class Spreadsheet::XLSX::Cell::Empty does Spreadsheet::XLSX::Cell {
    method value { Nil }

    multi method Str { '' }

    method sync-value-xml(LibXML::Document $, LibXML::Element $col --> Nil) {
        $col.removeChildNodes();
    }
}

#| Takes an XML node from shared strings and produces the correct kind of
#| Cell object from it.
sub cell-from-xml(LibXML::Element $element, Spreadsheet::XLSX::Root:D $root) is export {

    return Spreadsheet::XLSX::Cell::Empty.from-xml-element($element, :$root) unless $element.hasChildNodes;

    my LibXML::Attr $type-node = $element.getAttributeNode('t');
    given $type-node ?? $type-node.string-value !! '' {
        when '' {
            # Empty means number.
            Spreadsheet::XLSX::Cell::Number.from-xml-element($element, :profile{ :$root })
        }
        when 'inlineStr' | 's' {
            Spreadsheet::XLSX::Cell::Text.from-xml-element($element, :profile{ :$root });
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

