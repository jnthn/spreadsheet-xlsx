use LibXML::Document;
use Spreadsheet::XLSX::CellStyle;
use Spreadsheet::XLSX::Exceptions;

#| Commonalities of all kinds of cell.
role Spreadsheet::XLSX::Cell {
    #| Cell styling information.
    has Spreadsheet::XLSX::CellStyle $.style .= new;

    #| Cell formula.
    has Str $.formula;

    #| Row.
    has Int $.row is required;

    #| Column.
    has Int $.column is required;

    #| Sync the value to XML.
    method sync-value-xml(LibXML::Document $document, LibXML::Element $col --> Nil) { ... }

    #| Sync the formula to XML.
    #| Expects the target node to not contain an "f" element.
    method maybe-sync-formula-xml(LibXML::Document $document, LibXML::Element $col --> Nil) {
        with $!formula {
            my LibXML::Element $f = $document.createElement('f');
            $f.appendText($!formula);
            $col.add($f);
        }
    }
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
        self.maybe-sync-formula-xml($document, $col);
    }
}

#| An inline text cell.
class Spreadsheet::XLSX::Cell::Text does Spreadsheet::XLSX::Cell {
    #| The textual value of the cell.
    has Str $.value is required;

    #| Get a string representing the cell's value.
    multi method Str(::?CLASS:D:) {
        $!value
    }

    #| Sync the value to XML.
    method sync-value-xml(LibXML::Document $document, LibXML::Element $col --> Nil) {
        if $!formula {
            # Make sure the type is str.
            with $col.getAttributeNode('t') {
                .setValue('str')
            }
            else {
                $col.add($document.createAttribute('t', 'str'));
            }

            # Set the content (<v>cached value</v>).
            $col.removeChildNodes();
            my LibXML::Element $v = $document.createElement('v');
            $v.nodeValue = $!value;
            $col.add($v);
            self.maybe-sync-formula-xml($document, $col);
        }
        else {
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
}

#| A boolean cell.
class Spreadsheet::XLSX::Cell::Bool does Spreadsheet::XLSX::Cell {
    #| The boolean value
    has Bool $.value;

    #| Get a string representing the cell's value.
    multi method Str(::?CLASS:D:) {
        $!value ?? "TRUE" !! "FALSE"
    }

    #| Sync the formula to XML.
    method sync-value-xml(LibXML::Document $document, LibXML::Element $col --> Nil) {
        $col.removeChildNodes();
        my LibXML::Element $v = $document.createElement('v');
        $v.nodeValue = $!value ?? 1 !! 0;
        $col.add($v);
        self.maybe-sync-formula-xml($document, $col);
    }
}

enum Spreadsheet::XLSX::Cell::ErrorVal is export (
    'NULL' => 1,
    'DIV0',
    'VALUE',
    'REF',
    'NAME',
    'NUM',
    'NA',
    'GETTING_DATA',
);

#| An error cell.
class Spreadsheet::XLSX::Cell::Error does Spreadsheet::XLSX::Cell {
    #| The boolean value
    has Spreadsheet::XLSX::Cell::ErrorVal $.value;

    #| Get a string representing the cell's value.
    multi method Str(::?CLASS:D:) {
        given $!value {
            when NULL { "#NULL!" }
            when DIV0 { "#DIV/0!" }
            when VALUE { "#VALUE!" }
            when REF { "#REF!" }
            when NAME { "#NAME?" }
            when NUM { "#NUM!" }
            when NA { "#N/A" }
            when GETTING_DATA { "#GETTING_DATA" }
        }
    }

    #| Sync the formula to XML.
    method sync-value-xml(LibXML::Document $document, LibXML::Element $col --> Nil) {
        $col.removeChildNodes();
        my LibXML::Element $v = $document.createElement('v');
        $v.nodeValue = ~$!value;
        $col.add($v);
        self.maybe-sync-formula-xml($document, $col);
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
sub cell-from-xml(LibXML::Element $element, $shared-strings) is export {
    return Spreadsheet::XLSX::Cell::Empty.new unless $element.hasChildNodes;

    my LibXML::Attr $type-node = $element.getAttributeNode('t');
    my LibXML::Attr $address-attr = $element.getAttributeNode('r');
    unless $address-attr {
        die X::Spreadsheet::XLSX::Format.new:
                message => "Missing r cell attribute";
    }
    my ($row, $column) = parse-addr $address-attr.string-value;
    my $f-node = $element.childNodes.first: *.name eq 'f';
    my Str $formula = $_.string-value with $f-node;
    my $cell = do given $type-node ?? $type-node.string-value !! '' {
        when '' {
            # Empty means number.
            my LibXML::Element $v-node = $element.childNodes.first: *.name eq 'v';
            without $v-node {
                die X::Spreadsheet::XLSX::Format.new:
                    message => 'Number cell node missing v value tag';
            }
            Spreadsheet::XLSX::Cell::Number.new(value => +$v-node.string-value,
                                                :$row, :$column, :$formula);
        }
        when 's' {
            my LibXML::Element $shared-index-holder = $element.first;
            unless $shared-index-holder.nodeName eq 'v' {
                die X::Spreadsheet::XLSX::Format.new:
                        message => "Missing v node for shared cell value";
            }
            Spreadsheet::XLSX::Cell::Text.new(value => $shared-strings[$shared-index-holder.string-value.Int],
                                                :$row, :$column, :$formula);
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
            Spreadsheet::XLSX::Cell::Text.new(value => $content-node.string-value,
                                                :$row, :$column, :$formula);
        }
        when 'str' {
            my $v-node = $element.childNodes.first: *.name eq 'v';
            without $v-node {
                die X::Spreadsheet::XLSX::Format.new:
                        message => 'str cell missing v tag';
            }
            Spreadsheet::XLSX::Cell::Text.new(value => $v-node.string-value,
                                                :$row, :$column, :$formula);
        }
        when 'b' {
            my $v-node = $element.childNodes.first: *.name eq 'v';
            without $v-node {
                die X::Spreadsheet::XLSX::Format.new:
                        message => 'b cell missing v tag';
            }
            Spreadsheet::XLSX::Cell::Bool.new(value => +$v-node.string-value == 0 ?? False !! True,
                                                :$row, :$column, :$formula);
        }
        when 'e' {
            my $v-node = $element.childNodes.first: *.name eq 'v';
            without $v-node {
                die X::Spreadsheet::XLSX::Format.new:
                        message => 'b cell missing v tag';
            }
            my $err = do given $v-node.string-value {
                when "#NULL!"        { NULL }
                when "#DIV/0!"       { DIV0 }
                when "#VALUE!"       { VALUE }
                when "#REF!"         { REF }
                when "#NAME?"        { NAME }
                when "#NUM!"         { NUM }
                when "#N/A"          { NA }
                when "#GETTING_DATA" { GETTING_DATA }
                default {
                    die X::NYI.new(feature => "Excel cells of error type '$_'");
                }
            }
            Spreadsheet::XLSX::Cell::Error.new(value => $err,
                                                :$row, :$column, :$formula);
        }
        default {
            die X::NYI.new(feature => "Excel cells of type '$_'");
        }
    }
}

sub parse-addr(Str $addr) {
    my $pos = 0;
    my @parts = $addr.comb;
    my $val = 0;
    loop {
        my $v = ord(@parts[$pos]) - ord('A') + 1;
        last if $v <= 0;
        $val = $val * 26 + $v;
        $pos++;
    }
    $val--;
    ($addr.substr($pos).Int - 1, $val);
}

