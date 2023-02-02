use LibXML::Document;
use Spreadsheet::XLSX::Cell;
use Spreadsheet::XLSX::Root;

#| Since spreadsheets often contain a lot of repetition, a shared strings
#| table is used to extract the commonality.
class Spreadsheet::XLSX::SharedStrings does Positional {
    #| The root, used for resolutions at the document level.
    has Spreadsheet::XLSX::Root $.root is required;

    #| The backing document from the shared strings table, if any.
    has LibXML::Element $!backing is built;
    has Str @!string-cache is built;

    #| Create a string table from an XML document.
    method from-xml(Str $xml, Spreadsheet::XLSX::Root :$root! --> Spreadsheet::XLSX::SharedStrings) {
        my LibXML::Document $doc .= parse(:string($xml));
        my LibXML::Element $sst = $doc.documentElement();
        if $sst.nodeName ne 'sst' {
            die X::Spreadsheet::XLSX::Format.new: message =>
                    'Shared strings file did not start with tag sst';
        }
        my Str @string-cache = $sst.childNodes.map: -> $si {
            if $si.nodeName ne 'si' {
                die X::Spreadsheet::XLSX::Format.new: message =>
                        'Shared strings entry was not an si node';
            }
            my $element = $si.first;
            given $element.nodeName {
                when 't' {
                    $element.string-value
                }
                when 'r' {
                    # TODO: Deep RichText handling.
                    my $text;
                    for $element.childNodes -> $elem {
                        if $elem.nodeName eq 't' {
                            $text ~= $elem.string-value;
                        }
                    }
                    $text
                }
                default {
                    die X::NYI.new(feature => "Excel shared cells of type '$_'");
                }
            }
        }
        self.new(:$root, :backing($sst), :@string-cache)
    }

    #| Create a new, empty, string table.
    method empty(Spreadsheet::XLSX::Root :$root!--> Spreadsheet::XLSX::SharedStrings) {
        self.new(:$root)
    }

    #| Get the shared string entry at the given position.
    method AT-POS(Int $idx) {
        with $!backing {
            with @!string-cache[$idx] -> $val {
                return $val;
            }
        }
        fail "Could not resolve shared string entry $idx";
    }
}
