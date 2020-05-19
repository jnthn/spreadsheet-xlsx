use LibXML::Document;
use Spreadsheet::XLSX::Cell;
use Spreadsheet::XLSX::Root;

#| Since spreadsheets often contain a lot of repetition, a shared strings
#| table is used to extract the commonality.
class Spreadsheet::XLSX::SharedStrings does Positional {
    #| The root, used for resolutions at the document level.
    has Spreadsheet::XLSX::Root $.root is required;

    #| The backing document from the shared strings table, if any.
    has LibXML::Element $!backing;

    submethod TWEAK(:$!backing --> Nil) {}

    #| Create a string table from an XML document.
    method from-xml(Str $xml, Spreadsheet::XLSX::Root :$root! --> Spreadsheet::XLSX::SharedStrings) {
        my LibXML::Document $doc .= parse(:string($xml));
        my LibXML::Element $sst = $doc.documentElement();
        if $sst.nodeName ne 'sst' {
            die X::Spreadsheet::XLSX::Format.new: message =>
                    'Shared strings file did not start with tag sst';
        }
        self.new(:$root, :backing($sst))
    }

    #| Create a new, empty, string table.
    method empty(Spreadsheet::XLSX::Root :$root!--> Spreadsheet::XLSX::SharedStrings) {
        self.new(:$root)
    }

    #| Get the shared string entry at the given position. Creates a fresh
    #| object each time, which an given sheet can cache as a particular
    #| cell position.
    method AT-POS(Int $idx) {
        with $!backing {
            with $!backing.childNodes[$idx] -> LibXML::Element $si {
                if $si.nodeName ne 'si' {
                    die X::Spreadsheet::XLSX::Format.new: message =>
                            'Shared strings entry was not an si node';
                }
                return shared-cell-from-xml($si.first);
            }
        }
        fail "Could not resolve shared string entry $idx";
    }
}
