use LibXML::Document;
use Spreadsheet::XLSX::Cell;
use Spreadsheet::XLSX::Root;
use Spreadsheet::XLSX::XMLHelpers;
use Spreadsheet::XLSX::Types;

#| Since spreadsheets often contain a lot of repetition, a shared strings
#| table is used to extract the commonality.
class Spreadsheet::XLSX::SharedStrings is xml-sequence("sst", :si(Spreadsheet::XLSX::Types::CT_Rst)) {
    #| The root, used for resolutions at the document level.
    has Spreadsheet::XLSX::Root $.root is required;

    has UInt $.count is xml-attr;
    has UInt $.unique-count is xml-attr;

    #| Create a string table from an XML document.
    method from-xml(Str $xml, Spreadsheet::XLSX::Root :$root! --> Spreadsheet::XLSX::SharedStrings) {
        my LibXML::Document $doc .= parse(:string($xml));
        my LibXML::Element $sst = $doc.documentElement();
        if $sst.nodeName ne 'sst' {
            die X::Spreadsheet::XLSX::Format.new: message =>
                    'Shared strings file did not start with tag sst';
        }
        self.from-xml-element($sst, :profile{ :$root })
    }

    #| Create a new, empty, string table.
    method empty(Spreadsheet::XLSX::Root :$root!--> Spreadsheet::XLSX::SharedStrings) {
        self.new(:$root)
    }
}
