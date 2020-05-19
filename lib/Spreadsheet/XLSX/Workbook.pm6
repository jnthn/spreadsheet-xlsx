use LibXML::Document;
use Spreadsheet::XLSX::Exceptions;
use Spreadsheet::XLSX::Relationships;
use Spreadsheet::XLSX::Root;
use Spreadsheet::XLSX::SharedStrings;
use Spreadsheet::XLSX::Worksheet;

#| The XLSX workbook
class Spreadsheet::XLSX::Workbook {
    #| The root, used for resolutions at the document level.
    has Spreadsheet::XLSX::Root $.root is required;

    #| The relationships of the workbook.
    has Spreadsheet::XLSX::Relationships $.relationships;

    #| The shared strings of the workbook.
    has Spreadsheet::XLSX::SharedStrings $.shared-strings;

    #| The list of worksheets in the workbook.
    has @!worksheets;

    #| The backing XML document, if any.
    has LibXML::Document $!backing;

    submethod TWEAK(:$!backing, :@!worksheets --> Nil) {}

    #| Parse the XML content of a relationships file.
    method from-xml(Str $xml, Spreadsheet::XLSX::Root :$root!,
                    Spreadsheet::XLSX::Relationships :$relationships!) {
        my LibXML::Document $doc .= parse(:string($xml));
        my LibXML::Element $workbook = $doc.documentElement();
        if $workbook.nodeName ne 'workbook' {
            die X::Spreadsheet::XLSX::Format.new: message =>
                    'Workbook file did not start with tag workbook';
        }
        with $workbook.childNodes.list.first(*.name eq 'sheets') -> LibXML::Element $sheets-node {
            my @worksheets = $sheets-node.childNodes.map: -> LibXML::Element $sheet-node {
                my $name := self!get-attribute($sheet-node, 'name');
                my $sheet-rel-id := self!get-attribute($sheet-node, 'r:id');
                with $relationships.find-by-id($sheet-rel-id) -> $sheet-rel {
                    my $backing-path = $sheet-rel.target;
                    Spreadsheet::XLSX::Worksheet.new(:$root, :$name, :$backing-path)
                }
                else {
                    die X::Spreadsheet::XLSX::Format.new: message =>
                            "Could not resolve sheet relationship $sheet-rel-id";
                }
            }
            my $shared-strings = get-shared-strings($root, $relationships);
            self.new(:$root, :$relationships, :@worksheets, :$shared-strings, :backing($doc))
        }
        else {
            die X::Spreadsheet::XLSX::Format.new: message =>
                    'Required sheets element not found in workbook'
        }
    }

    sub get-shared-strings(Spreadsheet::XLSX::Root $root, Spreadsheet::XLSX::Relationships $relationships) {
        with $relationships.find-by-type('http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings').first -> $rel {
            with $root.get-file-from-archive($rel.target) {
                Spreadsheet::XLSX::SharedStrings.from-xml(.decode('utf-8'), :$root)
            }
            else {
                die X::Spreadsheet::XLSX::Format.new: message =>
                        "Could not find shared strings file $rel.target()"
            }
        }
        else {
            Spreadsheet::XLSX::SharedStrings.empty(:$root)
        }
    }

    method !get-attribute(LibXML::Element $entry, Str $name, :$optional --> Str) {
        with $entry.getAttributeNode($name) -> LibXML::Attr $attr {
            $attr.string-value
        }
        elsif $optional {
            Nil
        }
        else {
            die X::Spreadsheet::XLSX::Format.new: message =>
                    "Missing attribute '$name' on '$entry.nodeName()'";
        }
    }

    #| Create a new worksheet in this workbook.
    method create-worksheet(Str $name --> Spreadsheet::XLSX::Worksheet) {
        my $worksheet = Spreadsheet::XLSX::Worksheet.new(:$!root :$name);
        @!worksheets.push($worksheet);
        return $worksheet;
    }

    #| Get a list of the worksheets in this workbook.
    method worksheets(--> List) {
        @!worksheets.List
    }
}
