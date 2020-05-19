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

    submethod TWEAK(LibXML::Document :$!backing, :@!worksheets --> Nil) {}

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
                my $id := self!get-attribute($sheet-node, 'sheetId').Int;
                my $name := self!get-attribute($sheet-node, 'name');
                my $sheet-rel-id := self!get-attribute($sheet-node, 'r:id');
                with $relationships.find-by-id($sheet-rel-id) -> $sheet-rel {
                    my $backing-path = $sheet-rel.target;
                    Spreadsheet::XLSX::Worksheet.new(:$root, :$id, :$name, :$backing-path)
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
        # Give it a unique ID and file path (in all but weird cases,
        # the filename will just be sheetN where N is the ID, but we
        # try and be robust in the face of oddities).
        my $id = @!worksheets ?? @!worksheets.map(*.id).max + 1 !! 1;
        my $proposed-path = 'xl/worksheets/sheet' ~ $id ~ '.xml';
        while $!root.get-file-from-archive($proposed-path) {
            $proposed-path = 'xl/worksheets/sheet' ~ ++$id ~ '.xml'
        }

        # Create the worksheet.
        my $worksheet = Spreadsheet::XLSX::Worksheet.new(:$!root :$id, :$name, :$proposed-path);
        @!worksheets.push($worksheet);

        # Add the relationship.
        $!relationships.add:
                type => 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet',
                target => $proposed-path;

        return $worksheet;
    }

    #| Get a list of the worksheets in this workbook.
    method worksheets(--> List) {
        @!worksheets.List
    }
}
