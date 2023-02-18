use LibXML::Document;
use Spreadsheet::XLSX::Exceptions;
use Spreadsheet::XLSX::Relationships;
use Spreadsheet::XLSX::Root;
use Spreadsheet::XLSX::SharedStrings;
use Spreadsheet::XLSX::Styles;
use Spreadsheet::XLSX::Worksheet;

#| The XLSX workbook
class Spreadsheet::XLSX::Workbook {
    #| The root, used for resolutions at the document level.
    has Spreadsheet::XLSX::Root $.root is required;

    #| The relationships of the workbook.
    has Spreadsheet::XLSX::Relationships $.relationships;

    #| The shared strings of the workbook.
    has Spreadsheet::XLSX::SharedStrings $.shared-strings;

    #| The styles of the workbook.
    has Spreadsheet::XLSX::Styles $.styles;

    #| The list of worksheets in the workbook.
    has @!worksheets;

    #| The backing XML document, if any.
    has LibXML::Document $!backing;

    submethod TWEAK(LibXML::Document :$!backing, :@!worksheets --> Nil) {
        $!styles //= Spreadsheet::XLSX::Styles.new(:$!root);
    }

    #| Parse the XML content of a workbook file.
    method from-xml(Str $xml, Spreadsheet::XLSX::Root :$root!,
                    Spreadsheet::XLSX::Relationships :$relationships!) {
        my LibXML::Document $doc .= parse(:string($xml));
        my LibXML::Element $workbook = $doc.documentElement();
        if $workbook.nodeName ne 'workbook' {
            die X::Spreadsheet::XLSX::Format.new: message =>
                    'Workbook file did not start with tag workbook';
        }
        with $workbook.elements.list.first(*.name eq 'sheets') -> LibXML::Element $sheets-node {
            my @worksheets = $sheets-node.elements.map: -> LibXML::Element $sheet-node {
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
            my $styles = get-styles($root, $relationships);
            self.new(:$root, :$relationships, :@worksheets, :$shared-strings,
                    :$styles, :backing($doc))
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

    sub get-styles(Spreadsheet::XLSX::Root $root, Spreadsheet::XLSX::Relationships $relationships) {
        with $relationships.find-by-type('http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles').first -> $rel {
            with $root.get-file-from-archive($rel.target) {
                Spreadsheet::XLSX::Styles.from-xml(.decode('utf-8'), :$root)
            }
            else {
                die X::Spreadsheet::XLSX::Format.new: message =>
                        "Could not find styles file $rel.target()"
            }
        }
        else {
            Spreadsheet::XLSX::Styles.new(:$root)
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

        # Add content type override.
        $!root.content-types.add-override:
                part-name => "/$proposed-path",
                content-type => 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml';

        return $worksheet;
    }

    #| Get a list of the worksheets in this workbook.
    method worksheets(--> List) {
        @!worksheets.List
    }

    #| Synchronizes all changes to the internal representation of the
    #| archive, including all new or modified worksheets.
    method sync-to-archive(--> Nil) {
        # Save the XML of this workbook.
        $!root.set-file-in-archive($!relationships.for, self.to-xml());

        # And then the worksheets (which will sync the styles used by
        # cells also).
        for @!worksheets {
            .sync-to-archive;
        }

        # Sync the styles.
        my constant $style-type = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles';
        my $styles-path = do with $!relationships.find-by-type($style-type).first {
            .target
        }
        else {
            my $target = 'xl/styles.xml';
            $!relationships.add(:$target, :type($style-type));
            $!root.content-types.add-override:
                    part-name => "/$target",
                    content-type => 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml';
            $target
        }
        $!root.set-file-in-archive($styles-path, $!styles.to-xml());
    }

    #| Form an XML representation of the workbook. If the workbook was loaded
    #| from an existing XML document, then it will just change the parts of
    #| that backing storage that it understands, and aim to leave the rest of
    #| it intact.
    method to-xml(--> Blob) {
        # Create baseline backing storage, if there isn't any. Otherwise,
        # just locate the sheets part.
        my LibXML::Element $sheets = do with $!backing {
            # This will work out for sure, 'cus if it didn't we'd not have
            # successfully constructed this instance.
            my $workbook = $!backing.documentElement;
            $workbook.elements.list.first(*.name eq 'sheets')
        }
        else {
            $!backing .= new: :version('1.0'), :enc('UTF-8');
            $!backing.setStandalone(LibXML::Document::XmlStandaloneNo);
            my LibXML::Element $root = $!backing.createElementNS(
                    'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
                    'workbook');
            $root.addNamespace('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'r');
            $!backing.setDocumentElement($root);
            my LibXML::Element $sheets = $!backing.createElement('sheets');
            $root.add($sheets);
            $sheets
        }

        # Clear out the sheets that currently exist, and replace them
        # with those that we have.
        $sheets.removeChildNodes();
        my %sheet-path-to-id = $!relationships
                .find-by-type('http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet')
                .map({ .target => .id });
        for @!worksheets {
            my $sheet = $!backing.createElement('sheet');
            $sheet.add($!backing.createAttribute('name', .name));
            $sheet.add($!backing.createAttribute('sheetId', ~.id));
            $sheet.add($!backing.createAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id',
                    %sheet-path-to-id{.archive-path} // die "Missing reference for sheet {.archive-path}"));
            $sheets.add($sheet);
        }

        return $!backing.Blob;
    }
}
