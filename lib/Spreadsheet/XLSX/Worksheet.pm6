use LibXML::Document;
use Spreadsheet::XLSX::Cell;
use Spreadsheet::XLSX::Root;

#| A particular worksheet within an Excel workbook.
class Spreadsheet::XLSX::Worksheet {
    #| The root, used for resolutions at the document level.
    has Spreadsheet::XLSX::Root $.root is required;

    #| The worksheet ID within the workbook.
    has Int $.id is required;

    #| The name of the worksheet.
    has Str $.name is rw is required;

    #| If the sheet is loaded from an existing file, this is its path.
    #| We lazily load from this (meaning if we have a workbook with many
    #| sheets, we only load those we need to).
    has Str $!backing-path;

    #| Otherwise, this is the proposed path for it, where it will go when
    #| we save it.
    has Str $!proposed-path;

    #| Models the cells in the spreadsheet, providing access as a 2D array.
    class Cells does Positional {
        #| The enclosing worksheet.
        has Spreadsheet::XLSX::Worksheet $.worksheet is required;

        #| Sheet data element, if we're loaded from a file.
        has LibXML::Element $!backing;

        #| Quick lookup of rows, since the underying representation may be
        #| sparse.
        has LibXML::Element @!backing-rows;

        #| Cached cells (those we have created if we're making a new sheet,
        #| or those we have read and/or modified if we've based on a document
        #| that we've read).
        has Array[Spreadsheet::XLSX::Cell] @!rows;

        submethod TWEAK(LibXML::Element :$!backing --> Nil) {}

        multi method AT-POS(Int $row, Int $col) is raw {
            my @row := (@!rows[$row] //= Array[Spreadsheet::XLSX::Cell].new);
            @row[$col] //= self!maybe-load-from-backing($row, $col);
        }

        method !maybe-load-from-backing(Int $row, Int $col) {
            with self!lookup-backing-row($row) -> LibXML::Element $backing-row {
                my ($from, $to) = self!get-attribute($backing-row, "spans").split(':');
                if $from <= $col + 1 <= $to {
                    my LibXML::Element $doc-col = $backing-row.childNodes[$col - ($from - 1)];
                    if $doc-col && $doc-col.nodeName eq 'c' {
                        return self!load-cell($doc-col);
                    }
                }
            }
            return Spreadsheet::XLSX::Cell;
        }

        method !lookup-backing-row($row) {
            unless @!backing-rows {
                $!backing.childNodes.map: -> LibXML::Element $backing-row {
                    if $backing-row.nodeName eq 'row' {
                        my $row-str = self!get-attribute($backing-row, 'r');
                        @!backing-rows[$row-str.Int - 1] = $backing-row;
                    }
                }
            }
            @!backing-rows[$row]
        }

        method !load-cell(LibXML::Element $cell --> Spreadsheet::XLSX::Cell) {
            my $type = self!get-attribute($cell, 't', :optional) // '';
            if $type eq 's' {
                my LibXML::Element $shared-index-holder = $cell.first;
                unless $shared-index-holder.nodeName eq 'v' {
                    die X::Spreadsheet::XLSX::Format.new:
                            message => "Missing v node for shared cell value";
                }
                $!worksheet.root.shared-strings[$shared-index-holder.string-value.Int]
            }
            else {
                cell-from-xml($cell)
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
    }

    #| The cached backing document, if we have one.
    has LibXML::Document $!backing;

    #| The cached cells object, created lazily if we have a backing
    #| document.
    has Cells $!cells;

    submethod TWEAK(Str :$!backing-path, Str :$!proposed-path --> Nil) {}

    #| Get the cells model, which can be indexed using a 2-dimensional
    #| index (e.g. $worksheet.cells[1;2]).
    method cells(--> Cells) {
        $!cells //= self!setup-cells();
    }

    method !setup-cells() {
        with $!backing-path {
            my LibXML::Element $doc-root := self!get-backing-document().documentElement();
            with $doc-root.childNodes.list.first(*.nodeName eq 'sheetData') {
                Cells.new(:worksheet(self), :backing($_))
            }
            else {
                die X::Spreadsheet::XLSX::Format.new: message =>
                        'Worksheet file did not contain sheetData element';
            }
        }
        else {
            Cells.new(:worksheet(self))
        }
    }

    method !get-backing-document(--> LibXML::Document) {
        without $!backing {
            with $!root.get-file-from-archive($!backing-path) {
                my LibXML::Document $doc .= parse(:string(.decode('utf-8')));
                my LibXML::Element $root = $doc.documentElement();
                if $root.nodeName ne 'worksheet' {
                    die X::Spreadsheet::XLSX::Format.new: message =>
                            'Worksheet file did not start with tag worksheet';
                }
                $!backing = $doc;
            }
            else {
                die X::Spreadsheet::XLSX::Format.new:
                        message => "Missing worksheet file '$!backing-path'"
            }
        }
        $!backing
    }

    #| The path of the sheet in the XLSX archive.
    method archive-path(--> Str) {
        $!backing-path // $!proposed-path
    }
}
