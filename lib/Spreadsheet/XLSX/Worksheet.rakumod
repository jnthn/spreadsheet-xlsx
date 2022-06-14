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
            my $cell := @row[$col];
            $cell //= self!maybe-load-from-backing($row, $col);
            $cell
        }

        multi method ASSIGN-POS(Int $row, Int $col, Spreadsheet::XLSX::Cell $value) {
            my @row := (@!rows[$row] //= Array[Spreadsheet::XLSX::Cell].new);
            @row[$col] = $value
        }

        method !maybe-load-from-backing(Int $row, Int $col) {
            with self!lookup-backing-row($row) -> LibXML::Element $backing-row {
                my ($from, $to) = get-attribute($backing-row, "spans").split(':');
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
            with $!backing {
                unless @!backing-rows {
                    $!backing.childNodes.map: -> LibXML::Element $backing-row {
                        if $backing-row.nodeName eq 'row' {
                            my $row-str = get-attribute($backing-row, 'r');
                            @!backing-rows[$row-str.Int - 1] = $backing-row;
                        }
                    }
                }
            }
            @!backing-rows[$row]
        }

        method !load-cell(LibXML::Element $cell --> Spreadsheet::XLSX::Cell) {
            my $type = get-attribute($cell, 't', :optional) // '';
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

        #| Synchronize the values of cells we've set/changed with the
        #| sheet data XML document.
        method sync-sheet-data-xml(LibXML::Element $sheetData,
                Spreadsheet::XLSX::Styles $styles --> Nil) {
            # Look through the rows we have cell data for; these are ones we
            # have looked at or potentially modified.
            my LibXML::Document $document = $sheetData.getOwnerDocument;
            for @!rows.kv -> $row-idx, $cells {
                # If we never vivified this row, then either it has no data or
                # we didn't update what was already there.
                next without $cells;

                # Do we have a backing row for this cell? If not, create one and
                # add it.
                my LibXML::Element $row-xml = @!backing-rows[$row-idx] // do {
                    my LibXML::Element $row = $document.createElement('row');
                    $sheetData.add($row);
                    $row.add($document.createAttribute('r', ~($row-idx + 1)));
                    $row
                }

                # Build a map of existing cells in the row, identified by their
                # cell name (A2, B2, etc.)
                my %existing-cells = $row-xml.childNodes.map({
                    get-attribute($_, 'r') => $_
                });

                # Go over the cells we have a value for. These are the ones we
                # potentially want to update. Also keep track of the minimum and
                # maximum cell index we change.
                my $min-col = Inf;
                my $max-col = -Inf;
                for $cells.kv -> $col-idx, Spreadsheet::XLSX::Cell $cell {
                    with $cell {
                        # Update info we'll use for updating the span.
                        $min-col min= $col-idx;
                        $max-col max= $col-idx;

                        # Obtain the cell node from the existing document or
                        # make a new one.
                        my $cell-name = self!cell-name($row-idx, $col-idx);
                        my LibXML::Element $cell-xml = do with %existing-cells{$cell-name} {
                            $_
                        }
                        else {
                            my LibXML::Element $cell = $document.createElement('c');
                            $row-xml.add($cell);
                            $cell.add($document.createAttribute('r', $cell-name));
                            $cell
                        }

                        # Sync the data to it.
                        $cell.sync-value-xml($document, $cell-xml);

                        # Sync the style.
                        my $s-attr = $cell-xml.getAttributeNode('s');
                        with $cell.style {
                            with .sync-style-id($styles) -> $style-id {
                                with $s-attr {
                                    $s-attr.setValue(~$style-id);
                                }
                                else {
                                    $cell-xml.add($document.createAttribute('s', ~$style-id));
                                }
                            }
                        }
                        orwith $s-attr {
                            $cell-xml.removeChild($s-attr);
                        }
                    }
                }

                # Set or update the span.
                with $row-xml.getAttributeNode('spans') -> LibXML::Attr $spans {
                    my ($prev-min, $prev-max) = $spans.string-value.split(':').map(* - 1);
                    $min-col min= $prev-min;
                    $max-col max= $prev-max;
                    $spans.setValue($min-col + 1 ~ ':' ~ $max-col + 1);
                }
                else {
                    $row-xml.add($document.createAttribute('spans', $min-col + 1 ~ ':' ~ $max-col + 1));
                }
            }
        }

        #| Turns 0-based array indices for row and column into the cell
        #| name.
        method !cell-name(Int $row, Int $col) {
            ($col < 26 ?? chr(ord('A') + $col) !! self!big-col-name($col)) ~ ($row + 1)
        }

        method !big-col-name(Int $col) {
            my constant $offset = 'A'.ord - '0'.ord;
            $col.base(26).comb.map({ chr .ord + $offset }).join
        }
    }

    #| Properties associated with a column in a worksheet.
    class Column {
        has Real $.width;
        has Int $!style-id;
        has Bool $.hidden;
        has Bool $.best-fit;
        has Bool $.custom-width;
        has Bool $.phonetic;
        has Int $.outline-level;
        has Bool $.collapsed;

        submethod TWEAK(Int :$!style-id) { }

        #| Load the column from the XML col element.
        method from-xml(LibXML::Element $col --> Column) {
            my %properties;
            with get-attribute($col, 'width', :optional) {
                %properties<width> = +$_;
            }
            with get-attribute($col, 'style', :optional) {
                %properties<style-id> = .Int;
            }
            with get-boolean-attribute($col, 'hidden', :optional) {
                %properties<hidden> = $_;
            }
            with get-boolean-attribute($col, 'bestFit', :optional) {
                %properties<best-fit> = $_;
            }
            with get-boolean-attribute($col, 'customWidth', :optional) {
                %properties<custom-width> = $_;
            }
            with get-boolean-attribute($col, 'phonetic', :optional) {
                %properties<phonetic> = $_;
            }
            with get-attribute($col, 'outlineLevel', :optional) {
                %properties<outline-level> = .Int;
            }
            with get-boolean-attribute($col, 'collapsed', :optional) {
                %properties<collapsed> = $_;
            }
            self.new(|%properties)
        }

        #| Given a col XML element, set attributes on it based on what is
        #| configured in this object.
        method set-xml-attributes(LibXML::Document $backing, LibXML::Element $col-xml --> Nil) {
            $col-xml.add($backing.createAttribute('width', ~$_)) with $!width;
            $col-xml.add($backing.createAttribute('style', ~$_)) with $!style-id;
            $col-xml.add($backing.createAttribute('hidden', bool-xml($_))) with $!hidden;
            $col-xml.add($backing.createAttribute('bestFit', bool-xml($_))) with $!best-fit;
            $col-xml.add($backing.createAttribute('customWidth', bool-xml($_))) with $!custom-width;
            $col-xml.add($backing.createAttribute('phonetic', bool-xml($_))) with $!phonetic;
            $col-xml.add($backing.createAttribute('outlineLevel', ~$_)) with $!outline-level;
            $col-xml.add($backing.createAttribute('collapsed', bool-xml($_))) with $!collapsed;
        }

        sub bool-xml($value) {
            $value ?? 'true' !! 'false'
        }
    }

    #| The cached backing document, if we have one.
    has LibXML::Document $!backing;

    #| The cached cells object, created lazily if we have a backing
    #| document.
    has Cells $!cells;

    #| Array of columns. Populated from an existing worksheet on
    #| first access.
    has Array $!columns;

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

    #| Set the specified row and column to a text cell with the specified
    #| value, and apply the specified styles also. Any existing styles will
    #| be cleared.
    multi method set(Int $row, Int $col, Str $value, *%styles --> Spreadsheet::XLSX::Cell) {
        self!set($row, $col, Spreadsheet::XLSX::Cell::Text.new(:$value), %styles)
    }

    #| Set the specified row and column to a number cell with the specified
    #| value, and apply the specified styles also. Any existing styles will
    #| be cleared.
    multi method set(Int $row, Int $col, Real $value, *%styles --> Spreadsheet::XLSX::Cell) is default {
        self!set($row, $col, Spreadsheet::XLSX::Cell::Number.new(:value(+$value)), %styles)
    }

    method !set(Int $row, Int $col, Spreadsheet::XLSX::Cell $cell, %styles) {
        for %styles.kv -> $method, $value {
            $cell.style."$method"() = $value;
        }
        self.cells.ASSIGN-POS($row, $col, $cell);
    }

    #| Get column properties for the worksheet. Each column that has some
    #| associated properties will have an instance.
    method columns(--> Array) {
        $!columns //= self!setup-columns();
    }

    method !setup-columns() {
        my @columns;
        with $!backing-path {
            my LibXML::Element $doc-root := self!get-backing-document().documentElement();
            with $doc-root.childNodes.list.first(*.nodeName eq 'cols') -> LibXML::Element $cols {
                for $cols.childNodes -> LibXML::Element $col {
                    for get-attribute($col, 'min').Int .. get-attribute($col, 'max').Int {
                        @columns[$_ - 1] = Column.from-xml($col);
                    }
                }
            }
        }
        return @columns;
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

    #| Synchronizes all changes to the internal representation of the
    #| archive.
    method sync-to-archive(--> Nil) {
        # If we have a backing path, but never loaded from it, then we
        # don't need to update the archive.
        return if $!backing-path.defined && !$!backing.defined;

        # Otherwise, update the XML.
        $!root.set-file-in-archive($.archive-path, self.to-xml().encode('utf-8'));
    }

    #| Produce XML for the worksheet.
    method to-xml(--> Str) {
        # Create a stub worksheet document if we weren't loaded from
        # one.
        without $!backing {
            $!backing .= new: :version('1.0'), :enc('UTF-8');
            $!backing.setStandalone(LibXML::Document::XmlStandaloneNo);
            my LibXML::Element $root = $!backing.createElementNS(
                    'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
                    'worksheet');
            $root.addNamespace('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'r');
            $!backing.setDocumentElement($root);
            my LibXML::Element $sheetData = $!backing.createElement('sheetData');
            $root.add($sheetData);
        }

        # Update the sheet data.
        my @node-list := $!backing.documentElement.childNodes.list;
        my LibXML::Element $sheetData = @node-list.first(*.name eq 'sheetData');
        with $sheetData {
            .sync-sheet-data-xml($sheetData, $!root.styles) with $!cells;
        }
        else {
            die X::Spreadsheet::XLSX::Format.new:
                    message => 'Missing sheetData element';
        }

        # Update the column data.
        with $!columns {
            my $cols-xml = @node-list.first(*.name eq 'cols');
            if $!columns.first(*.defined) {
                # We have columns info to save.
                without $cols-xml {
                    $cols-xml = $!backing.createElement('cols');
                    $!backing.documentElement.insertBefore($cols-xml, $sheetData);
                }
                self!sync-columns-into($cols-xml);
            }
            else {
                # No columns. An empty cols node is disallowed per the
                # schema, so we must delete it if present.
                with $cols-xml {
                    $!backing.documentElement.removeChild($cols-xml);
                }
            }
        }

        return $!backing.Str;
    }

    method !sync-columns-into(LibXML::Element $cols) {
        for $!columns.kv -> Int $idx, $col {
            with $col {
                my $col-xml = $!backing.createElement('col');
                my $idx-str = ($idx + 1).Str;
                $col-xml.add($!backing.createAttribute('min', $idx-str));
                $col-xml.add($!backing.createAttribute('max', $idx-str));
                $col.set-xml-attributes($!backing, $col-xml);
                $cols.add($col-xml);
            }
        }
    }

    sub get-attribute(LibXML::Element $entry, Str $name, :$optional --> Str) {
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

    sub get-boolean-attribute(LibXML::Element $entry, Str $name, :$optional --> Bool) {
        with get-attribute($entry, $name, :$optional) {
            so $_ eq 'true' | '1'
        }
        else {
            Nil
        }
    }
}
