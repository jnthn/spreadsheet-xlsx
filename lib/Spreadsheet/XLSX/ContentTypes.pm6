use LibXML::Document;
use Spreadsheet::XLSX::Exceptions;

#| The content types map for the XLSX file.
class Spreadsheet::XLSX::ContentTypes {
    #| A default entry in the content types.
    class Default {
        has Str $.extension;
        has Str $.content-type;
    }

    #| An override entry in the content types.
    class Override {
        has Str $.part-name;
        has Str $.content-type;
    }

    #| List of extensions specified in the content types.
    has Default @.defaults;

    #| List of overrides specified in the content types.
    has Override @.overrides;

    #| Parse the XML content of a [Content_Types].xml.
    method from-xml(Str $xml) {
        my LibXML::Document $doc .= parse(:string($xml));
        my LibXML::Element $root = $doc.documentElement();
        my (@defaults, @overrides);
        if $root.nodeName ne 'Types' {
            die X::Spreadsheet::XLSX::Format.new: message =>
                    'Content types did not start with tag Types';
        }
        for $root.childNodes -> LibXML::Element $entry {
            if $entry.nodeName eq 'Default' {
                @defaults.push: Default.new:
                        extension => self!get-attribute($entry, 'Extension'),
                        content-type => self!get-attribute($entry, 'ContentType');
            }
            elsif $entry.nodeName eq 'Override' {
                @overrides.push: Override.new:
                        part-name => self!get-attribute($entry, 'PartName'),
                        content-type => self!get-attribute($entry, 'ContentType');
            }
        }
        self.bless(:@defaults, :@overrides)
    }

    method !get-attribute(LibXML::Element $entry, Str $name --> Str) {
        with $entry.getAttributeNode($name) -> LibXML::Attr $attr {
            $attr.string-value
        }
        else {
            die X::Spreadsheet::XLSX::Format.new: message =>
                    "Missing attribute '$name' on '$entry.nodeName()'";
        }
    }

    #| Creates a content types file with the default set of content types
    #| needed for a basic XLSX file.
    method new(--> Spreadsheet::XLSX::ContentTypes) {
        my constant @default-defaults =
                Default.new(extension => 'rels', content-type => 'application/vnd.openxmlformats-package.relationships+xml'),
                Default.new(extension => 'xml', content-type => 'application/xml'),;
        my constant @default-overrides =
                Override.new(part-name => '/xl/workbook.xml', content-type => 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml'),;
        self.bless(defaults => @default-defaults, overrides => @default-overrides)
    }

    #| Finds the (first) part name with a given content type.
    method part-name-for-content-type(Str $content-type --> Str) {
        self.part-names-for-content-type($content-type).first
    }

    #| Finds part names with a certain content type.
    method part-names-for-content-type(Str $content-type --> Seq) {
        @!overrides.grep(*.content-type eq $content-type).map(*.part-name)
    }

    #| Turn the content types into an XML string.
    method to-xml(--> Str) {
        # Create root element.
        my LibXML::Document $doc .= new: :version('1.0'), :enc('UTF-8');
        $doc.setStandalone(LibXML::Document::XmlStandaloneNo);
        my LibXML::Element $root = $doc.createElementNS(
                'http://schemas.openxmlformats.org/package/2006/content-types',
                'Types');
        $doc.setDocumentElement($root);

        # Add defaults.
        for @!defaults {
            my LibXML::Element $element = $doc.createElement('Default');
            $element.add($doc.createAttribute('Extension', .extension));
            $element.add($doc.createAttribute('ContentType', .content-type));
            $root.add($element);
        }

        # Add overrides.
        for @!overrides {
            my LibXML::Element $element = $doc.createElement('Override');
            $element.add($doc.createAttribute('PartName', .part-name));
            $element.add($doc.createAttribute('ContentType', .content-type));
            $root.add($element);
        }

        return $doc.Str;
    }
}
