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
        self.new(:@defaults, :@overrides)
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

    #| Finds the (first) part name with a given content type.
    method part-name-for-content-type(Str $content-type --> Str) {
        self.part-names-for-content-type($content-type).first
    }

    #| Finds part names with a certain content type.
    method part-names-for-content-type(Str $content-type --> Seq) {
        @!overrides.grep(*.content-type eq $content-type).map(*.part-name)
    }
}
