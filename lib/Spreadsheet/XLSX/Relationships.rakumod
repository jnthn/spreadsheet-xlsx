use LibXML::Document;
use Spreadsheet::XLSX::Exceptions;

#| A relationships file within an XLSX archive.
class Spreadsheet::XLSX::Relationships {
    #| An individual relationship within an archive.
    class Relationship {
        #| The ID of the relationship.
        has Str $.id is required;

        #| The type of relationship it is.
        has Str $.type is required;

        #| The target of the relationship. These are fully qualified
        #| within the archive, rather than what was directly read.
        has Str $.target;

        #| The source of the relationship.
        has Str $.source;
    }

    #| The file these relationships are for (the empty string for the
    #| root relationships).
    has Str $.for is required;

    #| The list of relationships.
    has @.relationships;

    #| By-ID lookup cache.
    has $!id-lookup;

    #| Maximum ID, for if we need to add further ones.
    has $!max-id;

    #| Parse the XML content of a relationships file.
    method from-xml(Str $xml, Str :$for!) {
        my LibXML::Document $doc .= parse(:string($xml));
        my LibXML::Element $root = $doc.documentElement();
        if $root.nodeName ne 'Relationships' {
            die X::Spreadsheet::XLSX::Format.new: message =>
                    'Relationships file did not start with tag Relationships';
        }
        self.new: :$for, relationships => $root.elements.map: -> LibXML::Element $entry {
            Relationship.new:
                    :id(self!get-attribute($entry, 'Id')),
                    :type(self!get-attribute($entry, 'Type')),
                    :target(self!qualify($for, self!get-attribute($entry, 'Target'))),
                    :source(self!get-attribute($entry, 'Source', :optional));
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

    method !qualify(Str $base, Str $relative) {
        with $base.rindex('/') -> $sep-pos {
            $base.substr(0, $sep-pos + 1) ~ $relative
        }
        else {
            $relative
        }
    }

    method !relativize(Str $absolute) {
        with $!for.rindex('/') -> $sep-pos {
            my $remove = $!for.substr(0, $sep-pos + 1);
            unless $absolute.starts-with($remove) {
                die "Confused trying to relativize '$absolute' relative to '$!for'";
            }
            $absolute.substr($remove.chars)
        }
        else {
            $absolute
        }
    }


    #| Find a relationship by ID. Returns a Failure if it is not found.
    method find-by-id(Str $id --> Relationship) {
        without $!id-lookup {
            $!id-lookup = @!relationships.map({ .id => $_ }).hash;
        }
        $!id-lookup{$id} // fail X::Spreadsheet::XLSX::NoSuchRelationship.new(:$id)
    }

    #| Finds all relationships with a specified type.
    method find-by-type(Str $type --> Seq) {
        @!relationships.grep(*.type eq $type)
    }

    #| Adds a new relationship, allocated it an ID.
    method add(Str :$target!, Str :$type!, Str :$source --> Relationship) {
        my $try-id-number = @!relationships.elems;
        my $id;
        repeat {
            $id = 'rId' ~ ++$try-id-number;
        } while self.find-by-id($id);
        my $relationship = Relationship.new(:$id, :$target, :$type, :$source);
        @!relationships.push: $relationship;
        return $relationship;
    }

    #| Turn the relationships into an XML blob.
    method to-xml(--> Blob) {
        # Create root element.
        my LibXML::Document $doc .= new: :version('1.0'), :enc('UTF-8');
        $doc.setStandalone(LibXML::Document::XmlStandaloneNo);
        my LibXML::Element $root = $doc.createElementNS(
                'http://schemas.openxmlformats.org/package/2006/relationships',
                'Relationships');
        $doc.setDocumentElement($root);

        # Add each of the relationships to it.
        for @!relationships {
            my LibXML::Element $element = $doc.createElement('Relationship');
            $element.add($doc.createAttribute('Id', .id));
            $element.add($doc.createAttribute('Type', .type));
            my $relative-target = self!relativize(.target);
            $element.add($doc.createAttribute('Target', $relative-target));
            with .source {
                $element.add($doc.createAttribute('Source', $_));
            }
            $root.add($element);
        }

        return $doc.Blob;
    }
}
