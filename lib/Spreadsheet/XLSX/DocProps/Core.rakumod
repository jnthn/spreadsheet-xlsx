use v6.d;
unit class Spreadsheet::XLSX::DocProps::Core;

use LibXML::Document;
use LibXML::Element;
use Spreadsheet::XLSX::Exceptions;

also does Associative;

has LibXML::Document $!backing is built;

#| Storage for already loaded properties
has %!properties{Mu};

#| Namespace names actually used by XML
has %!ns-map;

# Standard namespaces we expect. Keys represent conventional namespace IDs but they do not seem to be standardized,
# contrary to their URIs. Therefore the keys are used for convenience and we'd have to find the actual mapping of URI
# into its respective NS name from the XML.
my %cp-ns =
    cp       => "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
    dc       => "http://purl.org/dc/elements/1.1/",
    dcterms  => "http://purl.org/dc/terms/",
    dcmitype => "http://purl.org/dc/dcmitype/",
    xsi      => "http://www.w3.org/2001/XMLSchema-instance";

#| Represent keywords property of the core properties.
class Keywords {
    has $!default-lang = "en-US";
    has %!by-lang;

    multi method new(LibXML::Element:D $elem) {
        samewith()!FROM-XML($elem);
    }
    multi method new(\kwd) {
        given samewith() {
            .set(kwd);
            $_
        }
    }

    method add(Str:D $keyword, Str:D :$lang = $!default-lang) {
        %!by-lang{$lang}.push: $keyword;
    }

    method remove(Mu \kwd-match, Str:D :$lang = $!default-lang) {
        %!by-lang{$lang} = %!by-lang{$lang}.grep(* !~~ kwd-match)
    }

    proto method set(|) {*}
    multi method set(Array:D \keywords, Str:D :$lang = $!default-lang) {
        %!by-lang{$lang} = keywords;
    }
    multi method set(Positional:D \keywords, Str:D :$lang = $!default-lang) {
        %!by-lang{$lang} = [|keywords];
    }
    multi method set(Str:D $keyword, Str:D :$lang = $!default-lang) {
        %!by-lang{$lang} = [$keyword];
    }
    multi method set($kwd, Str:D :$lang = $!default-lang) {
        %!by-lang{$lang} = [~$kwd];
    }

    multi method list(::?CLASS:D:)             { %!by-lang.keys.map({ %!by-lang{$_}.values }).flat }
    multi method list(::?CLASS:D: Str:D $lang) { %!by-lang{$lang}.list // () }

    multi method pairs(::?CLASS:D:) { %!by-lang.keys.map({ $_ => %!by-lang{$_}.List }).flat }

    multi method List { self.list }

    method !FROM-XML(LibXML::Element:D $elem) {
        unless $elem.localName eq 'keywords' {
            die X::Spreadsheet::XLSX::Format.new:
                message => "Keywords must start with tag 'keywords', got '" ~ $elem.localName ~ "' instead"
        }
        $!default-lang = $_ with $elem.getAttribute('xml:lang');
        for $elem.childNodes -> $chld {
            my $lang = $!default-lang;
            if $chld ~~ LibXML::Element {
                unless $chld.localName eq 'value' {
                    die X::Spreadsheet::XLSX::Format.new:
                        message => "Keyword child must be a 'value' tag, got '" ~ $chld.localName ~ "'"
                }
                $lang = $_ with $chld.getAttribute('xml:lang');
            }
            %!by-lang{$lang}.append: $chld.textContent.trim.split(/\s+/);
        }
        self
    }
    # $elem is expected to be an empty <keywords />
    multi method to-xml(LibXML::Element:D $elem) {
        $elem.setAttribute('xml:lang', $!default-lang);
        for %!by-lang.keys -> $lang {
            if $lang eq $!default-lang {
                $elem.appendText: %!by-lang{$lang}.join(" ");
            }
            else {
                for %!by-lang{$lang}.list -> $keyword {
                    $elem.appendTextChild('value', $keyword).setAttribute('xml:lang', $lang);
                }
            }
        }
        $elem
    }
    multi method to-xml() {
        samewith LibXML::Element.new('keywords')
    }

    method Str { self.list.sort.join(" ") }
    method gist { %!by-lang.gist }
}

my class CPMeta {
    has Str $.ns;                   # a key from %cp-ns
    has Mu $.type is built(:bind);  # user-facing data type
    has Bool $.complex;             # if set the type must be coerced from corresponding XML node directly
    has %.xml-attrs;                # mandatory XML attributes
    my %cache;
    multi method new(Str:D :$ns!, Mu :$type = Str, *%props) {
        %props
            ?? nextsame()
            !! ( (%cache{$ns} //= my %{Mu}){$type} //= callwith(:$ns, :$type) )
    }
    multi method new(Str:D $ns, Mu $type = Str, Bool:D :$complex = False) {
        samewith :$ns, :$type, :$complex
    }
    method from-xml(LibXML::Element:D $elem) {
        $!complex ?? $!type.new($elem) !! $!type($elem.textContent)
    }
}

my %known-props =
    category        => CPMeta.new("cp"),
    contentStatus   => CPMeta.new("cp"),
    created         => CPMeta.new("dcterms", DateTime, xml-attrs => { "xsi:type" => "dcterms:W3CDTF" }),
    creator         => CPMeta.new("dc"),
    description     => CPMeta.new("dc"),
    identifier      => CPMeta.new("dc"),
    keywords        => CPMeta.new("cp", Keywords(), :complex),
    language        => CPMeta.new("dc"),
    lastModifiedBy  => CPMeta.new("cp"),
    lastPrinted     => CPMeta.new("cp", DateTime),
    modified        => CPMeta.new("dcterms", DateTime, xml-attrs => { "xsi:type" => "dcterms:W3CDTF" }),
    revision        => CPMeta.new("cp"),
    subject         => CPMeta.new("dc"),
    title           => CPMeta.new("dc"),
    version         => CPMeta.new("cp");

method !get-ns-name(Str:D $key) {
    # If there is no backing element then we use the defaults
    %!ns-map{$key} //=
        ( $!backing andthen $!backing.documentElement andthen
            ( .lookupNamespacePrefix(%cp-ns{$key})
                // die X::Spreadsheet::XLSX::Format.new:
                    message => "Namespace '{%cp-ns{$key}}' is missing in core properties XML" ) )
        orelse $key
}

method !is-prop-known($property) {
    %known-props{$property}:exists or die X::Spreadsheet::XLSX::NoSuchProperty.new: :$property
}

method !maybe-prop-from-backing($prop) {
    with $!backing {
        my $prop-ns = self!get-ns-name(%known-props{$prop}.ns);
        my $cp-ns = self!get-ns-name("cp");
        $!backing.documentElement.findnodes("//{$cp-ns}:coreProperties/{$prop-ns}:$prop").head // Nil
    }
    else {
        Nil
    }
}

method !load-prop($prop) {
    return Nil without $!backing;
    self!maybe-prop-from-backing($prop) andthen %known-props{$prop}.from-xml($_)
}

method EXISTS-KEY(Str:D $property) {
    self!is-prop-known($property);
    %!properties{$property}:exists
}

method AT-KEY(Str:D $prop) is raw {
    self!is-prop-known($prop);
    ( %!properties{$prop}:exists
        ?? %!properties{$prop}
        !! (%!properties{$prop} := self!load-prop($prop)))
}

method ASSIGN-KEY(Str:D $property, \value) is raw {
    self!is-prop-known($property);
    my $prop-meta = %known-props{$property};
    my \prop-type = $prop-meta.type;
    die X::TypeCheck.new(operation => "assigning to property '$property'", :got(value.WHAT), :expected(prop-type))
        unless value ~~ prop-type;
    %!properties{$property} := prop-type.^archetypes.coercive ?? prop-type.^coerce(value) !! value;
}

method DELETE-KEY(Str:D $property) is raw {
    self!is-prop-known($property);
    %!properties{$property}:delete
}

#| Return a list of core properties set for the document.
method keys {
    $!backing
        andthen .documentElement.childNodes.map(*.localname).list
        orelse ()
}

method list {
    die X::NYI.new: feature => "List of core properties"
}

method iterator {
    die X::NYI.new: feature => "Core properties iterator"
}

method !FROM-BACKING {
    my Str:D $cp-ns = self!get-ns-name('cp');

    if $!backing.documentElement.nodeName ne "{$cp-ns}:coreProperties" {
        die X::Spreadsheet::XLSX::Format.new: message => 'Core properties file did not start tag coreProperties'
    }

    self
}

method from-xml(Str:D $xml) {
    my LibXML::Document:D $backing .= parse($xml);

    self.bless(:$backing)!FROM-BACKING;
}

method to-xml(--> Blob) {
    my LibXML::Element:D $cpProps = do with $!backing {
        $!backing.documentElement
    }
    else {
        $!backing .= new: :version<1.0>, :enc<UTF-8>;
        $!backing.setStandalone(LibXML::Document::XmlStandaloneYes);

        my LibXML::Element:D $root = $!backing.createElement('coreProperties');
        $!backing.setDocumentElement($root);

        for %cp-ns.kv -> $prefix, $uri {
            $root.addNamespace($uri, $prefix);
        }

        $root.setNamespace(%cp-ns<cp>, 'cp');

        $root
    }

    for %!properties.kv -> $property, $value {
        my CPMeta $prop-meta = %known-props{$property};

        my LibXML::Element $prop-elem = self!maybe-prop-from-backing($property);

        with $prop-elem {
            .removeChildNodes;
        }
        else {
            # Setup a new element
            $cpProps.add: $prop-elem = $!backing.createElement($property);
            my $ns-prefix = $prop-meta.ns;
            $prop-elem.setNamespace(%cp-ns{$ns-prefix}, $ns-prefix);

            # If there are any mandatory attributes then set them
            for $prop-meta.xml-attrs.kv -> $attr-name, $attr-value {
                if $attr-name.contains(":") {
                    my ($ans, $aname) = $attr-name.split(":");
                    $prop-elem.setAttributeNS(%cp-ns{$ans}, $aname, $attr-value);
                }
                else {
                    $prop-elem.setAttribute($attr-name, $attr-value);
                }
            }
        }

        if $prop-meta.complex {
            $value.to-xml($prop-elem)
        }
        else {
            $prop-elem.add: $!backing.createTextNode(~$value);
        }
    }

    $!backing.Blob
}