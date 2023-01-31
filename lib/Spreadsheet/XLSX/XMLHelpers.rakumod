use v6.d;
unit module Spreadsheet::XLSX::XMLHelpers;
use experimental :will-complain;
use Spreadsheet::XLSX::Exceptions;
use LibXML::Attr;
use LibXML::Document;
use LibXML::Element;
use LibXML::Namespace;
use LibXML::Item :ast-to-xml;

role XMLSequence {...}

my sub nominalize-type(Mu \type) is raw is pure { type.^archetypes.nominalizable ?? type.^nominalize !! type }

# Mark type attribute objects representing a named XML element. Not to be directly consumed, consider XMLAttribute or
# XMLElement. See `xml-attr` or `xml-elem` traits.
my role XMLNamed {
    has Str $.xml-name;

    method xml-set-name(Str:D $!xml-name) {}
}

# For attributes mapping into XML element attributes
my role XMLAttribute does XMLNamed { }

# For attributes mapping into XML elements
my role XMLElement does XMLNamed { }

# XMLRepresentation implements object de- and serialization to/from XML. This is done by mapping class attributes
# into XML attributes or elements. A type object consuming this role would be mapping into an XML element. If optional
# $element-tag is provided then it would be used as the element tag unless altered by other means (see traits below).
#
# For the former XML attribute value is coerced into attribute's type. The only exception is for booleans where instead
# of coercion direct comparison to '1' or 'true' is used.
#
# For XML elements there are two methods:
#
# The first is for Raku attribute types which are not XMLRepresentation on their own. In this case element's `val`
# attribute is used if present, or element existence is considered as True for booleans.
#
# The second is for XMLRepresentation types in which case the XML element is de-serialized using the standard
# `from-xml-element` method of the type.
#
# Apparently, reverse approach is used for serializing into XML where either an attribute definite value is stringified
# into corresponding XML attribute; or an XML element created either using XMLRepresentation `to-xml-element` method,
# if available, or by producing a <property val="..." /> element.
#
# Special care is taken of unsupported XML elements or attributes when a source is de-serialized. Some of them are not
# supported yet, or are not part of the Open XML File specification. These are converted into LibXML AST (`$node.ast`)
# and pushed into object's `@!xml-unsupported` attribute. These are then de-serialized back into XML thus preserving
# (presumably) compatibility with the original file format. Namespaces are handled individually as we don't have direct
# use or mapping for them. Thus all namespaces from an element eventually end up in `@!xml-unsupported`.
#
# The role also provides specific support for XML elements declared in schema as xsd:sequence. If a class consumes
# the XMLSequence role (not every sequential element is actually a sequence, consider optional child elements) then
# serialization/deserialization also invokes corresponding private methods from XMLSequence. Also, elements not having
# direct mapping into attributes but known to be items of the sequence are not pushed into @!xml-unsupported in hope
# they'd be picked up later.
role XMLRepresentation[Str $element-tag?] is export {
    has @!xml-unsupported is built;

    method default-xml-tag is pure {
        $element-tag
    }

    method !profile-from-xml-element(LibXML::Element:D $elem, Bool :$local = True) {
        gather {
            # ASTs for attributes or elements we cannot map into class attributes
            my @unsupported;

            my sub coerce-into(::T, T() \coerced) is raw {
                coerced
            }

            my sub take-xml-value(Mu \type, Str:D $xml-value) {
                type ~~ Bool
                    ?? ?($xml-value eq "1" | "true")
                    !! coerce-into(type, $xml-value)
            }

            my proto sub by-attr-type(Any:U, LibXML::Element:D) {*}
            multi sub by-attr-type(::?ROLE \xml-elem, $property) {
                # A XMLRepresentation has 'from-xml-element' method
                xml-elem.from-xml-element($property)
            }
            multi sub by-attr-type(\xml-plain, $property) {
                $property.hasAttribute('val')
                    ?? take-xml-value(xml-plain, $property.getAttribute('val'))
                    !! xml-plain ~~ Bool
                ?? True # A boolean element with no attributes: <boolProp/>
                !! die X::Spreadsheet::XLSX::Format.new:
                    message => "Simple property element {$property.name} is missing required attribute 'val'"
            }


            # We don't consider namespaces as data but better preserve them
            for $elem.namespaces -> LibXML::Namespace:D $ns {
                @unsupported.push: $ns.ast;
            }

            for $elem.properties -> LibXML::Attr:D $xml-attr {
                with self.^xml-attr-for-attr($xml-attr.name) -> $attr {
                    take $attr.name.substr(2) => take-xml-value($attr.type, $xml-attr.value);
                }
                else {
                    @unsupported.push: $xml-attr.ast;
                }
            }

            my $is-sequential = self.^xml-is-sequential;
            my $seq-tags := self.^xml-sequence-tags.keys.Set;

            for $elem.childNodes -> LibXML::Element:D $xml-elem {
                my $elem-name = $xml-elem.nodeName;
                with self.^xml-attr-for-tag($elem-name) -> $attr {
                    take ($attr.name.substr(2) => by-attr-type(nominalize-type($attr.type), $xml-elem));
                }
                elsif !($is-sequential && $elem-name ∈ $seq-tags) {
                    @unsupported.push: $xml-elem.ast;
                }
            }

            take "xml-unsupported" => @unsupported if @unsupported;
        }
    }

    method !set-element-from-self(::?CLASS:D: LibXML::Document:D $doc, LibXML::Element:D $elem) {
        for @!xml-unsupported -> $ast {
            given ast-to-xml($ast) {
                when LibXML::Attr {
                    if .name eq 'xmlns' {
                        # Workaround for a problem where LibXML deserializes :xmlns(...) as an attribute
                        $elem.setNamespace(.value);
                    }
                    else {
                        $elem.setAttributeNode($_);
                    }
                }
                when LibXML::Namespace {
                    $elem.setNamespace( .declaredURI, .declaredPrefix );
                }
                default {
                    $elem.add: $_
                }
            }
        }

        for self.^attributes(:local).grep(XMLAttribute) -> Attribute:D $attr {
            my $xml-attr-name = $attr.xml-name;

            with $attr.get_value(self) -> $attr-value {
                my Str $xml-value;
                if $attr-value.HOW ~~ Metamodel::EnumHOW {
                    $xml-value = ~$attr-value.value;
                }
                elsif $attr-value ~~ Bool {
                    $xml-value = ~$attr-value.Int;
                }
                else {
                    $xml-value = ~$attr-value;
                }
                $elem.add: $doc.createAttribute($xml-attr-name, $xml-value);
            }
        }

        for self.^attributes(:local).grep(XMLElement) -> Attribute:D $attr {
            with $attr.get_value(self) -> $attr-value {
                if $attr-value.^can('to-xml-element') {
                    $elem.add: $attr-value.to-xml-element($doc, :tag($attr.xml-name));
                }
                else {
                    # Simple element with 'val' attribute
                    my $attr-elem = $doc.createElement($attr.xml-name);
                    if $attr-value.HOW ~~ Metamodel::EnumHOW {
                        $attr-elem.setAttribute('val', ~$attr-value.value);
                    }
                    elsif $attr-value ~~ Bool {
                        $attr-elem.setAttribute('val', "0") unless $attr-value;
                    }
                    else {
                        $attr-elem.setAttribute('val', ~$attr-value);
                    }
                    $elem.add: $attr-elem;
                }
            }
        }
        $elem
    }

    #| Produce an instance from XML element
    method from-xml-element(LibXML::Element:D $elem --> ::?CLASS:D) {
        # The method can be accidentally invoked on a nominalizable like a definite. Disrespect and proceed as normal.
        my $profile := self!profile-from-xml-element($elem).Capture;
        given self.new(|$profile) {
            $_!pull-from-sequence($elem) if .^xml-is-sequential;
            $_
        }
    }

    #| Serialize into a LibXML::Element
    method to-xml-element(::?CLASS:D: LibXML::Document:D $doc, LibXML::Element $elem?, Str :$tag ) {
        given $elem // $doc.createElement($tag // $element-tag) {
            self!set-element-from-self($doc, $_);
            self!set-element-from-sequence($doc, $_) if self.^xml-is-sequential;
            $_
        }
    }
}

# XMLSequence implements handling of sequential XML elements. It expects the class it is applied to to implement
# the standard `push` method.
#
# Special care is taken of union-like attributes (akin of <fill> of SpreadsheetML, apparently). Such elements
# must to "map" into a role (`xml-sequence` trait). If such mapping is encountered then the XMLSequence class is
# expected to implement `resolve-xml-element` method which takes an `LibXML::Element` and returns a deserialized
# representation of it. For example, in `<fill>` case either PatternFill or GradientFill instances would be produced.
# Note that objects produced by such a resolution are, in turn, expected to produce similar XML wrapping; i.e.
# `<fill><patternFill .../></fill>`, for example.
#
# Note that the role is expected to be applied to Array descendants by the `xml-sequence` trait. Since Array's `new`
# doesn't respect BUILDPLAN and we are likely to have attributes on our sequential class, the role implements own
# `new` method which explicitly calls `BUILDALL`. If this behavior is undesirable for classes directly consuming this
# role via `does`, then they must implement their own `new`.
#
# Also, it is typical for sequential XML elements to have attribute `count`. Role `XMLCounting` provides support for it.
# `XMLCounting` provides validation method `validate-count` which we try invoke after completing deserialization of a
# sequence.
role XMLSequence is export
{
    method !pull-from-sequence(::?CLASS:D: LibXML::Element:D $seq) {
        my %tags = self.^xml-sequence-tags;
        my Junction:D $any-of-seq = %tags.keys.any;
        for $seq.children.grep({ .nodeName eq $any-of-seq }) -> LibXML::Element:D $seq-elem {
            self.push: do given %tags{$seq-elem.nodeName} {
                # If current element maps into a role let the class itself resolve it into something concrete
                .^archetypes.parametric
                    ?? self.resolve-xml-element($seq-elem)
                    !! .from-xml-element($seq-elem)
            }
        }
        self.?validate-count($seq);
        self
    }

    method !set-element-from-sequence(::?CLASS:D: LibXML::Document:D $doc, LibXML::Element:D $elem) {
        my %tags = self.^xml-sequence-tags;

        my %types{Mu} = %tags.antipairs;

        for self -> $item {
            my %profile;
            with %types{$item.WHAT} {
                %profile<tag> = $_;
            }
            else {
                # If there is no mapping for this item type then it was likely produced by resolving a role (see Fill
                # for class Fills). Use either backward resolution or let the item itself decide what to do.
                %profile<tag> = $_ with self.?resolve-item-tag($item);
            }
            $elem.add($item.to-xml-element($doc, |%profile));
        }

        self
    }

    # Array constructor doesn't respect BUILDPLAN but we may need it.
    method new(*%named, |) {
        callsame().BUILDALL(Empty, %named)
    }
}

# XMLElementHOW provides functionality for working with attribute into XML name mapping and support for sequential
# elements.
my role XMLElementHOW {
    # Map XML tags into attributes
    has Map $!xml-tags;
    # Map XML element attributes into class attributes
    has Map $!xml-attrs;
    # List of tags allowed for a sequence
    has Map $!sequence-tags;

    method xml-attr-for-tag(Mu \obj, Str:D $tag) {
        without $!xml-tags {
            $!xml-tags :=
                Map.new: self.attributes(obj, :local).grep(XMLElement).map({ .xml-name => $_ });
        }
        $!xml-tags{$tag} // Nil
    }

    method xml-attr-for-attr(Mu \obj, Str:D $xml-attr) {
        without $!xml-attrs {
            $!xml-attrs :=
                Map.new: self.attributes(obj, :local).grep(XMLAttribute).map({ .xml-name => $_ });
        }
        $!xml-attrs{$xml-attr}
    }

    method xml-set-sequence-tags(Mu, \tags) { $!sequence-tags := tags.Map }

    method xml-is-sequential(Mu) { ? $!sequence-tags }

    method xml-sequence-tags(Mu) is raw { $!sequence-tags }
}

# Produce standard XML name for an attribute object. The rules used are following:
# If attribute type is an XMLRepresentation and its default tag is specified then use it;
# otherwise use attribute name either as-is, or if it's a kebab case then convert into lower camel case:
# has $.lowerCamel;           # <lowerCamel>
# has $.lower-camel;          # <lowerCamel>
# has TypeWithXMLRepr $.date; # "<" ~ TypeWithXMLRepr.default-xml-tag ~ ">"
my sub xml-name(Attribute:D $attr) {
    my \attr-type = nominalize-type($attr.type);
    (attr-type ~~ XMLRepresentation && attr-type.default-xml-tag)
        || ((my $aname = $attr.name.substr(2)).contains("-")
                ?? ($aname.split("-").cache andthen .head.lc ~ .tail(*-1).map(*.tclc).join)
                !! $aname)
}

my sub mark-attr-xml($attr, $name, Mu \xml-role) {
    my $pkg-how := $*PACKAGE.HOW;
    $pkg-how does XMLElementHOW unless $pkg-how ~~ XMLElementHOW;
    ($attr does xml-role).xml-set-name($name);
}

# `xml-attr` trait marks an attribute as mapping into an XML attribute. I.e.
# has $.my-attr is xml-attr; # <someElem myAttr="value">
# `is xml-attr<attrName>` sets explicit XML attribute name, overriding possible attribute's type provided (if the type
# is an XMLRepresentation), or implicit name produced by xml-name routine above.
multi sub trait_mod:<is>(Attribute:D $attr, Str:D :$xml-attr!) is export {
    unless $attr ~~ XMLAttribute {
        mark-attr-xml($attr, $xml-attr, XMLAttribute);
    }
}
multi sub trait_mod:<is>(Attribute:D $attr, Bool:D :$xml-attr!) is export {
    unless $attr ~~ XMLAttribute {
        mark-attr-xml($attr, xml-name($attr), XMLAttribute);
    }
}

# `xml-elem` trait marks an attribute as mapping into an XML element. Usage is similar to `xml-attr`.
multi sub trait_mod:<is>(Attribute:D $attr, Str:D :$xml-elem!) is export {
    unless $attr ~~ XMLElement {
        mark-attr-xml($attr, $xml-elem, XMLElement);
    }
}
multi sub trait_mod:<is>(Attribute:D $attr, Bool:D :xml-elem($)!) is export {
    unless $attr ~~ XMLElement {
        mark-attr-xml($attr, xml-name($attr), XMLElement);
    }
}

# `xml-sequence` trait marks a class as (de)serializing from/into an XML sequence element. Usage:
#
#      class Collection is xml-sequence("elementTag", ChildType1, :childTag(ChildType2)) { ... }
#
# Any child type must be either an XMLRepresentation class, or a role.
#
# The trait makes a class it is applied to into an Array[XMLChildTypes] descendant, where XMLChildTypes is a subset
# limiting elements of the array to the child types. Also, XMLSequence role is applied provding (de)serialization
# means and support for BUILDPLAN necessary to set class attributes via profile generated by XMLRepresentation, which is
# applied too.
#
# XMLElementHOW is mixed into class' meta-object.
#
# A class declared with this trait is normally used as:
#
#     has Collection:D $.collection is xml-elem .= new;
#
# Even though it is an Array. This is because `@`- and `$`-sigilled attributes are using different assignment semantics,
# not allowing for `@`-sigilled ones to preserve object attribute values.
multi sub trait_mod:<is>(Mu \type, :$xml-sequence! (Str:D $tag?, |child-types)) {

    die "Target type '{type.^name} of 'xml-sequence' trait must be a class"
        unless type.HOW ~~ Metamodel::ClassHOW;

    unless child-types.elems || child-types.keys {
        die "xml-sequence needs at least one XMLRepresentation type argument"
    }

    my @types;
    my %tags;

    my subset XMLReprType
        of Mu
        will complain{ "xml-sequence expects a role or a XMLRepresentation type object but got " ~ $_.gist }
        where { .^archetypes.parametric || (!.defined && $_ ~~ XMLRepresentation ) };

    for child-types.kv -> $child-tag, XMLReprType \child-type {
        @types.push: child-type;
        %tags{ $child-tag ~~ Int:D ?? child-type.default-xml-tag !! $child-tag } := child-type;
    }

    my @type-names = @types.map(*.^name);
    my $type-names-any = @type-names.join("|");
    my $any-of-type = @types.any;

    # Create a subset to validate sequence element types
    my \XMLChildTypes =
        Metamodel::SubsetHOW.new_type:
            :name("XMLSeqOf($type-names-any)"),
            :refinee(Any),
            :refinement({ $_ ~~ $any-of-type });

    # Make the subset produce meaningful error message on type check failure
    &trait_mod:<will>(:complain, XMLChildTypes,
                      { "expected any of " ~ @type-names.join(",") ~ " but got " ~ .^name ~ " ($_)" });

    # Finalize the target class
    type.^add_parent(Array[XMLChildTypes]);
    type.HOW does XMLElementHOW;
    type.^xml-set-sequence-tags(%tags);
    type.^add_role(XMLSequence);
    type.^add_role(XMLRepresentation[$tag]);
}

# XMLCounting role implements XML element 'count' attribute. Normally to be used with xml-sequence trait.
role XMLCounting is export {
    has UInt $.count is xml-attr;

    method validate-count(LibXML::Element:D $elem) {
        with $.count {
            (self.elems == $_) or die X::Spreadsheet::XLSX::Format.new:
                message => "Elemenent <{$elem.nodeName}> declared with count $.count but "
                    ~ self.elems ~ " child elements found"
        }
        self
    }
}
