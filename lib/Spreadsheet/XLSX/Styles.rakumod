use LibXML::Document;
use LibXML::Element;
use LibXML::Attr;
use Spreadsheet::XLSX::Exceptions;
use Spreadsheet::XLSX::XMLHelpers;

#| Models the styles in a workbook. Styles are stored with a high degree of
#| reuse of information, which is good for storage, but not a great API for
#| the module user. This object models the style storage model, while a
#| flat view is provided by Spreadsheet::XLSX::CellStyle.
class Spreadsheet::XLSX::Styles does XMLRepresentation["styleSheet"] {

    #| A color definition
    class Color does XMLRepresentation["color"] {
        has Bool $.auto is xml-attr;
        has UInt $.indexed is xml-attr;
        has Str $.rgb is xml-attr;
        has UInt $.theme is xml-attr;
        has Real $.tint is xml-attr;
    }

    #| Font families (common stroke and serif characteristics).
    enum FontFamily is export (:NotApplicable(0), :Roman(1), :Swiss(2), :Modern(3),
                               :Script(4), :Decorative(5));

    #| Vertical alignments.
    enum VerticalAlignRun is export (:Baseline<baseline>, :Subscript<subscript>, :Superscript<superscript>);

    #| Font scheme.
    enum FontScheme is export (:SchemeNone<none>, :SchemeMajor<major>, :SchemeMinor<minor>);

    #| Underlines
    enum UnderlineValues is export (:None<none>, :SingleLine<single>, :DoubleLine<double>,
                                    :SingleAccounting<singleAccounting>,
                                    :DoubleAccounting<doubleAccounting>);

    #| A font in the style store.
    class Font does XMLRepresentation["font"] {
        has Str $.name is xml-elem;
        has Int $.charset is xml-elem;
        has FontFamily $.family is xml-elem;
        has Bool $.bold is xml-elem<b>;
        has Bool $.italic is xml-elem<i>;
        has Bool $.strike is xml-elem;
        has Bool $.outline is xml-elem;
        has Bool $.shadow is xml-elem;
        has Bool $.condense is xml-elem;
        has Bool $.extend is xml-elem;
        has Color $.color is xml-elem;
        has Real $.size is xml-elem<sz>;
        has UnderlineValues $.underline is xml-elem<u>;
        has VerticalAlignRun $.vertical-align is xml-elem<vertAlign>;
        has FontScheme $.scheme is xml-elem;
    }

    class Fonts is xml-sequence(<fonts>, Font) does XMLCounting { }

    #| A fill in the style store.
    role Fill {
        #| Produce an XML element containing the fill information.
        method fill-xml-element(LibXML::Document:D $doc, LibXML::Element:D $fill-elem --> LibXML::Element:D) {
            given $doc.createElement('fill') {
                .add: $fill-elem;
                $_
            }
        }

    }

    #| The available types of patterns for a pattern fill.
    enum FillType is export (:NoFill<none>, :Solid<solid>, :MediumGray<mediumGray>,
                             :DarkGray<darkGray>, :LightGray<lightGray>,
                             :DarkHorizontal<darkHorizontal>, :DarkVertical<darkVertical>,
                             :DarkDown<darkDown>, :DarkUp<darkUp>, :DarkGrid<darkGrid>,
                             :DarkTrellis<darkTrellis>, :LightHorizontal<lightHorizontal>,
                             :LightVertical<lightVertical>, :LightDown<lightDown>,
                             :LightUp<lightUp>, :LightGrid<lightGrid>, :LightTrellis<lightTrellis>,
                             :Gray125<gray125>, :Gray0625<gray0625>);

    enum GradientType is export (:Linear<linear>, :Path<path>);

    #| A pattern fill in the style store.
    class PatternFill does XMLRepresentation["patternFill"] does Fill {
        has FillType $.fill-type is xml-attr<patternType>;
        has Color $.foreground-color is xml-elem<fgColor>;
        has Color $.background-color is xml-elem<bgColor>;

        method to-xml-element(::?CLASS:D: LibXML::Document:D $doc, *%c --> LibXML::Element:D) {
             self.fill-xml-element: $doc, self.XMLRepresentation::to-xml-element($doc, |%c)
        }
    }

    class GradientStop does XMLRepresentation[<stop>] {
        has Color:D $.color   is required is xml-elem;
        has Real:D $.position is required is xml-attr;
    }

    #| A gradient fill in the style store.
    class GradientFill is xml-sequence(<gradientFill>, :stop(GradientStop)) does Fill {
        has GradientType $.type is xml-attr;
        has Real $.degree is xml-attr;
        has Real $.left is xml-attr;
        has Real $.right is xml-attr;
        has Real $.top is xml-attr;
        has Real $.bottom is xml-attr;

        method to-xml-element(LibXML::Document:D $doc, *%c --> LibXML::Element:D) {
            self.fill-xml-element: $doc, self.XMLSequence::to-xml-element($doc, |%c)
        }
    }

    class Fills is xml-sequence(<fills>, :fill(Fill)) does XMLCounting {
        method resolve-xml-element(LibXML::Element:D $fill --> XMLRepresentation) {
            unless $fill.childNodes.elems == 1 {
                die X::Spreadsheet::XLSX::Format.new:
                    message => "Element <fill> must have exactly one child but " ~ $fill.childNodes.elems ~ " found"
            }
            my $fill-elem = $fill.childNodes.head;
            given $fill-elem.nodeName {
                when 'patternFill' {
                    PatternFill.from-xml-element($fill-elem)
                }
                when 'gradientFill' {
                    GradientFill.from-xml-element($fill-elem)
                }
                default {
                    die X::Spreadsheet::XLSX::Format.new: message => "Fill element child cannot be a <" ~ $_ ~ ">"
                }
            }
        }
    }

    #| Border style.
    enum BorderStyle is export (:NoBorder<none>, :Thin<thin>, :Medium<medium>, :Dashed<dashed>,
                                :Dotted<dotted>, :Thick<thick>, :Double<double>, :Hair<hair>,
                                :MediumDashed<mediumDashed>, :DashDot<dashDot>,
                                :MediumDashDot<mediumDashDot>, :DashDotDot<dashDotDot>,
                                :MediumDashDotDot<mediumDashDotDot>, :SlantDashDot<slantDashDot>);

    class BorderPr does XMLRepresentation {
        has BorderStyle $.style is xml-attr = NoBorder;
        has Color $.color is xml-elem;
    }

    #| A border in the style store.
    class Border does XMLRepresentation[<border>] {
        has BorderPr $.left         is xml-elem;
        has BorderPr $.right        is xml-elem;
        has BorderPr $.top          is xml-elem;
        has BorderPr $.bottom       is xml-elem;
        has BorderPr $.diagonal     is xml-elem;
        has BorderPr $.vertical     is xml-elem;
        has BorderPr $.horizontal   is xml-elem;
        has Bool $.diagonal-up      is xml-attr;
        has Bool $.diagonal-down    is xml-attr;
        has Bool:D $.outline        is xml-attr = True;
    }

    class Borders is xml-sequence(<borders>, Border) does XMLCounting { }

    #| Horizontal alignment for cells.
    enum HorizontalAlign is export (:GeneralAlign<general>, :LeftAlign<left>, :CenterAlign<center>,
                                    :RightAlign<right>, :FillAlign<fill>, :JustifyAlign<justify>,
                                    :CenterContinuousAlign<centerContinuous>, :DistributedAlign<distributed>);

    #| Vertical alignment for cells.
    enum VerticalAlign is export (:TopVerticalAlign<top>, :CenterVerticalAlign<center>, :BottomVerticalAlign<bottom>,
                                  :JustifyVerticalAlign<justify>, :DistributedVerticalAlign<distributed>,);

    #| A number format.
    class NumberFormat does XMLRepresentation[<numFmt>] {
        has UInt:D $.id  is required is xml-attr<numFmtId>;
        has Str:D $.code is required is xml-attr<formatCode>;
    }

    class NumberFormats is xml-sequence(<numFmts>, NumberFormat) does XMLCounting {}

    #| Cell alignment formatting. Part of a Format.
    class CellAlignment does XMLRepresentation["alignment"] {
        has HorizontalAlign $.horizontal is xml-attr;
        has VerticalAlign $.vertical is xml-attr;
        has Int $.text-rotation is xml-attr;
        has Bool $.wrap-text is xml-attr;
        has Int $.indent is xml-attr;
        has Int $.relative-indent is xml-attr;
        has Bool $.justify-last-line is xml-attr;
        has Bool $.shrink-to-fit is xml-attr;
        has Int $.reading-order is xml-attr;
    }

    #| A format record, bringing together various styling options.
    class Format does XMLRepresentation["xf"] {
        has CellAlignment $.alignment is xml-elem;
        has Int $.number-format-id is xml-attr<numFmtId>;
        has Int $.font-id is xml-attr;
        has Int $.fill-id is xml-attr;
        has Int $.border-id is xml-attr;
        has UInt $.x-format-id is xml-attr<xfId>;
        has Bool $.quote-prefix is xml-attr;
        has Bool $.pivot-button is xml-attr;
        has Bool $.apply-number-format is xml-attr;
        has Bool $.apply-font is xml-attr;
        has Bool $.apply-fill is xml-attr;
        has Bool $.apply-border is xml-attr;
        has Bool $.apply-alignment is xml-attr;
        has Bool $.apply-protection is xml-attr;
    }

    class Formats is xml-sequence(<cellXfs>, :xf(Format)) does XMLCounting {}

    #| A record representing the name and related formatting records for a named cell style in this workbook.
    class CellStyle does XMLRepresentation["cellStyle"] {
        has Str $.name                              is xml-attr;
        has UInt:D $.number-format-id is required   is xml-attr<xfId>;
        has UInt $.builtin-id                       is xml-attr;
        # ??? No idea what "i" stands for. "item" or "inside"?
        has UInt $.ilevel                           is xml-attr<iLevel>;
        has Bool $.hidden                           is xml-attr;
        has Bool $.custom-builtin                   is xml-attr;
    }

    class CellStyles is xml-sequence(<cellStyles>, CellStyle) does XMLCounting { }

    #| All font records in the styles.
    has Fonts:D $.fonts is xml-elem .= new(Font.new);

    #| All fill records in the styles.
    has Fills:D $.fills is xml-elem .= new(PatternFill.new);

    #| All border records in the styles.
    has Borders:D $.borders is xml-elem .= new(Border.new);

    #| All number format records.
    has NumberFormats:D $.number-formats is xml-elem .= new;

    #| Every number format has an ID, but some IDs are allocated with
    #| existing meanings. Thus, we track the maximum number of those.
    has Int $!max-number-format-id = $!number-formats.map(*.id).max max 166;

    #| All formatting records (referenced from cell formats).
    has Formats:D $.formatting-records is xml-elem<cellStyleXfs> .= new(Format.new);

    #| All cell formats.
    has Formats:D $.cell-formats is xml-elem .= new(Format.new);

    #| All cell styles
    has CellStyles:D $.cell-styles is xml-elem .= new(CellStyle.new(:number-format-id(0)));

    #| Obtain a style ID for the specified combination of stylings.
    method obtain-style-id-for(Font :$font, CellAlignment :$alignment,
                               Str :$number-format --> Int) {
        # Here we really should be clever and re-use existing IDs.
        # However, for now, we just add everything we're given afresh.
        my %ids;
        with $font {
            %ids<font-id> = $!fonts.elems;
            %ids<apply-font> = True;
            $!fonts.push($font);
        }
        with $alignment {
            %ids<alignment> = $_;
            %ids<apply-alignment> = True;
        }
        with $number-format {
            with $!number-formats.first(*.code eq $number-format) {
                %ids<number-format-id> = .id;
            }
            else {
                my $id = ++$!max-number-format-id;
                $!number-formats.push(NumberFormat.new(:$id, :code($number-format)));
                %ids<number-format-id> = $id;
            }
            %ids<apply-number-format> = True;
        }
        my $style-id = $!cell-formats.elems;
        $!cell-formats.push(Format.new(|%ids));
        return $style-id;
    }

    method from-xml(Str $xml --> Spreadsheet::XLSX::Styles) {
        my LibXML::Document:D $doc .= parse(:string($xml));
        my LibXML::Element:D $root = $doc.documentElement;

        if $root.localName ne 'styleSheet' {
            die X::Spreadsheet::XLSX::Format.new: message => "Styles file did not start with tag 'styleSheet'";
        }

        self.from-xml-element($root)
    }

    #| Persist the styles information to XML.
    method to-xml(--> Blob) {
        # Create root element.
        my LibXML::Document $doc .= new: :version('1.0'), :enc('UTF-8');
        $doc.setStandalone(LibXML::Document::XmlStandaloneNo);
        my LibXML::Element $root = $doc.createElementNS(
            'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
            'styleSheet');
        $doc.setDocumentElement($root);
        self.to-xml-element($doc, $root);

        return $doc.Blob;
    }
}
