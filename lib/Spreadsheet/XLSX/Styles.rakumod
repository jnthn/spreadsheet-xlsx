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
    use Spreadsheet::XLSX::Types;

    class Fonts is xml-sequence(<fonts>, Font) does XMLCounting { }

    #| The available types of patterns for a pattern fill.
    enum FillType is export (:NoFill<none>, :Solid<solid>, :MediumGray<mediumGray>,
                             :DarkGray<darkGray>, :LightGray<lightGray>,
                             :DarkHorizontal<darkHorizontal>, :DarkVertical<darkVertical>,
                             :DarkDown<darkDown>, :DarkUp<darkUp>, :DarkGrid<darkGrid>,
                             :DarkTrellis<darkTrellis>, :LightHorizontal<lightHorizontal>,
                             :LightVertical<lightVertical>, :LightDown<lightDown>,
                             :LightUp<lightUp>, :LightGrid<lightGrid>, :LightTrellis<lightTrellis>,
                             :Gray125<gray125>, :Gray0625<gray0625>);

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

    class NumberFormats is xml-sequence(<numFmts>, NumberFormat) does XMLCounting {

        my @std-number-formats = ( (0,  'General'),
                                   (1,  '0'),
                                   (2,  '0.00'),
                                   (3,  '#,##0'),
                                   (4,  '#,##0.00'),
                                   (9,  '0%'),
                                   (10, '0.00%'),
                                   (11, '0.00E+00'),
                                   (12, '# ?/?'),
                                   (13, '# ??/??'),
                                   (14, 'mm-dd-yy'),
                                   (15, 'd-mmm-yy'),
                                   (16, 'd-mmm'),
                                   (17, 'mmm-yy'),
                                   (18, 'h:mm AM/PM'),
                                   (19, 'h:mm:ss AM/PM'),
                                   (20, 'h:mm'),
                                   (21, 'h:mm:ss'),
                                   (22, 'm/d/yy h:mm'),
                                   (37, '#,##0 ;(#,##0)'),
                                   (38, '#,##0 ;[Red](#,##0)'),
                                   (39, '#,##0.00;(#,##0.00)'),
                                   (40, '#,##0.00;[Red](#,##0.00)'),
                                   (45, 'mm:ss'),
                                   (46, '[h]:mm:ss'),
                                   (47, 'mmss.0'),
                                   (48, '##0.0E+0'),
                                   (49, '@') ).map({ NumberFormat.new(id => .[0], code => .[1]) });

        my %std-map{UInt} = @std-number-formats.map({ .id => $_ });

        method by-id(::?CLASS:D: UInt:D $id) {
            @std-number-formats.first(*.id == $id)
                || self.first(*.id == $id)
                || Nil
        }

        method by-code(::?CLASS:D: Str:D $code) {
            @std-number-formats.first(*.code eq $code)
                || self.first(*.code eq $code)
                || Nil
        }

        method next-id(::?CLASS:D:) {
            # A custom number format ID mut not be less than 164. At least this is the number we find in Excel-produced
            # stylesheets.
            self.elems
                ?? self.max(*.id).id + 1
                !! 164
        }

        method is-standard(NumberFormat:D $fmt) {
            ? (%std-map{$fmt.id} andthen .code eq $fmt.code)
        }

        method add(::?CLASS:D: NumberFormat:D $fmt) {
            given $fmt.clone(id => self.next-id) {
                self.push: $_;
                $_
            }
        }
    }

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

        method use-alignment     { ? ($!alignment.defined        && $!apply-alignment    ) }
        method use-number-format { ? ($!number-format-id.defined && $!apply-number-format) }
        method use-font          { ? ($!font-id.defined          && $!apply-font         ) }
        method use-fill          { ? ($!fill-id.defined          && $!apply-fill         ) }
        method use-border        { ? ($!border-id.defined        && $!apply-border       ) }
    }

    class Formats is xml-sequence(<cellXfs>, :xf(Format)) does XMLCounting {}

    #| A record representing the name and related formatting records for a named cell style in this workbook.
    # Use XML name for the class to make clear distinction from Spreadsheet::XLSX::CellStyle
    class CT_CellStyle does XMLRepresentation["cellStyle"] {
        has Str $.name                              is xml-attr;
        has UInt:D $.format-id          is required is xml-attr<xfId>;
        has UInt $.builtin-id                       is xml-attr;
        # ??? No idea what "i" stands for. "item" or "inside"?
        has UInt $.ilevel                           is xml-attr<iLevel>;
        has Bool $.hidden                           is xml-attr;
        has Bool $.custom-builtin                   is xml-attr;
    }

    class CellStyles is xml-sequence(<cellStyles>, CT_CellStyle) does XMLCounting { }

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
    has CellStyles:D $.cell-styles is xml-elem .= new(CT_CellStyle.new(:format-id(0)));

    method !xml-default-doc {
        # Create root element.
        my LibXML::Document $doc .= new: :version('1.0'), :enc('UTF-8');
        $doc.setStandalone(LibXML::Document::XmlStandaloneNo);
        my LibXML::Element $root = $doc.createElementNS(
            'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
            'styleSheet');
        $doc.setDocumentElement($root);
        $doc
    }

    method !obtain-id-for(XMLRepresentation:D $what, XMLSequence:D $where, :$autogen) {
        return Nil without $what;
        my $doc = self!xml-default-doc;
        my $what-xml = $what.to-xml-element($doc);
        for ^$where.elems -> $id {
            return $id if $what-xml eq $where[$id].to-xml-element($doc);
        }
        $autogen ?? $where.push($what).end !! Nil
    }

    proto method id-for(|) {*}
    multi method id-for(Font:D $font) {
        self!obtain-id-for($font, $!fonts, |%_)
    }
    multi method id-for(Fill:D $fill) {
        self!obtain-id-for($fill, $!fills, |%_)
    }
    multi method id-for(Border:D $border) {
        self!obtain-id-for($border, $!borders, |%_)
    }
    multi method id-for(NumberFormat:D $num-fmt, :$autogen) {
        if self!obtain-id-for($num-fmt, $!number-formats, |%_)
            || $!number-formats.is-standard($num-fmt)
        {
            return $num-fmt.id;
        }

        return Nil unless $autogen;

        # First try to select from existing number formats and choose one with the same format code
        with $!number-formats.by-code($num-fmt.code) {
            return .id
        }

        # If no format code match found then add a new format code. $!number-formats will set new format's ID.
        $!number-formats.add($num-fmt).id
    }
    # Here we'd need to be explicit as we have $!formatting-records and we have $!cell-formats
    multi method id-for(Format:D $format, Formats:D $formats) {
        self!obtain-id-for($format, $formats, |%_)
    }
    multi method id-for(CT_CellStyle:D $cell-style) {
        self!obtain-id-for($cell-style, $!cell-styles, |%_)
    }

    method from-xml(Str $xml --> Spreadsheet::XLSX::Styles) {
        my LibXML::Document:D $doc .= parse(:string($xml));
        my LibXML::Element:D $style-sheet = $doc.documentElement;

        if $style-sheet.localName ne 'styleSheet' {
            die X::Spreadsheet::XLSX::Format.new: message => "Styles file did not start with tag 'styleSheet'";
        }

        self.from-xml-element($style-sheet)
    }

    #| Persist the styles information to XML.
    method to-xml(--> Blob) {
        my LibXML::Document:D $doc = self!xml-default-doc;
        self.to-xml-element($doc, $doc.documentElement);
        return $doc.Blob;
    }
}
