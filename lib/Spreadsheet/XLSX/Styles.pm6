use LibXML::Document;

#| Models the styles in a workbook. Styles are stored with a high degree of
#| reuse of information, which is good for storage, but not a great API for
#| the module user. This object models the style storage model, while a
#| flat view is provided by Spreadsheet::XLSX::CellStyle.
class Spreadsheet::XLSX::Styles {
    #| Font families (common stroke and serif characteristics).
    enum FontFamily is export (
        :NotApplicable(0), :Roman(1), :Swiss(2), :Modern(3),
        :Script(4), :Decorative(5)
    );

    #| Vertical alignments.
    enum VerticalAlignRun is export (
        :Baseline<baseline>, :Subscript<subscript>, :Superscript<superscript>
    );

    #| Font scheme.
    enum FontScheme is export (
        :SchemeNone<none>, :SchemeMajor<major>, :SchemeMinor<minor>
    );

    #| A font in the style store.
    class Font {
        has Str $.name;
        has Int $.charset;
        has FontFamily $.family;
        has Bool $.bold;
        has Bool $.italic;
        has Bool $.strike;
        has Bool $.outline;
        has Bool $.shadow;
        has Bool $.condense;
        has Bool $.extend;
        # TODO color, which is annoyingly involved
        has Real $.size;
        # TODO underline, which needs another enum
        has VerticalAlignRun $.vertical-align;
        has FontScheme $.scheme;

        #| Produce an XML element containing the font information.
        method to-xml-element(LibXML::Document $doc --> LibXML::Element) {
            my $font = $doc.createElement('font');
            $font.add($doc.createElement('b')) if $!bold;
            $font.add($doc.createElement('i')) if $!italic;
            $font.add(val-element($doc, 'sz', $_)) with $!size;
            # TODO
            #	<element name="name" type="CT_FontName" minOccurs="0" maxOccurs="1"/>
            #	<element name="charset" type="CT_IntProperty" minOccurs="0" maxOccurs="1"/>
            #	<element name="family" type="CT_IntProperty" minOccurs="0" maxOccurs="1"/>
            #	<element name="strike" type="CT_BooleanProperty" minOccurs="0" maxOccurs="1"/>
            #	<element name="outline" type="CT_BooleanProperty" minOccurs="0" maxOccurs="1"/>
            #	<element name="shadow" type="CT_BooleanProperty" minOccurs="0" maxOccurs="1"/>
            #	<element name="condense" type="CT_BooleanProperty" minOccurs="0" maxOccurs="1"/>
            #	<element name="extend" type="CT_BooleanProperty" minOccurs="0" maxOccurs="1"/>
            #	<element name="color" type="CT_Color" minOccurs="0" maxOccurs="1"/>
            #	<element name="u" type="CT_UnderlineProperty" minOccurs="0" maxOccurs="1"/>
            #	<element name="vertAlign" type="CT_VerticalAlignFontProperty" minOccurs="0" maxOccurs="1"/>
            #	<element name="scheme" type="CT_FontScheme" minOccurs="0" maxOccurs="1"/>
            return $font;
        }

        sub val-element(LibXML::Document $doc, Str $name, Str() $value) {
            my $element = $doc.createElement($name);
            $element.add($doc.createAttribute('val', $value));
            return $element;
        }
    }

    #| A fill in the style store.
    role Fill {
    }

    #| The available types of patterns for a pattern fill.
    enum FillType is export (
        :NoFill<none>, :Solid<solid>, :MediumGray<mediumGray>,
        :DarkGray<darkGray>, :LightGray<lightGray>,
        :DarkHorizontal<darkHorizontal>, :DarkVertical<darkVertical>,
        :DarkDown<darkDown>, :DarkUp<darkUp>, :DarkGrid<darkGrid>,
        :DarkTrellis<darkTrellis>, :LightHorizontal<lightHorizontal>,
        :LightVertical<lightVertical>, :LightDown<lightDown>,
        :LightUp<lightUp>, :LightGrid<lightGrid>, :LightTrellis<lightTrellis>,
        :Gray125<gray125>, :Gray0625<gray0625>
    );

    #| A pattern fill in the style store.
    class PatternFill does Fill {
        has FillType $.fill-type;
        # TODO colors

        #| Produce an XML element containing the fill information.
        method to-xml-element(LibXML::Document $doc --> LibXML::Element) {
            my $fill = $doc.createElement('fill');
            my $pattern-fill = $doc.createElement('patternFill');
            $pattern-fill.add($doc.createAttribute('patternType',
                    $!fill-type.defined ?? $!fill-type.value !! 'none'));
            # TODO colors
            $fill.add($pattern-fill);
            return $fill;
        }
    }

    #| A gradient fill in the style store.
    class GradientFill does Fill {
        # TODO
    }

    #| Border style.
    enum BorderStyle is export (
        :NoBorder<none>, :Thin<thin>, :Medium<medium>, :Dashed<dashed>,
        :Dotted<dotted>, :Thick<thick>, :Double<double>, :Hair<hair>,
        :MediumDashed<mediumDashed>, :DashDot<dashDot>,
        :MediumDashDot<mediumDashDot>, :DashDotDot<dashDotDot>,
        :MediumDashDotDot<mediumDashDotDot>, :SlantDashDot<slantDashDot>
    );

    #| A border in the style store.
    class Border {
        has BorderStyle $.left-style;
        has BorderStyle $.right-style;
        has BorderStyle $.top-style;
        has BorderStyle $.bottom-style;
        has BorderStyle $.diagonal-style;
        has BorderStyle $.vertical-style;
        has BorderStyle $.horizontal-style;
        # TODO colors
        has Bool $.diagonal-up;
        has Bool $.diagonal-down;
        has Bool $.outline;

        #| Produce an XML element containing the format information.
        method to-xml-element(LibXML::Document $doc --> LibXML::Element) {
            my $border = $doc.createElement('border');
            # TODO
            #   <element name="left" type="CT_BorderPr" minOccurs="0" maxOccurs="1"/>
            #	<element name="right" type="CT_BorderPr" minOccurs="0" maxOccurs="1"/>
            #	<element name="top" type="CT_BorderPr" minOccurs="0" maxOccurs="1"/>
            #	<element name="bottom" type="CT_BorderPr" minOccurs="0" maxOccurs="1"/>
            #	<element name="diagonal" type="CT_BorderPr" minOccurs="0" maxOccurs="1"/>
            #	<element name="vertical" type="CT_BorderPr" minOccurs="0" maxOccurs="1"/>
            #	<element name="horizontal" type="CT_BorderPr" minOccurs="0" maxOccurs="1"/>
            #	</sequence>
            #	<attribute name="diagonalUp" type="xsd:boolean" use="optional"/>
            #	<attribute name="diagonalDown" type="xsd:boolean" use="optional"/>
            #	<attribute name="outline" type="xsd:boolean" use="optional" default="true"/>
            return $border;
        }
    }

    #| Horizontal alignment for cells.
    enum HorizontalAlign is export (
        :GeneralAlign<general>, :LeftAlign<left>, :CenterAlign<center>,
        :RightAlign<right>, :FillAlign<fill>, :JustifyAlign<justify>,
        :CenterContinuousAlign<centerContinuous>,  :DistributedAlign<distributed>
    );

    #| Vertical alignment for cells.
    enum VerticalAlign is export (
        :TopVerticalAlign<top>, :CenterVerticalAlign<center>, :BottomVerticalAlign<bottom>,
        :JustifyVerticalAlign<justify>, :DistributedVerticalAlign<distributed>,
    );

    #| A number format.
    class NumberFormat {
        has Int $.id;
        has Str $.code;

        #| Produce an XML element containing the number format information.
        method to-xml-element(LibXML::Document $doc --> LibXML::Element) {
            my $numFmt = $doc.createElement('numFmt');
            $numFmt.add($doc.createAttribute('numFmtId', ~$!id));
            $numFmt.add($doc.createAttribute('formatCode', ~$!code));
            return $numFmt;
        }
    }

    #| Cell alignment formatting. Part of a Format.
    class CellAlignment {
        has HorizontalAlign $.horizontal;
        has VerticalAlign $.vertical;
        has Int $.text-rotation;
        has Bool $.wrap-text;
        has Int $.indent;
        has Int $.relative-indent;
        has Bool $.justify-last-line;
        has Bool $.shrink-to-fit;
        has Int $.reading-order;

        #| Produce an XML element containing the alignment information.
        method to-xml-element(LibXML::Document $doc --> LibXML::Element) {
            my $alignment = $doc.createElement('alignment');
            $alignment.add($doc.createAttribute('horizontal', .value)) with $!horizontal;
            $alignment.add($doc.createAttribute('vertical', .value)) with $!vertical;
            $alignment.add($doc.createAttribute('wrapText', '1')) if $!wrap-text;
            #	<attribute name="textRotation" type="xsd:unsignedInt" use="optional"/>
            #	<attribute name="indent" type="xsd:unsignedInt" use="optional"/>
            #	<attribute name="relativeIndent" type="xsd:int" use="optional"/>
            #	<attribute name="justifyLastLine" type="xsd:boolean" use="optional"/>
            #	<attribute name="shrinkToFit" type="xsd:boolean" use="optional"/>
            #	<attribute name="readingOrder" type="xsd:unsignedInt" use="optional"/>
            return $alignment;
        }
    }

    #| A format record, bringing together various styling options.
    class Format {
        has CellAlignment $.alignment;
        has Int $.number-format-id;
        has Int $.font-id;
        has Int $.fill-id;
        has Int $.border-id;
        has Int $.x-format-id;
        has Bool $.quote-prefix;
        has Bool $.pivot-button;
        has Bool $.apply-number-format;
        has Bool $.apply-font;
        has Bool $.apply-fill;
        has Bool $.apply-border;
        has Bool $.apply-alignment;
        has Bool $.apply-protection;

        #| Produce an XML element containing the format information.
        method to-xml-element(LibXML::Document $doc --> LibXML::Element) {
            my $xf = $doc.createElement('xf');
            $xf.add($doc.createAttribute('fontId', ~$_)) with $!font-id;
            $xf.add($doc.createAttribute('fillId', ~$_)) with $!fill-id;
            $xf.add($doc.createAttribute('borderId', ~$_)) with $!border-id;
            $xf.add($doc.createAttribute('numFmtId', ~$_)) with $!number-format-id;
            $xf.add($doc.createAttribute('applyFont', '1')) if $!apply-font;
            $xf.add($doc.createAttribute('applyFill', '1')) if $!apply-fill;
            $xf.add($doc.createAttribute('applyBorder', '1')) if $!apply-border;
            $xf.add($doc.createAttribute('applyAlignment', '1')) if $!apply-alignment;
            $xf.add($doc.createAttribute('applyNumberFormat', '1')) if $!apply-number-format;
            $xf.add(.to-xml-element($doc)) with $!alignment;
            # TODO
            #	<attribute name="xfId" type="ST_CellStyleXfId" use="optional"/>
            #	<attribute name="quotePrefix" type="xsd:boolean" use="optional" default="false"/>
            #	<attribute name="pivotButton" type="xsd:boolean" use="optional" default="false"/>
            #	<attribute name="applyProtection" type="xsd:boolean" use="optional"/>
            return $xf;
        }
    }

    #| All font records in the styles.
    has Font @.fonts;

    #| All fill records in the styles.
    has Fill @.fills;

    #| All border records in the styles.
    has Border @.borders;

    #| All number format records.
    has NumberFormat @.number-formats;

    #| Every number format has an ID, but some IDs are allocated with
    #| existing meanings. Thus, we track the maximum number of those.
    has Int $!max-number-format-id = @!number-formats.map(*.id).max max 166;

    #| All formatting records (referenced from cell formats).
    has Format @.formatting-records;

    #| All cell formats.
    has Format @.cell-formats;

    #| Load the styles information from XML.
    method from-xml(--> Spreadsheet::XLSX::Styles) {
        # TODO implement this
        Spreadsheet::XLSX::Styles.bless
    }

    method new() {
        # We need to have default, empty, instances of these as the first element,
        # otherwise Excel will be highly displeased.
        Spreadsheet::XLSX::Styles.bless:
                fonts => [Font.new],
                fills => [PatternFill.new],
                borders => [Border.new],
                cell-formats => [Format.new],
                formatting-records => [Format.new]
    }

    #| Obtain a style ID for the specified combination of stylings.
    method obtain-style-id-for(Font :$font, CellAlignment :$alignment,
                               Str :$number-format --> Int) {
        # Here we really should be clever and re-use existing IDs.
        # However, for now, we just add everything we're given afresh.
        my %ids;
        with $font {
            %ids<font-id> = @!fonts.elems;
            %ids<apply-font> = True;
            @!fonts.push($font);
        }
        with $alignment {
            %ids<alignment> = $_;
            %ids<apply-alignment> = True;
        }
        with $number-format {
            with @!number-formats.first(*.code eq $number-format) {
                %ids<number-format-id> = .id;
            }
            else {
                my $id = ++$!max-number-format-id;
                @!number-formats.push(NumberFormat.new(:$id, :code($number-format)));
                %ids<number-format-id> = $id;
            }
            %ids<apply-number-format> = True;
        }
        my $style-id = @!cell-formats.elems;
        @!cell-formats.push(Format.new(|%ids));
        return $style-id;
    }

    #| Persist the styles information to XML.
    method to-xml(--> Str) {
        # Create root element.
        my LibXML::Document $doc .= new: :version('1.0'), :enc('UTF-8');
        $doc.setStandalone(LibXML::Document::XmlStandaloneNo);
        my LibXML::Element $root = $doc.createElementNS(
                'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
                'styleSheet');
        $doc.setDocumentElement($root);

        # Store all the parts.
        self!add-part($doc, $root, 'numFmts', @!number-formats);
        self!add-part($doc, $root, 'fonts', @!fonts);
        self!add-part($doc, $root, 'fills', @!fills);
        self!add-part($doc, $root, 'borders', @!borders);
        self!add-part($doc, $root, 'cellStyleXfs', @!formatting-records);
        self!add-part($doc, $root, 'cellXfs', @!cell-formats);
        
        return $doc.Str;
    }

    method !add-part(LibXML::Document $doc, LibXML::Element $root, Str $tag-name, @items) {
        if @items {
            my $part = $doc.createElement($tag-name);
            $part.add($doc.createAttribute('count', ~@items.elems));
            for @items {
                $part.add(.to-xml-element($doc));
            }
            $root.add($part);
        }
    }
}
