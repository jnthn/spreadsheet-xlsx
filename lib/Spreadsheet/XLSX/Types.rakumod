use v6.d;
unit module Spreadsheet::XLSX::Types;
# Commonly used base types

use Spreadsheet::XLSX::XMLHelpers;

enum PhoneticType is export (:PTHalfWidthKatakana<halfwidthKatakana>, :PTFullWidthKatakana<fullwidthKatakana>,
                             :PTHiragana<Hiragana>, :PTNoConversion<noConversion>);

enum PhoneticAlignment is export (:PANoControl<noControl>, :PALeft<left>,
                                  :PACenter<center>, :PADistributes<distributed>);

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

#| A color definition
class Color does XMLRepresentation["color"] is export {
    has Bool $.auto is xml-attr;
    has UInt $.indexed is xml-attr;
    has Str $.rgb is xml-attr;
    has UInt $.theme is xml-attr;
    has Real $.tint is xml-attr;
}

#| A font in the style store.
class Font does XMLRepresentation["font"] is export {
    has Str $.name                          is xml-elem(:val);
    has Int $.charset                       is xml-elem(:val);
    has FontFamily $.family                 is xml-elem(:val);
    has Bool $.bold                         is xml-elem(<b>, :val);
    has Bool $.italic                       is xml-elem(<i>, :val);
    has Bool $.strike                       is xml-elem(:val);
    has Bool $.outline                      is xml-elem(:val);
    has Bool $.shadow                       is xml-elem(:val);
    has Bool $.condense                     is xml-elem(:val);
    has Bool $.extend                       is xml-elem(:val);
    has Color $.color                       is xml-elem(:val);
    has Real $.size                         is xml-elem(<sz>, :val);
    has UnderlineValues $.underline         is xml-elem(<u>, :val);
    has VerticalAlignRun $.vertical-align   is xml-elem(<vertAlign>, :val);
    has FontScheme $.scheme                 is xml-elem(:val);
}

class CT_RPrElt does XMLRepresentation {
    has Str $.rFont                         is xml-elem(:val);
    has Int $.charset                       is xml-elem(:val);
    has Int $.family                        is xml-elem(:val);
    has Bool $.bold                         is xml-elem(<b>, :val);
    has Bool $.italic                       is xml-elem(<i>, :val);
    has Bool $.strike                       is xml-elem(:val);
    has Bool $.outline                      is xml-elem(:val);
    has Bool $.shadow                       is xml-elem(:val);
    has Bool $.condense                     is xml-elem(:val);
    has Bool $.extend                       is xml-elem(:val);
    has Color $.color                       is xml-elem(:val);
    has Real $.size                         is xml-elem(<sz>, :val);
    has UnderlineValues $.underline         is xml-elem(<u>, :val);
    has VerticalAlignRun $.vertical-align   is xml-elem(<vertAlign>, :val);
    has FontScheme $.scheme                 is xml-elem(:val);
}

class CT_RElt does XMLRepresentation {
    has CT_RPrElt $.rPr is xml-elem;
    has Str:D $.t is xml-elem is required;
}

class CT_PhoneticRun does XMLRepresentation[<rPh>] {
    has Str:D $.t is xml-elem is required;
    has UInt:D $.sb is xml-attr is required;
    has UInt:D $.eb is xml-attr is required;
}


class CT_PhoneticPr does XMLRepresentation[<phoneticPr>] {
    has UInt $.font-id is xml-attr;
    has PhoneticType $.type is xml-attr = PTFullWidthKatakana;
    has PhoneticAlignment $.alignment is xml-attr = PALeft;
}

# This one is used with some other, unimplemented yet, data structure. So, a candidate to move out of this module.
class CT_Rst is xml-sequence(<is>, :rPh(CT_PhoneticRun), :r(CT_RElt)) is export {
    has Str $.t                             is xml-elem;
    has CT_PhoneticPr $.phonetic-properties is xml-elem<phoneticPr>;

    # A convenience user-facing method.
    method value {
        $!t // self.grep(CT_RElt).map(*.t).join
    }
}
