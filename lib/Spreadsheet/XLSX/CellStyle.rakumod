use Spreadsheet::XLSX::Styles;

#| Read and write styles on a cell. An XLSX stores styles in a form optimized
#| for re-use, but not especially convenient to work with. This class serves
#| as a facade to that data. When there is a style ID, styles are resolved
#| through the styles object into the appropriate value. When styles are set
#| on this class, they are stored here, and then saved into the XLSX-level
#| styles at the point of saving the sheet.
class Spreadsheet::XLSX::CellStyle {
    #| Changed values.
    has %!changed;

    #| The style ID, if we have one.
    has Int $!style-id;

    #| Metadata about a style property, describing their backing in the
    #| underlying style store.
    my class Property {
        has Mu $.type is required;
        has Str $.attr-name;
    }

    #| Style properties metadata.
    my constant %properties = %(
        font => %(
            'bold' => Property.new(type => Bool),
            'italic' => Property.new(type => Bool),
            'font-size' => Property.new(type => Int, attr-name => 'size'),
        ),
        alignment => %(
            'horizontal-align' => Property.new(type => HorizontalAlign, attr-name => 'horizontal'),
            'vertical-align' => Property.new(type => VerticalAlign, attr-name => 'vertical'),
            'wrap-text' => Property.new(type => Bool),
        ),
        format => %( number-format => Property.new(type => Str) )
    );

    submethod TWEAK(Int :$!style-id --> Nil) {}

    #| Should a bold font be used.
    method bold(--> Bool) is rw {
        self!property('font', 'bold')
    }

    #| Should an italic font be used.
    method italic(--> Bool) is rw {
        self!property('font', 'italic')
    }

    #| The size of font to use.
    method font-size(--> Int) is rw {
        self!property('font', 'font-size')
    }

    # The horizontal alignment of a cell.
    method horizontal-align(--> HorizontalAlign) is rw {
        self!property('alignment', 'horizontal-align')
    }

    # The vertical alignment of a cell.
    method vertical-align(--> VerticalAlign) is rw {
        self!property('alignment', 'vertical-align')
    }

    #| Whether text in the cell should be wrapped.
    method wrap-text(--> Bool) is rw {
        self!property('alignment', 'wrap-text')
    }

    #| The number format.
    method number-format(--> Str) is rw {
        self!property('format', 'number-format')
    }

    #| Produce a proxy for reading/writing the property.
    method !property(Str $group, Str $key) is rw {
        Proxy.new:
                FETCH => -> | {
                    %!changed{$key} // self!fetch($group, $key)
                },
                STORE => -> \p, $value {
                    %!changed{$key} = self!check-type($group, $key, $value)
                }
    }

    method !fetch(Str $group, Str $key) {
        with $!style-id {
            die X::NYI.new(feature => 'Reading styles');
        }
        else {
            %properties{$group}{$key}.type
        }
    }

    method !check-type(Str $group, Str $key, $value) {
        my $type = %properties{$group}{$key}.type;
        unless $value ~~ $type {
            die X::TypeCheck::Assignment.new:
                    got => $value,
                    expected => $type,
                    symbol => $key;
        }
        $value
    }

    #| Takes any changes to the styles and obtains a style ID that
    #| describes them. If there are no changes, then any existing style
    #| ID that was assigned to this element will be used.
    method sync-style-id(Spreadsheet::XLSX::Styles $styles --> Int) {
        if %!changed {
            with $!style-id {
                die X::NYI.new(feature => 'saving changes to existing styles');
            }
            my $font = self!build-group(Spreadsheet::XLSX::Styles::Font, 'font');
            my $alignment = self!build-group(Spreadsheet::XLSX::Styles::CellAlignment, 'alignment');
            my $number-format = %!changed<number-format> // Str;
            my $style-id = $styles.obtain-style-id-for(:$font, :$alignment, :$number-format);
            %!changed = ();
            $style-id
        }
        else {
            $!style-id
        }
    }

    method !build-group(Mu $type, Str $group) {
        my %props;
        for %properties{$group}.kv -> $prop, Property $metadata {
            with %!changed{$prop} {
                %props{$metadata.attr-name // $prop} = $_;
            }
        }
        return %props ?? $type.new(|%props) !! $type;
    }
}
