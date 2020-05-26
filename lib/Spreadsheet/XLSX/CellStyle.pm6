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
        has Mu $.type;
    }

    #| Style properties metadata.
    my constant %properties = %(
        'bold' => Property.new(type => Bool),
        'italic' => Property.new(type => Bool),
        'font-size' => Property.new(type => Int),
    );

    submethod TWEAK(Int :$!style-id --> Nil) {}

    #| Should a bold font be used.
    method bold(--> Bool) is rw {
        self!property('bold')
    }

    #| Should an italic font be used.
    method italic(--> Bool) is rw {
        self!property('italic')
    }

    #| The size of font to use.
    method font-size(--> Int) is rw {
        self!property('font-size')
    }

    #| Produce a proxy for reading/writing the property.
    method !property(Str $key) is rw {
        Proxy.new:
                FETCH => -> | {
                    %!changed{$key} // self!fetch($key)
                },
                STORE => -> \p, $value {
                    %!changed{$key} = self!check-type($key, $value)
                }
    }

    method !fetch($key) {
        with $!style-id {
            die X::NYI.new(feature => 'Reading styles');
        }
        else {
            %properties{$key}.type
        }
    }

    method !check-type($key, $value) {
        my $type = do %properties{$key}.type;
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
            my Spreadsheet::XLSX::Styles::Font $font = self!build-font();
            my $style-id = $styles.obtain-style-id-for(:$font);
            %!changed = ();
            $style-id
        }
        else {
            $!style-id
        }
    }

    method !build-font(--> Spreadsheet::XLSX::Styles::Font) {
        my %props;
        %props<bold> = $_ with %!changed<bold>;
        %props<italic> = $_ with %!changed<italic>;
        %props<size> = $_ with %!changed<font-size>;
        return %props
                ?? Spreadsheet::XLSX::Styles::Font.new(|%props)
                !! Nil;
    }
}
