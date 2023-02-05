use Spreadsheet::XLSX::Styles;
use Spreadsheet::XLSX::Types;
use Spreadsheet::XLSX::Root;

#| Read and write styles on a cell. An XLSX stores styles in a form optimized
#| for re-use, but not especially convenient to work with. This class serves
#| as a facade to that data. When there is a style ID, styles are resolved
#| through the styles object into the appropriate value. When styles are set
#| on this class, they are stored here, and then saved into the XLSX-level
#| styles at the point of saving the sheet.
class Spreadsheet::XLSX::CellStyle {
    use Spreadsheet::XLSX::XMLHelpers;

    # Local aliases for attributes (properties) on group objects
    my %prop-alias = %(
        font => %(
            :font-size<size>,
            :font-name<name>,
        ),
        alignment => %(
            :horizontal-align<horizontal>,
            :vertical-align<vertical>,
        ),
        format => %(
        ),
        numformat => %(
            :number-format<code>,
        ),
    );

    my role IndirectPropertyOf[::GroupType, Str:D $group-name] {
        has $!group is built;
        has %!changes;
        has $!pos-changes;

        has %!group-types{Mu};

        method of { ($!group // GroupType).WHAT }
        method group-name is pure { $group-name }

        method SETUP-SELF is implementation-detail {
            with $!group {
                # Caching property types at run-time because $!prop can be a descendant of GroupType
                for $!group.^attributes(:all) {
                    %!group-types{ .name.substr(2) } := .type;
                }

                # For sequences we'd need to keep track of changed sequence items too.
                $!pos-changes := $!group ~~ XMLSequence ?? Array[$!group.of].new !! Nil;
            }
        }

        method !check-type($name, $prop, Mu \value) is raw {
            unless value ~~ %!group-types{$name} {
                die X::TypeCheck::Assignment.new:
                    got => value,
                    expected => %!group-types{$name},
                    symbol => self.^name ~ "." ~ $prop;
            }
            value
        }

        method !prop(Str:D $prop) is rw {
            my $name = %prop-alias{$group-name}{$prop} // $prop;
            Proxy.new:
                FETCH => -> | {
                    %!changes{$name}:exists
                        ?? %!changes{$name}
                        !! ($!group andthen $!group."$name"() orelse Nil)
                },
                STORE => -> $, Mu \value {
                    %!changes{$name} := self!check-type: $name, $prop, value;
                }
        }

        method !sequential-only {
            $!group ~~ XMLSequence
                or die "Unable to use positional postcircumfix `[]` with " ~ self.^name;
        }

        method AT-POS(::?CLASS:D: Int:D $pos) {
            self!sequential-only;
            $!pos-changes[$pos]:exists
                ?? $!pos-changes.AT-POS($pos)
                !! $!group.AT-POS($pos)
        }

        method EXISTS-POS(::?CLASS:D: Int:D $pos) {
            self!sequential-only;
            $!pos-changes[$pos].EXISTS-POS($pos) || $!group.EXISTS-POS($pos)
        }

        method ASSIGN-POS(::?CLASS:D: Int:D $pos, \value) {
            self!sequential-only;
            $!pos-changes[$pos] = value
        }

        method !initial-push-append($method, \value) is raw {
            # No new elements were added yet beyond what we have in $!prop.
            # It wouldn't be wise to try to duplicate the work Array.push does. Then use a little trick, since
            # the slot at the position of the last element in $!prop is not used anyway (or else $!pos-changes would be
            # at least the length of $!prop).
            my $last-pos = $!group.end;
            $!pos-changes[$last-pos] = $!group[$last-pos]; # Make $!pos-changes the same length
            LEAVE $!pos-changes[$last-pos]:delete; # Release the slot again
            $!pos-changes."$method"(value)
        }

        method push(::?CLASS:D: \value) {
            self!sequential-only;

            $!pos-changes.elems < $!group.elems
                ?? self!initial-push-append("push", value)
                !! $!pos-changes.push(value)
        }

        method append(::?CLASS:D: \value) {
            self!sequential-only;

            $!pos-changes.elems < $!group.elems
                ?? self!initial-push-append("appen", value)
                !! $!pos-changes.append(value)
        }

        method reify(::?CLASS:D: --> GroupType) {
            my $changed := Nil;
            my $is-changed = %!changes || $!pos-changes;
            with $!group {
                # If group object is set and no changes has been made then nothing to be done. Otherwise we use
                # changes hash as twiddles for clone and re-assign changed positionals.
                return $_ unless $is-changed;
                $changed := .clone: |%!changes;
                if $_ ~~ XMLSequence {
                    # Migrate only changed elements
                    for ^$changed.elems -> $idx {
                        $changed[$idx] = $!pos-changes[$idx] if $!pos-changes.EXISTS-POS($idx);
                    }
                }
            }
            elsif $is-changed {
                die "Don't know how to produce an object of type " ~ GroupType.^name
                    unless GroupType ~~ XMLRepresentation;
                $changed := GroupType.new: |%!changes;
                if $changed ~~ XMLSequence && $!pos-changes {
                    for ^$!pos-changes.elems -> $idx {
                        $changed.ASSIGN-POS($idx, $!pos-changes.AT-POS($idx)) if $!pos-changes.EXISTS-POS($idx);
                    }
                }
            }
            %!changes = ();
            $!pos-changes = ();
            $!group := $changed
        }

        proto method new-from(|)                 {*}
        multi method new-from(Nil)               { self.new }
        multi method new-from(GroupType:U $) { self.new }
        multi method new-from(GroupType:D $group) { self.new(:$group) }

        method PROP-CAN(Str:D $method-name) is implementation-detail {
            (self andthen $!group orelse GroupType).^can($method-name)
        }

        # Allow an instance to be accessed as $indir-prop.property
        ::?CLASS.^add_fallback(
            -> \obj, $name {
                obj.PROP-CAN(%prop-alias{obj.group-name}{$name} // $name)
            },
            -> \obj, $name {
                my &meth = anon method (::?CLASS:D:) is rw {
                    self!prop: $name
                }
                &meth.set_name($name);
                ::?CLASS.^add_method($name, &meth);
                &meth
            });
    }
    my class IndirectProperty {

        multi method new(XMLRepresentation:D $group) {
            samewith :$group, |%_
        }

        submethod TWEAK {
            self.SETUP-SELF;
        }

        method ^parameterize(\obj, \of, Str:D $group-name) is raw {
            my \what := obj.^mixin(IndirectPropertyOf[of, $group-name]);
            what.^set_name("CellStyle." ~ $group-name);
            what
        }
    }

    has Spreadsheet::XLSX::Root $.root;

    #| The style ID, if we have one.
    has Int $.style-id;

    has IndirectProperty[Spreadsheet::XLSX::Styles::Format, "format"]           $.format
        handles <use-alignment use-number-format use-font use-fill use-border>;

    has IndirectProperty[Spreadsheet::XLSX::Types::Font, "font"]                $.font
        handles <bold italic font-size font-name>;
    has IndirectProperty[Spreadsheet::XLSX::Styles::Border, "border"]           $.border;
    has IndirectProperty[Spreadsheet::XLSX::Styles::Fill, "fill"]               $.fill;
    has IndirectProperty[Spreadsheet::XLSX::Styles::CellAlignment, "alignment"] $.alignment
        handles <horizontal-align vertical-align wrap-text>;
    has IndirectProperty[Spreadsheet::XLSX::Styles::NumberFormat, "numformat"]  $.numformat
        handles<number-format>;

    submethod TWEAK {
        self!RESET(:full);
    }

    method !RESET(:$full) {
        with $!root {
            my $styles = $!root.styles;
            if $full {
                $!format = $!format.new-from: ($!style-id andthen $styles.cell-formats[$_] orelse Nil);
            }
            given $!format {
                $!font      .= new-from: (.font-id          andthen $styles.fonts[$_]                orelse Nil);
                $!border    .= new-from: (.border-id        andthen $styles.borders[$_]              orelse Nil);
                $!fill      .= new-from: (.fill-id          andthen $styles.fills[$_]                orelse Nil);
                $!alignment .= new-from: (.alignment                                                 orelse Nil);
                $!numformat .= new-from: (.number-format-id andthen $styles.number-formats.by-id($_) orelse Nil);
            }
        }
        else {
            $!format    .= new;
            $!font      .= new;
            $!border    .= new;
            $!fill      .= new;
            $!alignment .= new;
            $!numformat .= new;
        }
        self
    }

    method set-id(Int:D $id) {
        $!style-id = $id;
        self!RESET(:full);
    }

    method set-format(Spreadsheet::XLSX::Styles::Format $format) {
        $!style-id = Nil;
        $!format .= new-from($format);
        self!RESET;
    }

    method set-font(Spreadsheet::XLSX::Types::Font $font) {
        $!style-id = Nil;
        $!font .= new-from($font);
        $.format.font-id = Nil;
        self
    }

    method set-border(Spreadsheet::XLSX::Styles::Border $border) {
        $!style-id = Nil;
        $!border .= new-from($border);
        $.format.border-id = Nil;
        self
    }

    method set-fill(Spreadsheet::XLSX::Styles::Fill $fill) {
        $!style-id = Nil;
        $!fill .= new-from($fill);
        $.format.fill-id = Nil;
        self
    }

    method set-alignment(Spreadsheet::XLSX::Styles::CellAlignment $alignment) {
        $!style-id = Nil;
        $!alignment .= new-from($alignment);
        $.format.alignment = Nil;
        self
    }

    #| Takes any changes to the styles and obtains a style ID that
    #| describes them. If there are no changes, then any existing style
    #| ID that was assigned to this element will be used.
    method sync-style-id(Spreadsheet::XLSX::Styles $styles --> Int) {
        with $!format {

            # For a newly created number format we need to set a first id because it is required. It's safe because
            # with :autogen `id-for` method candidate for number formats will generate and use a valid ID since we are
            # unlikely to match that number format if it exists. And even if we accidentally would the only thing it
            # means is that it is the actually requested format.
            # Using 164 as it seems to be de-facto default for the first custom format ID.
            $!numformat.id //= 164 if $!numformat.number-format;

            .font-id          = ($!font.reify      andthen $styles.id-for($_, :autogen) orelse Nil);
            .border-id        = ($!border.reify    andthen $styles.id-for($_, :autogen) orelse Nil);
            .number-format-id = ($!numformat.reify andthen $styles.id-for($_, :autogen) orelse Nil);
            .fill-id          = ($!fill.reify      andthen $styles.id-for($_, :autogen) orelse Nil);
            .alignment        = ($!alignment.reify                                      orelse Nil);

            $!style-id = $styles.id-for(.reify, $styles.cell-formats, :autogen);
            self!RESET;
        }
        $!style-id;
    }
}
