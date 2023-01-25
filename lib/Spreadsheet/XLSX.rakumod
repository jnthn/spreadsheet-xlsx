use Libarchive::Simple;
use Spreadsheet::XLSX::ContentTypes;
use Spreadsheet::XLSX::DocProps::Core;
use Spreadsheet::XLSX::Exceptions;
use Spreadsheet::XLSX::Relationships;
use Spreadsheet::XLSX::Root;
use Spreadsheet::XLSX::Workbook;
use Spreadsheet::XLSX::Worksheet;

class Spreadsheet::XLSX does Spreadsheet::XLSX::Root {
    #| Map of files in the decompressed archive we read from, if any.
    has Hash $!archive;

    #| The content types of the workbook.
    has Spreadsheet::XLSX::ContentTypes $.content-types;

    #| Map of loaded relationships for paths. (Those never used are not
    #| in here.)
    has Spreadsheet::XLSX::Relationships %!relationships;

    has Spreadsheet::XLSX::Relationships $!root-relationships;

    #| The workbook itself.
    has Spreadsheet::XLSX::Workbook $.workbook
            handles <create-worksheet worksheets shared-strings styles>;

    #| Document Core Properties
    has Spreadsheet::XLSX::DocProps::Core $!core-props;

    my $core-prop-type = 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties';

    #| Load an Excel workbook from the file path identified by the given string.
    multi method load(Str $file --> Spreadsheet::XLSX) {
        self.load($file.IO)
    }

    #| Load an Excel workbook in the specified file.
    multi method load(IO::Path $file --> Spreadsheet::XLSX) {
        self.load($file.slurp(:bin))
    }

    #| Load an Excel workbook from the specified blob. This is useful in
    #| the case it was sent over the network, and so never written to disk.
    multi method load(Blob $content --> Spreadsheet::XLSX) {
        my %archive = do for archive-read($content, :format<zip>) {
            .pathname => .data if .is-file
        }
        self.new(:%archive)
    }

    submethod TWEAK(Hash :$!archive) {
        # If we are being created based upon an archive, then we need to
        # parse that.
        with $!archive {
            # First, extract the content types, which we shall need to find
            # everything else.
            with $!archive{'[Content_Types].xml'} -> Blob $content-types {
                $!content-types = Spreadsheet::XLSX::ContentTypes.from-xml($content-types.decode('utf-8'));
            }
            else {
                die X::Spreadsheet::XLSX::Format.new: message =>
                    'Required [Content_Types].xml is missing'
            }

            # Locate the root relationships file, and using it, the workbook root.
            with $!root-relationships = self.find-relationships('') {
                with $!root-relationships
                        .find-by-type('http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument')
                        .first
                {
                    self!load-workbook-xml(.target);
                }
                else {
                    die X::Spreadsheet::XLSX::Format.new: message =>
                            'No workbook relation found'
                }
            }
            else {
                die X::Spreadsheet::XLSX::Format.new: message =>
                        'Required top-level rels are missing'
            }
        }
        else {
            # Create default set of content types (minimal needed).
            $!content-types = Spreadsheet::XLSX::ContentTypes.new;

            # Set up root relationships, indicating how the workbook is
            # found.
            my constant $workbook-path = 'xl/workbook.xml';

            $!root-relationships = self.find-relationships('', :create);
            $!root-relationships.add:
                    type => 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument',
                    target => $workbook-path;

            # Create an empty workbook.
            my $workbook-relationships = self.find-relationships($workbook-path, :create);
            $!workbook = Spreadsheet::XLSX::Workbook.new:
                    root => self,
                    relationships => $workbook-relationships;
        }
    }

    method !archive-file(Str:D $file, Str:D $what) {
        $!archive{$file}
            // die X::Spreadsheet::XLSX::Format.new: message => "$what file '$file' not found in archive"
    }

    #| Loads the workbook XML file from the archive.
    method !load-workbook-xml(Str $workbook-file) {
        $!workbook = Spreadsheet::XLSX::Workbook.from-xml:
            self!archive-file($workbook-file, "Workbook").decode('utf-8'),
            :root(self), relationships => self.find-relationships($workbook-file);
    }

    method !core-props-rel {
        $!root-relationships.find-by-type($core-prop-type).head
    }

    method !load-core-props {
        with self!core-props-rel {
            Spreadsheet::XLSX::DocProps::Core.from-xml:
                self!archive-file(.target, "Core properties").decode('utf-8')
        }
        else {
            $!root-relationships.add: :target('docProps/core.xml'), :type($core-prop-type);
            Spreadsheet::XLSX::DocProps::Core.new
        }
    }

    #| Get the relationships for a given path in the XLSX archive.
    method find-relationships(Str $path, Bool :$create = False --> Spreadsheet::XLSX::Relationships) {
        .return with %!relationships{$path};
        my $rel-path = self!rel-path($path);
        with $!archive{$rel-path} {
            %!relationships{$path} = Spreadsheet::XLSX::Relationships.from-xml(.decode('utf8'), :for($path))
        }
        elsif $create {
            %!relationships{$path} = Spreadsheet::XLSX::Relationships.new(:for($path))
        }
        else {
            Nil
        }
    }

    #| Calculate the path of the relations file for a given path.
    method !rel-path(Str $path --> Str) {
        if $path eq '' {
            '_rels/.rels';
        }
        else {
            my @parts = $path.split('/');
            my $file = @parts.pop;
            (|@parts, '_rels', $file ~ '.rels').join('/')
        }
    }

    #| Obtain a file from the archive. Will fail if we are not backed
    #| by an archive, or if there is no such file.
    method get-file-from-archive(Str $path --> Blob) {
        $!archive{$path} // fail "No such file '$path' in archive"
    }

    #| Set the content of a file in the archive, replacing any existing
    #| content.
    method set-file-in-archive(Str $path, Blob $content --> Nil) {
        $!archive{$path} = $content;
    }

    #| Serializes the current state of the spreadsheet into XLSX format
    #| and returns a Blob containing it.
    method to-blob(--> Blob) {
        # Get the archive hash updated with all changes.
        self.sync-to-archive();

        # Serialize it to a ZIP.
        my $buffer = Buf.new;
        given archive-write($buffer, format => 'zip') -> $archive {
            for $!archive.kv -> $path, $blob {
                $archive.write($path, $blob);
            }
            $archive.close;
        }
        return $buffer;
    }

    #| Saves the Excel workbook to the file path identified by the given string.
    multi method save(Str $file --> Nil) {
        self.save($file.IO)
    }

    #| Save an Excel workbook to the specified file.
    multi method save(IO::Path $file --> Nil) {
        $file.spurt(self.to-blob)
    }

    #| Synchronizes all changes to the internal representation of the
    #| archive. This is performed automatically before saving, and there
    #| is no need to explicitly perform it.
    method sync-to-archive(--> Nil) {
        # Sync the workbook, which will in turn handle sync of anything it owns.
        $!workbook.sync-to-archive();

        # Store core properties if were changed in any way.
        with $!core-props {
            $!archive{self!core-props-rel.target} = .to-xml;
        }

        # Sync any relationships objects we have; we only have these if we read
        # them, and so potentially modified them. Untouched ones won't need to
        # be updated.
        for %!relationships.values -> Spreadsheet::XLSX::Relationships $rels {
            $!archive{self!rel-path($rels.for)} = $rels.to-xml();
        }

        # Sync the content types; even if we didn't change these, they
        # need to be saved.
        $!archive //= {};
        $!archive{'[Content_Types].xml'} = $!content-types.to-xml();
    }

    method core-properties {
        $!core-props //= self!load-core-props;
    }
}

multi sub postcircumfix:<[; ]>( Spreadsheet::XLSX::Worksheet::Cells:D $c, @indicies, Bool:D :$exists! ) is export {
    $exists ?? $c.EXISTS-POS(|@indicies) !! !$c.EXISTS-POS(|@indicies)
}

multi sub postcircumfix:<[; ]>( Spreadsheet::XLSX::Worksheet::Cells:D $c, @indicies ) is export {
    $c.AT-POS: |@indicies
}

multi sub postcircumfix:<[ ]>(Spreadsheet::XLSX::Worksheet::Cells:D $c, Str:D $ref, Bool:D :$exists!) is export {
    $exists ?? $c.EXISTS-POS($ref) !! !$c.EXISTS-POS($ref)
}

multi sub postcircumfix:<[ ]>(Spreadsheet::XLSX::Worksheet::Cells:D $c, Str:D $ref) is export {
    $c.AT-POS($ref)
}

multi sub postcircumfix:<[; ]>( Spreadsheet::XLSX::Worksheet::Cells:D $c,
                               @indicies,
                               Spreadsheet::XLSX::Cell $value ) is export
{
    $c.ASSIGN-POS(|@indicies, $value)
}

multi sub postcircumfix:<[ ]>( Spreadsheet::XLSX::Worksheet::Cells:D $c,
                               Str:D $ref,
                               Spreadsheet::XLSX::Cell $value ) is export
{
    $c.ASSIGN-POS($ref, $value)
}
