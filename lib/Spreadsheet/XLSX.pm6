use Libarchive::Simple;
use Spreadsheet::XLSX::ContentTypes;
use Spreadsheet::XLSX::Exceptions;
use Spreadsheet::XLSX::Relationships;
use Spreadsheet::XLSX::Root;
use Spreadsheet::XLSX::Workbook;

class Spreadsheet::XLSX does Spreadsheet::XLSX::Root {
    #| Map of files in the decompressed archive we read from, if any.
    has Hash $!archive;

    #| The content types of the workbook.
    has Spreadsheet::XLSX::ContentTypes $.content-types;

    #| Map of loaded relationships for paths. (Those never used are not
    #| in here.)
    has Spreadsheet::XLSX::Relationships %!relationships;

    #| The workbook itself.
    has Spreadsheet::XLSX::Workbook $.workbook
            handles <create-worksheet worksheets shared-strings>;

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
            with self.find-relationships('') -> Spreadsheet::XLSX::Relationships $top-rel {
                with $top-rel
                        .find-by-type('http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument')
                        .first {
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
            my $root-relationships = Spreadsheet::XLSX::Relationships.new:
                    for => '';
            $root-relationships.add:
                    type => 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument',
                    target => $workbook-path;
            %!relationships{''} = $root-relationships;

            # Create an empty workbook.
            my $workbook-relationships = Spreadsheet::XLSX::Relationships.new:
                    for => $workbook-path;
            $!workbook = Spreadsheet::XLSX::Workbook.new:
                    root => self,
                    relationships => $workbook-relationships;
        }
    }

    #| Loads the workbook XML file from the archive.
    method !load-workbook-xml(Str $workbook-file) {
        with $!archive{$workbook-file} {
            $!workbook = Spreadsheet::XLSX::Workbook.from-xml:
                    .decode('utf-8'), :root(self),
                    relationships => self.find-relationships($workbook-file);
        }
        else {
            die X::Spreadsheet::XLSX::Format.new:
                    message => "Workbook file '$workbook-file' not found in archive";
        }
    }

    #| Get the relationships for a given path in the XLSX archive.
    method find-relationships(Str $path --> Spreadsheet::XLSX::Relationships) {
        .return with %!relationships{$path};
        my $rel-path = do if $path eq '' {
            '_rels/.rels';
        }
        else {
            my @parts = $path.split('/');
            my $file = @parts.pop;
            (|@parts, '_rels', $file ~ '.rels').join('/')
        }
        with $!archive{$rel-path} {
            %!relationships{$path} = Spreadsheet::XLSX::Relationships.from-xml(.decode('utf8'), :for($path))
        }
        else {
            Nil
        }
    }

    #| Obtain a file from the archive. Will fail if we are not backed
    #| by an archive, or if there is no such file.
    method get-file-from-archive(Str $path --> Blob) {
        $!archive{$path} // fail "No such file '$path' in archive"
    }
}
