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
            handles <create-worksheet worksheets>;

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

    submethod TWEAK(:$!archive) {
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
                with $top-rel.find-by-type('http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument').first {
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
            ... "Don't know how to create an empty file yet";
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
}
