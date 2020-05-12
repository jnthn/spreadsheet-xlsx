use Libarchive::Simple;
use Spreadsheet::XLSX::ContentTypes;
use Spreadsheet::XLSX::Exceptions;
use Spreadsheet::XLSX::Relationships;
use Spreadsheet::XLSX::Workbook;
use Spreadsheet::XLSX::Worksheet;

class Spreadsheet::XLSX does Spreadsheet::XLSX::Workbook {
    #| Map of files in the decompressed archive we read from, if any.
    has Hash $!archive;

    #| The content types of the workbook.
    has Spreadsheet::XLSX::ContentTypes $.content-types;

    #| The list of worksheets in the workbook.
    has @!worksheets;

    #| Map of loaded relationships for paths. (Those never used are not
    #| in here.)
    has Spreadsheet::XLSX::Relationships %!relationships;

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
        }
        else {
            ... "Don't know how to create an empty file yet";
        }
    }

    #| Create a new worksheet in this workbook.
    method create-worksheet(Str $name --> Spreadsheet::XLSX::Worksheet) {
        my $worksheet = Spreadsheet::XLSX::Worksheet.new(:workbook(self), :$name);
        @!worksheets.push($worksheet);
        return $worksheet;
    }

    #| Get a list of the worksheets in this workbook.
    method worksheets(--> List) {
        @!worksheets.List
    }

    #| Get the relationships for a given path in the XLSX archive.
    method find-relationships(Str $path --> Spreadsheet::XLSX::Relationships) {
        .return with %!relationships{$path};
        my $rel-path = $path eq '' ?? '_rels/.rels' !! die('NYI');
        with $!archive{$rel-path} {
            %!relationships{$path} = Spreadsheet::XLSX::Relationships.from-xml(.decode('utf8'))
        }
        else {
            Nil
        }
    }
}
