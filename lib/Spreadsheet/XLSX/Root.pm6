use Spreadsheet::XLSX::Relationships;

#| Operations provided by the root of the XLSX archive, such as reference
#| resolution.
role Spreadsheet::XLSX::Root {
    method find-relationships(Str $path --> Spreadsheet::XLSX::Relationships) { ... }
    method get-file-from-archive(Str $path --> Blob) { ... }
    method set-file-in-archive(Str $path, Blob $content --> Nil) { ... }
}
