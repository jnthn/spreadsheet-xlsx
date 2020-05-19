use Spreadsheet::XLSX;
use Test;

my Spreadsheet::XLSX $sheet;
lives-ok { $sheet = Spreadsheet::XLSX.new },
        'Can create a new SpreadSheet::XLSX instance';

given $sheet.content-types {
    isa-ok $_, Spreadsheet::XLSX::ContentTypes;
    is .defaults.elems, 2, '2 default content types created';
    is .defaults.map(*.extension).sort, <rels xml>,
            'Created expected set of defaults';
    is .overrides.elems, 1, '1 default override created';
    given .overrides[0] {
        is .part-name, '/xl/workbook.xml',
                'Correct workbook override part name';
        is .content-type, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
                'Correct workbook override content type';
    }
}

given $sheet.find-relationships('') {
    isa-ok $_, Spreadsheet::XLSX::Relationships,
        'Created a root relationships file';
    given .find-by-type('http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument') {
        is .elems, 1, 'Have a workbook type relationship in the root';
        is .[0].target, 'xl/workbook.xml', 'Relationship point to standard workbook location';
    }
}

done-testing;
