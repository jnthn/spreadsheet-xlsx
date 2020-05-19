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

given $sheet.workbook {
    isa-ok $_, Spreadsheet::XLSX::Workbook;
    is .worksheets.elems, 0, 'New workbook contains no worksheets';
    is .relationships.find-by-type('http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet').elems,
            0, 'There are no worksheet relations either';

    my $new-sheet-a;
    lives-ok { $new-sheet-a = .create-worksheet('Test A') },
        'Can create a new worksheet';
    isa-ok $new-sheet-a, Spreadsheet::XLSX::Worksheet,
        'The created worksheet is returned';
    is .worksheets.elems, 1,
        'Workbook now has 1 worksheet';
    is .worksheets[0].name, 'Test A',
        'Worksheet has expected name';
    given .relationships.find-by-type('http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet') {
        is .elems, 1, 'One new worksheet relation in the workbook relations';
        is .[0].target, $new-sheet-a.archive-path, 'Correct relation target';
    }

    my $new-sheet-b;
    lives-ok { $new-sheet-b = .create-worksheet('Test B') },
            'Can create another new worksheet';
    is .worksheets.elems, 2,
            'Workbook now has 2 worksheets';
    is .worksheets[1].name, 'Test B',
            'Worksheet has expected name';
    isnt $new-sheet-a.id, $new-sheet-b.id,
            'No ID conflict';
    isnt $new-sheet-a.archive-path, $new-sheet-b.archive-path,
            'No filename conflict';
}

done-testing;
