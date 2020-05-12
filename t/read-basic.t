use Spreadsheet::XLSX;
use Test;

given Spreadsheet::XLSX.load($*PROGRAM.parent.add('test-data/basic.xlsx')) {
    isa-ok $_, Spreadsheet::XLSX, 'Loaded gave a Spreadsheet::XLSX instance';

    given .content-types {
        is .defaults.elems, 3, 'Content type has 3 defaults';
        given .defaults[0] {
            is .extension, 'bin', 'Default has expected extension';
            is .content-type, 'application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings',
                'Default has expected content type';
        }
        is .overrides.elems, 8, 'Content type has 8 overrides';
        given .overrides[0] {
            is .part-name, '/xl/workbook.xml', 'Override has expected part name';
            is .content-type, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
                'Override has expected content type';
        }
        is-deeply .part-name-for-content-type('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml'),
                '/xl/workbook.xml',
                'Correct workbook part name';
        is-deeply .part-names-for-content-type('application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'),
                ('/xl/worksheets/sheet1.xml', '/xl/worksheets/sheet2.xml').Seq,
                'Correct worksheet part names';
    }

}

done-testing;
