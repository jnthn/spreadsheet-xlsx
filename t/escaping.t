use Spreadsheet::XLSX;
use Test;

my $blob;

subtest 'Encoding' => {
    my $new-workbook = Spreadsheet::XLSX.new;
    my $sheet = $new-workbook.create-worksheet('Test escaping');
    lives-ok { $sheet.set(0, 0, 'Hello & goodbye') },
            'Can set text containing &';
    lives-ok { $sheet.set(0, 1, 'I <3 XML') },
            'Can set text containing <';
    lives-ok { $blob = $new-workbook.to-blob },
            'Can save worksheet with text containing & and <';
}

subtest 'Decoding' => {
    my Spreadsheet::XLSX $loaded-workbook .= load($blob);
    my $sheet = $loaded-workbook.workbook.worksheets[0];
    is $sheet.cells[0; 0], 'Hello & goodbye',
        'Round-tripped cell containing &';
    is $sheet.cells[0; 1], 'I <3 XML',
            'Round-tripped cell containing <';
}

done-testing;
