use Spreadsheet::XLSX;
use Test;

my $workbook = Spreadsheet::XLSX.new;
my $sheet = $workbook.create-worksheet('Test sheet');

$sheet.set(0, 0, 'Answer', :bold);
given $sheet.cells[0;0] {
    isa-ok $_, Spreadsheet::XLSX::Cell::Text,
        'Convenience set form creates a text cell with a string';
    is-deeply .value, 'Answer', 'Correct cell value';
    ok .style.bold, 'Bold style set';
    nok .style.italic, 'Italic style not set';
}

$sheet.set(0, 1, 42, :number-format('#.#'));
given $sheet.cells[0;1] {
    isa-ok $_, Spreadsheet::XLSX::Cell::Number,
            'Convenience set form creates a number cell with a string';
    is-deeply .value, 42, 'Correct cell value';
    nok .style.bold, 'No bold style set';
    is .style.number-format, '#.#', 'Number format is set';
}

dies-ok { $sheet.set(0, 2, 42, :no-such-style) },
        'Trying to set a style that does not exist dies';

done-testing;
