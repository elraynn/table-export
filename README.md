# table-export

Export an array of rows to Excel or PDF with one consistent API, instead of hand-rolling PhpSpreadsheet/dompdf calls in every report controller.

If you've built reporting features in a PHP business app (invoices, purchase orders, attendance, whatever), you've probably written the same 40-line block more than once: create a spreadsheet, set column widths letter by letter, bold the header row, add borders, loop the rows, stream the download. Every report ends up with its own slightly different copy of that block. This wraps it into one call.

## Requirements

PHP 7.4+, [PhpSpreadsheet](https://github.com/PHPOffice/PhpSpreadsheet) and [dompdf](https://github.com/dompdf/dompdf) (both pulled in as dependencies).

## Install

```
composer require elrayn/table-export
```

## Usage

```php
use Elrayn\TableExport\Export;

Export::make($rows)
    ->title('Laporan Pembelian')
    ->columns([
        'kode'     => 'Kode',
        'tanggal'  => 'Tanggal',
        'supplier' => 'Supplier',
        'total'    => 'Total',
    ])
    ->asExcel('laporan-pembelian.xlsx');
```

`$rows` is any array of associative arrays or objects, `columns()` maps the key you read from each row to the column header you want. Same setup works for PDF:

```php
Export::make($rows)
    ->columns(['kode' => 'Kode', 'total' => 'Total'])
    ->asPdf('laporan.pdf');
```

Both methods send the right headers and stream the file directly, so a controller action can be as small as:

```php
public function exportExcel()
{
    Export::make($this->reportModel->all())
        ->title('Laporan Pembelian')
        ->columns(['kode' => 'Kode', 'tanggal' => 'Tanggal', 'total' => 'Total'])
        ->asExcel('pembelian.xlsx');
}
```

## What it doesn't do

This isn't trying to replace PhpSpreadsheet or dompdf for anything beyond a plain tabular export. No merged cells, no charts, no multi-sheet workbooks. If you need that, use the underlying libraries directly. This is for the 90% case: a table of rows going out as Excel or PDF, styled well enough to hand to someone.

## License

MIT
