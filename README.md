# table-export

Export an array of rows to Excel or PDF with one consistent API, instead of hand-rolling PhpSpreadsheet/dompdf calls in every report controller.

If you've built reporting features in a PHP business app (invoices, purchase orders, attendance, whatever), you've probably written the same 40-line block more than once: create a spreadsheet, set column widths letter by letter, bold the header row, add borders, loop the rows, stream the download. Every report ends up with its own slightly different copy of that block. This wraps it into one call.

## Requirements

PHP 7.4+ and [PhpSpreadsheet](https://github.com/PHPOffice/PhpSpreadsheet) (pulled in as a dependency).

`asPdf()` and `toPdfBinary()` need [dompdf](https://github.com/dompdf/dompdf) too, but it's not a hard dependency, install it yourself if you use either:

```
composer require dompdf/dompdf
```

On PHP 8.1+, Composer will give you dompdf 3.x. On PHP 7.4-8.0, it'll give you 2.0.8, the last release in that line (Composer's security audit will warn about it, there's no dompdf release that's both PHP <8.1 compatible and advisory-free right now).

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

Both methods send the right headers and stream the file directly, so from a plain script this is the whole thing:

```php
Export::make($rows)->columns([...])->asExcel('pembelian.xlsx');
// headers sent, file streamed, done
```

### Using this inside a framework

`asExcel()` and `asPdf()` call `header()` and echo the file straight to output — they assume nothing else is going to touch the response afterward. That's true for a plain script, but **not** true inside something like a Laravel controller: the framework builds its own `Response` after your action returns and sends *that*, which stomps the headers you just set back to `text/html`. The file contents still get echoed first, so you end up with a broken download instead of an error, which makes this easy to miss until someone opens the file.

Use `toXlsxBinary()` / `toPdfBinary()` instead — same output, but returned as a plain string with no headers sent and nothing echoed, so you can hand it to your framework's own response:

```php
public function exportExcel()
{
    $binary = Export::make($this->reportModel->all())
        ->title('Laporan Pembelian')
        ->columns(['kode' => 'Kode', 'tanggal' => 'Tanggal', 'total' => 'Total'])
        ->toXlsxBinary();

    return response($binary, 200, [
        'Content-Type' => 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Content-Disposition' => 'attachment;filename="pembelian.xlsx"',
    ]);
}

public function exportPdf()
{
    $binary = Export::make($this->reportModel->all())
        ->columns(['kode' => 'Kode', 'total' => 'Total'])
        ->toPdfBinary();

    return response($binary, 200, [
        'Content-Type' => 'application/pdf',
        'Content-Disposition' => 'attachment;filename="laporan.pdf"',
    ]);
}
```

## What it doesn't do

This isn't trying to replace PhpSpreadsheet or dompdf for anything beyond a plain tabular export. No merged cells, no charts, no multi-sheet workbooks. If you need that, use the underlying libraries directly. This is for the 90% case: a table of rows going out as Excel or PDF, styled well enough to hand to someone.

## License

MIT
