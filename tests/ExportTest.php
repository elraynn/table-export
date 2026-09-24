<?php

namespace Elrayn\TableExport\Tests;

use Elrayn\TableExport\Export;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PHPUnit\Framework\TestCase;

class ExportTest extends TestCase
{
    protected function sampleRows(): array
    {
        return [
            ['kode' => 'C001', 'nama' => 'Budi', 'total' => 100],
            ['kode' => 'C002', 'nama' => 'Ani', 'total' => 250],
        ];
    }

    public function testSpreadsheetHeaderRowMatchesColumnLabels(): void
    {
        $sheet = Export::make($this->sampleRows())
            ->columns(['kode' => 'Kode', 'nama' => 'Nama', 'total' => 'Total'])
            ->toSpreadsheet()
            ->getActiveSheet();

        $this->assertSame('Kode', $sheet->getCell('A1')->getValue());
        $this->assertSame('Nama', $sheet->getCell('B1')->getValue());
        $this->assertSame('Total', $sheet->getCell('C1')->getValue());
    }

    public function testSpreadsheetDataRowsFollowHeader(): void
    {
        $sheet = Export::make($this->sampleRows())
            ->columns(['kode' => 'Kode', 'nama' => 'Nama', 'total' => 'Total'])
            ->toSpreadsheet()
            ->getActiveSheet();

        $this->assertSame('C001', $sheet->getCell('A2')->getValue());
        $this->assertSame('Budi', $sheet->getCell('B2')->getValue());
        $this->assertSame('C002', $sheet->getCell('A3')->getValue());
    }

    public function testTitlePushesHeaderRowDown(): void
    {
        $sheet = Export::make($this->sampleRows())
            ->title('Laporan Test')
            ->columns(['kode' => 'Kode', 'nama' => 'Nama', 'total' => 'Total'])
            ->toSpreadsheet()
            ->getActiveSheet();

        $this->assertSame('Laporan Test', $sheet->getCell('A1')->getValue());
        $this->assertSame('Kode', $sheet->getCell('A3')->getValue());
        $this->assertSame('C001', $sheet->getCell('A4')->getValue());
    }

    public function testWorksWithObjectRowsNotJustArrays(): void
    {
        $rows = [(object) ['kode' => 'C001', 'nama' => 'Budi']];

        $sheet = Export::make($rows)
            ->columns(['kode' => 'Kode', 'nama' => 'Nama'])
            ->toSpreadsheet()
            ->getActiveSheet();

        $this->assertSame('Budi', $sheet->getCell('B2')->getValue());
    }

    public function testHtmlContainsHeaderAndRowValues(): void
    {
        $html = Export::make($this->sampleRows())
            ->title('Laporan Test')
            ->columns(['kode' => 'Kode', 'nama' => 'Nama', 'total' => 'Total'])
            ->toHtml();

        $this->assertStringContainsString('Laporan Test', $html);
        $this->assertStringContainsString('<th>Kode</th>', $html);
        $this->assertStringContainsString('<td>Budi</td>', $html);
    }

    public function testHtmlEscapesValues(): void
    {
        $rows = [['nama' => '<script>alert(1)</script>']];

        $html = Export::make($rows)->columns(['nama' => 'Nama'])->toHtml();

        $this->assertStringNotContainsString('<script>alert(1)</script>', $html);
        $this->assertStringContainsString('&lt;script&gt;', $html);
    }

    public function testXlsxBinaryProducesAValidReadableSpreadsheet(): void
    {
        $binary = Export::make($this->sampleRows())
            ->columns(['kode' => 'Kode', 'nama' => 'Nama', 'total' => 'Total'])
            ->toXlsxBinary();

        $this->assertStringStartsWith("PK\x03\x04", $binary, 'xlsx files are zip archives');

        $tempFile = tempnam(sys_get_temp_dir(), 'table-export-test') . '.xlsx';
        file_put_contents($tempFile, $binary);

        $sheet = IOFactory::load($tempFile)->getActiveSheet();
        unlink($tempFile);

        $this->assertSame('Kode', $sheet->getCell('A1')->getValue());
        $this->assertSame('Budi', $sheet->getCell('B2')->getValue());
    }

    public function testPdfBinaryProducesAValidPdf(): void
    {
        $binary = Export::make($this->sampleRows())
            ->columns(['kode' => 'Kode', 'nama' => 'Nama', 'total' => 'Total'])
            ->toPdfBinary();

        $this->assertStringStartsWith('%PDF', $binary);
    }
}
