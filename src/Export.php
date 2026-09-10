<?php

namespace Elrayn\TableExport;

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use Dompdf\Dompdf;

class Export
{
    protected array $rows;
    protected array $columns = [];
    protected ?string $title = null;

    protected function __construct(array $rows)
    {
        $this->rows = $rows;
    }

    public static function make(array $rows): self
    {
        return new self($rows);
    }

    // $columns = ['data_key' => 'Column Label', ...]
    public function columns(array $columns): self
    {
        $this->columns = $columns;
        return $this;
    }

    public function title(string $title): self
    {
        $this->title = $title;
        return $this;
    }

    public function asExcel(string $filename = 'export.xlsx'): void
    {
        // PhpSpreadsheet 1.x throws a notice on PHP 7.4+ every time a numeric
        // cell is set (DefaultValueBinder does an array offset check before
        // its is_string check). Harmless, but a stray notice here corrupts
        // the binary output since it's already being streamed. Suppress
        // display for this call only, restore right after.
        $previous = ini_set('display_errors', '0');
        $spreadsheet = $this->toSpreadsheet();
        ini_set('display_errors', $previous);

        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header('Content-Disposition: attachment;filename="' . $filename . '"');
        header('Cache-Control: max-age=0');

        (new Xlsx($spreadsheet))->save('php://output');
    }

    public function toSpreadsheet(): Spreadsheet
    {
        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();
        $row = 1;

        if ($this->title) {
            $sheet->setCellValue('A' . $row, $this->title);
            $sheet->getStyle('A' . $row)->getFont()->setBold(true)->setSize(14);
            $row += 2;
        }

        $col = 'A';
        $headerRow = $row;
        foreach ($this->columns as $label) {
            $sheet->setCellValue($col . $row, $label);
            $col++;
        }
        $lastCol = chr(ord('A') + count($this->columns) - 1);
        $sheet->getStyle("A{$headerRow}:{$lastCol}{$headerRow}")
            ->getFont()->setBold(true);
        $sheet->getStyle("A{$headerRow}:{$lastCol}{$headerRow}")
            ->getBorders()->getAllBorders()->setBorderStyle(Border::BORDER_THIN);
        $row++;

        foreach ($this->rows as $record) {
            $col = 'A';
            foreach (array_keys($this->columns) as $key) {
                $sheet->setCellValue($col . $row, $this->value($record, $key));
                $col++;
            }
            $row++;
        }

        $sheet->getStyle("A{$headerRow}:{$lastCol}" . ($row - 1))
            ->getBorders()->getAllBorders()->setBorderStyle(Border::BORDER_THIN);

        foreach (range('A', $lastCol) as $c) {
            $sheet->getColumnDimension($c)->setAutoSize(true);
        }

        return $spreadsheet;
    }

    public function asPdf(string $filename = 'export.pdf'): void
    {
        $dompdf = new Dompdf();
        $dompdf->loadHtml($this->toHtml());
        $dompdf->render();
        $dompdf->stream($filename, ['Attachment' => true]);
    }

    public function toHtml(): string
    {
        $html = '<style>table{width:100%;border-collapse:collapse;font-family:sans-serif;font-size:12px}';
        $html .= 'th,td{border:1px solid #999;padding:4px 6px;text-align:left}th{background:#eee}</style>';

        if ($this->title) {
            $html .= '<h3>' . htmlspecialchars($this->title) . '</h3>';
        }

        $html .= '<table><thead><tr>';
        foreach ($this->columns as $label) {
            $html .= '<th>' . htmlspecialchars($label) . '</th>';
        }
        $html .= '</tr></thead><tbody>';

        foreach ($this->rows as $record) {
            $html .= '<tr>';
            foreach (array_keys($this->columns) as $key) {
                $html .= '<td>' . htmlspecialchars((string) $this->value($record, $key)) . '</td>';
            }
            $html .= '</tr>';
        }
        $html .= '</tbody></table>';

        return $html;
    }

    protected function value($record, string $key)
    {
        return is_array($record) ? ($record[$key] ?? '') : ($record->$key ?? '');
    }
}
