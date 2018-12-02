<?php
/**
 * Created by PhpStorm.
 * User: Joshua
 * Date: 2018-12-02
 * Time: 12:51
 */

namespace Joshua\Helpers;


use Carbon\Carbon;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Alignment;

abstract class ExcelSpreadSheetWriter
{

    /**
     * @var null
     */
    protected $spreadsheet = null;
    /**
     * @var
     */
    protected $writer;
    /**
     * @var
     */
    protected $filename;

    /**
     * @param $sheet
     * @param $cell
     * @param $value
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    protected function cellValue($sheet, $cell, $value)
    {
        return $this->spreadsheet()->setActiveSheetIndex($sheet)->setCellValue($cell, $value);
    }

    /**
     * @param $sheet
     * @param $cell
     * @param $value
     * @param array $styleOverWrite
     * @return \PhpOffice\PhpSpreadsheet\Style\Style
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    protected function cellValueDefaultStyle($sheet, $cell, $value, $styleOverWrite = [])
    {
        return $this->cellValue($sheet, $cell, $value)->getStyle($cell)->applyFromArray($this->defaultCellFormat($styleOverWrite));
    }

    /**
     * @return Spreadsheet
     */
    protected function spreadsheet()
    {
        if (is_null($this->spreadsheet)) {
            $this->spreadsheet = new Spreadsheet();
        }
        return $this->spreadsheet;
    }

    /**
     * @param string $type
     * @return \PhpOffice\PhpSpreadsheet\Writer\IWriter
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    protected function writer($type = "Xlsx")
    {
        if (is_null($this->writer)) {
            $this->writer = IOFactory::createWriter($this->spreadsheet(), $type);
        }

        return $this->writer;
    }

    /**
     * @return mixed
     */
    protected abstract function path();

    /**
     * @param bool $withPath
     * @return mixed
     */
    protected abstract function filename($withPath = true);

    /**
     * @param int $mode
     */
    protected function makeDirectories($mode = 0777)
    {
        if (!file_exists($this->path())) {
            mkdir($this->path(), $mode, true);
        }
    }

    /**
     * @param string $format
     * @return string
     */
    protected function dateFormat($format = "YmdHis")
    {
        return Carbon::now()->format($format);
    }

    /**
     * @param array $overwrite
     * @return array
     */
    protected function defaultCellFormat($overwrite = [])
    {
        return array_merge($overwrite, ['numberformat' => ['code' => '@'], 'alignment' => ['horizontal' => Alignment::HORIZONTAL_LEFT, 'indent' => 1], 'fill' => ['type' => Fill::FILL_SOLID, 'color' => ['rgb' => 'D0D0D0']]]);
    }

    /**
     * @param string $contentType
     * @param bool $withPath
     * @return void
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    protected function outputToHttp($contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", $withPath = false)
    {
        // Redirect output to a client’s web browser (Xlsx)
        header('Content-Type:' . $contentType);
        header('Content-Disposition: attachment;filename="'. $this->filename($withPath) . '"');
        header('Cache-Control: max-age=0');
        // If you're serving to IE 9, then the following may be needed
        header('Cache-Control: max-age=1');

        // If you're serving to IE over SSL, then the following may be needed
        header('Expires: Mon, 26 Jul 1997 05:00:00 GMT'); // Date in the past
        header('Last-Modified: ' . gmdate('D, d M Y H:i:s') . ' GMT'); // always modified
        header('Cache-Control: cache, must-revalidate'); // HTTP/1.1
        header('Pragma: public'); // HTTP/1.0

        $this->writer();
        ob_end_clean();        // This is needed to remove junk characters
        $this->writer()->save('php://output');
        exit;         // This is needed to remove junk characters

    }
}