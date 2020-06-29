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
use PhpOffice\PhpSpreadsheet\Style\Fill;
use Symfony\Component\HttpFoundation\StreamedResponse;

abstract class ExcelSpreadSheetWriter extends BaseExcelSpreadSheet
{

    /**
     * @var
     */
    protected $writer;

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
     * @param $sheet
     * @param $column
     * @param $row
     * @param $value
     * @param array $styleOverWrite
     * @return \PhpOffice\PhpSpreadsheet\Style\Style
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    protected function cellValueRowColumnWithDefaultStyle($sheet, $column, $row, $value, $styleOverWrite = [])
    {
        return $this->cellValueRowColumn($sheet, $column, $row, $value)
            ->getStyleByColumnAndRow($column, $row)->applyFromArray($this->defaultCellFormat($styleOverWrite));
    }

    /**
     * @param $sheet
     * @param $column
     * @param $row
     * @param $value
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    protected function cellValueRowColumn($sheet, $column, $row, $value)
    {
        return $this->spreadsheet()->setActiveSheetIndex($sheet)->setCellValueByColumnAndRow($column, $row, $value);
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
        return array_merge($overwrite, [ 'numberformat' => ['code' => '@'], 'alignment' => ['horizontal' => Alignment::HORIZONTAL_LEFT, 'indent' => 1], 'fill' => ['type' => Fill::FILL_SOLID, 'color' => ['rgb' => 'D0D0D0']] ]);
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
        header(sprintf('Content-Disposition: attachment;filename="%s"', $this->filename($withPath)));
        header('Cache-Control: max-age=0');
        // If you're serving to IE 9, then the following may be needed
        header('Cache-Control: max-age=1');

        // If you're serving to IE over SSL, then the following may be needed
        header('Expires: Mon, 26 Jul 1997 05:00:00 GMT'); // Date in the past
        header(sprintf('Last-Modified: %s GMT', gmdate('D, d M Y H:i:s'))); // always modified
        header('Cache-Control: cache, must-revalidate'); // HTTP/1.1
        header('Pragma: public'); // HTTP/1.0

        $this->writer();
        ob_end_clean();        // This is needed to remove junk characters
        $this->writer()->save('php://output');
        exit;         // This is needed to remove junk characters

    }
        /**
     * @return mixed
     */
    public function output()
    {
        $this->writer();
        ob_end_clean();
        return $this->writer()->save('php://output');
    }

    /**
     * @param null $filename
     * @return StreamedResponse
     */
    public function response( $filename = null)
    {
        $filename = is_null($filename) ? $this->filename(false) : $filename;
        $stream = new StreamedResponse(function () {
            $this->output();
        });
        $stream->headers->set('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        $stream->headers->set('Content-Disposition', sprintf('attachment;filename="%s"', $filename));
        $stream->headers->set('Expires', ' Mon, 26 Jul 1997 05:00:00 GMT');
        $stream->headers->set('Last-Modified', gmdate('D, d M Y H:i:s') . ' GMT');
        $stream->headers->set('Cache-Control', 'cache, must-revalidate');
        $stream->headers->set('Pragma', 'public');;

        return $stream;
    }

    /**
     * @param null $filename
     * @return StreamedResponse
     */
    public function helperHttpResponse( $filename = null)
    {
        return $this->response($filename);
    }

    /**
     * @param null $filename
     * @return $this
     */
    public function helperStoreFileLocal( $filename = null)
    {
        if(is_null($filename )) {
            $filename = $this->filename(true);
        }
        $this->writer()->save($filename);
        return $this;
    }

    /**
     * @param bool $withPath
     * @return mixed
     */
    public function helperGetFilename( $withPath = true)
    {
        return $this->filename($withPath);
    }

    /**
     * @param $sheet
     * @param $title
     * @return $this
     */
    public function helperSetSheetTitle( $sheet, $title)
    {
        $this->spreadsheet()->setActiveSheetIndex($sheet)->setTitle($title);
        return $this;
    }
}