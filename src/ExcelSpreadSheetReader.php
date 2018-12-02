<?php
/**
 * Created by PhpStorm.
 * User: Joshua
 * Date: 2018-12-02
 * Time: 13:16
 */

namespace Joshua\Helpers;


use PhpOffice\PhpSpreadsheet\IOFactory;

/**
 * Class ExcelSpreadSheetReader
 * @package Joshua\Helpers
 */
abstract class ExcelSpreadSheetReader
{
    /**
     * @var null
     */
    protected $spreadsheet = null;
    /**
     * @var null
     */
    protected $workbook = null;
    /**
     * @var null
     */
    protected $sheet = null;

    /**
     * ExcelSpreadSheetReader constructor.
     * @param $filePath
     * @throws \PhpOffice\PhpSpreadsheet\Reader\Exception
     */
    public function __construct($filePath)
    {
        $this->loadSpreadSheet($filePath);

        $this->setCurrentSheet(0);

    }


    /**
     * @param $cell
     * @return mixed
     */
    protected function value($cell)
    {
        return $this->currentSheet()->getCell($cell)->getValue();
    }
    /**
     * @return mixed
     */
    protected function currentSheet()
    {
        return $this->sheet;
    }

    /**
     * @param int $sheet
     * @return $this
     */
    protected function setCurrentSheet($sheet = 0)
    {
        $this->sheet = $this->workbook->getSheet($sheet);

        return $this;
    }

    /**
     * @param $filePath
     * @throws \PhpOffice\PhpSpreadsheet\Reader\Exception
     */
    protected function loadSpreadSheet($filePath)
    {
        $this->spreadsheet = IOFactory::createReaderForFile($filePath);
        $this->spreadsheet->setReadDataOnly(true);
        $this->loadWorkBook($filePath);
    }

    /**
     * @param $filePath
     */
    protected function loadWorkBook($filePath)
    {
        $this->workbook = $this->spreadsheet->load($filePath);
    }

    protected function spreadsheet($factory = null)
    {
        if(is_null($factory) && is_null($this->spreadsheet)) {
            trigger_error(" There is no spreadsheet assigned");
        }

        if(is_null($factory)){
            return $this->spreadsheet;
        }

        $this->spreadsheet = $factory;

    }

}