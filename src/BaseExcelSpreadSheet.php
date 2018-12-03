<?php
/**
 * Created by PhpStorm.
 * User: Joshua
 * Date: 2018-12-03
 * Time: 18:01
 */

namespace Joshua\Helpers;


abstract class BaseExcelSpreadSheet
{
    /**
     * @var
     */
    protected $filename;

    /**
     * @return mixed
     */
    protected abstract function path();

    /**
     * @param bool $withPath
     * @return mixed
     */
    protected abstract function filename($withPath = true);
}