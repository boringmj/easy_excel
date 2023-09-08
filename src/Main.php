<?php

namespace Boringmj;

use Exception;
use Boringmj\EasyExcel\Excel;

class Main {

    /**
     * 运行程序
     */
    static public function run() {
        try {
            // 禁用错误报告
            // error_reporting(0);
            $excel_path=dirname(__DIR__).'/cache/temp.xlsx';
            $Excel=new Excel($excel_path,Excel::EXCEL_READ_WRITE);
            print_r($Excel->read());
        } catch (Exception $error) {
            echo  $error->getMessage().PHP_EOL;
        }
    }

}