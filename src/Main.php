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
            // 在指针位置写入数据并保存(规则为在表格最后一行的下一行第一个单元格写入,写入后指针移动到下一行第一个单元格)
            $Excel->write(1,"你好",True)->save();
            // 上面的代码的语法与下面的代码的语法是一样的(注意,传入参数不同)
            $Excel->write(['a',false,3])->save();
            // 在A6和G10写入数据并保存
            $Excel->write([
                'A6'=>'A6',
                'G10'=>'G10'
            ])->save();
            // 按行写入数据并保存
            $Excel->write([
                ['a','b','c'],
                ['d','e','f']
            ])->save();
            print_r($Excel->read());
        } catch (Exception $error) {
            echo  $error->getMessage().PHP_EOL;
        }
    }

}