<?php

namespace Boringmj\Excel;

use Boringmj\Excel\Excel;

class Main {

    /**
     * 运行程序
     */
    static public function run() {
        // Excel 文件路径
        $excel_path=dirname(__DIR__).'/cache/test.xlsx';
        // 读取 Excel 文件
        $data=Excel::read($excel_path);
        // 输出数据
        foreach ($data as $_=>$columns) {
            foreach ($columns as $_=>$value)
                echo $value."\t";
            echo PHP_EOL;
        }
        // 写入 Excel 文件
        $excel_path=dirname(__DIR__).'/cache/temp.xlsx';
        Excel::write($excel_path,...$data);
    }

}