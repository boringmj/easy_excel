<?php

namespace Boringmj\Excel;

use Exception;
use Boringmj\Excel\Excel;

class Main {

    /**
     * 运行程序
     */
    static public function run() {
        try {
            $excel_path=dirname(__DIR__).'/cache/temp.xlsx';
            // 在data数组中添加表头
            $data=array(
                array('center','姓名','性别','年龄'),
                array('center','张三','男','20'),
                array('center','李四','男','21'),
                array('center','王五','女','22')
            );
            Excel::write($excel_path,...$data);
            Excel::bold($excel_path,1);
            // 读取 Excel 文件
            $data=Excel::read($excel_path);
            // 输出数据
            foreach ($data as $_=>$columns) {
                foreach ($columns as $_=>$value)
                    echo $value."\t";
                echo PHP_EOL;
            }
        } catch (Exception $error) {
            echo $error->getMessage();
        }
    }

}