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
            $excel_path=dirname(__DIR__).'/cache/temp1.xlsx';
            $Excel=new Excel($excel_path,Excel::EXCEL_READ_WRITE);
            // 在指针位置写入数据并保存(规则为在表格最后一行的下一行第一个单元格写入,写入后指针移动到下一行第一个单元格)
            $Excel->write(1,"你好",True)->save();
            // 上面的代码的语法与下面的代码的语法是一样的(注意,传入参数不同)
            $Excel->write(['a',false,3])->save();
            // 重新加载一个新的 Excel 文件
            $Excel->load(dirname(__DIR__).'/cache/temp2.xlsx',Excel::EXCEL_READ_WRITE);
            // 在A6和G10写入数据并保存
            $Excel->write([
                'A6'=>'A6',
                'G10'=>'G10'
            ])->bold('A6')->save();
            // 按行写入数据并保存
            $Excel->write([
                ['a','b','c'],
                ['d','e','f']
            ])->bold(11,12)->save();
            // 下面是一个复杂的例子
            // $Excel->write('A1',array(
            //     'A2'=>'A2',
            //     'A3'
            // ),array(
            //     ['A4','B4'],
            //     'A5',
            //     'A6'=>'A6'
            // ))->save();
            print_r($Excel->read());
            // 如果需要返回空单元格,则使用下面的方式设置(注意,设置后只在读取时有效,读取完毕后会自动设置为 false)
            // Excel::returnNullCell(bool $return_null_cell=true);
            // 例如:
            // print_r($Excel->returnNullCell()->read());

            // 下面是一个切换Excel文件的例子
            // $Excel->load(dirname(__DIR__).'/cache/temp2.xlsx',Excel::EXCEL_READ_WRITE);

            // 新的Excel类使用了众多机制,例如:
            // 初始化Excel类时,如果不传入文件路径,则会自动进入虚拟模式
            // 虚拟模式下,保存文件时,如果不传入文件路径,则会抛出异常
            // load方法也可以加载一个虚拟的Excel文件,只需要在路径传入空字符“ '' ”即可(空字符不等同于“null”)
            // 例如:
            // $Excel=new Excel('',Excel::EXCEL_READ_WRITE); // $Excel=new Excel(); 也是可以的
            // 加载一个虚拟的Excel文件
            // $Excel->load('',Excel::EXCEL_READ_WRITE);
        } catch (Exception $error) {
            echo  $error->getMessage().PHP_EOL;
        }
    }

}