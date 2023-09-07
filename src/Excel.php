<?php

namespace Boringmj\Excel;

use  Exception;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;

/**
 * Excel 文件操作类
 * 
 * @package Boringmj\Excel
 * @version 1.0.0
 * @since 1.0.0
 */
class Excel {

    /**
     * 将数据写入 Excel 文件
     * 
     * @param string $path Excel文件路径
     * @param array ...$data 数据
     * @return void
     */
    static public function write(string $path,array ...$data) {
        $spreadsheet=new Spreadsheet();
        $sheet=$spreadsheet->getActiveSheet();
        foreach ($data as $row=>$columns) {
            foreach ($columns as $column=>$value) {
                $cellAddress=Coordinate::stringFromColumnIndex($column).($row+1);
                $sheet->setCellValue($cellAddress, $value);
            }
        }
        $writer=new Xlsx($spreadsheet);
        $writer->save($path);
    }

    /**
     * 从 Excel 文件中读取数据
     * 
     * @param string $path Excel文件路径
     * @return array
     * @throws Exception
     */
    static public function read(string $path):array {
        if (!is_file($path)) {
            throw new Exception("{$path} 文件不存在！");
        }
        $spreadsheet=IOFactory::load($path);
        $worksheet=$spreadsheet->getActiveSheet();
        $lastRow=$worksheet->getHighestRow();
        $lastColumn=$worksheet->getHighestColumn();
        $lastColumnIndex=Coordinate::columnIndexFromString($lastColumn);
        $data=[];
        for ($row=1;$row<=$lastRow;$row++) {
            for ($column=1;$column<=$lastColumnIndex;$column++) {
                $cellAddress=Coordinate::stringFromColumnIndex($column).$row;
                $cell=$worksheet->getCell($cellAddress);
                $value=$cell->getValue();
                $data[$row][$column]=$value;
            }
        }
        return $data;
    }
}