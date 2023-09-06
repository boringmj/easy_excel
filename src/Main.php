<?php

namespace Boringmj\Excel;

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;

class Main {
    static public function run() {
        if (!is_file('cache/test.xlsx')) {
            throw new \Exception(' cache/test.xlsx 文件不存在！');
        }
        $spreadsheet=IOFactory::load('cache/test.xlsx');
        $worksheet=$spreadsheet->getActiveSheet();
        $lastRow=$worksheet->getHighestRow();
        $lastColumn=$worksheet->getHighestColumn();
        print("最后一行：{$lastRow}，最后一列：{$lastColumn}\n");
        $lastColumnIndex=Coordinate::columnIndexFromString($lastColumn);
        $data=[];
        for ($row=1;$row<=$lastRow;$row++) {
            echo "| ";
            for ($column=1;$column<=$lastColumnIndex;$column++) {
                $cellAddress=Coordinate::stringFromColumnIndex($column).$row;
                $cell=$worksheet->getCell($cellAddress);
                $value=$cell->getValue();
                echo "{$cellAddress} -> {$value} | ";
                $data[$row][$column]=$value;
            }
            echo "\n\r";
        }
        $spreadsheet=new Spreadsheet();
        $sheet=$spreadsheet->getActiveSheet();
        foreach ($data as $row=>$columns) {
            foreach ($columns as $column=>$value) {
                $cellAddress=Coordinate::stringFromColumnIndex($column).$row;
                $sheet->setCellValue($cellAddress, $value);
            }
        }
        $writer=new Xlsx($spreadsheet);
        $writer->save('cache/example.xlsx');
        echo "Excel 文件已创建成功！";
    }
}