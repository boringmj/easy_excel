<?php

// copilot 帮我引入composer的autoload
require_once __DIR__.'/vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;

// 读取Excel文件
$spreadsheet=IOFactory::load('cache/test.xlsx');

// 获取第一个工作表
$worksheet=$spreadsheet->getActiveSheet();

// 获取所有数据
$lastRow=$worksheet->getHighestRow();
$lastColumn=$worksheet->getHighestColumn();
print("最后一行：{$lastRow}，最后一列：{$lastColumn}\n");

$lastColumnIndex=Coordinate::columnIndexFromString($lastColumn);
$data=[];
// 遍历单元格范围并输出值
for ($row=1;$row<=$lastRow;$row++) {
    for ($column=1;$column<=$lastColumnIndex;$column++) {
        $cellAddress=Coordinate::stringFromColumnIndex($column).$row;
        $cell=$worksheet->getCell($cellAddress);
        echo $cell->getValue()."\n";
        $data[$row][$column]=$cell->getValue();
    }
}

print_r($data);

// 创建一个新的电子表格对象
$spreadsheet=new Spreadsheet();

// 选择活动工作表
$sheet=$spreadsheet->getActiveSheet();

// 把data写入到表格中
foreach ($data as $row=>$columns) {
    foreach ($columns as $column=>$value) {
        $cellAddress=Coordinate::stringFromColumnIndex($column).$row;
        $sheet->setCellValue($cellAddress, $value);
    }
}

// 保存为 Excel 文件
$writer=new Xlsx($spreadsheet);
$writer->save('cache/example.xlsx');

echo "Excel 文件已创建成功！";