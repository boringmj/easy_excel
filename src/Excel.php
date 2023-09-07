<?php

namespace Boringmj\Excel;

use Exception;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
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
        $worksheet=$spreadsheet->getActiveSheet();
        $aligns=array();
        foreach ($data as $row=>$columns) {
            if(!is_array($columns))
                continue;
            foreach ($columns as $column=>$value) {
                if($column<0)
                    continue;
                if($column===0)
                    $aligns[$value][]=$row+1;
                else {
                    $cellAddress=Coordinate::stringFromColumnIndex($column).($row+1);
                    $worksheet->setCellValue($cellAddress,$value);
                }
            }
        }
        foreach ($aligns as $align=>$row)
            self::align($worksheet,$align,...$row);
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
        if (!is_file($path))
            throw new Exception("{$path} 文件不存在！");
        $spreadsheet=IOFactory::load($path);
        $worksheet=$spreadsheet->getActiveSheet();
        $lastRow=$worksheet->getHighestRow();
        $lastColumn=$worksheet->getHighestColumn();
        $lastColumnIndex=Coordinate::columnIndexFromString($lastColumn);
        $data=[];
        for ($row=1;$row<=$lastRow;$row++)
            for ($column=1;$column<=$lastColumnIndex;$column++) {
                $cellAddress=Coordinate::stringFromColumnIndex($column).$row;
                $cell=$worksheet->getCell($cellAddress);
                $value=$cell->getValue();
                $data[$row][$column]=$value;
            }
        return $data;
    }

    /**
     * 将指定行设置为粗体
     * 
     * @param string $path Excel文件路径
     * @param int ...$row 行号
     * @return void
     * @throws Exception
     */
    static public function bold(string $path,int ...$row) {
        if (!is_file($path))
            throw new Exception("{$path} 文件不存在！");
        $spreadsheet=IOFactory::load($path);
        $worksheet=$spreadsheet->getActiveSheet();
        foreach ($row as $r) {
            if ($r<1)
                continue;
            $rowStyle=$worksheet->getStyle('A'.$r.':'.$worksheet->getHighestColumn().$r);
            $font=$rowStyle->getFont();
            $font->setBold(true);
        }
        $writer=new Xlsx($spreadsheet);
        $writer->save($path);
    }

    /**
     * 为指定行设置对齐方式
     * 
     * @param string|Worksheet $excel Excel文件路径
     * @param string $align 对齐方式(center,left,right)
     * @param int ...$row 行号
     * @return void
     * @throws Exception
     */
    static public function align(string|Worksheet $excel,string $align,int ...$row) {
        if ($excel instanceof Worksheet) {
            self::alignByObject($excel,$align,...$row);
        } else {
            if (!is_file($excel))
                throw new Exception("{$excel} 文件不存在！");
            $spreadsheet=IOFactory::load($excel);
            $worksheet=$spreadsheet->getActiveSheet();
            self::alignByObject($worksheet,$align,...$row);
            $writer=new Xlsx($spreadsheet);
            $writer->save($excel);
        }
    }

    /**
     * 通过对象设置对齐方式
     * 
     * @param Worksheet $worksheet
     * @param string $align 对齐方式(center,left,right)
     * @param int ...$row 行号
     * @return void
     * @throws Exception
     */
    static public function alignByObject(Worksheet $worksheet,string $align,int ...$row) {
        if (!in_array($align,array('center','left','right')))
            throw new Exception("{$align} 不是有效的对齐方式！");
        foreach ($row as $r) {
            if ($r<1)
                continue;
            $rowStyle=$worksheet->getStyle('A'.$r.':'.$worksheet->getHighestColumn().$r);
            $rowStyle->getAlignment()->setHorizontal($align);
        }
    }

}