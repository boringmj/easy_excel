<?php

namespace Boringmj\EasyExcel;

use Boringmj\EasyExcel\Abstract\Excel as AbstractExcel;

/**
 * Excel 类
 * 
 * @package Boringmj\EasyExcel
 * @since 1.0.0
 * @version 1.0.0
 */
class Excel extends AbstractExcel {

    /**
     * 向 Excel 文件写入数据
     * 
     * @param array ...$data 要写入的数据
     * @return bool
     */
    public function write(array ...$data):bool {
        return true;
    }

    /**
     * 读取 Excel 文件全部内容
     * 
     * @return array
     */
    public function read():array {
        $lastRow=$this->worksheet->getHighestRow();
        $lastColumn=$this->worksheet->getHighestColumn();
        $lastColumnIndex=self::columnIndexFromString($lastColumn);
        $data=array();
        for ($row=1;$row<=$lastRow;$row++)
            for ($column=1;$column<=$lastColumnIndex;$column++) {
                $cellAddress=self::stringFromColumnIndex($column).$row;
                $cell=$this->worksheet->getCell($cellAddress);
                $value=$cell->getValue();
                $data[$row][$column]=$value;
            }
        return $data;
    }

    /**
     * 设置单元格加粗
     * 
     * @param array ...$data 数据
     * @return self
     */
    public function bold(array ...$data):self {
        return $this;
    }

    /**
     * 设置单元格对齐方式
     * 
     * @param string $align 对齐方式
     * @param array ...$data 数据
     * @return self
     */
    public function ailgn(string $align,array ...$data):self {
        return $this;
    }

}