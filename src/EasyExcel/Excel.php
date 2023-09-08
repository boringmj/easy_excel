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
     * @param mixed ...$data 要写入的数据
     * @return self
     */
    public function write(mixed ...$data):self {
        // 判断传入的参数是什么形式的
        foreach ($data as $value) {
            if (is_array($value)) {
                // 如果是数组
                foreach ($value as $_key=>$_value) {
                    if (is_string($_key))
                        // 如果是字符串,则在指定的单元格写入
                        $this->setCellValue($_key,$_value);
                    else {
                        // 判断值是不是数组
                        if (is_array($_value))
                            foreach ($_value as $_) {
                                $this->setCellValueByRowColumn($this->pointer['row'],$this->pointer['column'],$_);
                                $this->pointer['column']++;
                            }
                        else
                            // 如果是字符串,则在当前指针位置写入然后指针移动到下一行第一个单元格
                            $this->setCellValueByRowColumn($this->pointer['row'],$this->pointer['column'],$_value);
                    }
                    $this->reloadPointer();
                }
            } else
                // 如果是字符串,则在当前指针位置写入然后指针移动到下一行第一个单元格
                $this->setCellValueByRowColumn($this->pointer['row'],$this->pointer['column'],$value);
            $this->reloadPointer();
        }
        return $this;
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