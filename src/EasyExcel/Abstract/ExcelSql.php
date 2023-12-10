<?php

namespace Boringmj\EasyExcel\Abstract;

use ReflectionClass;
use Boringmj\EasyExcel\Abstract\Excel;
use Boringmj\EasyExcel\Interface\DatabaseExcel;

/**
 * Excel 伪数据库抽象类
 * 
 * @package Boringmj\EasyExcel\Abstract
 * @since 1.0.0
 * @version 1.0.0
 * @see \Boringmj\EasyExcel\Interface\DatabaseExcel
 */
abstract class ExcelSql implements DatabaseExcel {

    /**
     * Excel 对象
     * 
     * @var Excel
     */
    protected Excel $_excel;

    /**
     * 字段名称所在行
     * 
     * @var int
     */
    protected int $_field_row=1;

    /**
     * 构造函数
     * 
     * @param Excel $excel Excel 对象
     */
    public function __construct(Excel $excel) {
        $this->_excel=$excel;
    }

    /**
     * 设置字段名称所在行
     * 
     * @param int $row 字段名称所在行
     * @return self
     */
    public function fieldRow(int $row=1):self {
        $this->_field_row=$row;
        return $this;
    }

    /**
     * 获取字段名称所在行
     * 
     * @return int
     */
    public function getFieldRow():int {
        return $this->_field_row;
    }

    /**
     * 返回 Excel 对象
     * 
     * @return Excel
     */
    public function getExcel():Excel {
        return $this->_excel;
    }

}