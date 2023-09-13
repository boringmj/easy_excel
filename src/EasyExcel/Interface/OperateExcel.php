<?php

namespace Boringmj\EasyExcel\Interface;

/**
 * Excel 操作基础接口
 * 
 * @package Boringmj\EasyExcel\Interface
 * @since 1.0.0
 * @version 1.0.2
 * @method self write(mixed ...$data) 将数据写入 Excel 文件
 * @method bool save() 将数据保存到 Excel 文件
 * @method array read() 从 Excel 文件中读取数据
 * @method self bold(mixed ...$data) 设置单元格加粗
 * @method self ailgn(mixed ...$data) 设置单元格对齐方式
 */
interface OperateExcel {

    /**
     * 将数据写入 Excel 文件
     * 
     * @param string $path Excel文件路径
     * @param mixed ...$data 数据
     * @return self
     */
    public function write(mixed ...$data):self;
    
    /**
     * 将数据保存到 Excel 文件
     * 
     * @return bool
     */
    public function save():bool;

    /**
     * 从 Excel 文件中读取数据
     * 
     * @return array
     */
    public function read():array;

    /**
     * 设置单元格对齐方式
     * 
     * @param mixed ...$data 数据
     * @return self
     */
    public function bold(mixed ...$data):self;


    /**
     * 设置单元格对齐方式
     * 
     * @param mixed ...$data 数据
     * @return self
     */
    public function ailgn(mixed ...$data):self;

}