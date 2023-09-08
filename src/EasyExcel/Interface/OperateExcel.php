<?php

namespace Boringmj\EasyExcel\Interface;

/**
 * 操作 Excel 基础接口
 * 
 * @package Boringmj\EasyExcel\Interface
 * @since 1.0.0
 * @version 1.0.0
 * @method bool write(array ...$data) 将数据写入 Excel 文件
 * @method bool save() 将数据保存到 Excel 文件
 * @method array read() 从 Excel 文件中读取数据
 * @method self bold(array ...$data) 设置单元格对齐方式
 * @method self ailgn(string $align,array ...$data) 设置单元格对齐方式
 */
interface OperateExcel {

    /**
     * 将数据写入 Excel 文件
     * 
     * @param string $path Excel文件路径
     * @param array ...$data 数据
     * @return bool
     */
    public function write(array ...$data):bool;
    
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
     * @param array ...$data 数据
     * @return self
     */
    public function bold(array ...$data):self;


    /**
     * 设置单元格对齐方式
     * 
     * @param string $align 对齐方式
     * @param array ...$data 数据
     * @return self
     */
    public function ailgn(string $align,array ...$data):self;

}