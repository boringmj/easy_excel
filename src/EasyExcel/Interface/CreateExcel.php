<?php

namespace Boringmj\EasyExcel\Interface;

/**
 * 操作 Excel 基础接口
 * 
 * @package Boringmj\EasyExcel\Interface
 * @since 1.0.0
 * @version 1.0.0
 * @method bool create(string $path) 创建 Excel 文件
 */
interface CreateExcel {

    /**
     * 创建 Excel 文件
     * 
     * @param string $path Excel文件路径
     * @param array ...$data 数据
     * @return bool
     */
    public function create(string $path):bool;

}