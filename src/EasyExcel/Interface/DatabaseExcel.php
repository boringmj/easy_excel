<?php

namespace Boringmj\EasyExcel\Interface;

/**
 * Excel 伪数据库基础接口
 * 
 * @package Boringmj\EasyExcel\Interface
 * @since 1.0.0
 * @version 1.0.1
 * @method self where(array $where) 设置where条件
 * @method array select(mixed ...$fields) 查询多条数据
 * @method array find(mixed ...$fields) 查询一条数据
 * @method self field(array $fields) 设置需要查询的字段
 * @method self insert(mixed ...$data) 插入数据
 * @method self update(mixed ...$data) 更新数据
 * @method self delete(mixed ...$data) 删除数据
 * @method self fieldRow(int $row) 指定字段名称所在行
 */
interface DatabaseExcel {

    /**
     * 指定字段名称所在行
     * 
     * @param int $row 行数
     * @return self
     */
    public function fieldRow(int $row=1):self;

    /**
     * 设置where条件
     * 
     * @param array $where 条件
     * @return self
     */
    public function where(array $where):self;

    /**
     * 查询多条数据
     * 
     * @param mixed ...$fields 要查询的字段
     * @return array
     */
    public function select(mixed ...$fields):array;

    /**
     * 查询一条数据
     * 
     * @param mixed ...$fields 要查询的字段
     * @return array
     */
    public function find(mixed ...$fields):array;

    /**
     * 设置需要查询的字段
     * 
     * @param array $fields 字段
     * @return self
     */
    public function field(array $fields):self;

    /**
     * 插入数据
     * 
     * @param mixed ...$data 要插入的数据
     * @return self
     */
    public function insert(mixed ...$data):self;

    /**
     * 更新数据
     * 
     * @param mixed ...$data 要更新的数据
     * @return self
     */
    public function update(mixed ...$data):self;

    /**
     * 删除数据
     * 
     * @param mixed ...$data 要删除的数据
     * @return self
     */
    public function delete(mixed ...$data):self;

}