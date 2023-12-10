<?php

namespace Boringmj\EasyExcel;

use Boringmj\EasyExcel\Abstract\ExcelSql as AbstractExcelSql;

/**
 * Excel 类
 * 
 * @package Boringmj\EasyExcel
 * @since 1.0.0
 * @version 1.0.0
 */
class ExcelSql extends AbstractExcelSql {

    /**
     * where 条件
     * 
     * @var array
     */
    private array $where=array();

    public function where(array $where):self {
        return $this;
    }

    public function select(mixed ...$fields):array {
        return [];
    }

    public function find(mixed ...$fields):array {
        return [];
    }

    public function field(array $fields):self {
        return $this;
    }

    public function insert(mixed ...$data):self {
        return $this;
    }

    public function update(mixed ...$data):self {
        return $this;
    }

    public function delete(mixed ...$data):self {
        return $this;
    }

}