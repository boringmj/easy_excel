<?php

namespace Boringmj\EasyExcel\Exception;

/**
 * Excel 文件异常
 * 
 * @package Boringmj\EasyExcel\Exception
 * @since 1.0.0
 * @version 1.0.0
 * @see \Boringmj\EasyExcel\Exception\Exception
 */
class ExcelFileException extends Exception {

    private string $_excel_path; // Excel 文件路径

    /**
     * Excel 文件异常代码
     */
    const EXCEL_FILE_ERROR=10100;

    /**
     * Excel 文件不存在
     */
    const EXCEL_FILE_NOT_FOUND=10101;

    /**
     * Excel 文件不可读
     */
    const EXCEL_FILE_NOT_READABLE=10102;

    /**
     * Excel 文件不可写
     */
    const EXCEL_FILE_NOT_WRITABLE=10103;

    /**
     * Excel 文件以只读模式打开
     */
    const EXCEL_FILE_ONLY_READ=10104;

    /**
     * Excel 文件被其他程序占用
     */
    const EXCEL_FILE_LOCKED=10105;

    /**
     * Excel 文件异常代码对应的错误信息
     */
    public static array $error_message=[
        self::EXCEL_FILE_ERROR=>'Excel 文件错误',
        self::EXCEL_FILE_NOT_FOUND=>'Excel 文件不存在',
        self::EXCEL_FILE_NOT_READABLE=>'Excel 文件不可读',
        self::EXCEL_FILE_NOT_WRITABLE=>'Excel 文件不可写',
        self::EXCEL_FILE_ONLY_READ=>'Excel 以只读模式打开不可创建或写入',
        self::EXCEL_FILE_LOCKED=>'Excel 文件被其他程序占用'
    ];

    /**
     * 构造函数
     * 
     * @param string $excel_path Excel 文件路径
     * @param int $code 错误代码
     * @param callable $callback 回调函数
     * @param mixed ...$args 回调函数的参数
     */
    public function __construct(string $excel_path=null,int $code=0,?callable $callback=null,mixed ...$args) {
        $this->_excel_path=$excel_path;
        if (!isset(self::$error_message[$code]))
            $code=self::EXCEL_FILE_ERROR;
        $message=$excel_path?self::$error_message[$code].': '.$excel_path:self::$error_message[$code];
        parent::__construct($message,$code,$callback,...$args);
    }

    /**
     * 返回 Excel 文件路径
     * 
     * @return string
     */
    public function getExcelPath():string {
        return $this->_excel_path;
    }

}