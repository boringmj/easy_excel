<?php

namespace Boringmj\EasyExcel\Abstract;

use Exception;
use Boringmj\EasyExcel\Exception\ExcelFileException;
use Boringmj\EasyExcel\Interface\CreateExcel;
use Boringmj\EasyExcel\Interface\OperateExcel;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\Cell\Cell;

/**
 * 操作 Excel 抽象类
 * 
 * @package Boringmj\EasyExcel\Abstract
 * @since 1.0.0
 * @version 1.0.1
 * @see \Boringmj\EasyExcel\Interface\Excel
 * @property string $_excel_path Excel 文件路径
 */
abstract class Excel implements OperateExcel,CreateExcel {

    /**
     * 只读模式
     */
    const EXCEL_ONLY_READ=1;

    /**
     * 读写模式
     */
    const EXCEL_READ_WRITE=2;

    /**
     * Excel 文件路径
     * 
     * @var string
     */
    protected string $_excel_path;

    /**
     * Spreadsheet 对象
     * 
     * @var Spreadsheet
     */
    protected Spreadsheet $spreadsheet;

    /**
     * Worksheet 对象
     * 
     * @var Worksheet
     */
    protected Worksheet $worksheet;

    /**
     * 打开模式
     * 
     * @var int
     */
    protected int $open_mode;

    /**
     * 允许写入的模式
     * 
     * @var array
     */
    protected array $allow_write_mode=array(
        self::EXCEL_READ_WRITE
    );

    /**
     * 指针位置
     * 
     * @var array
     */
    protected $pointer=array(
        'row'=>1,
        'column'=>1
    );

    /**
     * 构造函数
     * 
     * @param string $excel_path Excel 文件路径
     * @param int $open_mode 打开模式(默认为读写模式,可选值为:EXCEL_ONLY_READ,EXCEL_READ_WRITE)
     * @throws ExcelFileException
     */
    public function __construct(protected string $excel_path,int $open_mode=self::EXCEL_READ_WRITE) {
        $this->_excel_path=$excel_path;
        $this->open_mode=$open_mode;
        $this->load();
    }

    /**
     * 加载 Excel 文件
     * 
     * @return self
     * @throws ExcelFileException
     */
    protected function load():self {
        // 判断 Excel 文件是否存在,不存在则创建
        if (!file_exists($this->_excel_path)) {
            // 判断是否允许写入
            if (!in_array($this->open_mode,$this->allow_write_mode))
                throw new ExcelFileException($this->_excel_path,ExcelFileException::EXCEL_FILE_ONLY_READ);
            $this->create($this->_excel_path);
        }
        else {
            // 判断 Excel 文件是否可读
            if (!is_readable($this->_excel_path))
                throw new ExcelFileException($this->_excel_path,ExcelFileException::EXCEL_FILE_NOT_READABLE);
            // 判断是否需要检查 Excel 文件是否可写
            if (in_array($this->open_mode,$this->allow_write_mode))
                // 判断 Excel 文件是否可写
                if (!is_writable($this->_excel_path))
                    throw new ExcelFileException($this->_excel_path,ExcelFileException::EXCEL_FILE_NOT_WRITABLE);
            $this->spreadsheet=IOFactory::load($this->_excel_path);
            $this->worksheet=$this->spreadsheet->getActiveSheet();
        }
        $this->reloadPointer();
        return $this;
    }

    /**
     * 创建 Excel 文件(创建新文件会导致当前加载的对象被覆盖)
     * 
     * @param string $excel_path Excel 文件路径
     * @return bool
     * @throws ExcelFileException
     */
    public function create(string $excel_path):bool {
        // 判断打开模式是否允许写入
        if (!in_array($this->open_mode,$this->allow_write_mode))
            throw new ExcelFileException($excel_path,ExcelFileException::EXCEL_FILE_ONLY_READ);
        // 判断路径的上级目录是否存在且可写
        $dir=dirname($excel_path);
        if (!is_dir($dir))
            // 这里使用 “@” 符号来抑制错误,如果创建失败会在下面的判断中抛出异常
            @mkdir($dir,0755,true);
        if (!is_writable($dir))
            throw new ExcelFileException($excel_path,ExcelFileException::EXCEL_FILE_NOT_WRITABLE);
        $this->spreadsheet=new Spreadsheet();
        $this->worksheet=$this->spreadsheet->getActiveSheet();
        $writer=new Xlsx($this->spreadsheet);
        $writer->save($excel_path);
        $this->reloadPointer();
        return true;
    }

    /**
     * 保存 Excel 文件
     * 
     * @return bool
     * @throws ExcelFileException
     */
    public function save():bool {
        try {
            // 判断是否允许写入
            if (in_array($this->open_mode,$this->allow_write_mode)) {
                // 尝试以可读模式打开 Excel 文件, 如果打开失败则抛出异常
                @$excel_file=fopen($this->_excel_path,'r+');
                if ($excel_file==false)
                    throw new ExcelFileException($this->_excel_path,ExcelFileException::EXCEL_FILE_LOCKED);
                fclose($excel_file);
                $writer=new Xlsx($this->spreadsheet);
                $writer->save($this->_excel_path);
                $this->reloadPointer();
                return true;
            }
            return false;
        } catch (Exception $error) {
            // 判断错误类是否为 ExcelFileException,如果是则抛出,否则抛出 ExcelFileException
            if ($error instanceof ExcelFileException)
                throw $error;
            throw new ExcelFileException($this->_excel_path,ExcelFileException::EXCEL_FILE_NOT_WRITABLE);
        }
    }

    /**
     * 通过字符串获取列索引
     * 
     * @param string $column_address 列地址
     * @return string
     */
    final static public function columnIndexFromString(string $column_address):string {
        return Coordinate::columnIndexFromString($column_address);
    }

    /**
     * 通过列索引获取字符串
     * 
     * @param int $column_letter 列索引
     * @return string
     */
    final static public function stringFromColumnIndex(int $column_letter):string {
        return Coordinate::stringFromColumnIndex($column_letter);
    }

    /**
     * 提供行列号设置某个单元格的值
     * 
     * @param int $row 行
     * @param int $column 列
     * @param mixed $value 值
     * @return void
     */
    final protected function setCellValueByRowColumn(int $row,int $column,mixed $value):void {
        $cellAddress=self::stringFromColumnIndex($column).$row;
        $this->worksheet->setCellValue($cellAddress,$value);
    }

    /**
     * 设置某个单元格的值
     * 
     * @param string $cell_address 单元格地址
     * @param mixed $value 值
     */
    final protected function setCellValue(string $cell_address,mixed $value):void {
        $this->worksheet->setCellValue($cell_address,$value);
    }

    /**
     * 重新计算指针位置
     * 
     * @return void
     */
    final protected function reloadPointer():void {
        $last_row=$this->getLastRow();
        // 如果最后一行等于1,则判断第一行是否有数据,如果有数据则指针移动到第二行第一个单元格,否则指针移动到第一行第一个单元格
        if ($last_row==1) {
            $last_column=$this->getLastColumn();
            $value=$this->getCellValueByRowColumn(1,$last_column);
            $last_row=$value==null?0:1;
        }
        $this->pointer['row']=$last_row+1;
        $this->pointer['column']=1;
    }

    /**
     * 获取指针位置
     * 
     * @param string $key 指针键名(默认为全部,可选值为:row,column)
     * @return array|int
     */
    final public function getPointer(string $key=null):array|int {
        if ($key==null)
            return $this->pointer;
        return $this->pointer[$key];
    }

    /**
     * 设置指针位置
     * 
     * @param int $row 行
     * @param int $column 列
     * @return self
     */
    final public function setPointer(int $row,int $column=1):self {
        $this->pointer['row']=$row>0?$row:1;
        $this->pointer['column']=$column>0?$column:1;
        return $this;
    }

    /**
     * 获取 Excel 最后一行的行号
     * 
     * @return int
     */
    final public function getLastRow():int {
        return $this->worksheet->getHighestRow();
    }

    /**
     * 获取 Excel 最后一列的列号
     * 
     * @return int
     */
    final public function getLastColumn():int {
        $last_column=$this->worksheet->getHighestColumn();
        return self::columnIndexFromString($last_column);
    }

    /**
     * 获取某个单元格
     * 
     * @param string $cell_address 单元格地址
     * @return Cell
     */
    final public function getCell(string $cell_address):Cell {
        return $this->worksheet->getCell($cell_address);
    }

    /**
     * 获取某个单元格的值
     * 
     * @param string $cell_address 单元格地址
     * @return mixed
     */
    final public function getCellValue(string $cell_address):mixed {
        return $this->getCell($cell_address)->getValue();
    }

    /**
     * 通过行列号获取某个单元格的值
     * 
     * @param int $row 行
     * @param int $column 列
     * @return mixed
     */
    final public function getCellValueByRowColumn(int $row,int $column):mixed {
        $cellAddress=self::stringFromColumnIndex($column).$row;
        return $this->getCellValue($cellAddress);
    }

}