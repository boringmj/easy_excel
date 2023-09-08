<?php

namespace Boringmj\EasyExcel\Exception;

use Exception as BaseException;
use ReflectionFunction;

/**
 * 基本异常类
 * 
 * @package Boringmj\EasyExcel\Exception
 * @since 1.0.0
 * @version 1.0.0
 * @see \Exception
 */
class Exception extends BaseException {

    protected string|int $error_code; // 错误代码
    protected mixed $callback; // 回调函数
    protected array $args; // 回调函数的参数

    /**
     * 构造函数
     * 
     * @param string $message 错误信息
     * @param string|int $code 错误代码(如果是整数则会传递给父类构造函数,否则传递0),可以是字符串,通过`self::getCodeString()`获取
     * @param callable $callback 回调函数(如果参数中含有`exception`则会传递当前错误对象,参数允许定义合理的默认值)
     * @param mixed ...$args 回调函数的参数
     */
    public function __construct(string $message='',string|int $code=0,?callable $callback=null,mixed ...$args) {
        $this->message=$message;
        $this->error_code=$code;
        $this->callback=$callback;
        $this->args=$args;
        parent::__construct($message,is_int($code)?$code:0);
        $this->_run_callback();
    }

    /**
     * 返回字符串形式的错误信息
     * 
     * @return string
     */
    public function __toString():string {
        return __CLASS__.": [{$this->error_code}]: {$this->message} in {$this->file}({$this->line})\nStack trace:\n{$this->getTraceAsString()}";
    }

    /**
     * 返回回调函数
     * 
     * @return callable
     */
    public function getCallback():callable {
        return $this->callback;
    }

    /**
     * 返回回调函数的参数
     * 
     * @return array
     */
    public function getArgs():array {
        return $this->args;
    }

    /**
     * 获取字符串形式的错误代码
     * 
     * @return string
     */
    public function getCodeString():string {
        return (string)$this->error_code;
    }

    /**
     * 运行回调函数
     * 
     * @return void
     */
    private function _run_callback():void {
        if (is_callable($this->callback)) {
            // 通过反射获取回调函数的参数
            $reflection=new ReflectionFunction($this->callback);
            $parameters=$reflection->getParameters();
            // 如果回调函数没有参数,则直接运行
            if (empty($parameters)) {
                call_user_func($this->callback);
            } else {
                // 如果回调函数有参数,则传递参数运行
                $args=array();
                foreach ($parameters as $parameter) {
                    $name=$parameter->getName();
                    switch ($name) {
                        case 'exception':
                            $args[$name]=$this;
                            break;
                        default:
                            $args[$name]=array_shift($this->args)??$parameter->getDefaultValue();
                    }
                }
                call_user_func_array($this->callback,$args);
            }
        }
    }

}