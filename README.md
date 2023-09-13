# Easy_Excel
## 写在前面
本项目只是基于 [phpoffice/phpspreadsheet](https://packagist.org/packages/phpoffice/phpspreadsheet) 的二次封装, 如果您需要更全面的 Excel 文件处理库, 可以前往 [phpoffice/phpspreadsheet](https://packagist.org/packages/phpoffice/phpspreadsheet) 查看详情
## 该如何开始?
在开始之前,您应该需要注意\
本项目使用了 `PHP 8.0.0` 的语法结构, 所以您需要保证您的 PHP 版本不低于 `8.0.0`\
如果您在开发中有新增文件, 请确保您已经使用 [composer](https://www.phpcomposer.com/) 更新 `composer update`, 新增的文件如果没有及时更新, 新增的程序将无法通过 autoload 自动加载, 常见的错误为 `class not found`


1. 您应该先下载或 `clone`(推荐) 本项目至本地
```
git clone https://github.com/boringmj/easy_excel.git
```
2. 通过 [composer](https://www.phpcomposer.com/) 安装依赖, 如果您没有下载 [composer](https://www.phpcomposer.com/), 您可以前往 [composer 官网](https://www.phpcomposer.com/) 获取帮助
```
// 前往项目路径
cd easy_excel
// composer 安装依赖
composer install
```
3. 启动 [php webserver](https://www.php.net/manual/zh/features.commandline.webserver.php), 我们提供了简单的快捷启动脚本
```
// 需要配置php环境变量且php>=5.4.0
php start
```
4. 访问 `localhost:8000`, 至此,您已经可以正常进行开发了

## 如何使用?
本项目在 `/src` 目录下提供了默认的 `Main.php` 文件, 这是项目的默认入口文件, 默认的php入口文件在 `/public` 目录下的 `index.php`\
请注意, 在使用前请先导入核心类 `use Boringmj\EasyExcel\Excel;`, 如果出现 `class not found` 错误, 请检查您的 `composer` 是否已经安装依赖, `composer install`\
你可以打开 `Main.php`, 然后在`\Boringmj:Main::run()`方法里编辑您的代码, 例如:
```
/**
 * 运行程序
 */
static public function run() {
    $excel_path=dirname(__DIR__).'/cache/temp1.xlsx';
    $Excel=new Excel($excel_path,Excel::EXCEL_READ_WRITE);
    $Excel->write(1,"你好",True)->save();
}
```
更多代码请参考 `/src/Main.php`

## 注意
本项目还处于开发阶段, 迭代速度较快, 请谨慎使用, 请勿用于生产环境, 本项目的所有代码均在 `PHP 8.0.0` 环境下测试通过, 请确保您的PHP版本不低于 `8.0.0`
