<?php

$autoload_file=__DIR__.'/../vendor/autoload.php';
if(!file_exists($autoload_file))
    die('Please run "composer install" first!');
require_once $autoload_file;

use Boringmj\Main;

Main::run();