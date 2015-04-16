<?php
error_reporting(E_ALL);
date_default_timezone_set('Europe/Kiev');

define('BASE_PATH', realpath(dirname(__FILE__) . '/..'));
define('TEST_DATA_PATH', BASE_PATH . "/tests/data");

require_once(BASE_PATH . '/vendor/autoload.php');
