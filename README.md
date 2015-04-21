# Xls Writer

[![Build Status](https://api.travis-ci.org/MAXakaWIZARD/xls-writer.png?branch=master)](https://travis-ci.org/MAXakaWIZARD/xls-writer) 
[![Scrutinizer Code Quality](https://scrutinizer-ci.com/g/MAXakaWIZARD/xls-writer/badges/quality-score.png?b=master)](https://scrutinizer-ci.com/g/MAXakaWIZARD/xls-writer/?branch=master)
[![Coverage Status](https://coveralls.io/repos/MAXakaWIZARD/xls-writer/badge.svg?branch=master)](https://coveralls.io/r/MAXakaWIZARD/xls-writer?branch=master)
[![SensioLabs Insight](https://img.shields.io/sensiolabs/i/9a9e7784-24a2-4b29-8b64-65f45306c34d.svg)](https://insight.sensiolabs.com/projects/9a9e7784-24a2-4b29-8b64-65f45306c34d)

[![GitHub tag](https://img.shields.io/github/tag/MAXakaWIZARD/xls-writer.svg?label=latest)](https://packagist.org/packages/maxakawizard/xls-writer) 
[![Packagist](https://img.shields.io/packagist/dt/maxakawizard/xls-writer.svg)](https://packagist.org/packages/maxakawizard/xls-writer)
[![Packagist](https://img.shields.io/packagist/dm/maxakawizard/xls-writer.svg)](https://packagist.org/packages/maxakawizard/xls-writer)
[![Minimum PHP Version](http://img.shields.io/badge/php-%3E%3D%205.3-8892BF.svg)](https://php.net/)
[![License](https://img.shields.io/packagist/l/maxakawizard/xls-writer.svg)](https://packagist.org/packages/maxakawizard/xls-writer)

Port of [PEAR Spreadsheet Excel Writer](http://pear.php.net/package/Spreadsheet_Excel_Writer).

This package is compliant with [PSR-0](http://www.php-fig.org/psr/0/), [PSR-1](http://www.php-fig.org/psr/1/), and [PSR-2](http://www.php-fig.org/psr/2/).
If you notice compliance oversights, please send a patch via pull request.

## Known limitations
* Supports only `XLS` format (BIFF5/7 and BIFF8)

## Usage
```php
require('vendor/autoload.php');

$workbook = new Xls\Writer('test.xls');

$worksheet = $workbook->addWorksheet('My first worksheet');

$worksheet->write(0, 0, 'Name');
$worksheet->write(0, 1, 'Age');
$worksheet->write(1, 0, 'John Smith');
$worksheet->write(1, 1, 30);
$worksheet->write(2, 0, 'Johann Schmidt');
$worksheet->write(2, 1, 31);
$worksheet->write(3, 0, 'Juan Herrera');
$worksheet->write(3, 1, 32);

$workbook->close();
```

## Documentation
Original docs can be found [here](https://pear.php.net/manual/en/package.fileformats.spreadsheet-excel-writer.php)

## License
This library is released under [MIT](http://www.tldrlegal.com/license/mit-license) license.
