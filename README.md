# Xls Writer

[![Build Status](https://api.travis-ci.org/MAXakaWIZARD/xls-writer.png?branch=master)](https://travis-ci.org/MAXakaWIZARD/xls-writer) 
[![Scrutinizer Code Quality](https://scrutinizer-ci.com/g/MAXakaWIZARD/xls-writer/badges/quality-score.png?b=master)](https://scrutinizer-ci.com/g/MAXakaWIZARD/xls-writer/?branch=master)
[![Coverage Status](https://coveralls.io/repos/MAXakaWIZARD/xls-writer/badge.svg?branch=master)](https://coveralls.io/r/MAXakaWIZARD/xls-writer?branch=master)
[![Latest Stable Version](https://poser.pugx.org/maxakawizard/xls-writer/v/stable.svg)](https://packagist.org/packages/maxakawizard/xls-writer) 
[![Total Downloads](https://poser.pugx.org/maxakawizard/xls-writer/downloads.svg)](https://packagist.org/packages/maxakawizard/xls-writer) 
[![License](https://poser.pugx.org/maxakawizard/xls-writer/license.svg)](https://packagist.org/packages/maxakawizard/xls-writer)

Port of [PEAR Spreadsheet Excel Writer](http://pear.php.net/package/Spreadsheet_Excel_Writer).

This package is compliant with [PSR-0](http://www.php-fig.org/psr/0/), [PSR-1](http://www.php-fig.org/psr/1/), and [PSR-2](http://www.php-fig.org/psr/2/).
If you notice compliance oversights, please send a patch via pull request.

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
